"""
PDF processing for tender documents.

The module intentionally does not convert PDF -> DOCX.  PDF is processed in-place:
1. text is extracted with coordinates (PyMuPDF);
2. scanned pages use GigaChat Vision for semantic OCR; Tesseract remains only a coordinate/last-resort fallback;
3. logical tender rows are reconstructed from the coordinates;
4. an algorithmic pass resolves each row from the extensible knowledge base
   (services.knowledge_base) plus the legacy DB rules, same as the DOCX flow;
5. GigaChat re-checks every row with the FULL document context (all rows,
   the algorithmic pass, the applicable knowledge base) and may only
   override a value when confident enough (see gigachat_client.AI_CONFIDENCE_THRESHOLD);
6. answers are drawn into the participant/answer column without rebuilding
   the PDF; values inferred by the AI from context (not from the DB) get a
   marker: ** normally, or a single * when the left-side requirement already
   contains *; DB constants never receive a marker;
7. a `<output>.audit.json` is written next to the result, compatible with
   tools/promote_audit.py.

Dependencies:
    pip install pymupdf pytesseract pillow sqlalchemy gigachat opencv-python numpy
    system: Tesseract OCR + Russian language pack (rus)
    Windows: if Tesseract is not in PATH, set TESSERACT_CMD to tesseract.exe.
"""

from __future__ import annotations

import json
import hashlib
import os
import platform
import re
import shutil
import uuid
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

import pymupdf
import pytesseract
from PIL import Image
from sqlalchemy.orm import Session

from models.tr_type import TrType
from services.gigachat_client import AI_CONFIDENCE_THRESHOLD, ask_json, ask_vision_json, open_client, verify_connection
from services.data_tenders_knowledge import get_data_tenders_knowledge, _param_key


_PROMPT_CONTEXT_CACHE: dict[int, tuple[int, str]] = {}

def _cached_context_json(context: dict) -> str:
    key = id(context)
    cached = _PROMPT_CONTEXT_CACHE.get(key)
    if cached is not None and cached[0] == len(context):
        return cached[1]
    value = json.dumps(context, ensure_ascii=False, separators=(",", ":"))
    if len(_PROMPT_CONTEXT_CACHE) > 8:
        _PROMPT_CONTEXT_CACHE.pop(next(iter(_PROMPT_CONTEXT_CACHE)))
    _PROMPT_CONTEXT_CACHE[key] = (len(context), value)
    return value

# Coordinates are PDF points.  Tender forms in this project use A4 pages.
LEFT_NUMBER_X = 125
LEFT_PARAM_X = 400
ANSWER_MIN_X = 450
OCR_DPI = int(os.getenv("OCR_DPI", "240"))
PDF_AI_MIN_CONFIDENCE = float(os.getenv("PDF_AI_MIN_CONFIDENCE", "0.50"))
PDF_CELL_RECHECK = os.getenv("PDF_CELL_RECHECK", "0") == "1"



@dataclass
class PdfItem:
    id: str
    page: int
    number: str
    param_name: str
    required_val: str
    rect: pymupdf.Rect
    answer_rect: Optional[pymupdf.Rect] = None
    existing_answer: str = ""


@dataclass
class PdfPageLayout:
    page: int
    answer_rect: Optional[pymupdf.Rect] = None
    items: list[PdfItem] = field(default_factory=list)


def extract_options_from_parentheses(text: str) -> list[str]:
    matches = re.findall(r"\((.*?)\)", text or "")
    result = []
    for match in matches:
        for part in re.split(r"[/,]| или ", match):
            part = part.strip().lower()
            if part:
                result.append(part)
    return result


def _norm(text: str) -> str:
    text = (text or "").replace("\u00a0", " ")
    text = re.sub(r"\s+", " ", text).strip()
    return text


def _is_param_number(text: str, x: float) -> bool:
    """Only left-column decimal numbers are considered logical row starts."""
    if x > LEFT_NUMBER_X:
        return False
    return bool(re.fullmatch(r"\d+(?:[.]\d+)*[.)]?", text.strip()))


def _line_words(words: list[tuple]) -> list[dict]:
    """Group PDF words into visual lines using their y coordinate."""
    words = sorted(words, key=lambda w: (w[1], w[0]))
    lines: list[dict] = []

    for w in words:
        x0, y0, x1, y1, text = w[:5]
        if not text.strip():
            continue
        if not lines or abs(y0 - lines[-1]["y"]) > 2.5:
            lines.append({"y": y0, "words": [w]})
        else:
            lines[-1]["words"].append(w)

    for line in lines:
        line["words"].sort(key=lambda w: w[0])
        line["text"] = _norm(" ".join(w[4] for w in line["words"]))
        line["x0"] = min(w[0] for w in line["words"])
        line["x1"] = max(w[2] for w in line["words"])
        line["y0"] = min(w[1] for w in line["words"])
        line["y1"] = max(w[3] for w in line["words"])
    return lines


def _detect_answer_column(page: pymupdf.Page, words: Optional[list[tuple]] = None) -> Optional[pymupdf.Rect]:
    """
    Find the participant column from its header or from table vertical borders.
    This is deliberately coordinate based: it works even when the answer cells
    themselves are empty.
    """
    if words is None:
        words = _page_text_words(page)

    # First use repeated table borders. This is more reliable than the header
    # text because the header is usually centered inside a wider cell.
    xs: dict[float, int] = {}
    for drawing in page.get_drawings():
        for item in drawing.get("items", []):
            if item[0] != "re":
                continue
            r = item[1]
            if r.y1 - r.y0 > 8 and r.x0 > page.rect.width * 0.55:
                for x in (round(r.x0, 1), round(r.x1, 1)):
                    xs[x] = xs.get(x, 0) + 1

    if len(xs) >= 2:
        candidates = [x for x, count in xs.items() if count >= 5]
        pairs = []
        for left in candidates:
            for right in candidates:
                if right <= left:
                    continue
                width = right - left
                if 50 <= width <= 180:
                    pairs.append((xs[left] + xs[right], left, right))
        if pairs:
            _, left, right = max(pairs)
            return pymupdf.Rect(left + 2, 0, right - 2, page.rect.height)

    # Fallback to the header if a PDF has no vector table borders.
    header_words = [
        w for w in words
        if any(token in w[4].lower() for token in (
            "предлагаем", "участником", "заполняется", "претендентом",
        ))
    ]
    if header_words:
        x0 = max(430, min(w[0] for w in header_words) - 10)
        x1 = min(page.rect.width - 15, max(w[2] for w in header_words) + 10)
        # For a two-line header, use the complete detected span.
        x1 = max(x1, x0 + 70)
        return pymupdf.Rect(x0, 0, x1, page.rect.height)

    return None


def _find_tesseract() -> Optional[str]:
    """Find Tesseract even when it is installed but not added to PATH."""
    configured = os.getenv("TESSERACT_CMD")
    if configured and os.path.isfile(configured):
        return configured

    found = shutil.which("tesseract")
    if found:
        return found

    candidates = []
    if platform.system() == "Windows":
        candidates = [
            os.path.expandvars(r"%ProgramFiles%\Tesseract-OCR\tesseract.exe"),
            os.path.expandvars(r"%ProgramFiles(x86)%\Tesseract-OCR\tesseract.exe"),
            os.path.expandvars(r"%LOCALAPPDATA%\Tesseract-OCR\tesseract.exe"),
            os.path.expandvars(r"%LOCALAPPDATA%\Programs\Tesseract-OCR\tesseract.exe"),
        ]
    else:
        candidates = ["/usr/bin/tesseract", "/usr/local/bin/tesseract"]

    for candidate in candidates:
        if os.path.isfile(candidate):
            return candidate
    return None


def _prepare_ocr_image(image: Image.Image) -> Image.Image:
    """Improve scanned tender pages before OCR."""
    gray = image.convert("L")
    # Light contrast enhancement; do not aggressively threshold because
    # handwritten marks and thin table text are important.
    from PIL import ImageOps, ImageFilter
    gray = ImageOps.autocontrast(gray, cutoff=1)
    gray = gray.filter(ImageFilter.SHARPEN)
    return gray


def _ocr_page(page: pymupdf.Page) -> list[tuple]:
    """OCR fallback. Returns PyMuPDF-like word tuples."""
    cmd = _find_tesseract()
    if not cmd:
        raise RuntimeError(
            "Tesseract OCR не установлен или не найден. "
            "Установите Tesseract OCR с русским языком (rus), "
            "либо задайте TESSERACT_CMD=C:\\\\Program Files\\\\Tesseract-OCR\\\\tesseract.exe"
        )

    pytesseract.pytesseract.tesseract_cmd = cmd

    # 300 DPI gives much better recognition for this type of scanned A3/A4
    # technical table than the original 220 DPI. The previous default here
    # ("100") silently contradicted the OCR_DPI module constant — fixed.
    dpi = int(os.getenv("OCR_DPI", str(OCR_DPI)))
    scale = dpi / 72
    clip_rect = clip if clip is not None else page.rect
    pix = page.get_pixmap(
        matrix=pymupdf.Matrix(scale, scale),
        clip=clip_rect,
        alpha=False,
    )
    image = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    image = _prepare_ocr_image(image)
    # Для таблиц OCR должен видеть текст, а не рамки таблицы.
    image = _remove_table_lines_for_ocr(image)

    lang = os.getenv("TESSERACT_LANG", "rus+eng")
    psm = os.getenv("TESSERACT_PSM", "11")

    data = pytesseract.image_to_data(
        image,
        lang=lang,
        config=f"--psm {psm}",
        output_type=pytesseract.Output.DICT,
    )

    words = []
    for i, text in enumerate(data["text"]):
        text = (text or "").strip()
        if not text:
            continue
        conf = float(data["conf"][i]) if str(data["conf"][i]).strip() not in ("", "-1") else -1
        # Keep low-confidence words too: technical abbreviations/numbers are
        # often marked with low confidence but are still useful.
        x = clip_rect.x0 + data["left"][i] / scale
        y = clip_rect.y0 + data["top"][i] / scale
        w = data["width"][i] / scale
        h = data["height"][i] / scale
        words.append((x, y, x + w, y + h, text, 0, 0, i))

    return words



def _group_projection_peaks(values, threshold: float, max_gap: int = 2):
    """Return contiguous groups of projection peaks."""
    import numpy as np
    idx = np.where(values >= threshold)[0]
    groups = []
    for i in idx:
        if not groups or i > groups[-1][-1] + max_gap:
            groups.append([int(i)])
        else:
            groups[-1].append(int(i))
    return groups


def _detect_scanned_table_grid(page: pymupdf.Page, dpi: int = 110):
    """Detect the six main table borders and horizontal row borders.

    The source document is a scan.  We therefore use the printed geometry,
    not the (often broken) OCR text layer.  Horizontal lines are searched over
    the whole table width; searching only the answer column fails on some
    pages because that column is interrupted by scan noise.
    """
    try:
        import cv2
        import numpy as np
    except ImportError:
        return None

    dpi = int(dpi or 150)
    scale = dpi / 72.0
    pix = page.get_pixmap(matrix=pymupdf.Matrix(scale, scale), alpha=False)
    image = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, 3)
    gray = cv2.cvtColor(image, cv2.COLOR_RGB2GRAY)
    bw = cv2.adaptiveThreshold(
        gray, 255, cv2.ADAPTIVE_THRESH_MEAN_C, cv2.THRESH_BINARY_INV, 31, 15
    )

    # Vertical table borders.
    v_kernel = cv2.getStructuringElement(
        cv2.MORPH_RECT, (1, max(35, pix.height // 30))
    )
    vertical = cv2.morphologyEx(bw, cv2.MORPH_OPEN, v_kernel)
    vproj = vertical.sum(axis=0) / 255.0
    v_groups = _group_projection_peaks(vproj, max(70, pix.height * 0.10))
    v_lines = []
    for g in v_groups:
        x = int(round(sum(g) / len(g)))
        ys = np.where(vertical[:, g[0]:g[-1] + 1].max(axis=1) > 0)[0]
        if len(ys) >= pix.height * 0.20:
            v_lines.append((x, int(ys.min()), int(ys.max()), len(ys)))

    # Ignore the page frame/title-block lines and find the six borders of the
    # tender table.  A participant table has: № | parameter | unit | required | answer.
    candidates = [v for v in v_lines if 35 <= v[0] <= pix.width - 35]
    sequences = []
    for i in range(len(candidates) - 5):
        run = candidates[i:i + 6]
        xs = [r[0] for r in run]
        widths = [xs[j + 1] - xs[j] for j in range(5)]
        # The parameter column is wide, unit is narrow, requirement/answer are
        # meaningful widths. This rejects title blocks and the page border.
        if widths[1] < pix.width * 0.28:
            continue
        if widths[2] > pix.width * 0.16:
            continue
        if widths[3] < pix.width * 0.10 or widths[4] < pix.width * 0.08:
            continue
        overlap0 = max(r[1] for r in run)
        overlap1 = min(r[2] for r in run)
        if overlap1 - overlap0 < pix.height * 0.25:
            continue
        score = (
            widths[1] * 2.0 + min(widths[3], widths[4]) * 1.2
            - abs(widths[0] - pix.width * 0.045) * 0.5
            - abs(widths[2] - pix.width * 0.08) * 0.4
        )
        sequences.append((score, run))

    if not sequences:
        return None
    _, run = max(sequences, key=lambda x: x[0])
    xs_px = [r[0] for r in run]

    # Horizontal rules. Search the entire six-column span and deliberately
    # exclude the lower title block (normally below ~70% of an A4 page).
    h_kernel = cv2.getStructuringElement(
        cv2.MORPH_RECT, (max(35, pix.width // 25), 1)
    )
    horizontal = cv2.morphologyEx(bw, cv2.MORPH_OPEN, h_kernel)
    crop = horizontal[:, max(0, xs_px[0] - 8):min(pix.width, xs_px[-1] + 8)]
    hproj = crop.sum(axis=1) / 255.0
    h_groups = _group_projection_peaks(hproj, max(45, (xs_px[-1] - xs_px[0]) * 0.25))
    ys_px = []
    max_y = int(pix.height * 0.70)
    for g in h_groups:
        y = int(round(sum(g) / len(g)))
        if 30 <= y <= max_y:
            ys_px.append(y)

    ys_px = sorted(set(ys_px))
    clean = []
    for y in ys_px:
        if not clean or y - clean[-1] >= max(6, int(scale * 3)):
            clean.append(y)
        else:
            clean[-1] = int(round((clean[-1] + y) / 2))
    ys_px = clean

    # The table header must be near the top of the page and there must be a
    # substantial sequence of rows. Remove isolated lines before the header.
    if len(ys_px) < 3:
        return None
    top_candidates = [y for y in ys_px if 35 <= y <= int(pix.height * 0.15)]
    if top_candidates:
        top = min(top_candidates, key=lambda y: abs(y - pix.height * 0.07))
        ys_px = [y for y in ys_px if y >= top]
    if len(ys_px) < 3:
        return None

    xs = [x / scale for x in xs_px]
    ys = [y / scale for y in ys_px]
    return xs, ys


def _ocr_crop_image(image: Image.Image, rect: pymupdf.Rect, scale: float, *, numeric: bool = False) -> str:
    """Локальный OCR одной ячейки.

    Используется только как второй проход для сомнительных ячеек. Перед OCR
    убираем линии рамки таблицы, иначе Tesseract часто принимает их за
    символы (особенно для «1», «I», «/», «—» и цифр).
    """
    left = max(0, int(round(rect.x0 * scale)))
    top = max(0, int(round(rect.y0 * scale)))
    right = min(image.width, int(round(rect.x1 * scale)))
    bottom = min(image.height, int(round(rect.y1 * scale)))
    if right <= left or bottom <= top:
        return ""

    crop = image.crop((left, top, right, bottom))
    # Обрезаем только рамку, но не содержимое ячейки.
    inset = max(3, int(round(2.0 * scale)))
    if crop.width > 2 * inset + 8 and crop.height > 2 * inset + 8:
        crop = crop.crop((inset, inset, crop.width - inset, crop.height - inset))

    gray = crop.convert("L")
    gray = __import__("PIL.ImageOps", fromlist=["ImageOps"]).autocontrast(gray, cutoff=1)

    # Увеличиваем маленькие ячейки. После увеличения тонкие линии таблицы
    # удаляются морфологически.
    gray = gray.resize((max(1, gray.width * 2), max(1, gray.height * 2)))

    try:
        import cv2
        import numpy as np
        arr = np.asarray(gray)
        bw = cv2.adaptiveThreshold(
            arr, 255,
            cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
            cv2.THRESH_BINARY_INV,
            21, 9,
        )
        # Вертикальные линии внутри маленькой ячейки НЕ удаляем:
        # у цифр 1/4/7 и у кириллических букв есть длинные вертикальные штрихи,
        # и агрессивная вертикальная морфология уничтожает их вместе с рамкой.
        # Вертикальные границы уже удалены inset-ом выше.
        h_len = max(16, arr.shape[1] // 2)
        horizontal = cv2.morphologyEx(
            bw, cv2.MORPH_OPEN,
            cv2.getStructuringElement(cv2.MORPH_RECT, (h_len, 1)),
        )
        lines = horizontal
        cleaned = cv2.bitwise_and(bw, cv2.bitwise_not(lines))
        ocr_image = cv2.bitwise_not(cleaned)
        ocr_image = Image.fromarray(ocr_image)
    except Exception:
        ocr_image = gray

    lang = os.getenv("TESSERACT_LANG", "rus+eng")
    if numeric:
        config = '--psm 7 -c tessedit_char_whitelist="0123456789.,;:/+-()*% АВСа-яА-ЯЁё"'
    else:
        config = "--psm 6"
    text = pytesseract.image_to_string(ocr_image, lang=lang, config=config)
    return _norm(text)

def _ocr_row_words(image: Image.Image, rect: pymupdf.Rect, scale: float) -> list[tuple]:
    """Совместимость со старым API.

    Новый скан-пайплайн НЕ вызывает эту функцию для каждой строки: OCR страницы
    выполняется один раз и затем слова раскладываются по физическим ячейкам.
    """
    left = max(0, int(round(rect.x0 * scale)))
    top = max(0, int(round(rect.y0 * scale)))
    right = min(image.width, int(round(rect.x1 * scale)))
    bottom = min(image.height, int(round(rect.y1 * scale)))
    if right <= left or bottom <= top:
        return []

    crop = image.crop((left, top, right, bottom))
    crop = _prepare_ocr_image(crop)
    data = pytesseract.image_to_data(
        crop,
        lang=os.getenv("TESSERACT_LANG", "rus+eng"),
        config="--psm 6",
        output_type=pytesseract.Output.DICT,
    )
    words = []
    for i, text in enumerate(data["text"]):
        text = (text or "").strip()
        if not text:
            continue
        x = rect.x0 + data["left"][i] / scale
        y = rect.y0 + data["top"][i] / scale
        w = data["width"][i] / scale
        h = data["height"][i] / scale
        words.append((x, y, x + w, y + h, text, 0, 0, i))
    return words


def _remove_table_lines_for_ocr(image: Image.Image) -> Image.Image:
    """Удаляет длинные линии таблицы, не удаляя штрихи букв и цифр.

    Для исходного скана именно линии рамки являются одним из главных источников
    ошибок Tesseract: вертикальные границы превращаются в «1/I», горизонтальные
    — в тире и случайные символы. Линии удаляются только морфологически длинными
    элементами, поэтому короткие штрихи текста сохраняются.
    """
    try:
        import cv2
        import numpy as np
        arr = np.asarray(image.convert("L"))
        bw = cv2.adaptiveThreshold(
            arr, 255,
            cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
            cv2.THRESH_BINARY_INV,
            31, 15,
        )
        h_len = max(45, arr.shape[1] // 15)
        v_len = max(45, arr.shape[0] // 20)
        horizontal = cv2.morphologyEx(
            bw, cv2.MORPH_OPEN,
            cv2.getStructuringElement(cv2.MORPH_RECT, (h_len, 1)),
        )
        vertical = cv2.morphologyEx(
            bw, cv2.MORPH_OPEN,
            cv2.getStructuringElement(cv2.MORPH_RECT, (1, v_len)),
        )
        lines = cv2.bitwise_or(horizontal, vertical)
        cleaned = cv2.bitwise_and(bw, cv2.bitwise_not(lines))
        return Image.fromarray(cv2.bitwise_not(cleaned))
    except Exception:
        return image


def _ocr_page_words_once(
    page: pymupdf.Page,
    dpi: int = 300,
    clip: Optional[pymupdf.Rect] = None,
) -> tuple[Image.Image, list[tuple]]:
    """Один OCR всей таблицы с координатами слов.

    OCR запускается один раз на страницу. Перед ним удаляются длинные линии
    таблицы. Это одновременно быстрее и заметно надёжнее, чем OCR каждой
    строки/ячейки отдельно.
    """
    cmd = _find_tesseract()
    if not cmd:
        raise RuntimeError(
            "Tesseract OCR не установлен или не найден. "
            "Установите Tesseract с языком rus либо задайте TESSERACT_CMD."
        )
    pytesseract.pytesseract.tesseract_cmd = cmd

    dpi = max(220, int(dpi or 300))
    scale = dpi / 72.0
    clip_rect = clip if clip is not None else page.rect
    pix = page.get_pixmap(
        matrix=pymupdf.Matrix(scale, scale),
        clip=clip_rect,
        alpha=False,
    )
    image = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    image = _prepare_ocr_image(image)
    image = _remove_table_lines_for_ocr(image)

    lang = os.getenv("TESSERACT_LANG", "rus+eng")
    psm = os.getenv("TESSERACT_PSM", "6")
    data = pytesseract.image_to_data(
        image,
        lang=lang,
        config=f"--psm {psm}",
        output_type=pytesseract.Output.DICT,
    )

    words: list[tuple] = []
    for i, text in enumerate(data["text"]):
        text = (text or "").strip()
        if not text:
            continue
        try:
            conf = float(data["conf"][i])
            x0 = float(data["left"][i])
            y0 = float(data["top"][i])
            w = float(data["width"][i])
            h = float(data["height"][i])
        except (TypeError, ValueError):
            continue
        x = clip_rect.x0 + x0 / scale
        y = clip_rect.y0 + y0 / scale
        words.append((x, y, x + w / scale, y + h / scale, text, conf, 0, len(words)))

    return image, words


def _cell_text_from_words(words: list[tuple], rect: pymupdf.Rect) -> str:
    """Извлекает текст ячейки по центрам bbox, сохраняя порядок строк."""
    inside = []
    for w in words:
        cx = (w[0] + w[2]) / 2.0
        cy = (w[1] + w[3]) / 2.0
        if rect.contains(pymupdf.Point(cx, cy)):
            inside.append(w)
    inside.sort(key=lambda w: (round(w[1], 1), w[0]))

    lines: list[list[tuple]] = []
    for w in inside:
        if not lines or abs(w[1] - lines[-1][0][1]) > 3.5:
            lines.append([w])
        else:
            lines[-1].append(w)
    result = []
    for line in lines:
        line.sort(key=lambda w: w[0])
        result.append(_norm(" ".join(w[4] for w in line)))
    return _norm(" ".join(x for x in result if x))


def _ocr_cell_precise(image: Image.Image, rect: pymupdf.Rect, scale: float, *, numeric: bool = False) -> str:
    """Второй OCR-проход только для реально сомнительной ячейки."""
    left = max(0, int(round(rect.x0 * scale)))
    top = max(0, int(round(rect.y0 * scale)))
    right = min(image.width, int(round(rect.x1 * scale)))
    bottom = min(image.height, int(round(rect.y1 * scale)))
    if right <= left or bottom <= top:
        return ""

    crop = image.crop((left, top, right, bottom))
    inset = max(4, int(round(2.5 * scale)))
    if crop.width > 2 * inset + 10 and crop.height > 2 * inset + 10:
        crop = crop.crop((inset, inset, crop.width - inset, crop.height - inset))

    # Увеличение особенно помогает цифрам в узкой колонке требований.
    crop = crop.resize((max(1, crop.width * 2), max(1, crop.height * 2)))
    crop = _prepare_ocr_image(crop)

    if numeric:
        config = '--psm 7 -c tessedit_char_whitelist="0123456789.,;:/+-()*%"'
    else:
        config = "--psm 6"
    return _norm(pytesseract.image_to_string(
        crop,
        lang=os.getenv("TESSERACT_LANG", "rus+eng"),
        config=config,
    ))


def _looks_suspicious_ocr(text: str, *, numeric: bool = False) -> bool:
    t = _norm(text)
    if not t:
        return True
    if numeric:
        digits = re.findall(r"\d", t)
        # Числовая ячейка без цифр почти всегда является ошибкой OCR.
        if not digits:
            return True
        # Лишнее длинное слово рядом с числом — повод перепроверить ячейку.
        words = re.findall(r"[A-Za-zА-Яа-я]{3,}", t)
        return len(words) > 1 or (len(words) == 1 and not re.search(r"(?:не|более|менее|требуется)", t, re.I))
    # Латиница внутри русской технической подписи допустима, но полностью
    # бессмысленный короткий латинский фрагмент — частый артефакт OCR.
    cyr = len(re.findall(r"[А-Яа-яЁё]", t))
    lat = len(re.findall(r"[A-Za-z]", t))
    return (cyr == 0 and len(t) >= 4) or (lat > cyr * 2 and cyr < 3)


def _is_numeric_requirement(param_name: str, unit: str) -> bool:
    p = _norm(param_name).lower()
    u = _norm(unit).lower()
    numeric_tokens = (
        "напряжен", "ток", "количество", "частот", "коэффициент", "нагруз", "высот",
        "температур", "сопротивлен", "мощност", "давлен", "массы", "масса", "дли",
        "путь утечки", "срок", "кратност", "уровн", "расход", "число",
    )
    unit_tokens = ("кв", "в", "а", "ка", "гц", "с", "м", "ма", "ва", "%", "ом", "н")
    return any(token in p for token in numeric_tokens) or any(u == token or u.startswith(token + "/") for token in unit_tokens)


def _normalize_numeric_requirement(text: str) -> str:
    """Убирает OCR-мусор из коротких числовых требований, не меняя диапазон."""
    t = _norm(text)
    if not t:
        return ""
    # Сохраняем типичные инженерные записи: 24,6/28,5; -45...+40; 1500 (750)*.
    direct = re.search(r"[-+]?\d+(?:[.,]\d+)?(?:\s*(?:\.\.\.|\.\.|/|[-–])\s*[-+]?\d+(?:[.,]\d+)?)?", t)
    if not direct:
        return t
    value = direct.group(0).replace(" ", "")
    # Если рядом явно есть проценты/звёздочка, не теряем их.
    if "%" in t:
        value += "%"
    if "*" in t and not value.endswith("*"):
        value += "*"
    return value


def _ocr_cell(page: pymupdf.Page, rect: pymupdf.Rect, dpi: int = 180, psm: int = 6) -> str:
    """Compatibility helper. Prefer _ocr_page once per page for speed."""
    words = _ocr_page(page)
    inside = []
    for w in words:
        cx = (w[0] + w[2]) / 2
        cy = (w[1] + w[3]) / 2
        if rect.contains(pymupdf.Point(cx, cy)):
            inside.append(w[4])
    return _norm(" ".join(inside))


def _page_text_words(page: pymupdf.Page) -> list[tuple]:
    """Get page text words in the same visual coordinate system as page.rect.

    Source tender PDFs have /Rotate=270 and a landscape media box. PyMuPDF
    returns text coordinates in the unrotated space, while get_pixmap() and
    drawing/insertion use the rotated visual page. Without this transform the
    table geometry and OCR/text coordinates refer to different cells.
    """
    raw = page.get_text("words")
    if not raw or page.rotation == 0:
        return raw

    result = []
    for w in raw:
        r = pymupdf.Rect(w[0], w[1], w[2], w[3]) * page.rotation_matrix
        result.append((r.x0, r.y0, r.x1, r.y1, w[4], *w[5:]))
    return result


def _words_in_rect(words: list[tuple], rect: pymupdf.Rect) -> str:
    vals = []
    for w in words:
        cx = (w[0] + w[2]) / 2
        cy = (w[1] + w[3]) / 2
        if rect.contains(pymupdf.Point(cx, cy)):
            vals.append(w)
    vals.sort(key=lambda w: (w[1], w[0]))
    return _norm(" ".join(w[4] for w in vals))


def _ocr_number(text: str) -> Optional[str]:
    """Extract a row number and repair common OCR confusions."""
    t = _norm(text).replace("|", "").replace("—", "-")
    m = re.search(r"(?<!\d)(\d{1,3}(?:\.\d+)*)", t)
    if m:
        return m.group(1)
    # Tesseract commonly confuses 1/2/5/9 in tiny numbered cells.
    substitutions = str.maketrans({"B": "8", "В": "8", "З": "3", "О": "0", "o": "0", "I": "1", "l": "1"})
    t2 = t.translate(substitutions)
    m = re.search(r"(?<!\d)(\d{1,3})(?!\d)", t2)
    return m.group(1) if m else None


def _detect_number_column_rows(page: pymupdf.Page, xs: list[float], dpi: int = 150) -> list[float]:
    """Return physical table row boundaries.

    The PDF contains merged cells in the № column (20, 22, 23, 24, 26, 42,
    43, ...).  Therefore the № column cannot be used to detect all answer
    cells: several physical rows share one number.  The only reliable source
    for answer-cell boundaries is the horizontal table grid across all five
    columns.
    """
    import cv2
    import numpy as np

    scale = dpi / 72.0
    pix = page.get_pixmap(matrix=pymupdf.Matrix(scale, scale), alpha=False)
    image = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, 3)
    gray = cv2.cvtColor(image, cv2.COLOR_RGB2GRAY)
    bw = cv2.adaptiveThreshold(
        gray, 255, cv2.ADAPTIVE_THRESH_MEAN_C, cv2.THRESH_BINARY_INV, 31, 15
    )

    x0 = max(0, int(xs[0] * scale) - 8)
    x1 = min(pix.width, int(xs[-1] * scale) + 8)
    crop = bw[:, x0:x1]
    kernel = cv2.getStructuringElement(
        cv2.MORPH_RECT, (max(40, int((xs[-1] - xs[0]) * scale * 0.18)), 1)
    )
    horizontal = cv2.morphologyEx(crop, cv2.MORPH_OPEN, kernel)
    projection = horizontal.sum(axis=1) / 255.0

    threshold = max(25, crop.shape[1] * 0.45)
    idx = np.where(projection >= threshold)[0]
    groups = []
    for y in idx:
        if not groups or y > groups[-1][-1] + 2:
            groups.append([int(y)])
        else:
            groups[-1].append(int(y))

    ys = [sum(g) / len(g) / scale for g in groups]
    ys = [y for y in ys if 30 <= y <= page.rect.height * 0.72]
    if len(ys) < 3:
        return []

    # Keep the table header top line and everything below it.
    top = min(ys, key=lambda y: abs(y - page.rect.height * 0.07))
    ys = [y for y in ys if y >= top - 2]

    clean = []
    for y in ys:
        if not clean or y - clean[-1] >= 4:
            clean.append(y)
        else:
            clean[-1] = (clean[-1] + y) / 2
    return clean

def _repair_pdf_param_name_dummy(text: str) -> str:
    repaired = _norm(text)
    replacements = {
        "cmaxdapm": "стандарт", "bud внешней": "вид внешней", "bud внутренней": "вид внутренней",
        "бнешней": "внешней", "бнутренней": "внутренней", "цторичный": "вторичный",
        "пербичный": "первичный", "бторичной": "вторичной", "пербичной": "первичной",
        "mpeoyemca": "требуется", "пребуется": "требуется", "пребиется": "требуется",
        "испытамельное": "испытательное", "электробинамической": "электродинамической",
        "термической сточкости": "термической стойкости", "стодкости": "стойкости",
    }
    for bad, good in replacements.items():
        repaired = re.sub(re.escape(bad), good, repaired, flags=re.I)

    # OCR иногда переставляет слова местами, но сохраняет ключевые термины.
    # Восстанавливаем только однозначные названия полей — это безопаснее,
    # чем пытаться «исправлять» все технические фразы целиком.
    low = repaired.lower()
    canonical_patterns = (
        (("класс", "напряж"), "Класс напряжения"),
        (("наибольшее", "напряж"), "Наибольшее рабочее напряжение"),
        (("климат",), "Климатическое исполнение/категория размещения"),
        (("температур", "окружающего"), "Температура окружающего воздуха"),
        (("внешней", "изоляц"), "Вид внешней изоляции"),
        (("внутренней", "изоляц"), "Вид внутренней изоляции"),
        (("номинальн", "первич", "ток"), "Номинальный ток первичной обмотки"),
        (("номинальн", "вторич", "ток"), "Номинальный ток вторичной обмотки"),
        (("номинальн", "частот"), "Номинальная частота"),
        (("ток", "короткого", "замыкания"), "Ток короткого замыкания"),
        (("термическ", "стойк"), "Ток термической стойкости"),
        (("электродинамическ", "стойк"), "Ток электродинамической стойкости"),
        (("испытательн", "напряжен", "грозов"), "Испытательное напряжение полного грозового импульса"),
    )
    for tokens, canonical in canonical_patterns:
        if all(token in low for token in tokens):
            return canonical
    return repaired


def _extract_scanned_page_items(page: pymupdf.Page, page_index: int) -> tuple[PdfPageLayout, str]:
    """Разбирает скан по геометрии таблицы + ОДНОМУ OCR всей страницы.

    Ключевой принцип:
      * сетка определяет физические ячейки;
      * Tesseract один раз распознаёт всю страницу с bbox;
      * bbox раскладываются по конкретным ячейкам;
      * второй OCR ячейки выполняется только если первый результат явно плохой.

    Поэтому распознанное значение никогда не «переезжает» в соседний столбец.
    """
    grid = _detect_scanned_table_grid(
        page,
        dpi=int(os.getenv("GRID_DPI", "150")),
    )
    if not grid:
        return PdfPageLayout(page=page_index), ""

    xs, row_ys = grid
    if len(xs) < 6 or len(row_ys) < 3:
        return PdfPageLayout(page=page_index), ""

    number_x0, number_x1 = xs[0], xs[1]
    param_x0, param_x1 = xs[1], xs[2]
    unit_x0, unit_x1 = xs[2], xs[3]
    req_x0, req_x1 = xs[3], xs[4]
    answer_x0, answer_x1 = xs[4], xs[5]

    ocr_dpi = max(200, int(os.getenv("OCR_DPI", str(OCR_DPI))))
    table_clip = pymupdf.Rect(
        max(0, xs[0] - 2),
        max(0, row_ys[0] - 2),
        min(page.rect.width, xs[-1] + 2),
        min(page.rect.height, row_ys[-1] + 2),
    )
    image, ocr_words = _ocr_page_words_once(
        page,
        dpi=ocr_dpi,
        clip=table_clip,
    )
    ocr_scale = ocr_dpi / 72.0
    embedded_words = _page_text_words(page)

    layout = PdfPageLayout(page=page_index)
    layout.answer_rect = pymupdf.Rect(
        answer_x0 + 1, row_ys[0] + 1,
        answer_x1 - 1, row_ys[-1] - 1,
    )

    # Полный OCR страницы — важнейший контекст для AI. Даже если одно слово
    # попало не в ту bbox-ячейку, модель получает исходную связную строку и может
    # восстановить смысл по соседним параметрам.
    raw_page_text = _norm(" ".join(w[4] for w in ocr_words if str(w[4]).strip()))
    page_text: list[str] = [raw_page_text] if raw_page_text else []
    previous_number = ""

    for row_idx in range(2, len(row_ys)):
        y0, y1 = row_ys[row_idx - 1], row_ys[row_idx]
        if y1 - y0 < 7:
            continue

        # Небольшой inset исключает линии рамки, но оставляет весь текст.
        number_rect = pymupdf.Rect(number_x0 + 1.5, y0 + 1.5, number_x1 - 1.5, y1 - 1.5)
        param_rect = pymupdf.Rect(param_x0 + 1.5, y0 + 1.5, param_x1 - 1.5, y1 - 1.5)
        unit_rect = pymupdf.Rect(unit_x0 + 1.5, y0 + 1.5, unit_x1 - 1.5, y1 - 1.5)
        req_rect = pymupdf.Rect(req_x0 + 1.5, y0 + 1.5, req_x1 - 1.5, y1 - 1.5)
        ans_rect = pymupdf.Rect(answer_x0 + 1.5, y0 + 1.5, answer_x1 - 1.5, y1 - 1.5)

        number_text = _cell_text_from_words(ocr_words, number_rect)
        row_number = _ocr_number(number_text)
        if not row_number and PDF_CELL_RECHECK:
            # Только номер — маленькая ячейка, поэтому допускаем локальную
            # перепроверку. Для merged №-ячеек используем предыдущий номер.
            row_number = _ocr_number(
                _ocr_cell_precise(image, number_rect, ocr_scale, numeric=True)
            )
        if row_number:
            previous_number = row_number
        elif previous_number:
            row_number = previous_number

        param_text = _cell_text_from_words(ocr_words, param_rect)
        unit_text = _cell_text_from_words(ocr_words, unit_rect)
        req_text = _cell_text_from_words(ocr_words, req_rect)

        # Локальная перепроверка только пустых/подозрительных ячеек.
        # Полноэкранный OCR обычно лучше локального crop для длинных русских
        # названий. Поэтому локальный проход используем только при полном
        # отсутствии текста, а не заменяем им уже распознанную строку.
        if not param_text and PDF_CELL_RECHECK:
            precise = _ocr_cell_precise(image, param_rect, ocr_scale, numeric=False)
            if precise:
                param_text = precise

        numeric_req = _is_numeric_requirement(param_text, unit_text)
        if PDF_CELL_RECHECK and _looks_suspicious_ocr(req_text, numeric=numeric_req):
            precise = _ocr_cell_precise(image, req_rect, ocr_scale, numeric=numeric_req)
            if precise:
                req_text = precise

        if numeric_req:
            req_text = _normalize_numeric_requirement(req_text)

        # Existing answer: для скана обычно отсутствует текстовый слой. Если
        # он всё-таки есть, берём его только из физической answer-ячейки.
        existing_answer = _cell_text_from_words(embedded_words, ans_rect)

        param_text = _norm(re.sub(r"[|]+", " ", param_text))
        unit_text = _norm(re.sub(r"[|]+", " ", unit_text))
        req_text = _norm(re.sub(r"[|]+", " ", req_text))
        param_text = _repair_pdf_param_name_dummy(param_text)

        # Если OCR плохо прочитал название, но требование информативно, не
        # выбрасываем строку: AI должен получить шанс восстановить смысл.
        if not param_text and not req_text:
            continue
        if len(param_text) < 2 and len(req_text) < 2:
            continue

        low = _norm(f"{param_text} {unit_text} {req_text}").lower()
        if any(token in low for token in (
            "наименование параметра",
            "технические требования",
            "требуемое значение",
            "заполняется претендентом",
        )) and not re.search(r"\d+(?:\.\d+)+", low):
            continue

        item_id = f"p{page_index + 1}_item_{len(layout.items)}"
        item = PdfItem(
            id=item_id,
            page=page_index,
            number=(row_number or f"p{page_index + 1}_r{row_idx:02d}"),
            param_name=_norm(f"{param_text} {unit_text}"),
            required_val=req_text,
            rect=pymupdf.Rect(number_x0, y0, answer_x1, y1),
            answer_rect=ans_rect,
            existing_answer=_norm(existing_answer),
        )
        layout.items.append(item)
        page_text.append(_norm(f"{item.param_name} {item.required_val}"))

    return layout, "\n".join(x for x in page_text if x)


def _is_scanned_page(page: pymupdf.Page) -> bool:
    """Detect rasterized/scan pages even when a stale OCR text layer exists."""
    infos = page.get_image_info()
    if not infos:
        return False
    page_area = max(page.mediabox.width * page.mediabox.height, 1.0)
    for info in infos:
        x0, y0, x1, y1 = info["bbox"]
        area = max(0.0, x1 - x0) * max(0.0, y1 - y0)
        if area / page_area >= 0.75:
            return True
    return False


def _vision_pdf_context(pdf_path: str, *, max_pages: int | None = None) -> list[dict]:
    """Vision для PDF только по действительно сканированным страницам.

    Результат кэшируется по пути/mtime/размеру файла, поэтому повторный запуск
    обычного заполнения не делает сетевые Vision-запросы заново.
    """
    if os.getenv("PDF_VISION_OCR", "1") != "1":
        return []
    cache_path = Path(pdf_path).with_suffix(Path(pdf_path).suffix + ".vision.json")
    try:
        max_pages = int(os.getenv("PDF_VISION_MAX_PAGES", "6")) if max_pages is None else int(max_pages)
    except ValueError:
        max_pages = 6

    try:
        st = os.stat(pdf_path)
        sig = f"{st.st_mtime_ns}:{st.st_size}"
    except OSError:
        sig = ""

    try:
        cached = json.loads(cache_path.read_text(encoding="utf-8"))
        if isinstance(cached, dict) and cached.get("signature") == sig and isinstance(cached.get("pages"), list):
            return cached["pages"][:max_pages]
    except Exception:
        pass

    doc = pymupdf.open(pdf_path)
    result: list[dict] = []
    try:
        need_live = os.getenv("PDF_VISION_LIVE", "0") == "1"
        if not need_live:
            return []
        with open_client() as client:
            dpi = max(180, int(os.getenv("PDF_VISION_DPI", "220")))
            for page_index, page in enumerate(doc):
                if page_index >= max_pages:
                    break
                # Не отправляем обычные текстовые страницы/логотипы в Vision.
                # Только крупный raster scan, либо явное включение всех image-pages.
                if not _is_scanned_page(page) and os.getenv("PDF_VISION_ALL_IMAGES", "0") != "1":
                    continue
                scale = dpi / 72.0
                pix = page.get_pixmap(matrix=pymupdf.Matrix(scale, scale), alpha=False, annots=True)
                image = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                buf = __import__("io").BytesIO()
                image.save(buf, format="PNG", optimize=True)
                prompt = (
                    "Извлеки все читаемые технические данные со страницы тендерного документа. "
                    "Особенно нужны номер строки, параметр, требование, ответ, марка/тип, "
                    "габариты, массы и единицы. Не додумывай. Верни строго JSON: "
                    '{"rows":[],"models":[],"parameters":[],"raw_text":""}'
                )
                try:
                    data = ask_vision_json(client=client, image_bytes=buf.getvalue(), user_prompt=prompt, mime_type="image/png")
                    data["page"] = page_index + 1
                    result.append(data)
                except Exception as exc:
                    print(f"[PDF VISION] page={page_index + 1}: {exc}")
    finally:
        doc.close()
    try:
        cache_path.write_text(json.dumps({"signature": sig, "pages": result}, ensure_ascii=False), encoding="utf-8")
    except Exception:
        pass
    return result


def extract_pdf_items(pdf_path: str, use_ocr: bool = True) -> tuple[list[PdfPageLayout], str]:
    """
    Extract logical tender rows and all document text.

    A row starts with a parameter number (1.1, 1.2, ...). Text until the next
    parameter number belongs to the same row. This is important for wrapped
    multi-line requirements and prevents values from shifting to another row.
    """
    if not os.path.exists(pdf_path):
        raise FileNotFoundError(pdf_path)

    doc = pymupdf.open(pdf_path)
    layouts: list[PdfPageLayout] = []
    all_text = []
    vision_pages = _vision_pdf_context(pdf_path) if use_ocr else []
    for vp in vision_pages:
        raw_vision = vp.get("raw_text", "")
        if raw_vision:
            all_text.append(_norm(raw_vision))
        for row in vp.get("rows", []) or []:
            if isinstance(row, dict):
                all_text.append(_norm(" ".join(str(row.get(k, "")) for k in ("number", "param_name", "unit", "required_val", "answer"))))
        for par in vp.get("parameters", []) or []:
            if isinstance(par, dict):
                all_text.append(_norm(" ".join(str(par.get(k, "")) for k in ("model", "name", "value", "unit"))))

    for page_index, page in enumerate(doc):
        words = _page_text_words(page)
        text = _norm(page.get_text())

        # Scanned tender pages are handled by table geometry + cell OCR.
        # This is critical for forms where the participant column is empty: OCR
        # of the whole page cannot tell us where an answer belongs.
        if use_ocr and _is_scanned_page(page):
            try:
                scanned_layout, scanned_text = _extract_scanned_page_items(page, page_index)
                if scanned_layout.items:
                    layouts.append(scanned_layout)
                    all_text.append(scanned_text)
                    continue
            except Exception as exc:
                print(f"[PDF GRID/OCR] page={page_index + 1}: {exc}")

            # Fallback to full-page OCR if table-grid detection failed.
            try:
                ocr_words = _ocr_page(page)
                if ocr_words:
                    words = ocr_words
                    text = _norm(" ".join(w[4] for w in words))
            except Exception as exc:
                print(f"[PDF OCR] page={page_index + 1}: {exc}")
                if len(text) < 30:
                    raise RuntimeError(
                        f"Не удалось прочитать сканированную PDF-страницу {page_index + 1}. "
                        f"OCR недоступен: {exc}"
                    ) from exc

        all_text.append(text)
        layout = PdfPageLayout(page=page_index)
        answer_col = _detect_answer_column(page, words)
        if answer_col:
            layout.answer_rect = answer_col

        lines = _line_words(words)

        starts = []
        for idx, line in enumerate(lines):
            for w in line["words"]:
                if _is_param_number(w[4], w[0]):
                    starts.append((idx, w[4], line["y0"]))
                    break

        # De-duplicate starts on the same visual line.
        uniq = []
        seen = set()
        for start in starts:
            key = (start[0], start[1])
            if key not in seen:
                uniq.append(start)
                seen.add(key)
        starts = uniq

        for pos, (start_idx, number, y0) in enumerate(starts):
            end_idx = starts[pos + 1][0] if pos + 1 < len(starts) else len(lines)
            row_lines = lines[start_idx:end_idx]
            if not row_lines:
                continue

            param_parts = []
            required_parts = []
            row_y0 = row_lines[0]["y0"]
            row_y1 = row_lines[-1]["y1"]

            for line in row_lines:
                left = []
                middle = []
                for w in line["words"]:
                    if w[0] < LEFT_PARAM_X:
                        left.append(w[4])
                    elif w[0] < (answer_col.x0 - 5 if answer_col else 480):
                        middle.append(w[4])
                # Remove the number itself from the parameter text.
                left_text = _norm(" ".join(left))
                left_text = re.sub(rf"^{re.escape(number)}\s*", "", left_text)
                if left_text:
                    param_parts.append(left_text)
                if middle:
                    required_parts.append(_norm(" ".join(middle)))

            param_name = _norm(" ".join(param_parts))
            required_val = _norm(" ".join(required_parts))

            # Existing participant answer, when present in the source file.
            existing_answer = ""
            if answer_col:
                ans_rect = pymupdf.Rect(answer_col.x0 + 2, max(0, row_y0 - 1), answer_col.x1 - 2, min(page.rect.height, row_y1 + 1))
                existing_answer = _words_in_rect(words, ans_rect)

            # Ignore document metadata/header numbers that are not tender rows.
            if not param_name and not required_val:
                continue

            item_id = f"p{page_index + 1}_item_{len(layout.items)}"
            answer_rect = None
            if answer_col:
                answer_rect = pymupdf.Rect(
                    answer_col.x0 + 2,
                    max(0, row_y0 - 1),
                    answer_col.x1 - 2,
                    min(page.rect.height, row_y1 + 1),
                )

            layout.items.append(
                PdfItem(
                    id=item_id,
                    page=page_index,
                    number=number,
                    param_name=param_name,
                    required_val=required_val,
                    rect=pymupdf.Rect(0, row_y0, page.rect.width, row_y1),
                    answer_rect=answer_rect,
                    existing_answer=existing_answer,
                )
            )

        layouts.append(layout)

    return layouts, "\n".join(t for t in all_text if t)


def _looks_like_climate_parameter(text: str) -> bool:
    """Устойчиво распознаёт «климатическое исполнение» после OCR-ошибок."""
    from difflib import SequenceMatcher

    t = _norm(text).lower()
    if "климат" in t or "климатич" in t:
        return True
    # Типичная ошибка Tesseract на скане: «клуматуческое».
    candidates = (
        "климатическое исполнение",
        "климатическое исполнение категория размещения",
        "климатическое исполнение/категория размещения",
    )
    return max((SequenceMatcher(None, t, x).ratio() for x in candidates), default=0.0) >= 0.58


def extract_requirements_from_pdf(layouts: list[PdfPageLayout]) -> dict:
    requirements = {
        "voltage": None,
        "climat": None,
        "isol_type": None,
        "isol_color": None,
        "raw_text": [],
    }

    for layout in layouts:
        for item in layout.items:
            param_name = _norm(item.param_name).lower()
            required_val = _norm(item.required_val).lower()
            requirements["raw_text"].append(f"{param_name} {required_val}")

            # Для тендерных ТЗ фактический класс часто находится в строке
            # «Класс напряжения», а «Наибольшее рабочее напряжение» (например,
            # 126 кВ) идёт следующей строкой. Берём класс первым и не даём 126
            # перезаписать 110.
            if ("класс напряжения" in param_name or "номинальное напряжение" in param_name) and requirements["voltage"] is None:
                nums = re.findall(r"\d+", required_val)
                if nums:
                    requirements["voltage"] = int(nums[0])
            elif "наибольшее рабочее напряжение" in param_name and requirements["voltage"] is None:
                nums = re.findall(r"\d+", required_val)
                if nums:
                    requirements["voltage"] = int(nums[0])
            elif _looks_like_climate_parameter(param_name):
                requirements["climat"] = required_val
            elif "изоляци" in param_name and "цвет" not in param_name and ("тип" in param_name or "вид" in param_name):
                options = extract_options_from_parentheses(param_name)
                options += extract_options_from_parentheses(required_val)
                requirements["isol_type"] = options if options else required_val
            elif "цвет" in param_name and "изоляци" in param_name:
                options = extract_options_from_parentheses(param_name)
                options += extract_options_from_parentheses(required_val)
                requirements["isol_color"] = options if options else required_val

    return requirements


def match_tr_type_by_pdf_rules(
    layouts: list[PdfPageLayout], session: Session, filename_hint: str = ""
) -> tuple[TrType, Optional[int]]:
    """Определяет паспорт изделия только по data_tenders и PDF-контексту."""
    data_kb = get_data_tenders_knowledge()
    reqs = extract_requirements_from_pdf(layouts)
    text_parts = list(reqs.get("raw_text", []))
    if filename_hint:
        text_parts.append(filename_hint)
    text = " ".join(text_parts)
    model, signature, candidates = data_kb.select_profile(text)
    if not model:
        model = data_kb.detect_model(text)
    if not model and filename_hint:
        low_name = _norm(filename_hint)
        for candidate in sorted(data_kb._profiles.keys(), key=len, reverse=True):
            if _norm(candidate) in low_name:
                model = candidate
                break
    voltage = signature.get("voltage_class") or (data_kb._model_voltage(model) if model else None)
    if not model and voltage:
        candidate = f"ТРГ-УЭТМ-{voltage}"
        if candidate in data_kb._profiles:
            model = candidate
    if not model:
        raise ValueError(f"Не удалось определить профиль PDF. Кандидаты data_tenders: {candidates[:10]}")
    print(
        "[PROFILE][PDF] " + str({
            "product_type": signature.get("product_type"),
            "model": model,
            "voltage_class": voltage,
            "insulation_type": signature.get("insulation_types", []),
            "insulation_color": signature.get("insulation_colors", []),
            "candidates": candidates[:10],
        })
    )
    return TrType(name=model), int(float(voltage)) if voltage is not None and str(voltage).replace('.', '', 1).isdigit() else voltage


def _db_context(tr_type_obj: TrType, voltage: Optional[int], session: Session) -> dict:
    """Совместимый API: технический контекст формируется только из data_tenders."""
    kb = get_data_tenders_knowledge()
    profile = kb.model_profile(tr_type_obj.name) if tr_type_obj else {}
    return {
        "source": "data_tenders",
        "manufacturer": "",
        "brand_name": "",
        "voltages": [str(profile.get("voltage"))] if profile.get("voltage") else [],
        "parameters": profile.get("parameters", {}),
    }

def _is_trivial_row(item: "PdfItem") -> bool:
    """Строки-заголовки разделов / шум OCR, для которых пустой ответ — норма.

    Раньше пустое значение по такой строке приводило к фатальной ошибке
    всего процесса (см. старую проверку `missing` в process_pdf_requirements).
    Теперь такие строки просто не считаются "непроставленными".
    """
    param = _norm(item.param_name)
    req = _norm(item.required_val)
    if not param and not req:
        return True
    cleaned = re.sub(r"[\[\]{}|~`_]", "", req).strip()
    if param and len(param) <= 3 and not re.search(r"\d", param) and not req:
        return True
    if not param and cleaned and len(cleaned) <= 3 and not re.search(r"\d", cleaned):
        return True
    return False


def _algorithm_fill_pdf(
    session: Session,
    items: list[dict],
    tr_type: TrType,
    voltage: Optional[int],
    legacy_context: dict,
    data_kb=None,
    document_text: str = "",
) -> None:
    """Первичный проход по единому паспорту data_tenders."""
    data_kb = data_kb or get_data_tenders_knowledge()
    model = tr_type.name if tr_type else data_kb.detect_model(document_text, items)
    signature = dict(items[0].get("_product_signature") or {}) if items else {}
    signature.update({"model": model, "voltage_class": str(voltage) if voltage is not None else signature.get("voltage_class")})
    for item in items:
        pkey = _param_key(item.get("param_name", ""))
        item["db_key"] = pkey
        value, source, score = data_kb.lookup(
            item.get("param_name", ""),
            model=model,
            voltage=str(voltage) if voltage is not None else None,
            required_val=item.get("required_val", ""),
            profile_signature=signature,
        )
        candidate = str(value or "").strip()
        required = str(item.get("required_val", "") or "").strip()
        param = _norm(item.get("param_name", ""))
        ambiguous = bool(
            candidate and required == "*" and
            (
                len(re.findall(r"\d+(?:[.,]\d+)?", candidate)) > 1
                or re.search(r"\b(?:или|/|;|,)\b", candidate, flags=re.IGNORECASE)
                or any(x in param for x in ("габарит", "масса трансформатора", "масса масла", "климатическ"))
            )
        )
        item["_product_signature"] = signature
        item["algorithm_candidate"] = candidate
        item["algorithm_candidate_source"] = source or "NONE"
        item["algorithm_score"] = score
        item["algorithm_value"] = "" if ambiguous else candidate
        item["algorithm_source"] = "NONE" if ambiguous else (source or "NONE")


def _build_pdf_ai_prompt(
    chunk: list[dict],
    global_summary: list[dict],
    data_tenders_context: dict,
    *,
    full_text: str = "",
    filename: str = "",
) -> str:
    return (
        "Ты — второй уровень контроля автозаполнения технического тендерного PDF-документа "
        "по трансформаторам тока/напряжения.\n\n"
        "ТЫ ВИДИШЬ ВЕСЬ КОНТЕКСТ РАБОТЫ: все строки документа (включая уже обработанные другими "
        "порциями) с результатом первого алгоритмического прохода, применимую базу знаний и "
        "допустимые справочные значения. В этом запросе тебе нужно вернуть ответ только для строк "
        "из блока \"СТРОКИ ДЛЯ ОТВЕТА\", но используй весь контекст для согласованности "
        "(единицы измерения, соседние обмотки, разделы, повторяющиеся параметры).\n\n"
        f"ИМЯ ИСХОДНОГО ФАЙЛА: {filename}\n\n"
        "ВАЖНО: PDF может быть сканом. Основное смысловое чтение страницы выполняет GigaChat Vision, "
        "а Tesseract используется только для геометрии ячеек и как резервный механизм. При конфликте "
        "между Vision и Tesseract доверяй связному Vision-контексту, если он согласуется с соседними "
        "строками, единицами и типом изделия. Не считай отдельную OCR-ошибку новым фактом.\n\n"
        f"ПОЛНЫЙ OCR/ТЕКСТ ДОКУМЕНТА:\n{full_text[:30000]}\n\n"
        f"ПОЛНЫЙ КОРПУС data_tenders (все эталонные файлы, таблицы и OCR изображений):\n{_cached_context_json(data_tenders_context)}\n\n"
        f"ВЕСЬ ДОКУМЕНТ (результат первого алгоритма по каждой строке):\n"
        f"{json.dumps(global_summary, ensure_ascii=False, separators=(",", ":"))}\n\n"
        f"СТРОКИ ДЛЯ ОТВЕТА:\n{json.dumps(chunk, ensure_ascii=False, separators=(",", ":"))}\n\n"
        "ПРАВИЛА:\n"
        "1) Изготовитель, заводской тип и марка сначала сверяются с data_tenders. Заполняй их даже при * слева; "
        'эти значения никогда не помечаются звёздочками.\n'
        "2) Значения, подтвержденные data_tenders, считаются константами и приоритетны. Если алгоритм "
        'правильно взял значение из data_tenders, сохрани его без изменения и source="DATA_TENDERS".\n'
        "2) Перепроверь строки с учетом всего документа: соседних строк, разделов, обмоток, "
        "номеров, единиц измерения и связей между параметрами.\n"
        '3) Если data_tenders не содержит однозначного значения, разрешено вывести значение только по явно достаточному '
        'контексту документа. Тогда source="AI_CONTEXT".\n'
        "4) confidence (0.0-1.0) должен честно отражать твою уверенность. Для строк без значения "
        "в БД всё равно предложи лучший обоснованный ответ по контексту и поставь mark=\"**\".\n"
        '5) Любое значение с source="AI_CONTEXT" должно быть помечено mark="**". data_tenders-значения mark="".\n'
        "6) Если алгоритм поставил значение, но оно явно противоречит требованию/контексту, "
        'исправь его: source="AI_CONTEXT", mark="**", reason с объяснением.\n'
        "7) Не выдумывай паспортные характеристики, которых нет ни в БД, ни в контексте. Если точного "
        "числа нет, укажи наиболее конкретный вывод, который подтверждается ТЗ.\n"
        "8) Если точного ответа нет, верни value=\"\". Не используй текстовые заглушки вроде 'Не определено по ТЗ'. "
        "Программа отдельно попробует взять явное требование из левой колонки как последний резерв.\n"
        "9) algorithm_value для source=HEURISTIC/REQUIREMENT/NONE — только предварительная подсказка программы. Самостоятельно перепроверь её по документу; если твой уверенный вывод лучше, верни его.\n"
        "10) Не переносить значения между строками. id — единственный идентификатор строки.\n\n"
        'ВЕРНИ СТРОГО JSON: {"<id>": {"value":"...", "source":"DB|AI_CONTEXT|NONE", "mark":"|*|**", '
        '"confidence":0.0, "reason":"..."}}. Верни объект для каждого id из блока "СТРОКИ ДЛЯ ОТВЕТА".'
    )


def _adaptive_pdf_chunk_size(items: list[dict], global_summary: list[dict], kb_context: dict, full_text: str, filename: str, max_prompt_chars: int, hard_max: int) -> int:
    n = len(items)
    if n <= 1:
        return n
    lo, hi, best = 1, min(n, hard_max), 1
    while lo <= hi:
        mid = (lo + hi) // 2
        try:
            probe = _build_pdf_ai_prompt(items[:mid], global_summary, kb_context, full_text=full_text, filename=filename)
            size = len(probe)
        except Exception:
            size = max_prompt_chars + 1
        if size <= max_prompt_chars:
            best = mid
            lo = mid + 1
        else:
            hi = mid - 1
    return max(1, best)


_PDF_AI_RESPONSE_CACHE: dict[str, dict[str, dict]] = {}

def _compact_pdf_kb_context(context: dict | None) -> dict:
    if not isinstance(context, dict):
        return {}
    out = {
        "source": context.get("source", "data_tenders"),
        "selection": context.get("selection", {}),
        "model_profiles": context.get("model_profiles", {}),
        "files": [],
        "images": context.get("images", []),
        "selection_rule": context.get("selection_rule", ""),
    }
    for f in context.get("files", []) or []:
        item = {"file": f.get("file", "")}
        if f.get("tables"):
            item["tables"] = f["tables"]
        if f.get("error"):
            item["error"] = f["error"]
        out["files"].append(item)
    return out

def _pdf_ai_cache_key(prompt: str) -> str:
    return hashlib.sha256(prompt.encode("utf-8")).hexdigest()

def _pdf_ai_review(
    client,
    items: list[dict],
    kb_context: dict,
    *,
    full_text: str = "",
    filename: str = "",
    chunk_size: int | None = None,
) -> dict[str, dict]:
    """Прогоняет только незакрытые строки через GigaChat, сохраняя весь контекст документа."""
    data_tenders_context = _compact_pdf_kb_context(kb_context)
    # Все строки остаются в global_summary для контекста зависимостей, но
    # в качестве targets отправляем только реально незаполненные поля.
    pending = [it for it in items if not str(it.get("algorithm_value", "") or "").strip()]
    if not pending:
        return {}
    global_summary = [
        {
            "id": it["id"],
            "number": it["number"],
            "param_name": it["param_name"],
            "required_val": it["required_val"],
            "algorithm_value": it["algorithm_value"],
            "algorithm_source": it["algorithm_source"],
        }
        for it in items
    ]
    merged: dict[str, dict] = {}
    if chunk_size is None:
        try:
            chunk_size = max(20, int(os.getenv("TENDER_AI_PDF_CHUNK_SIZE", "32")))
        except ValueError:
            chunk_size = 32

    # Сначала пытаемся закрыть весь документ одним запросом. Это устраняет
    # десятки сетевых обращений, когда число незаполненных строк умеренное.
    try:
        whole_chunk = [
            {
                "id": it["id"], "number": it["number"],
                "param_name": it["param_name"], "required_val": it["required_val"],
                "algorithm_value": it["algorithm_value"], "algorithm_source": it["algorithm_source"],
            }
            for it in pending
        ]
        whole_prompt = _build_pdf_ai_prompt(whole_chunk, global_summary, data_tenders_context, full_text=full_text, filename=filename)
        max_prompt_chars = int(os.getenv("TENDER_AI_MAX_PROMPT_CHARS", "240000"))
        if len(whole_prompt) <= max_prompt_chars:
            cache_key = _pdf_ai_cache_key(whole_prompt)
            cached = _PDF_AI_RESPONSE_CACHE.get(cache_key)
            if cached is not None:
                valid_ids = {x["id"] for x in items}
                return {k: v for k, v in cached.items() if k in valid_ids and isinstance(v, dict)}
            part = ask_json(client, whole_prompt)
            if isinstance(part, dict):
                valid_ids = {x["id"] for x in items}
                cleaned = {k: v for k, v in part.items() if k in valid_ids and isinstance(v, dict)}
                if len(_PDF_AI_RESPONSE_CACHE) >= 16:
                    _PDF_AI_RESPONSE_CACHE.pop(next(iter(_PDF_AI_RESPONSE_CACHE)))
                _PDF_AI_RESPONSE_CACHE[cache_key] = cleaned
                return cleaned
    except Exception as exc:
        print(f"[GigaChat][PDF] единый запрос не прошёл, переходим на чанки: {exc}")

    chunk_size = _adaptive_pdf_chunk_size(
        items, global_summary, data_tenders_context, full_text, filename,
        max_prompt_chars, max(chunk_size, 128),
    )

    for start in range(0, len(items), chunk_size):
        chunk = [
            {
                "id": it["id"],
                "number": it["number"],
                "param_name": it["param_name"],
                "required_val": it["required_val"],
                "algorithm_value": it["algorithm_value"],
                "algorithm_source": it["algorithm_source"],
            }
            for it in items[start : start + chunk_size]
        ]
        prompt = _build_pdf_ai_prompt(chunk, global_summary, data_tenders_context, full_text=full_text, filename=filename)
        try:
            part = ask_json(client, prompt)
            chunk_ids = {c["id"] for c in chunk}
            for key, value in part.items():
                if key in chunk_ids and isinstance(value, dict):
                    merged[key] = value
        except Exception as exc:
            print(f"[GigaChat] chunk {start + 1}-{start + len(chunk)} failed: {exc}")
    return merged


def _clean_pdf_value(value: str) -> str:
    """Удаляет хвостовые AI-маркеры, но не изменяет само инженерное значение."""
    value = str(value or "").strip()
    if not value or re.fullmatch(r"\*+", value):
        return ""
    value = re.sub(r"\s*\*+\s*$", "", value).strip()
    if value.lower() in {"не определено по тз", "не определено по т.з.", "неизвестно", "unknown"}:
        return ""
    return value


def _pdf_ai_mark(required_val: str) -> str:
    """Одна * если слева уже есть *, иначе стандартные **."""
    return "*" if "*" in str(required_val or "") else "**"


def _merge_algorithm_and_ai_pdf(item: dict, aid: dict) -> tuple[str, str, str, float, str]:
    """Сливает AI и первичный проход; подтвержденные data_tenders-константы неприкасаемы."""
    aid = aid or {}
    ai_value = _clean_pdf_value(aid.get("value", ""))
    try:
        confidence = max(0.0, min(1.0, float(aid.get("confidence", 0) or 0)))
    except (TypeError, ValueError):
        confidence = 0.0
    reason = str(aid.get("reason", "") or "").strip()
    alg_value = _clean_pdf_value(item.get("algorithm_value", ""))
    alg_source = str(item.get("algorithm_source", "NONE") or "NONE")
    alg_is_db = alg_source not in {"NONE", "HEURISTIC", "FALLBACK_ECHO", ""}
    ai_confident = bool(ai_value) and confidence >= PDF_AI_MIN_CONFIDENCE

    def equivalent(a: str, b: str) -> bool:
        a = _clean_pdf_value(a)
        b = _clean_pdf_value(b)
        a = _norm(a).replace(",", ".").replace(" ", "")
        b = _norm(b).replace(",", ".").replace(" ", "")
        strip = lambda x: re.sub(r"(?i)(кв|ка|а|в)$", "", x)
        return a == b or strip(a) == strip(b)

    # Сохранённый специалистом шаблон всегда побеждает AI.
    alg_is_template = alg_source == "DB_TEMPLATE"
    if alg_value and alg_is_template:
        return alg_value, "DB_TEMPLATE", "", confidence, reason

    if alg_value:
        if ai_value and equivalent(ai_value, alg_value):
            return alg_value, alg_source, "", confidence, reason
        if ai_confident:
            return ai_value, "AI_CONTEXT", _pdf_ai_mark(item.get("required_val", "")), confidence, reason
        return alg_value, alg_source, "", confidence, reason

    if ai_value:
        return ai_value, "AI_CONTEXT", _pdf_ai_mark(item.get("required_val", "")), confidence, reason
    return "", "NONE", "", confidence, reason

def _fallback_from_requirement(required_val: str) -> str:
    """Берёт буквальное значение из колонки требования как последний AI-fallback."""
    t = _norm(required_val)
    if not t or re.fullmatch(r"\*+", t):
        return ""
    low = t.lower()
    if "не требуется" in low:
        return "Не требуется"
    if re.search(r"\bтребуется\b", low):
        return "Да"
    if low in {"да", "нет"}:
        return t
    if low in {"обязательно", "обязателен", "обязательна", "обязательны"}:
        return "Да"
    if "не более" in low:
        m = re.search(r"[-+]?\d+(?:[.,]\d+)?", t)
        if m:
            return m.group(0).replace(",", ".")
    if "не менее" in low and len(re.findall(r"\d", t)) == 0:
        return ""
    cleaned = re.sub(r"\s+", " ", t).strip()
    if len(cleaned) <= 2 and not re.search(r"\d", cleaned):
        return ""
    return cleaned


def _deterministic_answer(item: PdfItem, context: dict, detected_voltage: Optional[int]) -> str:
    """Безопасный последний fallback, не притворяющийся значением БД."""
    p = _norm(item.param_name).lower()
    r = _norm(item.required_val)
    rl = r.lower()
    if not p and not r:
        return ""
    if "требуется" in rl and "не требуется" not in rl:
        return "Да"
    if "не требуется" in rl:
        return "Не требуется"
    if "не более" in rl:
        m = re.search(r"(\d+(?:[.,]\d+)?)", r)
        if m:
            return m.group(1).replace(",", ".")
    cleaned = re.sub(r"[\[\]{}|~`_]", "", r).strip()
    if len(cleaned) <= 3 and not re.search(r"\d", cleaned):
        return ""
    return r


def _clear_answer_cell(page: pymupdf.Page, rect: pymupdf.Rect) -> None:
    """Hide the old answer without applying PDF redactions.

    Redaction is intentionally NOT used for scanned PDFs: applying redactions
    repeatedly to the same page can force PyMuPDF to rebuild the embedded page
    image and may damage/rotate rasterized content. A white overlay inside the
    cell is safer and keeps the printed borders untouched.
    """
    inset = 1.2
    r = pymupdf.Rect(
        rect.x0 + inset, rect.y0 + inset,
        rect.x1 - inset, rect.y1 - inset,
    )
    if r.width > 5 and r.height > 4:
        page.draw_rect(r, color=None, fill=(1, 1, 1), overlay=True)


def _find_unicode_font() -> Optional[str]:
    """Find a font with Cyrillic glyphs for inserted Russian answers."""
    candidates = []
    if platform.system() == "Windows":
        candidates = [
            r"C:\Windows\Fonts\arial.ttf",
            r"C:\Windows\Fonts\times.ttf",
            r"C:\Windows\Fonts\calibri.ttf",
        ]
    else:
        candidates = [
            "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
            "/usr/share/fonts/dejavu/DejaVuSans.ttf",
        ]
    for candidate in candidates:
        if os.path.isfile(candidate):
            return candidate
    return None


def _write_answer(page: pymupdf.Page, rect: pymupdf.Rect, value: str) -> None:
    value = _norm(value)
    if not value:
        return

    fontsize = min(9.0, max(5.5, rect.height * 0.62))
    target = pymupdf.Rect(rect.x0 + 2, rect.y0 + 1, rect.x1 - 2, rect.y1 - 1)
    fontfile = _find_unicode_font()

    kwargs = dict(
        fontsize=fontsize,
        color=(0, 0, 0),
        align=0,
        overlay=True,
    )
    if fontfile:
        kwargs.update(fontname="TenderUnicode", fontfile=fontfile)
    else:
        # Keep a fallback for Latin-only installations.
        kwargs.update(fontname="helv")

    rc = page.insert_textbox(target, value, **kwargs)
    if rc < 0:
        # Long values: reduce font until they fit the cell.
        for size in (8, 7, 6, 5):
            kwargs["fontsize"] = size
            rc = page.insert_textbox(target, value, **kwargs)
            if rc >= 0:
                break


def _pil_to_png_bytes(image: Image.Image) -> bytes:
    import io
    buf = io.BytesIO()
    image.save(buf, format="PNG", optimize=True)
    return buf.getvalue()


def fill_pdf_answers(
    input_path: str,
    output_path: str,
    layouts: list[PdfPageLayout],
    results: dict,
) -> None:
    """Write answers into the detected cells.

    Scanned pages are flattened to a clean page image before writing. This is
    intentional: some engineering scans contain an invisible/incorrect OCR
    text layer. PyMuPDF may expose that layer after saving and it can suddenly
    become visible or render vertically. Flattening removes that broken layer
    while preserving the actual scanned drawing exactly as an image.

    `results` maps item id -> the FINAL display text, already including the
    "**" suffix when the value was inferred by the AI from context (see
    process_pdf_requirements).
    """
    src = pymupdf.open(input_path)
    out = pymupdf.open()

    layout_by_page = {layout.page: layout for layout in layouts}
    dpi = int(os.getenv("OUTPUT_DPI", os.getenv("OCR_DPI", str(OCR_DPI))))

    for page_index in range(len(src)):
        src_page = src[page_index]
        is_scanned = _is_scanned_page(src_page)

        if is_scanned:
            # Flatten only scanned pages. PyMuPDF's get_pixmap() already
            # applies page.rotation, so the returned pixels are in the same
            # visual coordinate system as page.rect. Convert the pixmap to a
            # fresh PNG through Pillow to remove any broken/hidden OCR layer.
            scale = dpi / 72.0
            pix = src_page.get_pixmap(
                matrix=pymupdf.Matrix(scale, scale),
                alpha=False,
                annots=True,
            )
            image = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

            page = out.new_page(width=src_page.rect.width, height=src_page.rect.height)
            page.insert_image(
                page.rect,
                stream=_pil_to_png_bytes(image),
                overlay=False,
            )
        else:
            page = out.new_page(width=src_page.rect.width, height=src_page.rect.height)
            page.show_pdf_page(page.rect, src, page_index)

        layout = layout_by_page.get(page_index)
        if not layout:
            continue

        for item in layout.items:
            if not item.answer_rect:
                continue
            value = results.get(item.id, "")
            if value in (None, ""):
                continue
            _clear_answer_cell(page, item.answer_rect)
            _write_answer(page, item.answer_rect, str(value))

    # Save a completely independent PDF. For scanned input this also removes
    # malformed/invisible OCR text layers that were present in the source.
    out.save(output_path, garbage=4, deflate=True)
    out.close()
    src.close()


def process_pdf_requirements(
    pdf_path: str,
    output_path: str,
    session: Session,
    *,
    use_ocr: bool = True,
) -> dict:
    """
    Main public API.

    Returns a diagnostic dictionary so the calling application can show what
    was detected and how many rows were filled/marked, and writes a
    `<output_path>.audit.json` compatible with tools/promote_audit.py.
    """
    layouts, full_text = extract_pdf_items(pdf_path, use_ocr=use_ocr)
    tr_type, detected_voltage = match_tr_type_by_pdf_rules(layouts, session, Path(pdf_path).name)
    data_kb = get_data_tenders_knowledge()
    # Паспорт изделия строим один раз до заполнения строк.
    profile_text = full_text + " " + Path(pdf_path).name
    _selected_model, _product_signature, _profile_candidates = data_kb.select_profile(profile_text, [
        {"param_name": str(x.param_name), "required_val": str(x.required_val), "current_value": str(getattr(x, "existing_answer", ""))}
        for layout in layouts for x in layout.items
    ])
    kb_context = {}

    items: list[dict] = []
    for layout in layouts:
        for item in layout.items:
            items.append(
                {
                    "id": item.id,
                    "page": item.page + 1,
                    "number": item.number,
                    "param_name": item.param_name,
                    "required_val": item.required_val,
                    "_product_signature": dict(_product_signature or {"model": tr_type.name, "voltage_class": str(detected_voltage) if detected_voltage is not None else None}),
                }
            )

    # 1) Алгоритмический проход: расширяемая БЗ -> легаси-справочники -> эвристики.
    _algorithm_fill_pdf(
        session, items, tr_type, detected_voltage, {},
        data_kb=data_kb, document_text=full_text,
    )

    # AI нужен только если после первичного data_tenders-прохода остались поля.
    pending_ai = any(not it.get("algorithm_value") for it in items)
    if os.getenv("TENDER_SKIP_AI", "0") == "1" or not pending_ai:
        print("[GigaChat] AI-проход отключен: TENDER_SKIP_AI=1")
        audit = {}
    else:
        with open_client() as client:
            if os.getenv("TENDER_VERIFY_GIGACHAT", "0") == "1":
                verify_connection(client)
            kb_context = data_kb.build_full_context(
                model=tr_type.name if tr_type else data_kb.detect_model(full_text),
                voltage=str(detected_voltage) if detected_voltage else None,
                include_images=True,
            )
            audit = _pdf_ai_review(client, items, kb_context, full_text=full_text, filename=Path(pdf_path).name)

    # 3) Слияние алгоритма и AI с порогом уверенности + маркировка "**".
    by_id = {it["id"]: it for it in items}
    normalized_results: dict[str, str] = {}
    audit_rows = []
    marked = 0
    unresolved: list[str] = []

    all_items = [item for layout in layouts for item in layout.items]
    for pdf_item in all_items:
        it = by_id[pdf_item.id]
        aid = audit.get(pdf_item.id, {})
        value, source, mark, confidence, reason = _merge_algorithm_and_ai_pdf(it, aid)

        # Уже заполненную колонку участника по умолчанию не стираем. При необходимости
        # принудительной перезаписи можно выставить TENDER_OVERWRITE_EXISTING=1.
        if pdf_item.existing_answer and os.getenv("TENDER_OVERWRITE_EXISTING", "0") != "1":
            value, source, mark = _clean_pdf_value(pdf_item.existing_answer), "EXISTING_ANSWER", ""

        # Существующий ответ в исходном PDF — последний детерминированный резерв
        # резерв перед совсем безопасным эвристическим fallback-ом.
        if not value and pdf_item.existing_answer:
            value, source, mark = _clean_pdf_value(pdf_item.existing_answer), "EXISTING_ANSWER", ""
        if not value and not _is_trivial_row(pdf_item):
            # Поле остается пустым: никаких текстовых заглушек.
            unresolved.append(pdf_item.id)

        if mark:
            marked += 1

        display_value = f"{value}{mark}" if value else ""
        normalized_results[pdf_item.id] = display_value
        audit_rows.append(
            {
                "id": pdf_item.id,
                "page": pdf_item.page + 1,
                "number": it["number"],
                "param_name": it["param_name"],
                "required_val": it["required_val"],
                "algorithm_value": it["algorithm_value"],
                "algorithm_source": it["algorithm_source"],
                "ai_value": _clean_pdf_value(aid.get("value", "")),
                "ai_source": str(aid.get("source", "") or ""),
                "ai_mark": str(aid.get("mark", "") or ""),
                "final_value": value,
                "source": source,
                "mark": mark,
                "confidence": confidence,
                "reason": reason,
                "decision": ("DATA_TENDERS_PRIORITY" if source == "DB" else "AI_OVERRIDE" if source == "AI_CONTEXT" else "ALGORITHM_OR_FALLBACK"),
            }
        )

    filled_count = sum(1 for v in normalized_results.values() if v)
    print(
        f"[PDF FILL] detected_items={len(all_items)} filled={filled_count} "
        f"ai_marked={marked} unresolved={len(unresolved)}"
    )
    if unresolved:
        print(f"[PDF FILL] строки без значения (вероятно, требуют внимания): {unresolved}")

    fill_pdf_answers(pdf_path, output_path, layouts, normalized_results)

    out = Path(output_path)
    audit_path = out.with_suffix(out.suffix + ".audit.json")
    audit_path.write_text(
        json.dumps(
            {
                "run_id": str(uuid.uuid4()),
                "input": str(pdf_path),
                "output": str(out),
                "detected_type": tr_type.name,
                "voltage": detected_voltage,
                "rows": audit_rows,
                "ai_marked_rows": marked,
                "unresolved_rows": unresolved,
                "rules": {
                    "DB_is_constant_source": False,
                    "DB_role": "templates_only",
                    "DATA_TENDERS_is_constant_source": True,
                    "AI_context_mark": "** (or one * when required value already contains *)",
                    "unknown_placeholders": "never",
                    "star_does_not_block_context_fill": True,
                    "ai_confidence_threshold": AI_CONFIDENCE_THRESHOLD,
                    "pdf_ai_min_confidence": PDF_AI_MIN_CONFIDENCE,
                },
            },
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )
    print(f"[OK] PDF сохранен: {out}")
    print(f"[OK] Аудит сохранен: {audit_path}")

    return {
        "input": pdf_path,
        "output": str(out),
        "pages": len(layouts),
        "items": len(items),
        "filled": filled_count,
        "ai_marked_rows": marked,
        "unresolved_rows": len(unresolved),
        "detected_type": tr_type.name,
        "detected_voltage": detected_voltage,
        "text_length": len(full_text),
    }
