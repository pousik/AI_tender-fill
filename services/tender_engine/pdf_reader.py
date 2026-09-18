from __future__ import annotations

import io
import os
import re
from dataclasses import dataclass
from pathlib import Path
from statistics import median
from typing import Iterable

try:
    import pymupdf
except ImportError:  # pragma: no cover
    import fitz as pymupdf

from PIL import Image

from services.gigachat_client import ask_vision_json_schema, open_client

from .schema import TenderRow
from .text import clean, norm
_WORD = tuple[float, float, float, float, str]


@dataclass(frozen=True)
class PdfGrid:
    """Геометрия технической таблицы на одной странице."""

    x_bounds: tuple[float, ...]
    y_bounds: tuple[float, ...]
    table_rect: pymupdf.Rect
    source: str

    @property
    def answer_rect(self) -> pymupdf.Rect:
        return pymupdf.Rect(self.x_bounds[-2], self.y_bounds[0], self.x_bounds[-1], self.y_bounds[-1])


@dataclass
class PdfRowLayout:
    row: TenderRow
    target_rect: pymupdf.Rect
    page_index: int


class PdfTenderReader:
    """Извлекает строки PDF по реальной геометрии таблицы.

    Приоритет:
    1. текстовый слой + векторные линии таблицы;
    2. OCR + векторные линии;
    3. OCR + линии таблицы, найденные на изображении;
    4. только в самом крайнем случае - координатная эвристика.

    Ключевой принцип: колонка ответа и границы строки определяются отдельно,
    а не процентами ширины страницы и высотой одной текстовой строки.
    """

    HEADER_ANSWER = "предлагаемое"
    MIN_GRID_COLUMNS = 4

    def __init__(self, *, ocr_dpi: int = 220, ocr_enabled: bool = True) -> None:
        self.ocr_dpi = ocr_dpi
        self.ocr_enabled = ocr_enabled

    def read(self, path: str | Path) -> tuple[list[PdfRowLayout], list[str]]:
        layouts: list[PdfRowLayout] = []
        pages_text: list[str] = []
        doc = pymupdf.open(str(path))
        try:
            for page_index, page in enumerate(doc):
                words, text = self._extract_words(page)
                if self.ocr_enabled and self._needs_ocr(words, text):
                    ocr_words = self._ocr_words(page)
                    if len(ocr_words) > max(10, len(words) // 3):
                        words = ocr_words
                        text = " ".join(w[4] for w in words)

                words = self._normalize_words(words)
                pages_text.append(clean(text or " ".join(w[4] for w in words)))

                grid = self._detect_grid(page, words)
                if grid is None:
                    # Даже при плохом/отсутствующем векторном слое пробуем
                    # определить линии из растрового изображения.
                    grid = self._detect_raster_grid(page, words)
                if grid is None:
                    print(f"[PDF][GRID] page={page_index + 1}: table grid not detected; using fallback")
                    layouts.extend(self._fallback_rows(page_index, page, words))
                else:
                    page_rows = self._rows_from_grid(page_index, page, words, grid)
                    if not page_rows:
                        print(f"[PDF][ROWS] page={page_index + 1}: grid found but rows not parsed; fallback")
                        page_rows = self._fallback_rows(page_index, page, words)
                    layouts.extend(page_rows)

            return layouts, pages_text
        finally:
            doc.close()

    @staticmethod
    def _extract_words(page) -> tuple[list, str]:
        try:
            words = page.get_text("words") or []
            text = page.get_text("text") or ""
            return words, text
        except Exception:
            return [], ""

    @staticmethod
    def _needs_ocr(words: list, text: str) -> bool:
        if not words or len(words) < 20:
            return True
        normalized = norm(text)
        # Плохой текстовый слой часто состоит из мусорных/однобуквенных
        # фрагментов. Наличие кириллицы + заголовка обычно означает нормальный слой.
        cyr = len(re.findall(r"[а-яё]", normalized, flags=re.I))
        if cyr >= 30 and "предлагаемое" in normalized:
            return False
        mean_len = sum(len(str(w[4])) for w in words if len(w) > 4) / max(len(words), 1)
        return cyr < 12 or mean_len < 2.2

    @staticmethod
    def _normalize_words(words: Iterable) -> list[_WORD]:
        result: list[_WORD] = []
        for word in words:
            if len(word) < 5:
                continue
            try:
                x0, y0, x1, y1 = map(float, word[:4])
                text = clean(word[4])
            except Exception:
                continue
            if not text:
                continue
            result.append((x0, y0, x1, y1, text))
        return result

    def _ocr_words(self, page) -> list[_WORD]:
        try:
            import pytesseract

            pix = page.get_pixmap(dpi=self.ocr_dpi, alpha=False)
            image = Image.open(io.BytesIO(pix.tobytes("png"))).convert("RGB")
            image = self._ocr_preprocess(image)
            language = self._ocr_language(pytesseract)
            data = pytesseract.image_to_data(
                image,
                lang=language,
                config="--psm 6",
                output_type=pytesseract.Output.DICT,
            )
            sx = page.rect.width / image.width
            sy = page.rect.height / image.height
            result: list[_WORD] = []
            for i, raw in enumerate(data.get("text", [])):
                text = clean(raw)
                if not text:
                    continue
                try:
                    x = float(data["left"][i]) * sx
                    y = float(data["top"][i]) * sy
                    w = float(data["width"][i]) * sx
                    h = float(data["height"][i]) * sy
                except (KeyError, IndexError, TypeError, ValueError):
                    continue
                result.append((x, y, x + w, y + h, text))
            return result
        except Exception as exc:
            print(f"[PDF][OCR] {exc}")
            return []

    @staticmethod
    def _ocr_preprocess(image: Image.Image) -> Image.Image:
        """Убирает линии таблицы перед OCR, чтобы Tesseract не склеивал буквы и границы."""
        try:
            import cv2
            import numpy as np

            gray = np.array(image.convert("L"))
            binary = cv2.threshold(gray, 210, 255, cv2.THRESH_BINARY_INV)[1]
            h_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (max(35, gray.shape[1] // 22), 1))
            v_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, max(35, gray.shape[0] // 22)))
            h_lines = cv2.morphologyEx(binary, cv2.MORPH_OPEN, h_kernel)
            v_lines = cv2.morphologyEx(binary, cv2.MORPH_OPEN, v_kernel)
            line_mask = cv2.bitwise_or(h_lines, v_lines)
            cleaned = gray.copy()
            cleaned[line_mask > 0] = 255
            return Image.fromarray(cleaned)
        except Exception:
            return image

    @staticmethod
    def _ocr_language(pytesseract) -> str:
        try:
            langs = set(pytesseract.get_languages(config=""))
        except Exception:
            return "rus+eng"
        if "rus" in langs and "eng" in langs:
            return "rus+eng"
        if "rus" in langs:
            return "rus"
        return "eng"

    def _detect_grid(self, page, words: list[_WORD]) -> PdfGrid | None:
        verticals: list[tuple[float, float, float]] = []
        horizontals: list[tuple[float, float, float]] = []
        try:
            for drawing in page.get_drawings():
                for item in drawing.get("items", []):
                    if not item:
                        continue
                    if item[0] == "l":
                        p1, p2 = item[1], item[2]
                        dx = abs(float(p2.x) - float(p1.x))
                        dy = abs(float(p2.y) - float(p1.y))
                        if dx <= 1.8 and dy >= 25:
                            x = (float(p1.x) + float(p2.x)) / 2
                            verticals.append((x, min(float(p1.y), float(p2.y)), max(float(p1.y), float(p2.y))))
                        elif dy <= 1.8 and dx >= 35:
                            y = (float(p1.y) + float(p2.y)) / 2
                            horizontals.append((y, min(float(p1.x), float(p2.x)), max(float(p1.x), float(p2.x))))
        except Exception:
            return None

        v_groups = self._group_lines(verticals, axis=0, tolerance=2.5)
        h_groups = self._group_lines(horizontals, axis=0, tolerance=2.5)
        if len(v_groups) < self.MIN_GRID_COLUMNS:
            return None

        header = self._find_answer_header(words)
        answer_x0 = answer_x1 = None
        if header is not None:
            hx0, hx1, hy = header
            before = [g for g in v_groups if g[0] < hx0]
            after = [g for g in v_groups if g[0] > hx1]
            if before and after:
                answer_x0 = before[-1][0]
                answer_x1 = after[0][0]

        if answer_x0 is None:
            # Ищем 5 почти вертикальных границ с максимальным общим диапазоном.
            candidates = sorted(v_groups, key=lambda x: (-x[1], x[0]))
            if len(candidates) >= 5:
                best = sorted(candidates[:8], key=lambda x: x[0])
                # В типовой таблице технических требований последняя колонка - ответ.
                answer_x0, answer_x1 = best[-2][0], best[-1][0]

        if answer_x0 is None or answer_x1 is None or answer_x1 - answer_x0 < 35:
            return None

        answer_parts = [g for g in v_groups if abs(g[0] - answer_x0) <= 2.5 or abs(g[0] - answer_x1) <= 2.5]
        if len(answer_parts) < 2:
            return None
        top = max(g[1] for g in answer_parts)
        bottom = min(g[2] for g in answer_parts)
        if bottom <= top:
            return None

        # Перед ответной колонкой должны находиться три границы:
        # левая граница №, правая граница №/левая параметра, правая параметра.
        # Отбрасываем вертикали других таблиц страницы по их пересечению с найденной
        # областью технической таблицы.
        left_candidates = []
        table_height = bottom - top
        for g in v_groups:
            if g[0] >= answer_x0 - 2.5:
                continue
            overlap = max(0.0, min(g[2], bottom) - max(g[1], top))
            if overlap >= max(25.0, table_height * 0.55):
                left_candidates.append(g)
        if len(left_candidates) < 3:
            return None
        left = left_candidates[-3:]
        x_bounds = tuple(sorted([g[0] for g in left] + [answer_x0, answer_x1]))
        if len(x_bounds) != 5:
            return None

        row_lines = [g[0] for g in h_groups if g[1] <= answer_x0 + 4 and g[2] >= answer_x1 - 4 and top - 3 <= g[0] <= bottom + 3]
        row_lines = self._dedupe(row_lines, 2.2)
        if len(row_lines) < 3:
            return None
        if abs(row_lines[0] - top) > 5:
            row_lines.insert(0, top)
        if abs(row_lines[-1] - bottom) > 5:
            row_lines.append(bottom)
        row_lines = tuple(sorted(self._dedupe(row_lines, 2.2)))
        if len(row_lines) < 3:
            return None
        return PdfGrid(x_bounds=x_bounds, y_bounds=row_lines, table_rect=pymupdf.Rect(x_bounds[0], row_lines[0], x_bounds[-1], row_lines[-1]), source="vector")

    def _detect_raster_grid(self, page, words: list[_WORD]) -> PdfGrid | None:
        try:
            import cv2
            import numpy as np
        except ImportError:
            return None

        try:
            pix = page.get_pixmap(dpi=self.ocr_dpi, alpha=False)
            image = Image.open(io.BytesIO(pix.tobytes("png"))).convert("L")
            arr = np.array(image)
            # Табличные линии значительно длиннее штрихов букв. Морфология удаляет текст,
            # оставляя устойчивые горизонтальные/вертикальные границы.
            binary = cv2.threshold(arr, 210, 255, cv2.THRESH_BINARY_INV)[1]
            h_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (max(35, arr.shape[1] // 18), 1))
            v_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, max(35, arr.shape[0] // 18)))
            h = cv2.morphologyEx(binary, cv2.MORPH_OPEN, h_kernel)
            v = cv2.morphologyEx(binary, cv2.MORPH_OPEN, v_kernel)

            sx = page.rect.width / arr.shape[1]
            sy = page.rect.height / arr.shape[0]
            verticals = self._raster_line_positions(v, vertical=True, scale=sx)
            horizontals = self._raster_line_positions(h, vertical=False, scale=sy)
            if len(verticals) < 4 or len(horizontals) < 3:
                return None

            # При скане проще всего выбрать сетку, внутри которой расположен OCR-заголовок
            # «Предлагаемое участником конкурса».
            header = self._find_answer_header(words)
            if header:
                hx0, hx1, _ = header
                before = [x for x in verticals if x < hx0]
                after = [x for x in verticals if x > hx1]
                if before and after:
                    ax0, ax1 = before[-1], after[0]
                else:
                    ax0, ax1 = verticals[-2], verticals[-1]
            else:
                ax0, ax1 = verticals[-2], verticals[-1]

            left = [x for x in verticals if x < ax0 - 2]
            if len(left) >= 3:
                left_bounds = left[-3:]
            elif len(left) == 2:
                left_bounds = left[-2:]
                # Скан нередко теряет левую рамку из-за обрезанного края.
                # Берём её из координаты номера первой заполненной строки.
                numeric_x = [w[0] for w in words if re.fullmatch(r"\d+(?:\.\d+)*\.?", clean(w[4]))]
                if not numeric_x:
                    return None
                left_bounds = [min(numeric_x) - 5.0] + left_bounds
            else:
                return None
            x_bounds = tuple(sorted(left_bounds + [ax0, ax1]))
            if len(x_bounds) != 5:
                return None
            table_horiz = self._select_horizontals_for_x(horizontals, ax0, ax1)
            if len(table_horiz) < 3:
                return None
            top, bottom = min(table_horiz), max(table_horiz)
            y_bounds = tuple(x for x in table_horiz if top <= x <= bottom)
            if len(y_bounds) < 3:
                return None
            return PdfGrid(x_bounds=x_bounds, y_bounds=y_bounds, table_rect=pymupdf.Rect(x_bounds[0], y_bounds[0], x_bounds[-1], y_bounds[-1]), source="raster")
        except Exception as exc:
            print(f"[PDF][GRID][RASTER] {exc}")
            return None

    @staticmethod
    def _raster_line_positions(mask, *, vertical: bool, scale: float) -> list[float]:
        import cv2

        contours, _ = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        positions: list[float] = []
        for contour in contours:
            x, y, w, h = cv2.boundingRect(contour)
            size = h if vertical else w
            if size < 100:
                continue
            pos = (x + w / 2) * scale if vertical else (y + h / 2) * scale
            positions.append(pos)
        return PdfTenderReader._dedupe(sorted(positions), 3.0)

    @staticmethod
    def _select_horizontals_for_x(lines: list[float], x0: float, x1: float) -> list[float]:
        # На этапе морфологии потерялась длина линии, поэтому оставляем плотную
        # последовательность; длинные строки таблицы обычно образуют основной кластер.
        if len(lines) <= 3:
            return lines
        groups: list[list[float]] = []
        for y in lines:
            if not groups or y - groups[-1][-1] > 8:
                groups.append([y])
            else:
                groups[-1].append(y)
        return [median(g) for g in groups]

    @staticmethod
    def _group_lines(lines: list[tuple[float, float, float]], *, axis: int, tolerance: float) -> list[tuple[float, float, float]]:
        if not lines:
            return []
        ordered = sorted(lines, key=lambda x: x[axis])
        groups: list[list[tuple[float, float, float]]] = []
        for item in ordered:
            if not groups or abs(item[axis] - groups[-1][-1][axis]) > tolerance:
                groups.append([item])
            else:
                groups[-1].append(item)
        result = []
        for group in groups:
            pos = median(x[axis] for x in group)
            start = min(x[1] for x in group)
            end = max(x[2] for x in group)
            result.append((pos, start, end))
        return result

    @staticmethod
    def _find_answer_header(words: list[_WORD]) -> tuple[float, float, float] | None:
        candidates = [w for w in words if norm(w[4]) == "предлагаемое"]
        if not candidates:
            candidates = [w for w in words if norm(w[4]).startswith("предлагаем")]
        if not candidates:
            return None
        w = candidates[0]
        return w[0], w[2], (w[1] + w[3]) / 2

    def _rows_from_grid(self, page_index: int, page, words: list[_WORD], grid: PdfGrid) -> list[PdfRowLayout]:
        rows: list[PdfRowLayout] = []
        parent = ""
        section = ""
        row_counter = 0

        for y0, y1 in zip(grid.y_bounds, grid.y_bounds[1:]):
            if y1 - y0 < 5:
                continue
            band_words = [w for w in words if y0 - 1.5 <= self._center_y(w) <= y1 + 1.5]
            if not band_words:
                continue

            if len(grid.x_bounds) < 5:
                continue
            x0, x1, x2, x3, x4 = grid.x_bounds[:5]
            number = self._cell_text(band_words, x0, x1)
            middle_left, middle_right = x1, x2
            requirement = self._cell_text(band_words, x2, x3)
            answer_raw = self._cell_text(band_words, x3, x4)

            parameter, parent, nested = self._parameter_and_parent(
                band_words, middle_left, middle_right, parent, y0, y1, page
            )

            # Пропускаем пустые и заголовочные строки.
            if not parameter and not requirement and not answer_raw:
                continue
            parameter_norm = norm(parameter)
            if "предлагаемое участником" in parameter_norm or "технические требования к оборудованию" in parameter_norm:
                continue

            is_header = self._is_header_row(number, parameter, requirement, answer_raw)
            if is_header:
                num_clean = number.rstrip(".")
                if num_clean and re.fullmatch(r"\d+(?:\.\d+)?", num_clean):
                    section = num_clean
                if parameter and not nested:
                    parent = ""
                continue

            answer = self._strip_placeholder(answer_raw)
            field_parameter = parameter
            if not field_parameter and requirement:
                # Для редких таблиц с отсутствующим текстом в средней колонке.
                field_parameter = requirement
                requirement = ""

            key = self._field_key(field_parameter, parent)
            row = TenderRow(
                f"p{page_index}_r{row_counter}",
                0,
                page_index,
                number,
                field_parameter,
                requirement,
                answer,
                0,
                parent if nested or key in {"accuracy_class", "secondary_load"} else "",
                section,
                key,
                is_header=False,
            )
            rect = self._answer_rect(grid, y0, y1, page.rect)
            rows.append(PdfRowLayout(row, rect, page_index))
            row_counter += 1

        return rows

    def _parameter_and_parent(self, words, left: float, right: float, parent: str, y0: float, y1: float, page):
        local = self._local_internal_verticals(page, left, right, y0, y1)
        if local:
            split = local[0]
            parent_text = self._cell_text(words, left, split)
            subparam = self._cell_text(words, split, right)
            if parent_text:
                parent = parent_text
            return subparam or parent_text, parent, True
        parameter = self._cell_text(words, left, right)
        # Для обычной строки родительскую обмотку не переносим дальше.
        return parameter, parent if self._looks_winding_parameter(parameter) else "", False

    @staticmethod
    def _looks_winding_parameter(value: str) -> bool:
        low = norm(value)
        return "обмотка" in low and ("учет" in low or "измер" in low or "защит" in low)

    @staticmethod
    def _local_internal_verticals(page, left: float, right: float, y0: float, y1: float) -> list[float]:
        xs: list[float] = []
        try:
            for drawing in page.get_drawings():
                for item in drawing.get("items", []):
                    if not item or item[0] != "l":
                        continue
                    p1, p2 = item[1], item[2]
                    if abs(float(p2.x) - float(p1.x)) <= 1.8:
                        x = (float(p1.x) + float(p2.x)) / 2
                        a, b = min(float(p1.y), float(p2.y)), max(float(p1.y), float(p2.y))
                        overlap = max(0.0, min(b, y1) - max(a, y0))
                        if left + 15 < x < right - 15 and overlap >= max(4.0, (y1 - y0) * 0.65):
                            xs.append(x)
        except Exception:
            pass
        return PdfTenderReader._dedupe(sorted(xs), 2.5)

    @staticmethod
    def _center_y(word: _WORD) -> float:
        return (word[1] + word[3]) / 2

    @staticmethod
    def _cell_text(words: list[_WORD], x0: float, x1: float) -> str:
        selected = [w for w in words if x0 - 1.0 <= (w[0] + w[2]) / 2 <= x1 + 1.0]
        if not selected:
            return ""
        selected.sort(key=lambda w: (round(w[1], 1), w[0]))
        return clean(" ".join(w[4] for w in selected))

    @staticmethod
    def _strip_placeholder(value: str) -> str:
        value = clean(value)
        if not value:
            return ""
        if re.fullmatch(r"[*\s/]+", value):
            return ""
        return value

    @staticmethod
    def _is_header_row(number: str, parameter: str, requirement: str, answer: str) -> bool:
        low = norm(parameter)
        if not parameter:
            return False
        if not requirement and not answer:
            if re.fullmatch(r"\d+(?:\.\d+)?\.?", clean(number)) and (
                low.startswith("основные технические")
                or low.startswith("технические требования")
                or low.startswith("требования к конструкции")
                or low.startswith("для тр")
                or low.startswith("массо-габаритные")
                or low.startswith("климатическое исполнение")
                or low.startswith("требования по")
                or low.startswith("требования к серв")
                or low.startswith("комплектность")
                or low.startswith("гарантии изготовителя")
                or low.startswith("параметры вторичных")
                or low.endswith(":")
            ):
                return True
            if low in {"№ п/п", "технические требования к оборудованию", "требования", "значение параметра"}:
                return True
        return False

    def _fallback_rows(self, page_index: int, page, words: list[_WORD]) -> list[PdfRowLayout]:
        """Последний fallback. Здесь нет привязки к 34%/72%.

        Колонка ответа вычисляется из заголовка, остальные границы - по его соседним
        координатам. Строки группируются по реальным y текстового слоя/OCR.
        """
        if not words:
            return []
        header = self._find_answer_header(words)
        if header:
            hx0, hx1, _ = header
            # Для распространённой формы задаём только ответную границу;
            # остальные границы выводим из минимально разумных координат.
            answer_left = min(w[0] for w in words if w[0] > hx0 - 80 and w[0] < hx0 + 10) if any(w[0] > hx0 - 80 and w[0] < hx0 + 10 for w in words) else hx0 - 8
        else:
            answer_left = page.rect.width * 0.70
        lines = self._text_lines(words)
        result: list[PdfRowLayout] = []
        counter = 0
        for line in lines:
            text = clean(" ".join(w[4] for w in line))
            m = re.match(r"^(\d+(?:\.\d+)*\.?)\s+(.+)$", text)
            if not m:
                continue
            number = m.group(1)
            parameter = clean(m.group(2))
            if norm(parameter).startswith("технические требования"):
                continue
            row = TenderRow(f"p{page_index}_r{counter}", 0, page_index, number, parameter, "", "", 0, "", "", self._field_key(parameter, ""))
            result.append(PdfRowLayout(row, pymupdf.Rect(answer_left, min(w[1] for w in line), page.rect.width - 4, max(w[3] for w in line) + 2), page_index))
            counter += 1
        return result

    @staticmethod
    def _text_lines(words: list[_WORD]) -> list[list[_WORD]]:
        if not words:
            return []
        words = sorted(words, key=lambda w: (w[1], w[0]))
        lines: list[list[_WORD]] = []
        for word in words:
            if not lines:
                lines.append([word])
                continue
            prev = lines[-1]
            base = sum(w[1] + w[3] for w in prev) / (2 * len(prev))
            height = median(w[3] - w[1] for w in prev)
            if abs(((word[1] + word[3]) / 2) - base) <= max(3.0, height * 0.55):
                prev.append(word)
            else:
                lines.append([word])
        return [sorted(line, key=lambda w: w[0]) for line in lines]

    @staticmethod
    def _answer_rect(grid: PdfGrid, y0: float, y1: float, page_rect: pymupdf.Rect) -> pymupdf.Rect:
        left = grid.x_bounds[-2]
        right = min(grid.x_bounds[-1], page_rect.width)
        # Не закрываем линии рамки - только внутреннюю область ячейки.
        return pymupdf.Rect(left + 2.0, min(page_rect.height, y0 + 1.2), max(left + 8, right - 2.0), max(y0 + 6, min(page_rect.height, y1 - 1.2)))

    @staticmethod
    def _field_key(parameter: str, parent: str) -> str:
        from .rules import field_key
        return field_key(parameter, parent)

    @staticmethod
    def _dedupe(values: Iterable[float], tolerance: float) -> list[float]:
        result: list[float] = []
        for value in sorted(values):
            if not result or abs(value - result[-1]) > tolerance:
                result.append(value)
            else:
                result[-1] = (result[-1] + value) / 2
        return result


