from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Any

import pymupdf


# ============================================================
# PDF TEXT/OCR
# ============================================================

def extract_pdf_text_with_coordinates(
    path: str | Path,
) -> list[dict[str, Any]]:
    """
    Извлекает текст PDF вместе с координатами.

    Сначала используется встроенный текстовый слой PDF.
    OCR сюда намеренно не зашит: если PDF сканированный,
    вызывается OCR-функция проекта.

    Формат результата:

    [
        {
            "page": 0,
            "text": "ТРГ-110",
            "x0": 100.0,
            "y0": 200.0,
            "x1": 180.0,
            "y1": 220.0,
        }
    ]
    """

    path = Path(path)

    if not path.exists():
        raise FileNotFoundError(path)

    result: list[dict[str, Any]] = []

    doc = pymupdf.open(str(path))

    try:
        for page_number, page in enumerate(doc):

            words = page.get_text("words")

            for word in words:
                if len(word) < 5:
                    continue

                x0, y0, x1, y1, text = word[:5]

                text = str(text).strip()

                if not text:
                    continue

                result.append({
                    "page": page_number,
                    "text": text,
                    "x0": float(x0),
                    "y0": float(y0),
                    "x1": float(x1),
                    "y1": float(y1),
                })

    finally:
        doc.close()

    return result


# ============================================================
# NORMALIZATION
# ============================================================

def normalize_text(value: Any) -> str:
    if value is None:
        return ""

    value = str(value)

    value = value.replace("\xa0", " ")
    value = value.replace("\u200b", "")

    value = re.sub(r"\s+", " ", value)

    return value.strip()


def normalize_value(value: Any) -> str:
    """
    Нормализация значения перед сравнением.
    """

    value = normalize_text(value)

    value = value.lower()

    # OCR часто путает символы
    value = value.replace("ё", "е")

    # Разные виды тире
    value = value.replace("–", "-")
    value = value.replace("—", "-")

    # Не считаем разный регистр изменением
    return value


# ============================================================
# GENERIC STATE
# ============================================================

def normalize_state(state: dict[str, Any] | None) -> dict[str, Any]:
    if not state:
        return {
            "document_type": None,
            "voltage": None,
            "product_type": None,
            "items": [],
        }

    result = {
        "document_type": state.get("document_type"),
        "voltage": state.get("voltage"),
        "product_type": state.get("product_type"),
        "items": [],
    }

    items = state.get("items", [])

    if not isinstance(items, list):
        return result

    for item in items:
        if not isinstance(item, dict):
            continue

        name = normalize_text(
            item.get("name")
            or item.get("field")
            or item.get("label")
        )

        value = normalize_text(
            item.get("value")
        )

        if not name:
            continue

        result["items"].append({
            "name": name,
            "value": value,
            "page": item.get("page"),
            "bbox": item.get("bbox"),
            "confidence": item.get("confidence"),
            "source": item.get("source"),
        })

    return result


# ============================================================
# ITEM MATCHING
# ============================================================

def _field_key(name: str) -> str:
    name = normalize_value(name)

    # Убираем мусор
    name = re.sub(r"[^a-zа-я0-9]+", " ", name)

    return name.strip()


def _build_item_index(
    items: list[dict[str, Any]],
) -> dict[str, dict[str, Any]]:

    result = {}

    for item in items:
        name = item.get("name", "")

        key = _field_key(name)

        if not key:
            continue

        result[key] = item

    return result


# ============================================================
# STATE COMPARISON
# ============================================================

def compare_pdf_states(
    old_state: dict[str, Any] | None,
    new_state: dict[str, Any] | None,
) -> dict[str, Any]:

    old_state = normalize_state(old_state)
    new_state = normalize_state(new_state)

    old_items = _build_item_index(old_state["items"])
    new_items = _build_item_index(new_state["items"])

    changes = []

    # --------------------------------------------------------
    # Сначала сравниваем известные поля
    # --------------------------------------------------------

    for key, old_item in old_items.items():

        new_item = new_items.get(key)

        if not new_item:
            continue

        old_value = normalize_text(old_item.get("value"))
        new_value = normalize_text(new_item.get("value"))

        if not old_value and not new_value:
            continue

        if normalize_value(old_value) == normalize_value(new_value):
            continue

        changes.append({
            "field": old_item.get("name") or new_item.get("name"),
            "old": old_value,
            "new": new_value,
            "page_old": old_item.get("page"),
            "page_new": new_item.get("page"),
            "bbox_old": old_item.get("bbox"),
            "bbox_new": new_item.get("bbox"),
            "old_confidence": old_item.get("confidence"),
            "new_confidence": new_item.get("confidence"),
        })

    # --------------------------------------------------------
    # Общие изменения
    # --------------------------------------------------------

    return {
        "document_type": new_state.get("document_type")
        or old_state.get("document_type"),

        "product_type": new_state.get("product_type")
        or old_state.get("product_type"),

        "voltage": new_state.get("voltage")
        or old_state.get("voltage"),

        "changes": changes,

        "changed_count": len(changes),
    }


# ============================================================
# PDF STATE CAPTURE
# ============================================================

def capture_pdf_state(
    path: str | Path,
    *,
    ai_extractor=None,
    ocr_extractor=None,
) -> dict[str, Any]:
    """
    Получает структурированное состояние PDF.

    Приоритет:

        1. OCR/существующий PDF extractor проекта
        2. GigaChat
        3. встроенный текст PDF

    ai_extractor должен вернуть dict.

    ocr_extractor должен вернуть либо:
        str
    либо:
        list[dict]
    """

    path = Path(path)

    if not path.exists():
        raise FileNotFoundError(path)

    # --------------------------------------------------------
    # 1. Используем существующий OCR проекта
    # --------------------------------------------------------

    ocr_data = None

    if ocr_extractor is not None:

        try:
            ocr_data = ocr_extractor(str(path))
        except Exception as exc:
            print(f"[PDF STATE] OCR ошибка: {exc}")

    # --------------------------------------------------------
    # 2. Если OCR ничего не дал — встроенный текст PDF
    # --------------------------------------------------------

    if not ocr_data:

        words = extract_pdf_text_with_coordinates(path)

        ocr_data = words

    # --------------------------------------------------------
    # 3. GigaChat превращает распознанный документ
    #    в структурированные данные
    # --------------------------------------------------------

    if ai_extractor is not None:

        try:

            state = ai_extractor(
                ocr_data,
                str(path),
            )

            if isinstance(state, dict):
                return normalize_state(state)

        except Exception as exc:
            print(
                f"[PDF STATE] GigaChat ошибка: {exc}"
            )

    # --------------------------------------------------------
    # 4. Fallback
    # --------------------------------------------------------

    text_parts = []

    if isinstance(ocr_data, str):
        text_parts.append(ocr_data)

    elif isinstance(ocr_data, list):

        for item in ocr_data:

            if isinstance(item, dict):
                text = item.get("text")

                if text:
                    text_parts.append(str(text))

            else:
                text_parts.append(str(item))

    text = "\n".join(text_parts)

    return {
        "document_type": None,
        "product_type": None,
        "voltage": None,

        "items": [
            {
                "name": "__raw_text__",
                "value": text,
                "source": "fallback",
            }
        ],
    }


# ============================================================
# SAVE SNAPSHOT
# ============================================================

def save_pdf_snapshot(
    path: str | Path,
    state: dict[str, Any],
) -> Path:

    path = Path(path)

    snapshot_path = path.with_suffix(
        path.suffix + ".state.json"
    )

    with snapshot_path.open(
        "w",
        encoding="utf-8",
    ) as f:

        json.dump(
            state,
            f,
            ensure_ascii=False,
            indent=2,
        )

    return snapshot_path


def load_pdf_snapshot(
    path: str | Path,
) -> dict[str, Any] | None:

    path = Path(path)

    snapshot_path = path.with_suffix(
        path.suffix + ".state.json"
    )

    if not snapshot_path.exists():
        return None

    try:
        with snapshot_path.open(
            "r",
            encoding="utf-8",
        ) as f:

            data = json.load(f)

        if isinstance(data, dict):
            return data

    except Exception as exc:
        print(
            f"[PDF STATE] Не удалось загрузить snapshot: {exc}"
        )

    return None


def delete_pdf_snapshot(path: str | Path):
    path = Path(path)

    snapshot_path = path.with_suffix(
        path.suffix + ".state.json"
    )

    try:
        snapshot_path.unlink(missing_ok=True)
    except Exception:
        pass