from __future__ import annotations

import json
from typing import Any

from services.gigachat_client import ask_json


SYSTEM_PROMPT = """
Ты анализируешь тендерный документ трансформатора.

Тебе передаётся результат OCR PDF.

Нужно восстановить смысл таблицы.

Не придумывай значения.

Верни ТОЛЬКО JSON следующего формата:

{
  "document_type": "...",
  "product_type": "...",
  "voltage": "...",
  "items": [
    {
      "name": "...",
      "value": "...",
      "page": null,
      "bbox": null,
      "confidence": null,
      "source": "ocr"
    }
  ]
}

Правила:

1. name — название характеристики.
2. value — фактическое значение.
3. Не объединяй разные характеристики.
4. Не меняй значения.
5. Не исправляй OCR самостоятельно, если нет уверенности.
6. Если значение отсутствует — value = "".
7. Не создавай значения, которых нет в документе.
8. Если рядом есть номер строки/позиции, не считай его характеристикой.
"""


def _ocr_to_text(ocr_data: Any) -> str:

    if isinstance(ocr_data, str):
        return ocr_data

    if not isinstance(ocr_data, list):
        return str(ocr_data)

    result = []

    for item in ocr_data:

        if not isinstance(item, dict):
            result.append(str(item))
            continue

        text = item.get("text", "")

        if not text:
            continue

        page = item.get("page")
        x0 = item.get("x0")
        y0 = item.get("y0")

        result.append(
            f"[page={page};x={x0};y={y0}] "
            f"{text}"
        )

    return "\n".join(result)


def extract_pdf_structure(
    ocr_data,
    path: str,
) -> dict:

    ocr_text = _ocr_to_text(
        ocr_data
    )

    prompt = f"""
{SYSTEM_PROMPT}

Файл:

{path}

OCR:

{ocr_text}
"""

    result = ask_json(
        prompt,
    )

    if not isinstance(result, dict):

        raise RuntimeError(
            "GigaChat вернул не JSON-object"
        )

    return result