from __future__ import annotations

import re
from difflib import SequenceMatcher


def clean(value: object) -> str:
    text = re.sub(r"\s+", " ", str(value or "")).strip()
    # DOCX/OCR иногда возвращают непарные UTF-16 surrogate-символы.
    return "".join(ch for ch in text if not 0xD800 <= ord(ch) <= 0xDFFF)


def norm(value: object) -> str:
    text = clean(value).lower().replace("ё", "е")
    text = text.replace("®", "").replace("™", "")
    text = text.replace("–", "-").replace("—", "-")
    text = re.sub(r"[\u00a0\u2009\u200b]", " ", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def numbers(value: object) -> list[float]:
    text = str(value or "")
    # Русские словесные знаки часто встречаются в ТЗ: «плюс 40», «минус 55».
    text = re.sub(r"\bминус\s*(?=\d)", "-", text, flags=re.I)
    text = re.sub(r"\bплюс\s*(?=\d)", "+", text, flags=re.I)
    out: list[float] = []
    for token in re.findall(r"[-+]?\d+(?:[,.]\d+)?", text):
        try:
            out.append(float(token.replace(",", ".")))
        except ValueError:
            pass
    return out


def number_strings(value: object) -> list[str]:
    return re.findall(r"[-+]?\d+(?:[,.]\d+)?", str(value or ""))


def has_number(value: object, target: float, tol: float = 1e-9) -> bool:
    return any(abs(x - target) <= tol for x in numbers(value))


def similarity(left: object, right: object) -> float:
    a, b = norm(left), norm(right)
    if not a or not b:
        return 0.0
    if a == b:
        return 1.0
    if a in b or b in a:
        return 0.96
    return SequenceMatcher(None, a, b).ratio()


def strip_marks(value: object) -> str:
    text = clean(value)
    return re.sub(r"\s*\*+\s*$", "", text).strip()


def is_placeholder(value: object) -> bool:
    text = clean(value)
    return not text or bool(re.fullmatch(r"\*+", text))


def is_numbered(value: object) -> bool:
    return bool(re.fullmatch(r"\d+(?:\.\d+)*\.?", clean(value)))


def split_lines(value: object) -> list[str]:
    return [clean(x) for x in re.split(r"\n+", str(value or "")) if clean(x)]
