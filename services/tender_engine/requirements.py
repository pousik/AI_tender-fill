from __future__ import annotations

import re

from .schema import TenderRow
from .text import clean, norm, strip_marks


class RequirementResolver:
    """Извлекает только безопасные значения непосредственно из требования."""

    def resolve(self, row: TenderRow) -> str:
        req = clean(row.requirement)
        if not req or re.fullmatch(r"\*+", req):
            return ""
        low = norm(req)
        if low in {"да", "нет"}:
            return req
        # «Да, обязательно ...»/«Нет, ...» — это всё равно однозначный ответ.
        leading = re.match(r"^(да|нет)(?:\s*[,;:]|\s+)", low)
        if leading:
            return leading.group(1).capitalize()
        if low in {"обязательно", "обязателен", "обязательна", "обязательны"}:
            return "Да"
        if low in {"а/м транспорт", "автотранспорт"}:
            return req
        if re.fullmatch(r"(?:плюс|минус)\s*[-+]?\d+(?:[,.]\d+)?", low):
            return req
        if low.startswith("не менее") or low.startswith("не более"):
            # Если есть число — это порог. Без числа это может быть готовая
            # текстовая величина, например «не менее гарантийного срока».
            m = re.search(r"[-+]?\d+(?:[,.]\d+)?", low)
            return m.group(0).replace(".", ",") if m else req
        if "в соответствии с" in low or low.startswith("согласно") or low.startswith("указать"):
            return ""
        if re.fullmatch(r"[\w\-/.%,+]+", req, flags=re.UNICODE):
            return strip_marks(req)
        return ""
