from __future__ import annotations

import re
from typing import Iterable

from .schema import ProductContext, TenderRow
from .text import clean, norm


class ProductDetector:
    """Определяет вид изделия, модель и напряжение из самого документа."""

    MODEL_RE = re.compile(r"ТРГ\s*[-–]?\s*УЭТМ\s*[®™]?\s*[-–]?\s*(35|110|220|330|500|750)", re.I)
    VOLTAGE_RE = re.compile(r"(?:номинальн\w*\s+)?напряжен\w*[^0-9]{0,20}(35|110|220|330|500|750)\s*кВ", re.I)

    def detect(self, rows: Iterable[TenderRow], paragraphs: Iterable[str] = ()) -> ProductContext:
        text_parts = list(paragraphs)
        for row in rows:
            text_parts.extend((row.number, row.parameter, row.requirement, row.current_value, row.parent_context))
        text = " ".join(clean(x) for x in text_parts if clean(x))

        model = self._model(text)
        voltage = self._voltage(text, model)
        kind = "трансформатор тока" if "трансформатор" in norm(text) and ("тока" in norm(text) or model) else ""
        if model:
            kind = "трансформатор тока"
        signature = {"product_type": "ТРГ-УЭТМ" if model else None, "model": model, "voltage_class": voltage}
        return ProductContext(kind, model, voltage, "ООО «Эльмаш (УЭТМ)»" if model else None, signature)

    def _model(self, text: str) -> str | None:
        match = self.MODEL_RE.search(text)
        if match:
            return f"ТРГ-УЭТМ-{match.group(1)}"
        voltage = self._voltage(text, None)
        return f"ТРГ-УЭТМ-{voltage}" if voltage else None

    def _voltage(self, text: str, model: str | None) -> str | None:
        if model:
            m = re.search(r"(?:35|110|220|330|500|750)$", model)
            if m:
                return m.group(0)
        match = self.VOLTAGE_RE.search(text)
        if match:
            return match.group(1)
        # В техническом ТЗ заголовок часто содержит «110 кВ».
        match = re.search(r"\b(35|110|220|330|500|750)\s*кВ\b", text, re.I)
        return match.group(1) if match else None
