from __future__ import annotations

import json
import os
from typing import Any

from services.gigachat_client import ask_json_schema, open_client

from .schema import ProductContext, TenderRow
from .text import clean


class TenderAIReviewer:
    """Один пакетный запрос GigaChat только для реально нерешённых строк."""

    def __init__(self, enabled: bool | None = None):
        self.enabled = bool(int(os.getenv("TENDER_AI_ENABLED", "1"))) if enabled is None else enabled
        self.max_chars = int(os.getenv("TENDER_AI_PROMPT_MAX", "28000"))
        self.model_calls = 0

    def review(self, rows: list[TenderRow], context: ProductContext, full_context: str) -> dict[str, dict[str, Any]]:
        unresolved = [r for r in rows if not r.proposed_value]
        if not self.enabled or not unresolved:
            return {}
        result: dict[str, dict[str, Any]] = {}
        chunks = self._chunks(unresolved, full_context)
        with open_client() as client:
            for chunk in chunks:
                prompt = self._prompt(chunk, context, full_context)
                try:
                    response = ask_json_schema(client, prompt, self._schema(), partial_array_key="results")
                    self.model_calls += 1
                except Exception as exc:
                    print(f"[GigaChat] пакет не выполнен: {exc}")
                    continue
                for item in response.get("results", []) or []:
                    row_id = clean(item.get("id"))
                    if row_id:
                        result[row_id] = dict(item)
        return result

    def _chunks(self, rows: list[TenderRow], full_context: str) -> list[list[TenderRow]]:
        static_chars = len(full_context) + 5000
        limit = max(4, (self.max_chars - static_chars) // 180)
        limit = min(max(1, limit), 32)
        return [rows[i:i + limit] for i in range(0, len(rows), limit)]

    def _prompt(self, rows: list[TenderRow], context: ProductContext, full_context: str) -> str:
        targets = []
        for row in rows:
            candidates = " | ".join(row.candidate_values[:6]) or "нет"
            targets.append({
                "id": row.row_id,
                "number": row.number,
                "parent": row.parent_context,
                "parameter": row.parameter,
                "requirement": row.requirement,
                "data_tenders_candidates": candidates,
            })
        return (
            "Ты заполняешь колонку 'Предлагаемое участником конкурса' в технической таблице. "
            "Каждая строка — отдельная задача. Никогда не переносить значение соседней строки. "
            "Поле id является строгим якорем ответа; отвечай только тем id, для которого найдено основание. "
            "Сначала используй подтвержденные сведения data_tenders/БД, затем точное требование текущей строки. "
            "Не выдумывай номера, марки, сертификаты, размеры или составы. Для '*' без основания оставь value пустым. "
            "Для составных строк выбирай только компонент, относящийся к названию целевого параметра. "
            "Для параметров в скобках используй допустимый вариант, если он явно задан в текущем ТЗ. "
            "Если точного значения нет — верни пустое value. Программа сама поставит маркер и желтую подсветку для AI-значений.\n\n"
            f"Паспорт изделия: {json.dumps({'kind': context.product_kind, 'model': context.model, 'voltage': context.voltage, 'manufacturer': context.manufacturer}, ensure_ascii=True)}\n\n"
            f"Независимый контекст документа:\n{full_context}\n\n"
            f"Целевые строки:\n{json.dumps(targets, ensure_ascii=True, indent=2)}\n\n"
            "Для каждой цели верни только JSON: id, value, confidence, evidence, reason. evidence — краткое основание из контекста."
        )

    @staticmethod
    def _schema() -> dict[str, Any]:
        return {
            "type": "object",
            "additionalProperties": False,
            "properties": {
                "results": {
                    "type": "array",
                    "items": {
                        "type": "object",
                        "additionalProperties": False,
                        "properties": {
                            "id": {"type": "string"},
                            "value": {"type": "string"},
                            "confidence": {"type": "number", "minimum": 0, "maximum": 1},
                            "evidence": {"type": "string"},
                            "reason": {"type": "string"},
                        },
                        "required": ["id", "value", "confidence", "evidence", "reason"],
                    },
                }
            },
            "required": ["results"],
        }
