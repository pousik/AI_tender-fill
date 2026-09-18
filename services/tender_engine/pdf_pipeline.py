from __future__ import annotations

import re
from pathlib import Path

try:
    import pymupdf
except ImportError:  # pragma: no cover
    import fitz as pymupdf

from services.gigachat_client import ask_vision_json_schema, open_client

from .context import ProductDetector
from .data_source import DataTendersRepository
from .pdf_reader import PdfRowLayout, PdfTenderReader
from .pdf_writer import PdfTenderWriter
from .resolver import ResolverConfig, ValueResolver
from .schema import ProductContext, TenderRow
from .text import clean, strip_marks
from .validator import TenderValidator


class PdfTenderFillingEngine:
    """Гибридное заполнение PDF: БД/ТЗ -> AI только для unresolved -> валидация -> запись."""

    def __init__(self, data_tenders_root: str | Path, *, ai_enabled: bool = True, ocr_enabled: bool = True) -> None:
        self.reader = PdfTenderReader(ocr_enabled=ocr_enabled)
        self.repo = DataTendersRepository(data_tenders_root)
        self.resolver = ValueResolver(self.repo, ResolverConfig(overwrite_existing=False))
        self.detector = ProductDetector()
        self.validator = TenderValidator()
        self.ai_enabled = ai_enabled

    def fill(self, input_path: str | Path, output_path: str | Path, *, overwrite_existing: bool = False) -> dict:
        layouts, pages = self.reader.read(input_path)
        rows = [item.row for item in layouts]
        if not rows:
            raise ValueError("В PDF не обнаружена техническая таблица с заполняемыми строками.")

        context = self.detector.detect(rows, pages)
        if not context.model:
            raise ValueError("Не удалось определить модель ТТ в PDF.")

        unresolved: list[TenderRow] = []
        stats = {"existing": 0, "database": 0, "requirement": 0, "ai": 0, "unresolved": 0}
        self.resolver.config = ResolverConfig(overwrite_existing=overwrite_existing)

        for row in rows:
            if row.is_header:
                continue
            decision = self.resolver.resolve_initial(row, context)
            row.proposed_value = decision.value
            row.source = decision.source
            row.mark = decision.mark
            row.confidence = decision.confidence
            row.reason = decision.reason
            row.evidence = list(decision.evidence)
            if not row.proposed_value:
                unresolved.append(row)
            elif decision.source == "EXISTING_ANSWER":
                stats["existing"] += 1
            elif decision.source == "REQUIREMENT":
                stats["requirement"] += 1
            else:
                stats["database"] += 1

        if self.ai_enabled and unresolved:
            self._vision_ai(unresolved, input_path, context)

        for row in rows:
            if row.is_header or not row.proposed_value:
                continue
            ok, reason = self.validator.validate(row, row.proposed_value, rows)
            if not ok:
                row.proposed_value = ""
                row.mark = ""
                row.source = "UNRESOLVED"
                row.reason = reason

        stats["ai"] = sum(1 for row in rows if row.source.startswith("AI"))
        stats["unresolved"] = sum(1 for row in rows if not row.proposed_value and not row.is_header)

        target = Path(output_path).expanduser().resolve()
        target.parent.mkdir(parents=True, exist_ok=True)
        temp = target.with_suffix(target.suffix + ".tmp.pdf")
        base = pymupdf.open(str(input_path))
        try:
            base.save(str(temp), garbage=4, deflate=True)
        finally:
            base.close()
        PdfTenderWriter().write(temp, layouts)
        temp.replace(target)

        return {
            "output": str(target),
            "filled": sum(bool(r.proposed_value) for r in rows),
            "unresolved": [r.row_id for r in rows if not r.proposed_value and not r.is_header],
            "detected_type": context.model,
            "voltage": context.voltage,
            "pages": len(pages),
            "rows": len(rows),
            "stats": stats,
        }

    def _vision_ai(self, unresolved: list[TenderRow], pdf_path: str | Path, context: ProductContext) -> None:
        try:
            kb_context = self.repo.ai_context(context.model, max_chars=18000)
            with open_client() as client:
                doc = pymupdf.open(str(pdf_path))
                try:
                    page_numbers = sorted({self._page_from_row_id(row.row_id) for row in unresolved})
                    for page_no in page_numbers:
                        rows = [r for r in unresolved if self._page_from_row_id(r.row_id) == page_no]
                        if not rows:
                            continue
                        pix = doc[page_no].get_pixmap(dpi=170, alpha=False)
                        image = pix.tobytes("png")
                        payload = [
                            {
                                "id": row.row_id,
                                "number": row.number,
                                "parameter": row.parameter,
                                "requirement": row.requirement,
                                "parent": row.parent_context,
                                "candidates": row.candidate_values[:8],
                            }
                            for row in rows
                        ]
                        response = ask_vision_json_schema(client, image, self._prompt(context, payload, kb_context), self._schema())
                        results = response.get("results", []) if isinstance(response, dict) else []
                        by_id = {row.row_id: row for row in rows}
                        for item in results or []:
                            row = by_id.get(str(item.get("id", "")))
                            if not row:
                                continue
                            try:
                                confidence = float(item.get("confidence", 0) or 0)
                            except (TypeError, ValueError):
                                confidence = 0.0
                            value = clean(item.get("value", ""))
                            if confidence < 0.55 or not value:
                                continue
                            row.proposed_value = strip_marks(value)
                            row.source = "AI_VISION"
                            row.mark = "*" if "*" in clean(row.requirement) else "**"
                            row.confidence = confidence
                            row.reason = clean(item.get("reason", ""))
                            row.evidence.append("GigaChat Vision: страница PDF")
                finally:
                    doc.close()
        except Exception as exc:
            print(f"[PDF][AI] пропуск: {exc}")

    @staticmethod
    def _page_from_row_id(row_id: str) -> int:
        match = re.match(r"p(\d+)_", row_id)
        if not match:
            raise ValueError(f"Некорректный PDF row_id: {row_id}")
        return int(match.group(1))

    @staticmethod
    def _prompt(context: ProductContext, payload: list[dict], kb_context: str) -> str:
        return (
            "Ты заполняешь только незаполненные ячейки колонки 'Предлагаемое участником конкурса' в PDF. "
            "Никогда не переноси число или текст из соседней строки. Для каждой строки используй только её id, "
            "номер, название параметра, требование, родительскую обмотку и изображение таблицы. "
            "Приоритет: точный факт из КБ -> явное требование -> очевидный контекст документа. "
            "Если достоверного значения нет, верни пустое value. Не пиши пояснения вместо значения. "
            f"Модель: {context.model}; напряжение: {context.voltage}.\n\n"
            f"КОНТЕКСТ БД:\n{kb_context}\n\n"
            f"СТРОКИ:\n{payload}"
        )

    @staticmethod
    def _schema() -> dict:
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
                            "confidence": {"type": "number"},
                            "reason": {"type": "string"},
                        },
                        "required": ["id", "value", "confidence", "reason"],
                    },
                }
            },
            "required": ["results"],
        }
