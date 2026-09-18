"""Совместимый вход для старого кода. Реальная логика находится в services.tender_engine."""
from __future__ import annotations

from pathlib import Path
from typing import Any

from docx import Document

from services.tender_engine import TenderFillingEngine
from services.tender_engine.docx_reader import DocxTableReader, snapshot_context
from services.tender_engine.context import ProductDetector
from services.tender_engine.data_source import get_data_tenders_knowledge


def extract_doc_items(doc: Document) -> tuple[list[dict], dict[str, Any]]:
    snapshot = DocxTableReader().read(doc)
    return [row.to_dict() for row in snapshot.rows], snapshot_context(snapshot)


def _determine_type(session, items: list[dict], document_context: dict[str, Any] | None = None):
    rows = []
    for item in items:
        from services.tender_engine.schema import TenderRow
        rows.append(TenderRow(
            row_id=str(item.get("id", "")),
            table_index=int(item.get("table", 0)),
            row_index=int(item.get("row", 0)),
            number=str(item.get("num", "")),
            parameter=str(item.get("param_name", "")),
            requirement=str(item.get("required_val", "")),
            current_value=str(item.get("current_value", "")),
            target_cell_index=item.get("target_cell_index"),
            parent_context=str(item.get("parent_context", "")),
            section=str(item.get("section", "")),
            field_key=str(item.get("field_key", "")),
        ))
    context = ProductDetector().detect(rows, (document_context or {}).get("paragraphs", []))
    from models.tr_type import TrType
    if not context.model:
        raise ValueError("Не удалось определить модель изделия.")
    return TrType(name=context.model), context.voltage


def process_docx_requirements(docx_path: str, output_path: str, session) -> dict:
    engine = TenderFillingEngine(overwrite_existing=bool(int(__import__("os").getenv("TENDER_OVERWRITE_EXISTING", "0"))))
    return engine.fill_docx(docx_path, output_path, session)


def capture_docx_file(path: str | Path, session, specialist_name: str | None = None) -> dict[str, int]:
    return TenderFillingEngine(ai_enabled=False).capture_final_docx(path, session, specialist_name)


__all__ = [
    "extract_doc_items",
    "_determine_type",
    "process_docx_requirements",
    "capture_docx_file",
    "get_data_tenders_knowledge",
]
