"""Совместимый API шаблонной БЗ. Реализация перенесена в TenderTemplateStore."""
from __future__ import annotations

from pathlib import Path

from services.tender_engine.template_store import TemplateStore
from services.tender_engine.docx_reader import DocxTableReader
from services.tender_engine.context import ProductDetector
from services.tender_engine.schema import TenderRow
from services.tender_engine.text import clean, norm, strip_marks
from services.tender_engine.rules import field_key as _param_key


def find_best_template(session, items, tr_type_id=None, voltage=None, min_coverage=0.55, min_anchors=4, *, filename=None, model=None):
    rows = [_row_from_dict(x) for x in items]
    template, mapping, coverage = TemplateStore().find(session, rows, filename=filename or "", model=model)
    return template, coverage, mapping


def apply_template(session, items, template, matches):
    filled = 0
    for item in items:
        match = matches.get(str(item.get("id")))
        if not match:
            continue
        value = strip_marks(match.proposed_value)
        if not value or clean(item.get("current_value")):
            continue
        item["template_value"] = value
        item["template_source"] = "DB_TEMPLATE"
        item["algorithm_value"] = value
        item["algorithm_source"] = "DB_TEMPLATE"
        filled += 1
    return filled


def save_document_to_knowledge(session, items, *, filename="", tr_type=None, voltage=None, specialist_name=None):
    rows = [_row_from_dict(x) for x in items]
    for row, item in zip(rows, items):
        row.proposed_value = clean(item.get("final_value") or item.get("current_value"))
        row.mark = clean(item.get("final_mark") or item.get("mark"))
    return TemplateStore().save(session, rows, filename=filename, model=tr_type, voltage=voltage, specialist_name=specialist_name)


def apply_voltage_profile(*args, **kwargs) -> int:
    return 0


def capture_docx_file(session, path: str | Path, specialist_name: str | None = None):
    doc = __import__("docx").Document(str(path))
    snapshot = DocxTableReader().read(doc)
    if not snapshot.answer_columns:
        return {"saved_rows": 0, "template_created": 0, "template_id": 0}
    context = ProductDetector().detect(snapshot.rows, snapshot.paragraphs)
    for row in snapshot.rows:
        row.proposed_value = clean(row.current_value)
        row.mark = ""
    return TemplateStore().save(session, snapshot.rows, filename=Path(path).name, model=context.model, voltage=context.voltage, specialist_name=specialist_name)


def _row_from_dict(item):
    return TenderRow(
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
        field_key=str(item.get("field_key") or _param_key(str(item.get("param_name", "")))),
    )
