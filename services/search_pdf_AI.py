"""Совместимый вход PDF. Основная логика находится в tender_engine.pdf_pipeline."""
from __future__ import annotations

from pathlib import Path

from services.tender_engine.pdf_pipeline import PdfTenderFillingEngine


def process_pdf_requirements(pdf_path: str, output_path: str, session=None, *, use_ocr: bool = True) -> dict:
    del session
    engine = PdfTenderFillingEngine(
        Path(__file__).resolve().parents[1] / "data_tenders",
        ai_enabled=True,
        ocr_enabled=use_ocr,
    )
    return engine.fill(pdf_path, output_path)


def extract_pdf_items(pdf_path: str, use_ocr: bool = True):
    from services.tender_engine.pdf_pipeline import PdfTenderReader
    layouts, pages = PdfTenderReader().read(pdf_path)
    return layouts, "\n".join(pages)
