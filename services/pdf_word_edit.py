"""PDF <-> Microsoft Word bridge for editable tender documents."""
from __future__ import annotations

import sys
from pathlib import Path


def _word_available() -> bool:
    if not sys.platform.startswith("win"):
        return False
    try:
        import win32com.client  # type: ignore  # noqa: F401
        return True
    except Exception:
        return False


def pdf_to_editable_docx(pdf_path: str | Path, docx_path: str | Path | None = None) -> Path:
    """Convert a generated PDF to a DOCX that can be edited in Microsoft Word."""
    if not _word_available():
        raise RuntimeError(
            "Редактирование PDF через Word требует Windows, Microsoft Word и pywin32."
        )

    import win32com.client  # type: ignore

    source = Path(pdf_path).expanduser().resolve()
    if not source.exists():
        raise FileNotFoundError(source)

    target = (
        Path(docx_path).expanduser().resolve()
        if docx_path
        else source.with_name(f"{source.stem}_для_редактирования.docx")
    )
    target.parent.mkdir(parents=True, exist_ok=True)

    word = win32com.client.DispatchEx("Word.Application")
    document = None
    try:
        word.Visible = False
        word.DisplayAlerts = 0
        document = word.Documents.Open(
            str(source),
            ConfirmConversions=False,
            ReadOnly=False,
            AddToRecentFiles=False,
        )
        # PDF imported into Word becomes an editable Word document.
        document.SaveAs2(
            str(target),
            FileFormat=16,  # wdFormatXMLDocument (.docx)
            AddToRecentFiles=False,
        )
        if not target.exists():
            raise RuntimeError(f"Word не создал редактируемый DOCX: {target}")
        return target
    finally:
        if document is not None:
            try:
                document.Close(False)
            except Exception:
                pass
        try:
            word.Quit()
        except Exception:
            pass


def docx_to_pdf(docx_path: str | Path, pdf_path: str | Path) -> Path:
    """Convert the edited DOCX back to PDF using Microsoft Word."""
    if not _word_available():
        raise RuntimeError(
            "Сохранение PDF через Word требует Windows, Microsoft Word и pywin32."
        )

    import win32com.client  # type: ignore

    source = Path(docx_path).expanduser().resolve()
    target = Path(pdf_path).expanduser().resolve()
    if not source.exists():
        raise FileNotFoundError(source)
    target.parent.mkdir(parents=True, exist_ok=True)

    word = win32com.client.DispatchEx("Word.Application")
    document = None
    try:
        word.Visible = False
        word.DisplayAlerts = 0
        document = word.Documents.Open(
            str(source),
            ConfirmConversions=False,
            ReadOnly=True,
            AddToRecentFiles=False,
        )
        document.ExportAsFixedFormat(
            OutputFileName=str(target),
            ExportFormat=17,  # wdExportFormatPDF
            OpenAfterExport=False,
            OptimizeFor=0,
            CreateBookmarks=1,
        )
        if not target.exists():
            raise RuntimeError(f"Word не создал PDF: {target}")
        return target
    finally:
        if document is not None:
            try:
                document.Close(False)
            except Exception:
                pass
        try:
            word.Quit()
        except Exception:
            pass


def prepare_pdf_for_word_editing(pdf_path: str | Path, editable_docx_path: str | Path | None = None) -> Path:
    return pdf_to_editable_docx(pdf_path, editable_docx_path)
