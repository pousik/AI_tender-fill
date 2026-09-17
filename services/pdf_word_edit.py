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


def set_docx_header_specialist(docx_path: str | Path, specialist_name: str) -> None:
    """Записывает ``Заполнил специалист: ФИО`` в верхний колонтитул DOCX.

    Вызов выполняется до открытия Word и повторно после закрытия, поэтому ФИО
    не зависит от QFileSystemWatcher. Обрабатываются обычный, first-page и
    even-page headers; существующее содержимое не удаляется.
    """
    from docx import Document
    from docx.shared import Pt
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    name = str(specialist_name or "").strip()
    if not name:
        raise ValueError("ФИО специалиста не задано")

    path = Path(docx_path).expanduser().resolve()
    if not path.exists():
        raise FileNotFoundError(path)

    doc = Document(str(path))
    marker = "Заполнил специалист:"
    text = f"{marker} {name}"

    def write_header(header):
        # Обновляем существующую строку специалиста.
        for paragraph in list(header.paragraphs):
            if marker.lower() in (paragraph.text or "").lower():
                paragraph.text = text
                paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
                for run in paragraph.runs:
                    run.font.size = Pt(9)
                return

        # В том числе для header, содержащего таблицы/логотипы, добавляем отдельный
        # обычный paragraph и ставим его первым элементом XML.
        paragraph = header.add_paragraph()
        paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
        run = paragraph.add_run(text)
        run.font.size = Pt(9)

        p_el = paragraph._p
        header._element.remove(p_el)
        children = list(header._element)
        insert_at = 0
        for i, child in enumerate(children):
            if child.tag.endswith('}p') or child.tag.endswith('}tbl'):
                insert_at = i
                break
            insert_at = i + 1
        header._element.insert(insert_at, p_el)

    processed = set()
    for section in doc.sections:
        headers = [section.header]
        if section.different_first_page_header_footer:
            headers.append(section.first_page_header)
        try:
            if doc.settings.odd_and_even_pages_header_footer:
                headers.append(section.even_page_header)
        except Exception:
            pass
        for header in headers:
            key = id(header._element)
            if key in processed:
                continue
            processed.add(key)
            write_header(header)

    doc.save(str(path))

    # Контроль сохранённого файла: читаем его заново.
    check = Document(str(path))
    found = []
    for section in check.sections:
        headers = [section.header, section.first_page_header]
        try:
            headers.append(section.even_page_header)
        except Exception:
            pass
        for header in headers:
            found.extend((p.text or '').strip() for p in header.paragraphs)
    if text not in found:
        raise RuntimeError(
            f"ФИО специалиста не найдено после сохранения DOCX: {text}"
        )

