"""Единый ввод/вывод документов тендеров.

DOCX обрабатывается напрямую. Старые DOC, а также DOCM/ODT/RTF,
нормализуются в DOCX через Microsoft Word (если доступен), либо LibreOffice.
После заполнения результат при необходимости возвращается в исходный формат.
PDF обрабатывается отдельным процессором без конвертации.
"""
from __future__ import annotations

import shutil
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import Optional

DIRECT_DOCX = {".docx"}
PDF_FORMATS = {".pdf"}
OFFICE_TO_DOCX = {".doc", ".docm", ".odt", ".rtf"}
SUPPORTED = DIRECT_DOCX | PDF_FORMATS | OFFICE_TO_DOCX


def detect_format(path: str | Path) -> str:
    suffix = Path(path).suffix.lower()
    if suffix not in SUPPORTED:
        raise ValueError(
            f"Неподдерживаемый формат {suffix or '<без расширения>'}. "
            f"Поддерживаются: {', '.join(sorted(SUPPORTED))}."
        )
    return suffix


def _find_soffice() -> Optional[str]:
    return shutil.which("soffice") or shutil.which("libreoffice")


def _win32_word_available() -> bool:
    if not sys.platform.startswith("win"):
        return False
    try:
        import win32com.client  # type: ignore
        return True
    except Exception:
        return False


def _convert_with_word(source: Path, out_dir: Path, target_ext: str) -> Optional[Path]:
    """Конвертация через установленный Microsoft Word.

    Это основной путь для старого .doc в Windows: Word корректно понимает
    бинарный формат Word 97-2003, чего python-docx не умеет.
    """
    if not _win32_word_available():
        return None

    import win32com.client  # type: ignore

    # Word FileFormat constants: DOC=0, DOCX=16, DOCM=13, ODT=23, RTF=6.
    fmt_by_ext = {".docx": 16, ".doc": 0, ".docm": 13, ".odt": 23, ".rtf": 6}
    file_format = fmt_by_ext[target_ext]

    out_dir.mkdir(parents=True, exist_ok=True)
    destination = out_dir / f"{source.stem}{target_ext}"
    word = None
    document = None
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        # Open ReadOnly=True, ConfirmConversions=False.
        document = word.Documents.Open(
            str(source),
            ConfirmConversions=False,
            ReadOnly=True,
            AddToRecentFiles=False,
        )
        document.SaveAs2(str(destination), FileFormat=file_format, AddToRecentFiles=False)
        return destination if destination.exists() else None
    finally:
        if document is not None:
            try:
                document.Close(False)
            except Exception:
                pass
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass


def _convert_with_libreoffice(source: Path, out_dir: Path, target_ext: str) -> Optional[Path]:
    soffice = _find_soffice()
    if not soffice:
        return None

    out_dir.mkdir(parents=True, exist_ok=True)
    proc = subprocess.run(
        [
            soffice,
            "--headless",
            "--convert-to",
            target_ext.lstrip("."),
            "--outdir",
            str(out_dir),
            str(source),
        ],
        capture_output=True,
        text=True,
        timeout=180,
    )
    destination = out_dir / f"{source.stem}{target_ext}"
    if proc.returncode == 0 and destination.exists():
        return destination
    return None


def normalize_for_processing(
    input_path: str | Path,
    work_root: str | Path | None = None,
) -> tuple[Path, str, Path | None]:
    """Возвращает (рабочий_файл, исходное_расширение, временная_директория)."""
    source = Path(input_path).expanduser().resolve()
    if not source.exists():
        raise FileNotFoundError(source)
    ext = detect_format(source)

    if ext in DIRECT_DOCX or ext in PDF_FORMATS:
        return source, ext, None

    base = Path(work_root).resolve() if work_root else Path(tempfile.mkdtemp(prefix="tender_normalize_"))
    base.mkdir(parents=True, exist_ok=True)

    # На Windows первым используем Word — особенно важно для старого .doc.
    converted = _convert_with_word(source, base, ".docx")
    if converted is None:
        converted = _convert_with_libreoffice(source, base, ".docx")

    if converted is None:
        raise RuntimeError(
            f"Не удалось открыть {source.name}. Для формата {ext} нужен "
            "Microsoft Word (Windows) или LibreOffice. "
            "При наличии Word установите pywin32: pip install pywin32."
        )
    return converted, ext, base


def convert_docx_output(
    docx_path: str | Path,
    requested_output: str | Path,
    original_ext: str,
) -> Path:
    """Сохраняет внутренний DOCX обратно в исходный legacy-формат."""
    output = Path(requested_output).expanduser().resolve()
    if original_ext not in OFFICE_TO_DOCX:
        return output

    temp_dir = Path(tempfile.mkdtemp(prefix="tender_export_"))
    try:
        generated = _convert_with_word(Path(docx_path).resolve(), temp_dir, original_ext)
        if generated is None:
            generated = _convert_with_libreoffice(Path(docx_path).resolve(), temp_dir, original_ext)
        if generated is None:
            raise RuntimeError(
                f"Не удалось сохранить результат в {original_ext}. "
                "Нужен Microsoft Word (Windows) или LibreOffice."
            )
        output.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(generated, output)
        return output
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)
