from pathlib import Path
import argparse
import os
import sys

from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker

from bootstrap import bootstrap, ROOT
from services.search_docx_AI import process_docx_requirements
from services.search_pdf_AI import process_pdf_requirements
from services.document_io import convert_docx_output, normalize_for_processing, detect_format

engine = create_engine(f"sqlite:///{ROOT / 'transformers.db'}")
Session = sessionmaker(bind=engine)


def process_document(
    input_path: str,
    output_path: str,
    session,
    *,
    use_ocr: bool = True
) -> dict:
    """Единая точка входа для DOCX/PDF/DOC/DOCM/ODT/RTF."""

    input_file = Path(input_path).expanduser().resolve()
    output_file = Path(output_path).expanduser().resolve()

    original_ext = detect_format(input_file)
    work_root = ROOT / ".tmp_normalized"

    working_input, _, temp_root = normalize_for_processing(
        input_file,
        work_root
    )

    try:
        # ==========================================================
        # PDF
        # ==========================================================
        if original_ext == ".pdf":
            result = process_pdf_requirements(
                pdf_path=str(working_input),
                output_path=str(output_file),
                session=session,
                use_ocr=use_ocr,
            )

            return result or {}

        # ==========================================================
        # DOCX / DOC / DOCM / ODT / RTF
        # ==========================================================

        # Для любого старого формата результат ВСЕГДА DOCX.
        #
        # Это важно для дальнейшего:
        #   1. открытия инженером;
        #   2. QFileSystemWatcher;
        #   3. python-docx;
        #   4. обучения базы знаний.
        if original_ext == ".docx":
            working_output = output_file
        else:
            working_output = output_file.with_suffix(".docx")

        result = process_docx_requirements(
            docx_path=str(working_input),
            output_path=str(working_output),
            session=session,
        )

        # НЕ конвертируем DOCX обратно в DOC.
        #
        # Старый вариант:
        # convert_docx_output(...)
        #
        # здесь больше не нужен.

        return {
            "output": str(working_output)
        }

    finally:
        # Удаляем только временный нормализованный файл.
        if (
            temp_root is not None
            and working_input.exists()
            and working_input.parent == work_root
        ):
            try:
                working_input.unlink()
            except OSError:
                pass

def main() -> int:
    parser = argparse.ArgumentParser(description="Автозаполнение тендерных DOCX/DOC/PDF/ODT/RTF")
    parser.add_argument("input", nargs="?", help="Входной документ")
    parser.add_argument("output", nargs="?", help="Выходной документ")
    parser.add_argument("--no-ocr", action="store_true", help="Не использовать OCR для сканированных PDF")
    args = parser.parse_args()

    input_doc = args.input or os.getenv("TENDER_INPUT", '30.11.2021 ТехТреб ТРГ-110 поз.20 Поляково.doc')
    if not input_doc:
        parser.error("Укажите входной файл аргументом или через TENDER_INPUT.")

    input_path = Path(input_doc).expanduser().resolve()
    if not input_path.exists():
        print(f"[ОШИБКА] Файл не найден: {input_path}", file=sys.stderr)
        return 1

    output_arg = args.output or os.getenv("TENDER_OUTPUT")
    if output_arg:
        output_path = Path(output_arg).expanduser().resolve()
    else:
        output_path = input_path.with_name("Заполненный_" + input_path.name)
    if output_path == input_path:
        parser.error("Выходной файл не должен совпадать с входным.")
    output_path.parent.mkdir(parents=True, exist_ok=True)

    try:
        bootstrap()
        with Session() as session:
            result = process_document(
                str(input_path), str(output_path), session, use_ocr=not args.no_ocr
            )
        if result:
            print(result)
        return 0
    except (FileNotFoundError, RuntimeError, ValueError) as exc:
        print(f"[ОШИБКА] {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
