"""Явно импортирует исправленный инженером DOCX в БЗ шаблонов."""
from __future__ import annotations

import argparse
from pathlib import Path

from bootstrap import Session, bootstrap
from services.tender_engine.pipeline import TenderFillingEngine


def main() -> int:
    parser = argparse.ArgumentParser(description="Импорт исправленного тендерного DOCX в БЗ")
    parser.add_argument("document", help="Исправленный DOCX")
    parser.add_argument("--specialist", default=None, help="ФИО специалиста")
    args = parser.parse_args()
    path = Path(args.document).expanduser().resolve()
    if not path.exists():
        raise FileNotFoundError(path)
    bootstrap()
    with Session() as session:
        result = TenderFillingEngine(ai_enabled=False).capture_final_docx(path, session, args.specialist)
    print(result)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
