"""Переносит подтвержденное вручную AI-значение из аудита в базу знаний.

После этого следующий запуск использует значение уже на первом
(детерминированном) проходе, и "**" для него больше не потребуется.
Работает как с `<файл>.docx.audit.json`, так и с `<файл>.pdf.audit.json` —
оба процессора пишут аудит одного формата.

Пример:
    python tools/promote_audit.py "Заполненный_файл.docx.audit.json" \
        --id item_17 --key sf6_pressure
"""

import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker

from models.tr_type import Base, TrType
from services.knowledge_base import upsert_knowledge

ROOT = Path(__file__).resolve().parents[1]
DB = ROOT / "transformers.db"


def main() -> None:
    parser = argparse.ArgumentParser(description="Подтвердить AI-значение и занести его в БЗ")
    parser.add_argument("audit", help="Путь к <файл>.audit.json")
    parser.add_argument("--id", required=True, help="id строки из audit.json (например item_17 или p2_item_5)")
    parser.add_argument("--key", required=True, help="Канонический ключ БЗ, под которым сохранить значение")
    parser.add_argument("--category", default="product")
    parser.add_argument("--source", default="human_verified")
    parser.add_argument("--notes", default="Подтверждено оператором по аудиту.")
    args = parser.parse_args()

    data = json.loads(Path(args.audit).read_text(encoding="utf-8"))
    row = next((x for x in data.get("rows", []) if x.get("id") == args.id), None)
    if not row or not row.get("final_value"):
        raise SystemExit(f"Строка id='{args.id}' не найдена в {args.audit} или не содержит итогового значения.")

    engine = create_engine(f"sqlite:///{DB}")
    Base.metadata.create_all(engine)
    Session = sessionmaker(bind=engine)

    with Session() as session:
        tr_id = None
        tr_name = data.get("detected_type")
        if tr_name:
            tr = session.query(TrType).filter_by(name=tr_name).first()
            tr_id = tr.id if tr else None
        voltage = data.get("voltage")
        saved = upsert_knowledge(
            session,
            category=args.category,
            key=args.key,
            value=str(row["final_value"]),
            tr_type_id=tr_id,
            voltage=str(voltage) if voltage is not None else None,
            source=args.source,
            notes=args.notes,
        )
        print(f"Знание подтверждено и сохранено: id={saved.id}, {saved.category}/{saved.key}={saved.value}")


if __name__ == "__main__":
    main()
