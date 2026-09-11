"""CLI для роста базы знаний без изменения кода/схемы.

Примеры:
    python tools/add_knowledge.py --key nominal_current --value "500-1000" \
        --tr-type ТРГ --voltage 110 --source passport

    python tools/add_knowledge.py --json knowledge_seed.json
"""

import argparse
import sys
from pathlib import Path

# Позволяет запускать файл и как `python tools/add_knowledge.py`, и как модуль.
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker

from models.tr_type import Base, TrType
from services.knowledge_base import load_knowledge_json, upsert_knowledge

# БАГ В ИСХОДНИКЕ: было `Path(__file__).resolve().parent` — это каталог
# tools/, а не корень проекта, поэтому CLI писал в другую (несуществующую)
# transformers.db, а не в ту, которую читает main.py. Исправлено на parents[1],
# как это уже было корректно сделано в tools/promote_audit.py.
ROOT = Path(__file__).resolve().parents[1]
DB = ROOT / "transformers.db"


def main() -> None:
    parser = argparse.ArgumentParser(description="Добавление/обновление знания в БД тендерного модуля")
    parser.add_argument("--key")
    parser.add_argument("--value")
    parser.add_argument("--category", default="product")
    parser.add_argument("--voltage")
    parser.add_argument("--tr-type", dest="tr_type")
    parser.add_argument("--source", default="manual")
    parser.add_argument("--notes", default="")
    parser.add_argument("--aliases", default="", help="Через запятую: alias1,alias2")
    parser.add_argument("--json", help="Путь к JSON-файлу с пакетным обновлением")
    args = parser.parse_args()

    engine = create_engine(f"sqlite:///{DB}")
    Base.metadata.create_all(engine)
    Session = sessionmaker(bind=engine)

    with Session() as session:
        if args.json:
            count = load_knowledge_json(session, args.json)
            print(f"Импортировано записей: {count}")
            return

        if not args.key or args.value is None:
            parser.error("Для ручного добавления нужны --key и --value")

        tr_id = None
        if args.tr_type:
            tr = session.query(TrType).filter_by(name=args.tr_type).first()
            if not tr:
                tr = TrType(name=args.tr_type)
                session.add(tr)
                session.flush()
            tr_id = tr.id

        aliases = [x.strip() for x in args.aliases.split(",") if x.strip()]
        row = upsert_knowledge(
            session,
            category=args.category,
            key=args.key,
            value=args.value,
            aliases=aliases,
            tr_type_id=tr_id,
            voltage=args.voltage,
            source=args.source,
            notes=args.notes,
        )
        print(f"Сохранено знание id={row.id}: {row.category}/{row.key}={row.value} (БД: {DB})")


if __name__ == "__main__":
    main()
