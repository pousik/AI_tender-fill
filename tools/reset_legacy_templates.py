"""Безопасная очистка старых шаблонов, созданных до структурной фиксации ячеек.

Перед удалением создается копия transformers.db.
Новые шаблоны с field_key/table_index не затрагиваются.
"""
from pathlib import Path
import shutil
from datetime import datetime

from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker

from bootstrap import ROOT, DB_PATH
from models.tr_type import Tender, TenderParameter


def main():
    db = Path(DB_PATH)
    if not db.exists():
        print("База transformers.db не найдена")
        return

    backup = db.with_name(f"transformers_backup_{datetime.now():%Y%m%d_%H%M%S}.db")
    shutil.copy2(db, backup)
    print(f"[BACKUP] {backup}")

    engine = create_engine(f"sqlite:///{db}")
    Session = sessionmaker(bind=engine)

    with Session() as session:
        removed_tenders = 0
        removed_rows = 0
        for tender in session.query(Tender).all():
            rows = session.query(TenderParameter).filter_by(tender_id=tender.id).all()
            # Старый шаблон: ни одна строка не имеет структурного ключа.
            if rows and not any(r.field_key and r.table_index is not None for r in rows):
                removed_rows += len(rows)
                for row in rows:
                    session.delete(row)
                session.delete(tender)
                removed_tenders += 1

        session.commit()
        print(f"[CLEAN] шаблонов удалено: {removed_tenders}, строк: {removed_rows}")


if __name__ == "__main__":
    main()
