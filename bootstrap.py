from pathlib import Path
from sqlalchemy import create_engine, inspect, text
from sqlalchemy.orm import sessionmaker
from models.tr_type import Base
from models.fill_data_db import init_database
from services.knowledge_base import load_knowledge_json
from services.knowledge_migration import migrate_legacy_tender_db

ROOT = Path(__file__).resolve().parent
DB_PATH = ROOT / "transformers.db"
engine = create_engine(f"sqlite:///{DB_PATH}")
Session = sessionmaker(bind=engine)


def migrate_tender_parameter_columns():
    """Добавляет структурные поля в старую SQLite БД без удаления данных."""
    inspector = inspect(engine)
    cols = {c["name"] for c in inspector.get_columns("tender_parameters")}
    additions = {
        "table_index": "INTEGER",
        "target_cell_index": "INTEGER",
        "field_key": "VARCHAR(150)",
    }
    with engine.begin() as conn:
        for name, typ in additions.items():
            if name not in cols:
                conn.execute(text(f"ALTER TABLE tender_parameters ADD COLUMN {name} {typ}"))


def bootstrap():
    init_database()
    Base.metadata.create_all(engine)
    migrate_tender_parameter_columns()
    # Импортируем старую шаблонную БЗ один раз. После первого импорта
    # источником является единая transformers.db.
    legacy = ROOT / "tender.db"
    if legacy.exists():
        with Session() as session:
            migrated = migrate_legacy_tender_db(session, legacy)
            if migrated["tenders"]:
                print(f"[KB] Импортирована старая шаблонная БЗ: {migrated}")

    seed = ROOT / "knowledge_seed.json"
    if seed.exists():
        with Session() as session:
            load_knowledge_json(session, str(seed))
