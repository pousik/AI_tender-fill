from pathlib import Path
from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker

try:
    from .tr_type import Base, TrType, VoltageClass, Climat, IsolType, IsolColor, AccuracyClass, TrTypeRule
except ImportError:
    from tr_type import Base, TrType, VoltageClass, Climat, IsolType, IsolColor, AccuracyClass, TrTypeRule

ROOT = Path(__file__).resolve().parents[1]
DB_FILE = ROOT / "transformers.db"


def get_or_create(session, model, **kwargs):
    obj = session.query(model).filter_by(**kwargs).first()
    if obj is None:
        obj = model(**kwargs)
        session.add(obj)
        session.flush()
    return obj


def init_database():
    engine = create_engine(f"sqlite:///{DB_FILE}")
    Base.metadata.create_all(engine)
    Session = sessionmaker(bind=engine)
    with Session() as session:
        volts = [get_or_create(session, VoltageClass, value=x) for x in ["35", "110", "220", "330", "500", "750"]]
        climats = [get_or_create(session, Climat, name=x) for x in ["УХЛ1", "ХЛ1", "У1", "Т1"]]
        colors = [get_or_create(session, IsolColor, name=x) for x in ["Серый", "Коричневый"]]
        isols = [get_or_create(session, IsolType, name=x) for x in ["Фарфор", "Полимер"]]
        accs = [
            get_or_create(session, AccuracyClass, name="0,2", is_measuring=True),
            get_or_create(session, AccuracyClass, name="0,2S", is_measuring=True),
            get_or_create(session, AccuracyClass, name="0,5", is_measuring=True),
            get_or_create(session, AccuracyClass, name="0,5S", is_measuring=True),
            get_or_create(session, AccuracyClass, name="5P", is_measuring=False),
            get_or_create(session, AccuracyClass, name="10P", is_measuring=False),
            get_or_create(session, AccuracyClass, name="5PR", is_measuring=False),
            get_or_create(session, AccuracyClass, name="10PR", is_measuring=False),
            get_or_create(session, AccuracyClass, name="TPY", is_measuring=False),
            get_or_create(session, AccuracyClass, name="TPZ", is_measuring=False),
        ]
        trg = get_or_create(session, TrType, name="ТРГ")
        rule = session.query(TrTypeRule).filter_by(tr_type_id=trg.id).first()
        if rule is None:
            rule = TrTypeRule(tr_type_id=trg.id)
            session.add(rule)
        rule.voltage_classes = [x.id for x in volts]
        rule.climats = [x.id for x in climats]
        rule.isol_types = [x.id for x in isols]
        rule.isol_colors = [x.id for x in colors]
        rule.accuracy_classes = [x.id for x in accs]
        session.commit()
    return DB_FILE


if __name__ == "__main__":
    print(f"БД инициализирована: {init_database()}")
