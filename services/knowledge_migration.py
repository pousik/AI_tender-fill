"""Однократный импорт старой БЗ tender.db в текущую transformers.db."""
from __future__ import annotations

import sqlite3
from pathlib import Path
from sqlalchemy.orm import Session

from models.tr_type import Tender, TenderParameter, ParameterProfile, ProfileParameter


def migrate_legacy_tender_db(session: Session, legacy_path: str | Path) -> dict[str, int]:
    path = Path(legacy_path)
    if not path.exists():
        return {"tenders": 0, "parameters": 0, "profiles": 0, "profile_parameters": 0}

    # Не импортируем повторно, если в новой БД уже есть шаблоны.
    if session.query(Tender).count() > 0:
        return {"tenders": 0, "parameters": 0, "profiles": 0, "profile_parameters": 0}

    conn = sqlite3.connect(str(path))
    conn.row_factory = sqlite3.Row
    result = {"tenders": 0, "parameters": 0, "profiles": 0, "profile_parameters": 0}
    try:
        for r in conn.execute("select * from tenders"):
            obj = Tender(
                filename=r["filename"], object_name=r["object_name"],
                quantity=r["quantity"], delivery_date=r["delivery_date"],
                delivery_address=r["delivery_address"], created_at=r["created_at"],
            )
            session.add(obj); session.flush()
            result["tenders"] += 1
            for p in conn.execute("select * from tender_parameters where tender_id=?", (r["id"],)):
                session.add(TenderParameter(
                    tender_id=obj.id,
                    section_number=p["section_number"], parameter_name=p["parameter_name"],
                    required_value=p["required_value"], proposed_value=p["proposed_value"],
                    row_index=p["row_index"], is_mandatory=bool(p["is_mandatory"]),
                ))
                result["parameters"] += 1

        for r in conn.execute("select * from parameter_profiles"):
            obj = ParameterProfile(name=r["name"], voltage=r["voltage"], description=r["description"])
            session.add(obj); session.flush()
            result["profiles"] += 1
            for p in conn.execute("select * from profile_parameters where profile_id=?", (r["id"],)):
                session.add(ProfileParameter(profile_id=obj.id, parameter_name=p["parameter_name"], value=p["value"]))
                result["profile_parameters"] += 1
        session.commit()
    finally:
        conn.close()
    return result
