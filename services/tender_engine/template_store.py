from __future__ import annotations

import re
from datetime import datetime
from difflib import SequenceMatcher
from pathlib import Path
from typing import Iterable

from sqlalchemy.orm import Session

from models.tr_type import Tender, TenderParameter
from .schema import TenderRow
from .text import clean, norm, strip_marks


class TemplateStore:
    """Работа с ранее исправленными инженером таблицами."""

    def find(self, session: Session, rows: list[TenderRow], *, filename: str, model: str | None) -> tuple[Tender | None, dict[str, TenderParameter], float]:
        candidates = self._candidate_tenders(session, filename, model)
        if not candidates:
            return None, {}, 0.0
        best: tuple[Tender | None, dict[str, TenderParameter], float, float] = (None, {}, 0.0, 0.0)
        for tender in candidates:
            stored = session.query(TenderParameter).filter(
                TenderParameter.tender_id == tender.id,
                TenderParameter.proposed_value.isnot(None),
                TenderParameter.proposed_value != "",
            ).all()
            mapping: dict[str, TenderParameter] = {}
            score = 0.0
            strong = 0
            for row in rows:
                match = self._match_row(row, stored)
                if match:
                    mapping[row.row_id] = match
                    score += min(self._row_score(row, match), 1.0)
                    strong += 1
            coverage = score / max(1, len([r for r in rows if r.parameter]))
            filename_score = self._filename_score(filename, tender.filename)
            model_score = self._model_score(model, tender.object_name)
            rank = coverage + 0.50 * filename_score + 0.35 * model_score + min(strong, 10) * 0.01
            if self._accepted(coverage, strong, filename_score, model_score) and rank > best[3]:
                best = (tender, mapping, coverage, rank)
        return best[0], best[1], best[2]

    def save(self, session: Session, rows: Iterable[TenderRow], *, filename: str, model: str | None, voltage: str | None, specialist_name: str | None = None) -> dict[str, int]:
        rows = list(rows)
        trusted = [r for r in rows if r.proposed_value and "**" not in r.mark and "**" not in r.proposed_value]
        if not trusted:
            return {"saved_rows": 0, "template_created": 0, "template_id": 0}
        tender, _, _ = self.find(session, trusted, filename=filename, model=model)
        created = 0
        if tender is None:
            tender = Tender(
                filename=filename,
                object_name=model or "",
                quantity="",
                delivery_date="",
                delivery_address="",
                created_at=datetime.now().isoformat(timespec="seconds"),
                specialist_name=clean(specialist_name),
            )
            session.add(tender)
            session.flush()
            created = 1
        else:
            tender.filename = filename or tender.filename
            tender.object_name = model or tender.object_name
            if specialist_name:
                tender.specialist_name = clean(specialist_name)

        old = session.query(TenderParameter).filter(TenderParameter.tender_id == tender.id).all()
        index = {self._key_from_db(x): x for x in old}
        present = set()
        saved = 0
        for row in rows:
            if not row.parameter:
                continue
            value = strip_marks(row.proposed_value)
            key = self._key(row)
            present.add(key)
            if not value:
                if key in index:
                    session.delete(index[key])
                    index.pop(key, None)
                continue
            entry = index.get(key)
            if entry is None:
                entry = TenderParameter(tender_id=tender.id)
                session.add(entry)
                index[key] = entry
            entry.section_number = row.number
            entry.parameter_name = row.parameter
            entry.required_value = strip_marks(row.requirement)
            entry.proposed_value = value
            entry.row_index = row.row_index
            entry.table_index = row.table_index
            entry.target_cell_index = row.target_cell_index
            entry.field_key = row.field_key
            entry.parent_context = row.parent_context
            entry.is_mandatory = "*" in row.requirement
            saved += 1
        session.commit()
        return {"saved_rows": saved, "template_created": created, "template_id": int(tender.id)}

    @staticmethod
    def _candidate_tenders(session: Session, filename: str, model: str | None) -> list[Tender]:
        result = session.query(Tender).order_by(Tender.id.desc()).all()
        target_file = TemplateStore._canon_filename(filename)
        target_model = norm(model)
        filtered = []
        for tender in result:
            file_score = TemplateStore._filename_score(filename, tender.filename)
            model_score = TemplateStore._model_score(model, tender.object_name)
            if file_score >= 0.70 or model_score >= 0.80 or not tender.object_name:
                filtered.append(tender)
        return filtered

    def _match_row(self, row: TenderRow, stored: list[TenderParameter]) -> TenderParameter | None:
        exact = [x for x in stored if self._same_anchor(row, x) and self._is_value(x.proposed_value)]
        if exact:
            return max(exact, key=lambda x: self._row_score(row, x))
        # Fallback допускается только при совпадающем canonical field и parent.
        candidates = [x for x in stored if self._same_semantics(row, x) and self._is_value(x.proposed_value)]
        scored = [(self._row_score(row, x), x) for x in candidates]
        scored = [x for x in scored if x[0] >= 0.90]
        return max(scored, default=(0.0, None), key=lambda x: x[0])[1]

    @staticmethod
    def _same_anchor(row: TenderRow, db: TenderParameter) -> bool:
        num_a = clean(row.number).rstrip(".")
        num_b = clean(db.section_number).rstrip(".")
        if num_a or num_b:
            return num_a == num_b
        return TemplateStore._same_semantics(row, db)

    @staticmethod
    def _same_semantics(row: TenderRow, db: TenderParameter) -> bool:
        return (
            norm(row.field_key) == norm(db.field_key)
            and norm(row.parent_context) == norm(db.parent_context)
        )

    def _row_score(self, row: TenderRow, db: TenderParameter) -> float:
        score = 0.0
        if norm(row.field_key) == norm(db.field_key):
            score += 1.0
        else:
            score += SequenceMatcher(None, norm(row.parameter), norm(db.parameter_name)).ratio()
        if clean(row.number).rstrip(".") == clean(db.section_number).rstrip("."):
            score += 0.20
        if norm(row.parent_context) == norm(db.parent_context):
            score += 0.15
        if self._req_equal(row.requirement, db.required_value):
            score += 0.12
        return min(score, 1.40)

    @staticmethod
    def _req_equal(left: str, right: str) -> bool:
        a, b = norm(strip_marks(left)), norm(strip_marks(right))
        if a == b:
            return True
        if not a or not b or a == "*" or b == "*":
            return a == b
        return SequenceMatcher(None, a, b).ratio() >= 0.90

    @staticmethod
    def _key(row: TenderRow) -> tuple[str, str, str, str]:
        return (clean(row.number).rstrip("."), norm(row.field_key), norm(row.parent_context), norm(strip_marks(row.requirement)))

    @staticmethod
    def _key_from_db(row: TenderParameter) -> tuple[str, str, str, str]:
        return (clean(row.section_number).rstrip("."), norm(row.field_key or ""), norm(row.parent_context or ""), norm(strip_marks(row.required_value or "")))

    @staticmethod
    def _is_value(value: object) -> bool:
        text = strip_marks(value)
        return bool(text and not re.fullmatch(r"\*+", text))

    @staticmethod
    def _canon_filename(value: str | None) -> str:
        text = clean(value)
        text = re.sub(r"^заполненный[_\s-]*", "", text, flags=re.I)
        text = re.sub(r"[_\s-]*для_редактирования(?=\.[^.]+$)", "", text, flags=re.I)
        text = re.sub(r"\.(docx|doc|pdf)$", "", text, flags=re.I)
        return norm(text)

    @classmethod
    def _filename_score(cls, left: str | None, right: str | None) -> float:
        a, b = cls._canon_filename(left), cls._canon_filename(right)
        if not a or not b:
            return 0.0
        if a == b or a in b or b in a:
            return 1.0
        return SequenceMatcher(None, a, b).ratio()

    @staticmethod
    def _model_score(left: str | None, right: str | None) -> float:
        a, b = norm(left), norm(right)
        if not a or not b:
            return 0.0
        if a == b:
            return 1.0
        if a in b or b in a:
            return 0.90
        return 0.0

    @staticmethod
    def _accepted(coverage: float, strong: int, file_score: float, model_score: float) -> bool:
        if strong == 0:
            return False
        return (
            (file_score >= 0.92 and strong >= 1)
            or (model_score >= 0.90 and coverage >= 0.02)
            or (strong >= 4 and coverage >= 0.55)
        )
