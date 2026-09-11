"""Расширяемая база знаний тендерного модуля.

Таблица ``knowledge_entries`` хранит произвольные пары ключ/значение
(константы: изготовитель, марка, номинальный ток и т.д.) с привязкой к
типу трансформатора и классу напряжения. Новый параметр добавляется без
изменения схемы и без изменения кода — только записью в БД (вручную,
через ``tools/add_knowledge.py`` или пакетно из JSON).

Таблица ``field_rules`` сопоставляет реальные формулировки строк ТЗ
("Номинальное напряжение обмотки, кВ") с каноническим ключом БЗ
("nominal_voltage"), так что документы с разной формулировкой одного и
того же параметра распознаются одинаково.
"""

import json
import re
from pathlib import Path
from typing import Any

from sqlalchemy.orm import Session

from models.tr_type import (
    AccuracyClass,
    Climat,
    FieldRule,
    IsolColor,
    IsolType,
    KnowledgeEntry,
    TrType,
    TrTypeRule,
    VoltageClass,
)


def norm(value: str | None) -> str:
    """Нормализует строку для сравнения: нижний регистр, ё->е, схлопнутые пробелы."""
    value = (value or "").strip().lower().replace("ё", "е")
    return re.sub(r"\s+", " ", value)


def canonicalize_field(session: Session, param_name: str) -> str | None:
    """Находит канонический ключ БЗ даже при умеренных OCR-ошибках."""
    from difflib import SequenceMatcher

    name = norm(param_name)
    if not name:
        return None
    rules = session.query(FieldRule).filter_by(active=True).order_by(FieldRule.priority.desc()).all()

    # 1. Обычное быстрое совпадение.
    for rule in rules:
        candidates = [rule.canonical_name, *(rule.aliases or [])]
        for candidate in candidates:
            cand = norm(candidate)
            if cand and re.search(r"(?<![\w])" + re.escape(cand) + r"(?![\w])", name, flags=re.UNICODE):
                return rule.db_key or rule.canonical_name

    # 2. Fuzzy-поиск нужен прежде всего для сканов: Tesseract может превращать
    #    «наибольшее» -> «haudonbwee», «вид» -> «Bud» и т.п.
    #    Сравниваем не всю длинную строку, а лучшие близкие окна.
    words = name.split()
    best = (0.0, None)
    for rule in rules:
        candidates = [rule.canonical_name, *(rule.aliases or [])]
        for candidate in candidates:
            cand = norm(candidate)
            if not cand:
                continue
            key = rule.db_key or rule.canonical_name
            # Короткие семантические поля нельзя выводить fuzzy-поиском из
            # длинного предложения: например, «производителем» не означает
            # поле «Изготовитель». DOCX-проход дополнительно обрабатывает эти
            # поля явными правилами.
            if key in {"manufacturer", "brand"}:
                continue
            cand_words = cand.split()
            if len(words) >= len(cand_words):
                window_len = len(cand_words)
                windows = (
                    " ".join(words[i:i + window_len])
                    for i in range(max(1, len(words) - window_len + 1))
                )
            else:
                windows = (name,)
            score = max(SequenceMatcher(None, cand, w).ratio() for w in windows)
            # Длинные инженерные названия допускают чуть более низкий порог,
            # но короткие слова требуют строгого совпадения.
            threshold = 0.62 if len(cand) >= 18 else 0.78
            if score >= threshold and score > best[0]:
                best = (score, rule.db_key or rule.canonical_name)
    return best[1]


def _score_entry(entry: KnowledgeEntry, tr_type_id: int | None, voltage: str | None) -> int:
    """Оценивает применимость записи БЗ к текущему документу.

    Возвращает -1, если запись несовместима (привязана к другому типу
    трансформатора или другому напряжению), иначе — числовой приоритет:
    более специфичные записи (привязанные к типу/напряжению) побеждают
    общие ("для любого типа").
    """
    score = 0
    if entry.tr_type_id is None:
        score += 10
    elif entry.tr_type_id == tr_type_id:
        score += 50
    else:
        return -1
    if entry.voltage:
        if voltage and str(entry.voltage) == str(voltage):
            score += 40
        else:
            return -1
    else:
        score += 5
    return score


def find_knowledge(
    session: Session, key: str, tr_type_id: int | None = None, voltage: str | None = None
) -> list[KnowledgeEntry]:
    key_n = norm(key)
    rows = session.query(KnowledgeEntry).filter_by(active=True).all()
    scored = []
    for row in rows:
        if norm(row.key) != key_n and not any(norm(a) == key_n for a in (row.aliases or [])):
            continue
        score = _score_entry(row, tr_type_id, voltage)
        if score >= 0:
            scored.append((score, row))
    return [x[1] for x in sorted(scored, key=lambda x: x[0], reverse=True)]


def get_best_value(
    session: Session, key: str, tr_type_id: int | None = None, voltage: str | None = None
) -> tuple[str | None, str | None]:
    rows = find_knowledge(session, key, tr_type_id, voltage)
    if not rows:
        return None, None
    return rows[0].value, rows[0].source


def _values_match(required: str, allowed: str, key: str) -> bool:
    """Безопасное сравнение значения требования со справочником."""
    req = norm(required).replace(",", ".")
    item = norm(allowed).replace(",", ".")
    if req == item:
        return True
    if key == "climate":
        # В ТЗ климат часто пишут как УХЛ, а в справочнике — УХЛ1.
        return (req.endswith("1") and req[:-1] == item) or (
            item.endswith("1") and item[:-1] == req
        )
    try:
        return float(req) == float(item)
    except ValueError:
        return False


def resolve_db_field(
    session: Session,
    key: str,
    required_val: str,
    tr_type_id: int | None = None,
    voltage: str | None = None,
) -> tuple[str | None, str | None]:
    """Возвращает константу из расширяемой БЗ или из исторических справочников БД."""
    value, source = get_best_value(session, key, tr_type_id, voltage)
    if value is not None:
        return value, source or "knowledge_db"

    if not tr_type_id:
        return None, None
    rule = session.query(TrTypeRule).filter_by(tr_type_id=tr_type_id).first()
    if not rule:
        return None, None

    def names(model, ids):
        column = getattr(model, "name", None)
        if column is None:
            column = getattr(model, "value", None)
        if column is None:
            return []
        return [x[0] for x in session.query(column).filter(model.id.in_(ids or [])).all()]

    allowed = None
    if key == "nominal_voltage":
        allowed = names(VoltageClass, rule.voltage_classes)
    elif key == "climate":
        allowed = names(Climat, rule.climats)
    elif key in {"internal_insulation", "external_insulation"}:
        allowed = names(IsolType, rule.isol_types)
    elif key == "external_insulation_color":
        allowed = names(IsolColor, rule.isol_colors)
    elif key == "accuracy_class":
        allowed = names(AccuracyClass, rule.accuracy_classes)

    if not allowed:
        return None, None
    req = norm(required_val)
    if req in {"", "*"}:
        # Для одиночного допустимого значения его можно вернуть как
        # детерминированную константу. Несколько вариантов нельзя склеивать:
        # это не ответ на конкретное требование.
        if len(allowed) == 1:
            return allowed[0], "legacy_db"
        return None, None
    for item in allowed:
        if _values_match(req, item, key):
            return item, "legacy_db"
    return None, None


def build_knowledge_context(session: Session, tr_type_id: int | None, voltage: str | None) -> dict[str, Any]:
    """Собирает применимые записи БЗ для передачи нейросети как источника констант."""
    rows = session.query(KnowledgeEntry).filter_by(active=True).all()
    context = []
    for row in rows:
        score = _score_entry(row, tr_type_id, voltage)
        if score < 0:
            continue
        context.append(
            {
                "category": row.category,
                "key": row.key,
                "value": row.value,
                "aliases": row.aliases or [],
                "source": row.source,
                "notes": row.notes or "",
                "scope_score": score,
            }
        )
    context.sort(key=lambda x: x["scope_score"], reverse=True)
    return {"entries": context}


def upsert_knowledge(
    session: Session,
    *,
    category: str,
    key: str,
    value: str,
    aliases: list[str] | None = None,
    tr_type_id: int | None = None,
    voltage: str | None = None,
    source: str = "manual",
    notes: str = "",
) -> KnowledgeEntry:
    """Добавляет или обновляет запись БЗ. Это единственная точка роста базы знаний."""
    existing = (
        session.query(KnowledgeEntry)
        .filter_by(category=category, key=key, tr_type_id=tr_type_id, voltage=voltage)
        .first()
    )
    if existing:
        existing.value = value
        existing.aliases = aliases or existing.aliases or []
        existing.source = source
        existing.notes = notes
        existing.active = True
        row = existing
    else:
        row = KnowledgeEntry(
            category=category,
            key=key,
            value=value,
            aliases=aliases or [],
            tr_type_id=tr_type_id,
            voltage=voltage,
            source=source,
            notes=notes,
            active=True,
        )
        session.add(row)
    session.commit()
    return row


def load_knowledge_json(session: Session, path: str) -> int:
    """Пакетно загружает знания и правила сопоставления полей из JSON-файла."""
    data = json.loads(Path(path).read_text(encoding="utf-8"))
    count = 0
    types = {norm(x.name): x for x in session.query(TrType).all()}
    for item in data.get("knowledge", []):
        tr_type_id = item.get("tr_type_id")
        tr_type_name = item.get("tr_type")
        if tr_type_id is None and tr_type_name:
            found = types.get(norm(tr_type_name))
            tr_type_id = found.id if found else None
        upsert_knowledge(
            session,
            category=item.get("category", "product"),
            key=item["key"],
            value=str(item["value"]),
            aliases=item.get("aliases", []),
            tr_type_id=tr_type_id,
            voltage=None if item.get("voltage") is None else str(item.get("voltage")),
            source=item.get("source", "manual"),
            notes=item.get("notes", ""),
        )
        count += 1
    for rule in data.get("field_rules", []):
        existing = session.query(FieldRule).filter_by(canonical_name=rule["canonical_name"]).first()
        if existing:
            existing.aliases = rule.get("aliases", existing.aliases or [])
            existing.data_type = rule.get("data_type", existing.data_type)
            existing.db_key = rule.get("db_key", existing.db_key)
            existing.section = rule.get("section", existing.section)
            existing.priority = rule.get("priority", existing.priority)
            existing.active = True
        else:
            session.add(
                FieldRule(
                    canonical_name=rule["canonical_name"],
                    aliases=rule.get("aliases", []),
                    data_type=rule.get("data_type", "string"),
                    db_key=rule.get("db_key"),
                    section=rule.get("section"),
                    priority=rule.get("priority", 100),
                    active=True,
                )
            )
        count += 1
    session.commit()
    return count
