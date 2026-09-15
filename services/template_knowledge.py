"""Шаблонная база знаний тендеров.

Хранит не только отдельные константы KnowledgeEntry, но и целые ранее
обработанные таблицы. При новом документе сначала ищется похожий шаблон.
Если строка найдена в шаблоне и имеет доверенное значение, GigaChat для нее
не вызывается.

После сохранения документа в БЗ попадают только значения без маркера **.
Таким образом AI-значения, требующие ручной проверки, не загрязняют шаблонную БЗ.
"""
from __future__ import annotations

import re
from datetime import datetime
from difflib import SequenceMatcher
from typing import Any
from pathlib import Path

from sqlalchemy.orm import Session

from models.tr_type import Tender, TenderParameter, ParameterProfile, ProfileParameter, TrType
from services.knowledge_base import canonicalize_field, norm


def _clean(value: Any) -> str:
    return re.sub(r"\s+", " ", str(value or "")).strip()


def _without_star(value: Any) -> str:
    value = _clean(value)
    return re.sub(r"\s*\*+\s*$", "", value).strip()


def _is_real_value(value: Any) -> bool:
    v = _clean(value)
    return bool(v) and not re.fullmatch(r"\*+", v)


def _numbers(value: str) -> list[float]:
    result = []
    for x in re.findall(r"[-+]?\d+(?:[,.]\d+)?", value or ""):
        try:
            result.append(float(x.replace(",", ".")))
        except ValueError:
            pass
    return result


def _req_similar(a: str, b: str) -> bool:
    a, b = norm(_without_star(a)), norm(_without_star(b))
    if not a or not b or a == "*" or b == "*":
        return a == b or not a or not b
    if a == b or a in b or b in a:
        return True
    na, nb = _numbers(a), _numbers(b)
    if na and nb and any(x == y for x in na for y in nb):
        return True
    return SequenceMatcher(None, a, b).ratio() >= 0.88


def _field_key(session: Session, name: str) -> str:
    # canonicalize_field читает правила из БД, поэтому кэшируем результат
    # на время одного SQLAlchemy-сеанса. Это критично при сравнении сотен строк.
    cache = session.info.setdefault("template_field_key_cache", {})
    raw = _clean(name)
    cache_key = norm(raw)
    if cache_key in cache:
        return cache[cache_key]
    value = canonicalize_field(session, raw) or norm(raw)
    cache[cache_key] = value
    return value


def _template_rows(session: Session, tender_id: int) -> list[TenderParameter]:
    return session.query(TenderParameter).filter(
        TenderParameter.tender_id == tender_id,
        TenderParameter.proposed_value.isnot(None),
        TenderParameter.proposed_value != "",
    ).all()


def _row_match(session: Session, item: dict, row: TenderParameter) -> float:
    """Строгое сопоставление строки. Нельзя брать соседнюю строку только
    потому, что название похоже: номер/раздел и требование являются якорями."""
    key_a = _field_key(session, item.get("param_name", ""))
    key_b = _field_key(session, row.parameter_name or "")
    num_a = _clean(item.get("num", ""))
    num_b = _clean(row.section_number or "")

    # Если оба номера содержат нумерацию (1.3, 2.14 и т.п.), различный
    # номер означает другую строку. Это главное средство защиты от
    # подстановки значения ближайшего одноименного поля.
    if num_a and num_b and re.fullmatch(r"\d+(?:\.\d+)*\.?", num_a) and re.fullmatch(r"\d+(?:\.\d+)*\.?", num_b):
        na = num_a.rstrip(".")
        nb = num_b.rstrip(".")
        if na != nb:
            return 0.0

    if key_a == key_b:
        score = 1.0
    else:
        score = SequenceMatcher(None, norm(item.get("param_name", "")), norm(row.parameter_name or "")).ratio()

    if score < 0.82:
        return 0.0

    req_a = item.get("required_val", "")
    req_b = row.required_value or ""
    if _req_similar(req_a, req_b):
        score += 0.12
    elif _clean(req_a) and _clean(req_b) and _clean(req_a) != "*" and _clean(req_b) != "*":
        return 0.0

    if num_a and num_b and num_a.rstrip(".") == num_b.rstrip("."):
        score += 0.18

    if key_a == key_b:
        score += 0.10
    return min(score, 1.35)


def find_best_template(
    session: Session,
    items: list[dict],
    tr_type_id: int | None = None,
    voltage: str | None = None,
    min_coverage: float = 0.55,
    min_anchors: int = 4,
):
    """Быстрый поиск шаблона.

    Убираем N+1 запросов и полный перебор всех строк для каждой строки ТЗ.
    Сначала строим индексы точных ключей, а SequenceMatcher используется
    только для остаточных кандидатов.
    """
    current = [x for x in items if _clean(x.get("param_name")) and not re.fullmatch(r"\*+", _clean(x.get("param_name")))]
    if not current:
        return None, 0.0, {}

    candidates = session.query(Tender).order_by(Tender.id.desc()).all()
    if tr_type_id is not None:
        tr = session.get(TrType, tr_type_id)
        if tr is not None:
            candidates = [t for t in candidates if not t.object_name or norm(t.object_name) == norm(tr.name)]
    if not candidates:
        return None, 0.0, {}

    tender_ids = [t.id for t in candidates]
    all_rows = session.query(TenderParameter).filter(
        TenderParameter.tender_id.in_(tender_ids),
        TenderParameter.proposed_value.isnot(None),
        TenderParameter.proposed_value != "",
    ).all()
    rows_by_tender: dict[int, list[TenderParameter]] = {}
    for row in all_rows:
        rows_by_tender.setdefault(row.tender_id, []).append(row)

    # Поля текущего документа считаем один раз.
    prepared_items = []
    for item in current:
        prepared_items.append({
            "item": item,
            "key": _field_key(session, item.get("param_name", "")),
            "num": _clean(item.get("num", "")),
            "req": _clean(item.get("required_val", "")),
        })

    best = (None, 0.0, {})
    for tender in candidates:
        rows = rows_by_tender.get(tender.id, [])
        if not rows or not any(r.section_number is not None for r in rows):
            continue

        unused = {r.id for r in rows}
        exact_index: dict[tuple[str, str, str], list[TenderParameter]] = {}
        field_index: dict[str, list[TenderParameter]] = {}
        prepared_rows = []
        for row in rows:
            key = _field_key(session, row.parameter_name or "")
            num = _clean(row.section_number or "")
            req = _clean(row.required_value or "")
            exact_index.setdefault((key, num.rstrip("."), norm(_without_star(req))), []).append(row)
            field_index.setdefault(key, []).append(row)
            prepared_rows.append((row, key, num, req))

        matches = {}
        score_sum = 0.0
        strong = 0
        ranked = []
        for prepared in prepared_items:
            item, key_a, num_a_raw, req_a = prepared["item"], prepared["key"], prepared["num"], prepared["req"]
            num_a = num_a_raw.rstrip(".")
            exact = exact_index.get((key_a, num_a, norm(_without_star(req_a))), [])
            if not exact and num_a:
                exact = [r for r in field_index.get(key_a, []) if _clean(r.section_number or "").rstrip(".") == num_a and r.id in unused]
            opts = [(1.28, r) for r in exact if r.id in unused and _is_real_value(r.proposed_value)]
            if not opts:
                # Fuzzy — только строки того же canonical field, а не вся таблица.
                for row in field_index.get(key_a, []):
                    if row.id not in unused or not _is_real_value(row.proposed_value):
                        continue
                    score = _row_match(session, item, row)
                    if score >= 0.90:
                        opts.append((score, row))
            opts.sort(key=lambda x: x[0], reverse=True)
            ranked.append((len(opts), item, opts))

        ranked.sort(key=lambda x: (x[0] if x[0] else 999, -max([v[0] for v in x[2]], default=0)))
        for count, item, opts in ranked:
            if not opts:
                continue
            selected = next(((score, r) for score, r in opts if r.id in unused), None)
            if selected is None:
                continue
            score, row = selected
            matches[str(item.get("id"))] = (row, score)
            unused.remove(row.id)
            score_sum += min(score, 1.0)
            strong += 1

        coverage = score_sum / max(1, len(current))
        if strong >= min_anchors and coverage >= min_coverage and (best[0] is None or coverage > best[1]):
            best = (tender, coverage, matches)

    if best[0] is None:
        return None, 0.0, {}
    return best[0], best[1], best[2]


def apply_template(
    session: Session,
    items: list[dict],
    template: Tender,
    matches: dict[int, tuple[TenderParameter, float]],
) -> int:
    """Заполняет строки из найденного шаблона и помечает их DB_TEMPLATE."""
    filled = 0
    for item in items:
        if not _clean(item.get("param_name")) or re.fullmatch(r"\*+", _clean(item.get("param_name"))):
            continue
        match = matches.get(str(item.get("id")))
        if not match:
            continue
        row, score = match
        value = _without_star(row.proposed_value)
        if not value:
            continue
        if _clean(item.get("current_value")):
            continue
        item["template_value"] = value
        item["template_source"] = "DB_TEMPLATE"
        item["template_tender_id"] = template.id
        item["template_score"] = score
        item["algorithm_value"] = value
        item["algorithm_source"] = "DB_TEMPLATE"
        filled += 1
    return filled


def save_document_to_knowledge(
    session: Session,
    items: list[dict],
    *,
    filename: str = "",
    tr_type: str | None = None,
    voltage: str | None = None,
) -> dict[str, int]:
    """Сохраняет ФИНАЛЬНЫЙ документ инженера. Каждая запись привязана к
    номеру строки, каноническому ключу и физической позиции таблицы."""
    trusted = []
    for item in items:
        value = _clean(item.get("final_value", "")) or _clean(item.get("current_value", ""))
        mark = _clean(item.get("final_mark", "") or item.get("mark", ""))
        if not value or "**" in mark or "**" in value:
            continue
        value = _without_star(value)
        if value:
            trusted.append((item, value))

    if not trusted:
        return {"saved_rows": 0, "template_created": 0, "template_id": 0}

    # Для обучения существующий шаблон ищем только среди документов того же типа.
    template, coverage, _ = _find_existing_for_save(session, trusted, tr_type=tr_type)
    created = 0
    if template is None:
        template = Tender(
            filename=filename, object_name=tr_type or "", quantity="",
            delivery_date="", delivery_address="",
            created_at=datetime.now().isoformat(timespec="seconds"),
        )
        session.add(template)
        session.flush()
        created = 1
    elif filename:
        template.filename = filename
        template.object_name = tr_type or template.object_name

    existing = {}
    for row in _template_rows(session, template.id):
        k = (
            _clean(row.section_number).rstrip("."),
            _field_key(session, row.parameter_name),
            norm(_without_star(row.required_value or "")),
        )
        existing[k] = row

    # ВАЖНО: сохранение документа инженера — это полное состояние строк,
    # а не только список непустых значений. Если инженер удалил значение
    # из уже известного шаблона, старую запись необходимо удалить/обнулить,
    # иначе find_best_template() восстановит её при следующем запуске.
    current_keys = set()
    for item in items:
        name = _clean(item.get("param_name"))
        if not name or re.fullmatch(r"\*+", name):
            continue
        key = (
            _clean(item.get("num", "")).rstrip("."),
            _field_key(session, name),
            norm(_without_star(item.get("required_val", ""))),
        )
        current_keys.add(key)

        final_value = _clean(item.get("final_value", ""))
        current_value = _clean(item.get("current_value", ""))
        effective_value = _without_star(final_value or current_value)
        if not effective_value:
            old_row = existing.get(key)
            if old_row is not None:
                session.delete(old_row)
                existing.pop(key, None)
                print(
                    f"[KB TEMPLATE] удалено значение по строке: "
                    f"{name} / {item.get('num', '')}"
                )

    saved = 0
    for item, value in trusted:
        name = _clean(item.get("param_name"))
        if not name:
            continue
        field_key = _field_key(session, name)
        k = (
            _clean(item.get("num", "")).rstrip("."),
            field_key,
            norm(_without_star(item.get("required_val", ""))),
        )
        row = existing.get(k)
        if row is None:
            row = TenderParameter(
                tender_id=template.id,
                section_number=_clean(item.get("num")),
                parameter_name=name,
                required_value=_clean(item.get("required_val")),
                proposed_value=value,
                row_index=item.get("row"),
                table_index=item.get("table"),
                target_cell_index=item.get("target_cell_index"),
                field_key=field_key,
                is_mandatory="*" in _clean(item.get("required_val")),
            )
            session.add(row)
            existing[k] = row
        else:
            # Обновляем именно эту строку, а не ближайшую одноименную.
            row.proposed_value = value
            row.row_index = item.get("row")
            row.table_index = item.get("table")
            row.target_cell_index = item.get("target_cell_index")
            row.field_key = field_key
        saved += 1

    session.commit()
    return {"saved_rows": saved, "template_created": created, "template_id": int(template.id)}


def _find_existing_for_save(session: Session, trusted: list[tuple[dict, str]], tr_type: str | None = None):
    best = (None, 0.0, {})
    for tender in session.query(Tender).all():
        if tr_type and tender.object_name and norm(tender.object_name) != norm(tr_type):
            continue
        rows = _template_rows(session, tender.id)
        if not rows:
            continue
        unused = set(r.id for r in rows)
        matches = {}
        score = 0.0
        total = len(trusted)
        for item, value in trusted:
            opts = []
            for row in rows:
                if row.id not in unused:
                    continue
                s = _row_match(session, item, row)
                if s >= 0.90:
                    opts.append((s, row))
            if not opts:
                continue
            s, row = max(opts, key=lambda x: x[0])
            unused.remove(row.id)
            matches[id(item)] = (row, s)
            score += 1.0
        coverage = score / max(1, total)
        if coverage > best[1]:
            best = (tender, coverage, matches)
    if best[0] is not None and best[1] >= 0.70:
        return best
    return None, 0.0, {}


def apply_voltage_profile(
    session: Session,
    items: list[dict],
    voltage: str | None,
) -> int:
    """Заполняет пустые строки из старого профиля по напряжению.

    Профиль используется только после шаблона и только для точного/сильного
    сопоставления названия поля. Поэтому он не заменяет индивидуальный шаблон.
    """
    if not voltage:
        return 0
    profiles = session.query(ParameterProfile).filter(
        ParameterProfile.voltage == str(voltage)
    ).all()
    if not profiles:
        return 0
    profile = profiles[0]
    rows = session.query(ProfileParameter).filter_by(profile_id=profile.id).all()
    if not rows:
        return 0
    filled = 0
    for item in items:
        if _clean(item.get("current_value")) or _clean(item.get("algorithm_value")):
            continue
        best = None
        best_score = 0.0
        for row in rows:
            if not _is_real_value(row.value):
                continue
            key_a = _field_key(session, item.get("param_name", ""))
            key_b = _field_key(session, row.parameter_name or "")
            if key_a == key_b:
                score = 1.15
            else:
                score = SequenceMatcher(None, norm(item.get("param_name", "")), norm(row.parameter_name or "")).ratio()
            if score > best_score:
                best_score, best = score, row
        if best is not None and best_score >= 0.90:
            value = _without_star(best.value)
            if value:
                item["algorithm_value"] = value
                item["algorithm_source"] = "DB_PROFILE"
                item["profile_id"] = profile.id
                filled += 1
    return filled


def capture_docx_file(session: Session, path: str | Path) -> dict[str, int]:
    """Обучает БЗ только по финальному заполняемому DOCX.

    Исходное техническое задание без колонки «Предлагаемое участником
    конкурса» не является шаблоном ответа и намеренно игнорируется.
    """
    from docx import Document
    from services.search_docx_AI import extract_doc_items, _determine_type

    doc = Document(str(path))
    items, extra = extract_doc_items(doc)
    document = extra.get("document", {})

    if not document.get("has_answer_column"):
        print(
            f"[KB] {Path(path).name}: явная колонка ответа не найдена — "
            "файл не заносится в шаблонную БЗ"
        )
        return {"saved_rows": 0, "template_created": 0, "template_id": 0}

    if not items:
        return {"saved_rows": 0, "template_created": 0, "template_id": 0}

    tr_type, voltage = _determine_type(session, items, document)

    for item in items:
        item["final_value"] = item.get("current_value", "")
        item["final_mark"] = ""

    return save_document_to_knowledge(
        session,
        items,
        filename=Path(path).name,
        tr_type=tr_type.name,
        voltage=voltage,
    )

