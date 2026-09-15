"""Гибридное автозаполнение тендерных DOCX.

Пайплайн:
1. Алгоритм: для каждой строки ищем канонический ключ БЗ (FieldRule) и
   берем значение из расширяемой базы знаний (KnowledgeEntry) или из
   исторических справочников (TrTypeRule/VoltageClass/...). Такие значения
   — константы, они не помечаются.
2. AI-аудит: GigaChat получает ВЕСЬ контекст документа (все параграфы, все
   таблицы, результат первого прохода, применимую БЗ) и перепроверяет
   каждую строку.
3. Слияние: значение нейросети принимается как "выведенное из контекста"
   (source=AI_CONTEXT, mark="**") только если оно отличается от значения
   алгоритма И уверенность модели не ниже AI_CONFIDENCE_THRESHOLD. Иначе
   побеждает алгоритмическое/БД значение.
4. Рядом с результатом сохраняется `<output>.audit.json` с источником,
   уверенностью и причиной решения по каждой строке — по нему работает
   tools/promote_audit.py для переноса подтвержденного значения в БЗ.
"""

from __future__ import annotations

import json
import os
import re
import uuid
from copy import deepcopy
from pathlib import Path
from typing import Any

from docx import Document
from sqlalchemy.orm import Session

from models.tr_type import (
    AccuracyClass, Climat, IsolColor, IsolType, KnowledgeEntry, TrType, TrTypeRule, VoltageClass
)
from services.gigachat_client import (
    AI_CONFIDENCE_THRESHOLD,
    ask_json,
    check_gigachat_connection,
    open_client,
    verify_connection,
)
from services.knowledge_base import build_knowledge_context, canonicalize_field, norm, resolve_db_field
from services.template_knowledge import find_best_template, apply_template, apply_voltage_profile, save_document_to_knowledge
from services.data_tenders_knowledge import get_data_tenders_knowledge, _param_key

__all__ = ["process_docx_requirements", "check_gigachat_connection"]


def _is_header_row(param_name: str, num: str = "", required_val: str = "") -> bool:
    p = norm(param_name) if isinstance(param_name, str) else ""
    r = norm(required_val) if isinstance(required_val, str) else ""
    n = norm(num) if isinstance(num, str) else ""
    # В некоторых Word-ТЗ заголовок раздела физически занимает все колонки,
    # поэтому параметр == требуемому значению. Такой ряд нельзя автозаполнять.
    section_header = bool(p and not n and r and p == r)
    return (
        section_header
        or "наименование параметра" in p
        or "технические требования к оборудованию" in p
        or p.replace(".", "").strip().isdigit()
        or (not p and n and n.replace(".", "").strip().isdigit())
        or p.endswith(":")
    )


def extract_doc_items(doc: Document) -> tuple[list[dict], dict[str, Any]]:
    """
    Извлекает строки ТЗ.

    Важный момент:
    если строка имеет вид:

        1 | Тип внешней изоляции (фарфор, полимер) | * | <пусто>

    то:
        required_val = "*"
        current_value = ""

    Звёздочка означает свободное поле и НИКОГДА не считается
    фактическим значением.
    """

    items: list[dict] = []
    mapping: dict[str, str] = {}
    counter = 0

    document_context: dict[str, Any] = {
        "paragraphs": [],
        "tables": [],
        "has_answer_column": False,
    }

    # ---------------------------------------------------------
    # Параграфы
    # ---------------------------------------------------------
    for paragraph in doc.paragraphs:
        text = paragraph.text.strip()
        if text:
            document_context["paragraphs"].append(text)

    # ---------------------------------------------------------
    # Таблицы
    # ---------------------------------------------------------
    for table_idx, table in enumerate(doc.tables):
        table_rows = []

        # Word может возвращать одну и ту же физическую ячейку несколько
        # раз через row.cells при вертикальном/горизонтальном merge.
        # Для записи это КРИТИЧНО: row.cells[-1] может оказаться не той
        # физической ячейкой, которую пользователь считает последней.
        # Поэтому сначала строим карту уникальных <w:tc>.
        unique_cells = []
        seen_tc = set()
        for row in table.rows:
            for cell in row.cells:
                tc_id = id(cell._tc)
                if tc_id not in seen_tc:
                    seen_tc.add(tc_id)
                    unique_cells.append(cell)

        # Определяем колонку ответа по заголовку таблицы. Если явного
        # заголовка нет, для тендерных таблиц ответом считается последняя
        # ФИЗИЧЕСКАЯ колонка, а не последний элемент row.cells.
        # Сначала ищем НЕ двусмысленные заголовки ответа. Слово
        # «значение» отдельно использовать нельзя: оно встречается и в
        # заголовке колонки требований.
        strong_answer_keywords = (
            "предлагаемое участником", "предлагаемое значение",
            "предложение участника", "ответ участника",
            "заполняется участником", "данные участника",
        )
        weak_answer_keywords = ("предложение", "ответ")
        answer_col = None

        for probe_row in table.rows[:8]:
            for col_idx, cell in enumerate(probe_row.cells):
                text = norm(cell.text)
                if text and any(k in text for k in strong_answer_keywords):
                    answer_col = col_idx
                    break
            if answer_col is not None:
                break

        if answer_col is None:
            for probe_row in table.rows[:8]:
                for col_idx, cell in enumerate(probe_row.cells):
                    text = norm(cell.text)
                    if text and any(k in text for k in weak_answer_keywords):
                        answer_col = col_idx
                        break
                if answer_col is not None:
                    break

        if answer_col is not None:
            document_context["has_answer_column"] = True

        for row_idx, row in enumerate(table.rows):

            cells = [
                re.sub(r"\s+", " ", str(cell.text or "")).strip()
                for cell in row.cells
            ]

            if not any(cells):
                continue

            table_rows.append({
                "row": row_idx,
                "cells": cells,
            })

            # Минимум:
            # номер | параметр | значение
            if len(cells) < 3:
                continue

            # -------------------------------------------------
            # Определяем номер
            # -------------------------------------------------
            num = cells[0].strip()

            # -------------------------------------------------
            # Определяем название параметра
            # -------------------------------------------------
            param_name = cells[1].strip()

            # -------------------------------------------------
            # Определяем требование и целевую ячейку.
            #
            # Для стандартной таблицы:
            #
            # [номер, параметр, требование, ответ]
            #
            # получаем:
            # required = cells[-2]
            # target   = cells[-1]
            #
            # Для нестандартных таблиц дополнительно ищем
            # последнюю пустую ячейку.
            # -------------------------------------------------

            required = ""
            target = ""

            # Сначала работаем с ФИЗИЧЕСКИМИ ячейками строки.
            # Дубликаты, появившиеся из-за merge, не считаются отдельными
            # колонками. Это устраняет запись "в соседнюю ячейку".
            physical = []
            seen = set()
            for cell in row.cells:
                tc_id = id(cell._tc)
                if tc_id not in seen:
                    seen.add(tc_id)
                    physical.append(cell)

            physical_text = [
                re.sub(r"\s+", " ", str(cell.text or "")).strip()
                for cell in physical
            ]

            target_cell_index = None
            if answer_col is not None and answer_col < len(row.cells):
                answer_cell = row.cells[answer_col]
                answer_tc = id(answer_cell._tc)
                for pi, cell in enumerate(physical):
                    if id(cell._tc) == answer_tc:
                        target_cell_index = pi
                        break

            # Стандартные тендерные таблицы: номер | параметр |
            # требование | предложение. Если заголовка ответа нет,
            # последняя физическая ячейка — запасной вариант.
            if target_cell_index is None and physical:
                target_cell_index = len(physical) - 1

            if len(physical_text) >= 4:
                req_idx = target_cell_index - 1 if target_cell_index is not None and target_cell_index > 1 else len(physical_text) - 2
                required = physical_text[req_idx].strip() if req_idx >= 0 else ""
            elif len(physical_text) >= 3:
                required = ""

            if target_cell_index is not None and target_cell_index < len(physical_text):
                target = physical_text[target_cell_index].strip()

            # -------------------------------------------------
            # Критически важно:
            #
            # "*" в колонке требования НЕ является значением.
            # Это свободное поле.
            # Поэтому target остаётся пустым.
            # -------------------------------------------------
            if target == "*":
                # Иногда Word после объединения ячеек может
                # физически сдвинуть звёздочку в последнюю
                # ячейку.
                required = "*"
                target = ""

            # Если в последней ячейке стоит несколько звёздочек
            if re.fullmatch(r"\*+", target or ""):
                required = target
                target = ""

            if _is_header_row(
                param_name,
                num,
                required
            ):
                continue

            # -------------------------------------------------
            # Отбрасываем реально пустые визуальные строки.
            #
            # Но строки с названием параметра оставляем,
            # даже если required="*" и target="".
            # -------------------------------------------------
            if (
                not param_name
                and not num
                and not required
                and not target
            ):
                continue

            item_id = f"item_{counter}"
            counter += 1

            item = {
                "id": item_id,
                "table": table_idx,
                "row": row_idx,
                "target_cell_index": target_cell_index,

                "num": num,
                "param_name": param_name,

                "required_val": required,
                "current_value": target,

                # Для диагностики структуры исходного DOCX
                "_table_cells": cells,
                "_cell_count": len(cells),
            }

            items.append(item)

            mapping[item_id] = target

        document_context["tables"].append(table_rows)

    # ---------------------------------------------------------
    # Сохраняем полный список строк документа.
    # ---------------------------------------------------------
    document_context["items"] = [
        {
            "id": x["id"],
            "table": x["table"],
            "row": x["row"],
            "num": x["num"],
            "param_name": x["param_name"],
            "required_val": x["required_val"],
            "current_value": x["current_value"],
            "target_cell_index": x.get("target_cell_index"),
        }
        for x in items
    ]

    return items, {
        "document": document_context,
        "mapping": mapping,
    }

def _determine_type(session: Session, items: list[dict], document_context: dict[str, Any] | None = None) -> tuple[TrType, str | None]:
    """Определяет тип только при наличии достаточного основания.

    Старый ``rstrip('.0')`` был ошибочным: для строки ``110`` он давал ``11``.
    Здесь напряжение сравнивается численно, а при нескольких подходящих типах
    неоднозначность не скрывается выбором первой строки БД.
    """
    parts = [
        f"{x.get('param_name', '')} {x.get('required_val', '')}"
        for x in items
    ]
    if document_context:
        parts.extend(document_context.get("paragraphs", []))
        for table in document_context.get("tables", []):
            for row in table:
                parts.extend(row.get("cells", []))
    text = " ".join(parts).lower()
    voltage = None
    for item in items:
        p = norm(item.get("param_name", ""))
        if "номинальн" in p and "напряж" in p:
            m = re.search(r"\d+(?:[,.]\d+)?", str(item.get("required_val", "")))
            if m:
                voltage = m.group(0).replace(",", ".")
                break

    def voltage_equal(a: str, b: str) -> bool:
        try:
            return float(a.replace(",", ".")) == float(b.replace(",", "."))
        except (ValueError, AttributeError):
            return norm(a) == norm(b)

    candidates = []
    rules = session.query(TrTypeRule).all()
    for rule in rules:
        tr = session.query(TrType).filter_by(id=rule.tr_type_id).first()
        if not tr:
            continue
        score = 0
        if norm(tr.name) and norm(tr.name) in norm(text):
            score += 100
        volts = {
            str(x.value)
            for x in session.query(VoltageClass)
            .filter(VoltageClass.id.in_(rule.voltage_classes or []))
            .all()
        }
        if voltage and any(voltage_equal(voltage, v) for v in volts):
            score += 50
        candidates.append((score, tr, rule))

    if voltage is None:
        # Напряжение может находиться в заголовке документа, а не в строке ТЗ.
        m = re.search(r"(?i)\b(\d+(?:[,.]\d+)?)\s*(?:кв|kv)\b", text)
        if m:
            voltage = m.group(1)

    # Пересчитать кандидатов после извлечения напряжения из общего контекста.
    if voltage is not None:
        rescored = []
        for _, tr, rule in candidates:
            score = 0
            if norm(tr.name) and norm(tr.name) in norm(text):
                score += 100
            volts = {
                str(x.value)
                for x in session.query(VoltageClass)
                .filter(VoltageClass.id.in_(rule.voltage_classes or []))
                .all()
            }
            if any(voltage_equal(voltage, v) for v in volts):
                score += 50
            rescored.append((score, tr, rule))
        candidates = rescored

    if not candidates:
        raise ValueError("Не удалось определить тип трансформатора: в БД нет правил.")

    candidates.sort(key=lambda x: x[0], reverse=True)
    best_score = candidates[0][0]
    best = [x for x in candidates if x[0] == best_score]
    if best_score <= 0 and len(candidates) > 1:
        raise ValueError(
            "Не удалось однозначно определить тип трансформатора: "
            "в документе нет названия типа или распознанного напряжения."
        )
    if len(best) > 1:
        names = ", ".join(x[1].name for x in best)
        raise ValueError(f"Тип трансформатора определен неоднозначно: {names}.")
    return best[0][1], voltage


def _algorithm_fill(session: Session, items: list[dict], tr_type: TrType, voltage: str | None, data_kb=None, document_text: str = "") -> list[dict]:
    result = []
    # Модель и напряжение определяем один раз на весь документ. Ранее
    # detect_model() повторно проходил весь контекст для каждой строки.
    data_model = data_kb.detect_model(document_text, items) if data_kb is not None else None
    effective_model = data_model or (tr_type.name if tr_type and "ТРГ-УЭТМ" in str(tr_type.name).upper() else None)
    effective_voltage = voltage or (data_kb.detect_voltage(document_text, items) if data_kb is not None else None)
    for item in items:
        # При активном data_tenders не делаем SQL/fuzzy-поиск FieldRule на каждой строке.
        db_key = _param_key(item["param_name"]) if data_kb is not None else canonicalize_field(session, item["param_name"])
        param_text = norm(item.get("param_name", ""))
        item["db_key"] = db_key
        value, source = (None, None)

        # Изготовитель и заводской тип/марка — обязательные константы проекта.
        # Для них дополнительно используем распознавание по названию строки,
        # чтобы даже при неидеальном совпадении FieldRule значение не потерялось.
        is_manufacturer = (
            re.match(r"^(?:изготовител(?:ь|я|ем)?|производител(?:ь|я|ем)?)\b", param_text) is not None
            or "завод-изготовитель" in param_text
        )
        is_brand = (
            "заводской тип" in param_text
            or "заводской тип марка" in param_text
            or ("тип" in param_text and "марка" in param_text)
        )
        # Константы берем только из data_tenders. SQLite БД в этом проходе
        # намеренно не используется. Составные значения передаются AI как
        # candidate для второго уровня и не фиксируются вслепую.
        if data_kb is not None:
            value, source, _score = data_kb.get_constant(
                item["param_name"],
                text_context=document_text,
                model=effective_model,
                voltage=effective_voltage,
                required_val=item.get("required_val", ""),
            )
        else:
            value, source = None, None

        candidate = str(value or "").strip()
        required = str(item.get("required_val", "") or "").strip()
        param_lower = param_text
        ambiguous = bool(
            candidate and required == "*" and
            (
                "габарит" in param_lower
                or "масса трансформатора" in param_lower
                or "масса масла" in param_lower
                or "климатическ" in param_lower
                or len(re.findall(r"\d+(?:[.,]\d+)?", candidate)) > 1
                or re.search(r"(?:\bили\b|/|;|,)", candidate, flags=re.IGNORECASE)
            )
        )
        item["algorithm_candidate"] = candidate
        item["algorithm_candidate_source"] = source or "NONE"
        item["algorithm_value"] = "" if ambiguous else candidate
        item["algorithm_source"] = "NONE" if ambiguous else (source or "NONE")
        result.append(item.copy())
    return result


def _prompt_items(items: list[dict]) -> list[dict]:
    """Возвращает только JSON-сериализуемые поля для AI-промпта.

    Внутри рабочего item хранится python-docx Cell для записи результата.
    Такие объекты нельзя передавать в json.dumps и, главное, нет смысла
    отправлять модели. Это также защищает промпт от случайного появления
    служебных объектов в будущем.
    """
    fields = (
        "id", "table", "row", "num", "param_name", "required_val",
        "current_value", "db_key", "algorithm_value", "algorithm_source",
    )
    result = []
    for item in items:
        row = {key: item.get(key) for key in fields}
        if item.get("local_context"):
            row["local_context"] = item["local_context"]
        result.append(row)
    return result


def _is_db_source(source: str) -> bool:
    """Определяет, является ли значение подтвержденным алгоритмической БЗ."""
    return bool(source) and source not in {"NONE", "HEURISTIC", "FALLBACK_ECHO"}


def _clean_ai_value(value: str) -> str:
    """Очищает ответ AI от служебных звездочек/маркеров.

    Маркер добавляет только программа. Поэтому 0,2**, 0,2*** и другие
    варианты от самой модели превращаются в обычное значение 0,2.
    """
    value = str(value or "").strip()
    if not value:
        return ""
    if re.fullmatch(r"\*+", value):
        return ""
    value = re.sub(r"\s*\*+\s*$", "", value).strip()
    return value


def _ai_mark_for_required(required_val: str) -> str:
    """Возвращает корректный маркер для значения, выведенного нейросетью.

    Если в левой колонке уже есть *, достаточно одной * в правой колонке.
    Иначе значение, действительно выведенное AI и отсутствующее в БД,
    получает стандартную отметку **.
    """
    raw = str(required_val or "").strip()
    return "*" if "*" in raw else "**"


def _fallback_from_requirement(required_val: str) -> str:
    """Извлекает минимально пригодный ответ непосредственно из ТЗ.

    Это аварийный, но детерминированный fallback для ситуации, когда GigaChat
    не вернул строку. Мы не оставляем явное требование пустым: если слева
    указано конкретное значение или текстовое требование, оно переносится
    вправо. Служебные звездочки из левой колонки удаляются — маркер добавляет
    только _ai_mark_for_required().
    """
    raw = str(required_val or "").replace("\r", "").strip()
    if not raw:
        return ""

    # '*' и варианты вроде '**' — это признак свободного поля, а не значение.
    cleaned = re.sub(r"\s*\*+\s*$", "", raw).strip()
    cleaned = re.sub(r"\s{2,}", " ", cleaned)
    cleaned = cleaned.strip(" /")
    if not cleaned or re.fullmatch(r"\*+", cleaned):
        return ""

    # Если слева есть явный текст/число/единица — его можно использовать как
    # исходный ответ. Не придумываем ничего сверх текста самого требования.
    return cleaned


def _values_equivalent(left: str, right: str) -> bool:
    """Сравнивает ответы без ложных расхождений из-за формата записи."""
    left = _clean_ai_value(left)
    right = _clean_ai_value(right)

    def normalize(value: str) -> str:
        value = norm(value).replace(" ", "").replace(",", ".")
        value = re.sub(r"(?i)(кв|ка|а|в)$", "", value)
        return value
    return normalize(left) == normalize(right)


def _legacy_db_values(session: Session, tr_type: TrType, key: str) -> set[str]:
    """Возвращает все допустимые константы старых справочников БД."""
    rule = session.query(TrTypeRule).filter_by(tr_type_id=tr_type.id).first()
    if not rule:
        return set()

    def names(model, ids):
        column = getattr(model, "name", None) or getattr(model, "value", None)
        if column is None:
            return []
        return [str(x[0]) for x in session.query(column).filter(model.id.in_(ids or [])).all()]

    if key == "nominal_voltage":
        return set(names(VoltageClass, rule.voltage_classes))
    if key == "climate":
        return set(names(Climat, rule.climats))
    if key in {"internal_insulation", "external_insulation"}:
        return set(names(IsolType, rule.isol_types))
    if key == "external_insulation_color":
        return set(names(IsolColor, rule.isol_colors))
    if key == "accuracy_class":
        return set(names(AccuracyClass, rule.accuracy_classes))
    return set()


def _is_known_db_value(session: Session, item: dict, value: str, tr_type: TrType, voltage: str | None) -> bool:
    """Проверяет, есть ли значение среди применимых констант БД.

    Проверка намеренно шире canonical field строки: если AI вывел, например,
    0,2 / 0,2S / 10PR / 220 / У1, а такое значение уже есть в БД текущего
    проекта, маркер ** ставить нельзя.
    """
    value = _clean_ai_value(value)
    if not value:
        return False

    keys = []
    if item.get("db_key"):
        keys.append(item["db_key"])
    keys.extend([
        "nominal_voltage", "climate", "internal_insulation",
        "external_insulation", "external_insulation_color", "accuracy_class",
    ])

    # KnowledgeEntry — расширяемая БД проекта.
    for row in session.query(KnowledgeEntry).filter_by(active=True).all():
        if row.tr_type_id not in (None, tr_type.id):
            continue
        if row.voltage and (not voltage or str(row.voltage) != str(voltage)):
            continue
        if row.value and _values_equivalent(value, str(row.value)):
            return True

    # Legacy-справочники.
    for key in dict.fromkeys(keys):
        for known in _legacy_db_values(session, tr_type, key):
            if known and _values_equivalent(value, known):
                return True
    return False


def _is_blank_target(item: dict) -> bool:
    return not str(item.get("current_value", "") or "").strip()


def _constraint_text(text: str) -> bool:
    t = norm(text)
    return any(x in t for x in (
        "не менее", "не более", "в соответствии", "согласно",
        "должен", "должны", "обязательно", "указать", "определить проектом",
    ))


def _fact_value(value: str, required: str = "") -> str:
    """Возвращает значение, пригодное как внутренний факт документа.

    Не используем строки вида «не менее 1000», «указать», «обязательно» как
    фактические значения. Точные значения вроде 110, 50, У1, Фарфор сохраняем.
    """
    value = re.sub(r"\s+", " ", str(value or "")).strip()
    required = re.sub(r"\s+", " ", str(required or "")).strip()
    if not value or _constraint_text(value):
        return ""
    if value.lower() in {"да", "нет", "не требуется", "не относится"}:
        return value
    if _constraint_text(required) and _values_equivalent(value, required):
        return ""
    return value


def _tokens(text: str) -> set[str]:
    stop = {
        "и", "в", "на", "для", "по", "из", "с", "к", "а", "от", "до",
        "при", "не", "кв", "гц", "а", "ва", "мм", "м", "с", "шт",
    }
    return {
        x for x in re.findall(r"[а-яa-z0-9]+", norm(text), flags=re.I)
        if len(x) > 2 and x not in stop
    }

def _extract_allowed_values(text: str) -> list[str]:
    """
    Извлекает допустимые значения из скобок параметра.

    Примеры:
        '(фарфор, полимер)'       -> ['фарфор', 'полимер']
        '(белый/коричневый)'      -> ['белый', 'коричневый']
        '(да/нет)'                -> ['да', 'нет']
        '(1, 5)'                  -> ['1', '5']
        '(220/380 В)'             -> ['220/380 В']

    Важно:
    slash внутри составного числового значения вроде 220/380 В
    не считается разделителем вариантов.
    """
    text = str(text or "")
    result: list[str] = []

    for match in re.finditer(r"\(([^()]*)\)", text):
        content = match.group(1).strip()
        if not content:
            continue

        # Основные разделители вариантов.
        parts = re.split(r"\s*[;,]\s*|\s+\bили\b\s+", content, flags=re.I)

        # '/' считаем разделителем только когда это действительно
        # короткие самостоятельные варианты:
        # белый/коричневый, да/нет, 1/5.
        if len(parts) == 1 and "/" in content:
            slash_parts = [x.strip() for x in content.split("/") if x.strip()]

            # 220/380 В — единое техническое значение.
            # Белый/коричневый — два варианта.
            if (
                len(slash_parts) >= 2
                and not re.search(r"\b(?:в|кв|ка|а|мм|см|м|гц|кн)\b", content, re.I)
                and all(len(x) <= 40 for x in slash_parts)
            ):
                parts = slash_parts

        for part in parts:
            value = re.sub(r"^\s*[-•]\s*", "", part).strip()
            value = re.sub(r"\s+", " ", value)

            if value and value not in result:
                result.append(value)

    return result


def _allowed_values(item: dict) -> list[str]:
    """
    Возвращает допустимые значения конкретного поля.

    Варианты могут быть указаны как в названии параметра,
    так и непосредственно в колонке требований.
    """
    values: list[str] = []

    for source in (
        item.get("param_name", ""),
        item.get("required_val", ""),
    ):
        for value in _extract_allowed_values(source):
            if norm(value) not in {norm(x) for x in values}:
                values.append(value)

    return values


def _matches_allowed_value(value: str, allowed_values: list[str]) -> bool:
    """
    Проверяет, соответствует ли ответ одному из допустимых вариантов.
    """
    value = _clean_ai_value(value)
    if not value or not allowed_values:
        return True

    value_norm = norm(value)

    for allowed in allowed_values:
        allowed_norm = norm(allowed)

        if value_norm == allowed_norm:
            return True

        # Например:
        # AI -> "Фарфор"
        # вариант -> "фарфор"
        if value_norm.replace(" ", "") == allowed_norm.replace(" ", ""):
            return True

        # Для значений с единицами допускаем:
        # "110 кВ" == "110"
        if _values_equivalent(value, allowed):
            return True

    return False

def _param_similarity(a: str, b: str) -> float:
    from difflib import SequenceMatcher
    ta, tb = _tokens(a), _tokens(b)
    if not ta or not tb:
        return 0.0
    overlap = len(ta & tb) / max(1, min(len(ta), len(tb)))
    seq = SequenceMatcher(None, norm(a), norm(b)).ratio()
    return max(overlap, seq)


def _has_engineering_conflict(a: str, b: str) -> bool:
    """Запрещает перенос между разными семействами одной характеристики."""
    a, b = norm(a), norm(b)
    families = [
        ("удельная длина пути утечки",),
        ("номинальный ток первичной", "первичный ток"),
        ("номинальный вторичный ток", "вторичный ток"),
        ("ток термической стойкости", "термической стойкости"),
        ("ток динамической стойкости", "динамической стойкости"),
        ("ток отключения", "номинальный ток отключения"),
        ("номинальное напряжение", "номинальная частота", "испытательное напряжение", "наибольшее напряжение"),
        ("высота установки", "высота металлоконструкций"),
        ("тип внешней изоляции", "вид изоляции", "удельная длина пути утечки"),
    ]
    # Сильный маркер встречается только в одном из двух названий — строки
    # относятся к разным полям и не могут служить аналогом.
    for group in families:
        hits_a = [x for x in group if x in a]
        hits_b = [x for x in group if x in b]
        if bool(hits_a) != bool(hits_b):
            return True
    return False


def _build_self_context(items: list[dict], extra: dict[str, Any]) -> dict[str, Any]:
    """Строит компактный профиль только из фактов самого документа."""
    facts = []
    for x in items:
        if not _is_blank_target(x):
            value = _fact_value(x.get("current_value", ""), x.get("required_val", ""))
            if value:
                facts.append({
                    "id": x["id"],
                    "table": x["table"],
                    "row": x["row"],
                    "num": x.get("num", ""),
                    "param_name": x.get("param_name", ""),
                    "value": value,
                })

    paragraphs = [re.sub(r"\s+", " ", str(x)).strip() for x in extra.get("paragraphs", []) if str(x).strip()]
    # В профиль попадают только информативные фразы: названия объекта,
    # оборудования, типы, напряжения, производители и т.п.
    important = []
    for text in paragraphs:
        low = norm(text)
        if any(k in low for k in (
            "трансформатор", "выключатель", "тфнд", "трг", "110 кв", "220 кв",
            "изготовител", "производител", "пс ", "подстанц",
        )):
            important.append(text[:1000])

    # Глобальный профиль: только ключевые факты, а не весь документ.
    priority_terms = (
        "изготовител", "производител", "заводской тип", "марка", "трг",
        "номинальное напряжение", "номинальная частота", "климатическое",
        "изоляц", "первичный ток", "вторичный ток", "вторичных обмоток",
        "сейсмостойкость", "высота установки",
    )
    global_facts = []
    used = set()
    for fact in facts:
        key = norm(fact.get("param_name", ""))
        if any(term in key for term in priority_terms):
            if fact["id"] not in used:
                global_facts.append(fact)
                used.add(fact["id"])
    for fact in facts:
        if len(global_facts) >= 24:
            break
        if fact["id"] not in used:
            global_facts.append(fact)
            used.add(fact["id"])
    return {
        "document_facts": facts,
        "global_facts": global_facts[:24],
        "key_paragraphs": important[:15],
    }


def _find_self_context_value(
    item: dict,
    all_items: list[dict],
    *,
    min_score: float = 0.92
) -> tuple[str, str, list[str]]:
    """
    Ищет однозначный аналог параметра внутри ЭТОГО ЖЕ документа.

    Дополнительно учитывает варианты, указанные в скобках.

    Например:

        1.9. Тип внешней изоляции -> Фарфор

        2.14. Тип внешней изоляции (фарфор, полимер) -> *

    Результат:

        Фарфор
    """
    target = norm(item.get("param_name", ""))
    if not target:
        return "", "", []

    allowed = _allowed_values(item)

    candidates = []

    for x in all_items:
        if x.get("id") == item.get("id") or _is_blank_target(x):
            continue

        value = _fact_value(
            x.get("current_value", ""),
            x.get("required_val", "")
        )

        if not value:
            continue

        # Если у целевого поля есть список вариантов,
        # значение источника ОБЯЗАНО попасть в этот список.
        if allowed and not _matches_allowed_value(value, allowed):
            continue

        if _has_engineering_conflict(
            target,
            x.get("param_name", "")
        ):
            continue

        raw_score = _param_similarity(
            target,
            x.get("param_name", "")
        )

        score = raw_score

        # Совпадение DB-key усиливает связь.
        if (
            norm(item.get("db_key", ""))
            and norm(item.get("db_key", ""))
            == norm(x.get("db_key", ""))
            and raw_score >= 0.72
        ):
            score += 0.20

        # Одинаковый последний номер подпункта.
        t_num = str(item.get("num", ""))
        x_num = str(x.get("num", ""))

        sub_same = bool(
            t_num
            and x_num
            and t_num.split(".")[-1:] == x_num.split(".")[-1:]
        )

        if sub_same and raw_score >= 0.72:
            score += 0.12

        # Если явно указаны допустимые варианты,
        # совпадение параметра должно быть достаточно сильным.
        if allowed:
            min_allowed_score = max(min_score, 0.90)
        else:
            min_allowed_score = min_score

        if score >= min_allowed_score:
            candidates.append(
                (score, x, value)
            )

    candidates.sort(
        key=lambda z: z[0],
        reverse=True
    )

    if not candidates:
        return "", "", []

    best_score = candidates[0][0]

    best = [
        x for x in candidates
        if x[0] >= best_score - 0.03
    ]

    values = {
        norm(x[2])
        for x in best
    }

    # Если разные строки дают разные ответы,
    # нельзя выбирать случайный.
    if len(values) != 1:
        return "", "", []

    source = best[0][1]

    return (
        best[0][2],
        "SELF_CONTEXT",
        [source["id"]],
    )

def _build_target_context(
    item: dict,
    all_items: list[dict],
    self_profile: dict[str, Any]
) -> dict[str, Any]:
    target = norm(item.get("param_name", ""))

    allowed_values = _allowed_values(item)

    scored = []

    for fact in self_profile.get("document_facts", []):
        fact_value = fact.get("value", "")

        # Если у поля есть варианты, показываем AI только
        # релевантные факты.
        if (
            allowed_values
            and not _matches_allowed_value(
                fact_value,
                allowed_values
            )
        ):
            continue

        score = _param_similarity(
            target,
            fact.get("param_name", "")
        )

        if score >= 0.25:
            scored.append((score, fact))

    scored.sort(
        key=lambda x: x[0],
        reverse=True
    )

    related = [
        fact
        for _, fact in scored[:6]
    ]

    same_section = []

    table = item.get("table")
    row = item.get("row", 0)

    for x in all_items:
        if (
            x.get("id") == item.get("id")
            or x.get("table") != table
            or _is_blank_target(x)
        ):
            continue

        if abs(
            int(x.get("row", 0))
            - int(row)
        ) <= 4:

            value = _fact_value(
                x.get("current_value", ""),
                x.get("required_val", "")
            )

            if not value:
                continue

            if (
                allowed_values
                and not _matches_allowed_value(
                    value,
                    allowed_values
                )
            ):
                continue

            same_section.append({
                "id": x["id"],
                "num": x.get("num", ""),
                "param_name": x.get("param_name", ""),
                "value": value,
            })

    return {
        "allowed_values": allowed_values,
        "related_facts": related[:4],
        "nearby_facts": same_section[:4],
    }

def _prepare_ai_targets(
    items: list[dict],
    self_profile: dict[str, Any]
) -> tuple[list[dict], list[dict]]:
    """
    Формирует список полей для AI.

    Пустое поле с required_val="*" является полноценным
    полем для заполнения.

    Например:

        Тип внешней изоляции (фарфор, полимер) | * | ""

    обязательно попадает в ai_targets.
    """

    ai_targets = []
    self_filled = []

    for item in items:

        # Уже заполненное поле не отправляем в AI.
        if not _is_blank_target(item):
            continue

        param = str(item.get("param_name", "") or "").strip()
        num = str(item.get("num", "") or "").strip()
        required = str(item.get("required_val", "") or "").strip()
        db_key = str(item.get("db_key", "") or "").strip()

        # ---------------------------------------------
        # Только полностью пустая техническая строка
        # считается визуальной строкой.
        # ---------------------------------------------
        if not param and not num and not required and not db_key:
            item["pre_ai_value"] = ""
            item["pre_ai_source"] = "SKIP_NON_FIELD"
            item["pre_ai_evidence"] = []
            continue

        # ---------------------------------------------
        # Значение уже найдено детерминированным проходом:
        # KnowledgeEntry / legacy DB / сохраненный шаблон.
        # Такой параметр НЕ передаем в AI.
        # ---------------------------------------------
        if str(item.get("algorithm_value", "") or "").strip():
            item["pre_ai_value"] = str(item["algorithm_value"]).strip()
            item["pre_ai_source"] = str(item.get("algorithm_source", "DB") or "DB")
            item["pre_ai_evidence"] = []
            self_filled.append(item)
            continue

        # ---------------------------------------------
        # Сначала ищем точный аналог в текущем документе.
        # ---------------------------------------------
        value, source, evidence = _find_self_context_value(
            item,
            items,
        )

        if value:
            item["pre_ai_value"] = value
            item["pre_ai_source"] = source
            item["pre_ai_evidence"] = evidence

            self_filled.append(item)
            continue

        # ---------------------------------------------
        # Если аналога нет — отправляем AI.
        #
        # "*" здесь НЕ является причиной пропуска.
        # ---------------------------------------------
        item["pre_ai_value"] = ""
        item["pre_ai_source"] = ""
        item["pre_ai_evidence"] = []

        ai_targets.append(item)

    return ai_targets, self_filled

def _compact_ai_profile(self_profile: dict[str, Any], max_facts: int = 80) -> dict[str, Any]:
    return {
        "key_paragraphs": self_profile.get("key_paragraphs", [])[:20],
        "document_facts": self_profile.get("document_facts", [])[:max_facts],
    }


def _build_ai_prompt(
    target_items: list[dict],
    all_items: list[dict],
    self_profile: dict[str, Any],
    kb_context: dict,
    document_context: dict[str, Any] | None = None,
) -> str:

    document_context = document_context or {}

    # =========================================================
    # 1. ПОЛНЫЙ ДОКУМЕНТ
    # =========================================================

    full_document = {
        "paragraphs": document_context.get("paragraphs", []),
        "tables": document_context.get("tables", []),
    }

    # =========================================================
    # 2. ВСЕ ИЗВЛЕЧЁННЫЕ СТРОКИ
    # =========================================================

    document_items = []

    for item in all_items:
        document_items.append({
            "id": item.get("id", ""),
            "table": item.get("table", ""),
            "row": item.get("row", ""),
            "num": item.get("num", ""),
            "param_name": item.get("param_name", ""),
            "required_val": item.get("required_val", ""),
            "current_value": item.get("current_value", ""),
            "db_key": item.get("db_key", ""),
            "algorithm_value": item.get("algorithm_value", ""),
            "algorithm_source": item.get("algorithm_source", ""),
            "algorithm_candidate": item.get("algorithm_candidate", ""),
            "algorithm_candidate_source": item.get("algorithm_candidate_source", ""),
        })

    # =========================================================
    # 3. ЦЕЛЕВЫЕ ПОЛЯ
    # =========================================================

    targets = []

    for item in target_items:

        allowed_values = _allowed_values(item)

        targets.append({
            "id": item.get("id", ""),
            "table": item.get("table", ""),
            "row": item.get("row", ""),
            "num": item.get("num", ""),
            "param_name": item.get("param_name", ""),
            "required_val": item.get("required_val", ""),
            "current_value": item.get("current_value", ""),
            "allowed_values": allowed_values,

            "context": _build_target_context(
                item,
                all_items,
                self_profile,
            ),
        })

    # =========================================================
    # 4. ПОЛНЫЙ КОРПУС data_tenders
    # =========================================================

    data_tenders_context = kb_context

    # =========================================================
    # 5. Результат алгоритмического прохода
    # =========================================================

    algorithm_pass = document_context.get(
        "algorithm_pass",
        []
    )

    # =========================================================
    # 6. ОПРЕДЕЛЁННЫЙ ТИП
    # =========================================================

    detected = document_context.get(
        "detected",
        {}
    )

    # =========================================================
    # 7. PROMPT
    # =========================================================

    data_tenders_json = json.dumps(data_tenders_context, ensure_ascii=False, indent=2)

    return f"""
Ты инженер по высоковольтному электрооборудованию.

Твоя задача — заполнить ТОЛЬКО пустые поля тендерной таблицы.

=========================================================
КРИТИЧЕСКИЕ ПРАВИЛА
=========================================================

1. "*" означает, что поле свободное и требует заполнения.

2. "*" НИКОГДА не является фактическим значением.

3. Если required_val="*", это НЕ означает, что value="*".

4. required_val — это требование технического задания,
   а не обязательно готовый ответ.

5. Если в названии параметра указаны варианты:

   Тип внешней изоляции (фарфор, полимер)

   то допустимые значения:

   ["фарфор", "полимер"]

6. Если указано:

   Цвет внешней изоляции (белый/коричневый)

   допустимые значения:

   ["белый", "коричневый"]

7. Для таких полей разрешено вернуть ТОЛЬКО один
   из перечисленных вариантов.

8. НЕЛЬЗЯ автоматически выбирать первый вариант.

9. Если в текущем документе уже есть такой же параметр
   с конкретным значением — используй его.

10. Если значение отсутствует в строке, ищи его:
    а) в других строках этого же документа;
    б) в других таблицах документа;
    в) в параграфах документа;
    г) в результате алгоритмического прохода;
    д) в БЗ.

11. Нельзя использовать соседнюю строку как значение,
    если она относится к другому параметру.

12. Например:

    Тип внешней изоляции (фарфор, полимер) | * | ""
    Цвет внешней изоляции (белый/коричневый) | * | ""

    Эти две строки являются РАЗНЫМИ характеристиками.

13. Если найдено значение:

    1.9 Тип внешней изоляции -> Фарфор

    а целевое поле:

    2.1 Тип внешней изоляции (фарфор, полимер) -> *

    нужно вернуть:

    "Фарфор"

14. Если найдено:

    Цвет внешней изоляции -> Белый

    то для:

    Цвет внешней изоляции (белый/коричневый)

    нужно вернуть:

    "Белый".

15. Если точного основания нет — value="".

16. Никогда не возвращай "*" в value.

17. Никогда не возвращай "**" в value.

18. evidence должен содержать ID строк, на основании
    которых принято решение.

=========================================================
ОПРЕДЕЛЁННЫЙ ОБЪЕКТ
=========================================================

{json.dumps(
    detected,
    ensure_ascii=False,
    indent=2
)}

=========================================================
ПОЛНЫЙ ТЕКСТ ДОКУМЕНТА
=========================================================

{json.dumps(
    full_document,
    ensure_ascii=False,
    indent=2
)}

=========================================================
ВСЕ СТРОКИ ТАБЛИЦ
=========================================================

{json.dumps(
    document_items,
    ensure_ascii=False,
    indent=2
)}

=========================================================
ПЕРВЫЙ АЛГОРИТМИЧЕСКИЙ ПРОХОД
=========================================================

{json.dumps(
    algorithm_pass,
    ensure_ascii=False,
    indent=2
)}

=========================================================
ФАКТЫ ДОКУМЕНТА
=========================================================

{json.dumps(
    self_profile.get("document_facts", [])[:100],
    ensure_ascii=False,
    indent=2
)}

=========================================================
КЛЮЧЕВЫЕ ФАКТЫ
=========================================================

{json.dumps(
    self_profile.get("global_facts", [])[:40],
    ensure_ascii=False,
    indent=2
)}

=========================================================
ПОЛНЫЙ КОРПУС data_tenders
=========================================================

Это основной технический эталон для констант и связанных
характеристик. В контексте присутствуют ВСЕ найденные DOCX
из data_tenders, все их таблицы без дублирования merged-ячеек
и OCR-текст встроенных изображений/чертежей.

{data_tenders_json}

=========================================================
ПОЛЯ, КОТОРЫЕ НУЖНО ЗАПОЛНИТЬ
=========================================================

{json.dumps(
    targets,
    ensure_ascii=False,
    indent=2
)}

=========================================================
ФОРМАТ ОТВЕТА
=========================================================

Верни строго JSON.

Для каждого id обязательно верни объект:

{{
  "item_10": {{
    "value": "Фарфор",
    "source": "DATA_TENDERS|SELF_CONTEXT|AI_CONTEXT|NONE",
    "confidence": 0.98,
    "evidence": ["item_5"],
    "reason": "Значение найдено в строке item_5 текущего документа."
  }}
}}

Если значение определить нельзя:

{{
  "item_10": {{
    "value": "",
    "source": "NONE",
    "confidence": 0.0,
    "evidence": [],
    "reason": "Подтвержденное значение отсутствует."
  }}
}}

ВАЖНО:
- не возвращай "*";
- не возвращай "**";
- не добавляй текст вне JSON;
- возвращай ВСЕ переданные id.
"""

def _ai_review(
    client,
    items: list[dict],
    self_profile: dict[str, Any],
    kb_context: dict,
    document_context: dict[str, Any] | None = None,
    chunk_size: int | None = None,
) -> dict[str, dict]:

    document_context = document_context or {}

    pending = [
        x
        for x in items
        if (
            _is_blank_target(x)
            and x.get("pre_ai_source") != "SKIP_NON_FIELD"
            and not x.get("pre_ai_value")
        )
    ]

    if not pending:
        return {}

    merged: dict[str, dict] = {}

    # Обычно один запрос закрывает весь остаток. Если prompt слишком большой,
    # делаем небольшое число крупных чанков вместо старых порций по 12 строк.
    if chunk_size is None:
        try:
            chunk_size = max(20, int(os.getenv("TENDER_AI_CHUNK_SIZE", "32")))
        except ValueError:
            chunk_size = 32
    try:
        max_prompt_chars = int(os.getenv("TENDER_AI_MAX_PROMPT_CHARS", "240000"))
    except ValueError:
        max_prompt_chars = 240000

    # Если весь запрос помещается в лимит, отправляем его ровно один раз.
    try:
        whole_prompt = _build_ai_prompt(pending, items, self_profile, kb_context, document_context)
        if len(whole_prompt) <= max_prompt_chars:
            try:
                part = ask_json(client, whole_prompt)
                if isinstance(part, dict):
                    valid_ids = {x["id"] for x in pending}
                    return {k: v for k, v in part.items() if k in valid_ids and isinstance(v, dict)}
            except Exception as exc:
                print(f"[GigaChat] единый запрос не прошёл, переходим на чанки: {exc}")
    except Exception as exc:
        print(f"[GigaChat] не удалось собрать единый prompt: {exc}")

    for start in range(0, len(pending), chunk_size):

        chunk = pending[start:start + chunk_size]

        prompt = _build_ai_prompt(
            chunk,
            items,
            self_profile,
            kb_context,
            document_context,
        )

        try:
            part = ask_json(
                client,
                prompt,
            )

            if not isinstance(part, dict):
                print(
                    f"[GigaChat] chunk "
                    f"{start + 1}-{start + len(chunk)}: "
                    f"неверный формат ответа"
                )
                continue

            ids = {
                x["id"]
                for x in chunk
            }

            for key, value in part.items():

                if key not in ids:
                    continue

                if not isinstance(value, dict):
                    continue

                # ---------------------------------------------
                # Никогда не принимаем "*" как value
                # ---------------------------------------------
                clean_value = _clean_ai_value(
                    value.get("value", "")
                )

                value["value"] = clean_value

                merged[key] = value

            missing = ids - set(merged)

            if missing:
                print(
                    f"[GigaChat] chunk "
                    f"{start + 1}-{start + len(chunk)}: "
                    f"без ответа {sorted(missing)}"
                )

        except Exception as exc:

            print(
                f"[GigaChat] chunk "
                f"{start + 1}-{start + len(chunk)} "
                f"failed: {exc}"
            )

    return merged

def _is_constraint_like_required(required_val: str, param_name: str = "") -> bool:
    """True для требований, которые нельзя механически копировать как предложение."""
    raw = norm(required_val)
    p = norm(param_name)
    if not raw or raw == "*":
        return True
    if any(x in raw for x in ("не менее", "не более", "в соответствии", "согласно", "должен", "должны", "обязательно", "требуется")):
        return True
    if "предел" in p or "сопротивлен" in p or "ток" in p or "нагруз" in p:
        # Числовые поля с ограничением лучше заполнять фактическим значением
        # из профиля/контекста, а не просто повторять порог ТЗ.
        return bool(re.search(r"\d", raw))
    return False


def _same_requirement_value(required_val: str, ai_value: str) -> bool:
    r = _clean_ai_value(required_val)
    a = _clean_ai_value(ai_value)
    return bool(r and a and _values_equivalent(r, a))


def _context_validation(item: dict, value: str, all_items: list[dict], *, document_context: dict | None = None) -> tuple[bool, str]:
    """Проверяет очевидные логические противоречия ответа с остальным ТЗ."""
    value = _clean_ai_value(value)
    if not value:
        return True, ""
    p = norm(item.get("param_name", ""))

    def first_number(text: str) -> float | None:
        m = re.search(r"[-+]?\d+(?:[,.]\d+)?", str(text))
        if not m:
            return None
        try:
            return float(m.group(0).replace(",", "."))
        except ValueError:
            return None

    # Связка "номинальное напряжение" / "наибольшее напряжение".
    if "наибольшее" in p and "напряж" in p:
        nominal = next((x for x in all_items if "номинальн" in norm(x.get("param_name", "")) and "напряж" in norm(x.get("param_name", ""))), None)
        n = first_number((nominal or {}).get("algorithm_value") or (nominal or {}).get("required_val", "")) if nominal else None
        a = first_number(value)
        if n is not None and a is not None and a < n:
            return False, f"Наибольшее напряжение {a:g} не может быть меньше номинального {n:g}."

    # Температуры: верхняя не должна быть ниже нижней.
    if "верхнее рабочее" in p and "температур" in p:
        low = next((x for x in all_items if "нижнее рабочее" in norm(x.get("param_name", "")) and "температур" in norm(x.get("param_name", ""))), None)
        lo = first_number((low or {}).get("algorithm_value") or (low or {}).get("required_val", "")) if low else None
        hi = first_number(value)
        if lo is not None and hi is not None and hi < lo:
            return False, f"Верхняя температура {hi:g} не может быть ниже нижней {lo:g}."

    if "нижнее рабочее" in p and "температур" in p:
        high = next((x for x in all_items if "верхнее рабочее" in norm(x.get("param_name", "")) and "температур" in norm(x.get("param_name", ""))), None)
        hi = first_number((high or {}).get("algorithm_value") or (high or {}).get("required_val", "")) if high else None
        lo = first_number(value)
        if hi is not None and lo is not None and hi < lo:
            return False, f"Нижняя температура {lo:g} не может быть выше верхней {hi:g}."

    # Для полей с явным ограничением не принимаем дословный повтор порога,
    # если нет дополнительного подтверждающего контекста.
    if _is_constraint_like_required(item.get("required_val", ""), item.get("param_name", "")) and _same_requirement_value(item.get("required_val", ""), value):
        req = norm(item.get("required_val", ""))
        if any(x in req for x in ("не менее", "не более", "в соответствии", "согласно")):
            return False, "Ответ дословно повторяет ограничение ТЗ, а не фактическое значение характеристики."
    # ---------------------------------------------------------
    # ПРОВЕРКА ДОПУСТИМЫХ ЗНАЧЕНИЙ ИЗ СКОБОК
    # ---------------------------------------------------------
    allowed_values = _allowed_values(item)

    if allowed_values:
        if not _matches_allowed_value(
            value,
            allowed_values
        ):
            return (
                False,
                "Значение не входит в допустимые варианты: "
                + ", ".join(allowed_values)
            )

    return True, ""


def _set_cell_text(cell, value: str, mark: str) -> None:
    """Записывает значение в ячейку, сохраняя базовое оформление."""
    value = _clean_ai_value(value)
    text = f"{value}{mark}" if value and mark else value
    paragraphs = cell.paragraphs
    if not paragraphs:
        cell.text = text
        return

    paragraph = paragraphs[0]
    template_rpr = None
    if paragraph.runs and paragraph.runs[0]._r.rPr is not None:
        template_rpr = deepcopy(paragraph.runs[0]._r.rPr)

    for p in paragraphs:
        for run in list(p.runs):
            p._p.remove(run._r)
    for p in paragraphs[1:]:
        p._p.getparent().remove(p._p)

    if text:
        run = paragraph.add_run(text)
        if template_rpr is not None:
            run._r.insert(0, deepcopy(template_rpr))



def _merge_algorithm_and_ai(
    item: dict,
    aid: dict,
    *,
    session: Session,
    tr_type: TrType,
    voltage: str | None,
) -> tuple[str, str, str, float, str]:
    """Сливает БД и AI с жестким правилом источника.

    ** ставится ТОЛЬКО для значения, которое пришло от AI и не найдено
    среди применимых констант БД. Если значение есть в БД — маркер снимается.
    """
    ai_value = _clean_ai_value(aid.get("value", ""))
    try:
        confidence = max(0.0, min(1.0, float(aid.get("confidence", 0) or 0)))
    except (TypeError, ValueError):
        confidence = 0.0
    reason = str(aid.get("reason", "") or "").strip()
    evidence = aid.get("evidence", []) or []
    if not isinstance(evidence, list):
        evidence = [str(evidence)]

    valid_ai, validation_reason = _context_validation(item, ai_value, item.get("_all_items", []), document_context=None)
    if not valid_ai:
        ai_value = ""
        confidence = 0.0
        reason = (reason + " " if reason else "") + "AI-значение отклонено проверкой контекста: " + validation_reason
    elif ai_value and str(aid.get("source", "AI_CONTEXT")) == "AI_CONTEXT" and confidence >= AI_CONFIDENCE_THRESHOLD and not evidence:
        ai_value = ""
        confidence = 0.0
        reason = (reason + " " if reason else "") + "AI-значение отклонено: модель не указала контекстное основание (evidence)."

    alg_value = _clean_ai_value(item.get("algorithm_value", ""))
    alg_source = str(item.get("algorithm_source", "NONE") or "NONE")
    alg_is_db = _is_db_source(alg_source)

    # Значение, которое первым проходом подтверждено БД, неприкасаемо:
    # AI не должен подменять один DB-констант другим допустимым значением.
    if alg_value and alg_is_db:
        return alg_value, "DB", "", confidence, reason

    # Если AI самостоятельно вернул значение, которое уже есть в БД проекта,
    # маркер также запрещен.
    if ai_value and _is_known_db_value(session, item, ai_value, tr_type, voltage):
        return ai_value, "DB", "", confidence, reason

    if alg_value:
        if _values_equivalent(ai_value, alg_value):
            return alg_value, ("DB" if alg_is_db else alg_source), "", confidence, reason
        # Любое уверенное AI-решение имеет приоритет над heuristic/requirement/fallback.
        # Подтвержденная БД отсечена выше и сюда не попадает.
        if ai_value and confidence >= AI_CONFIDENCE_THRESHOLD:
            return ai_value, "AI_CONTEXT", _ai_mark_for_required(item.get("required_val", "")), confidence, reason
        return alg_value, ("DB" if alg_is_db else alg_source), "", confidence, reason

    # Пустое поле: сначала используем точное значение из самого документа,
    # затем вывод AI. Оба источника имеют приоритет над аварийным переносом requirement.
    if ai_value:
        source_name = str(aid.get("source", "AI_CONTEXT") or "AI_CONTEXT")
        if source_name == "SELF_CONTEXT":
            return ai_value, "SELF_CONTEXT", "**", confidence, reason
        return ai_value, "AI_CONTEXT", _ai_mark_for_required(item.get("required_val", "")), confidence, reason

    # Если AI не вернул строку совсем, а в левой колонке есть явное требование,
    # переносим его как запасной ответ и помечаем по обычному правилу. Это не
    # выдумка модели: значение целиком взято из поля "Требуемое значение".
    requirement_value = _fallback_from_requirement(item.get("required_val", ""))
    if requirement_value:
        req_mark = _ai_mark_for_required(item.get("required_val", ""))
        return requirement_value, "REQUIREMENT", req_mark, 1.0, "Значение взято непосредственно из поля 'Требуемое значение', поскольку AI не вернул ответ."

    return "", "NONE", "", confidence, reason


def _fallback_from_requirement(required_val: str) -> str:
    """Безопасный fallback: явное значение из колонки требований."""
    value = str(required_val or "").strip()
    if not value or re.fullmatch(r"\*+", value):
        return ""
    # Список вариантов сам по себе не является ответом.
    # Нельзя выбрать первый вариант наугад.
    if _extract_allowed_values(value):
        return ""
    low = value.lower()
    if "не требуется" in low:
        return "Не требуется"
    if re.search(r"\bтребуется\b", low):
        return "Да"
    if low in {"да", "нет"}:
        return value
    if low in {"обязательно", "обязателен", "обязательна", "обязательны"}:
        return "Да"
    if "не более" in low:
        m = re.search(r"[-+]?\d+(?:[.,]\d+)?", value)
        if m:
            return m.group(0).replace(",", ".")
    cleaned = re.sub(r"\s+", " ", value).strip()
    if len(cleaned) <= 2 and not re.search(r"\d", cleaned):
        return ""
    return cleaned


def _physical_row_cells(row):
    """Возвращает только уникальные физические Word-ячейки строки."""
    result = []
    seen = set()
    for cell in row.cells:
        tc_id = id(cell._tc)
        if tc_id in seen:
            continue
        seen.add(tc_id)
        result.append(cell)
    return result


def _resolve_target_cell(doc: Document, item: dict):
    """Находит ту же физическую ячейку, которая была определена при разборе."""
    table = doc.tables[int(item["table"])]
    row = table.rows[int(item["row"])]
    cells = _physical_row_cells(row)
    idx = item.get("target_cell_index")
    if idx is None or int(idx) < 0 or int(idx) >= len(cells):
        return None
    return cells[int(idx)]


def process_docx_requirements(docx_path: str, output_path: str, session: Session) -> str:
    if not Path(docx_path).exists():
        raise FileNotFoundError(docx_path)

    doc = Document(docx_path)
    items, extra = extract_doc_items(doc)
    tr_type, voltage = _determine_type(session, items, extra["document"])

    # ---------------------------------------------------------
    # 1. Сначала ищем готовый шаблон. Если он найден, он является
    #    единственным источником заполнения: пустые строки НЕ заполняются.
    # ---------------------------------------------------------
    data_kb = get_data_tenders_knowledge()
    document_text = " ".join(extra["document"].get("paragraphs", [])) + " " + " ".join(
        c for table in extra["document"].get("tables", []) for row in table for c in row.get("cells", [])
    )

    # Сначала создаём минимальные элементы без алгоритмического заполнения.
    template_items = [dict(item) for item in items]
    for item in template_items:
        item.setdefault("algorithm_value", "")
        item.setdefault("algorithm_source", "")

    template, template_score, template_matches = find_best_template(
        session,
        template_items,
        tr_type_id=tr_type.id,
        voltage=voltage,
    )
    template_filled = 0
    profile_filled = 0

    if template is not None:
        template_filled = apply_template(
            session,
            template_items,
            template,
            template_matches,
        )
        algorithm_items = template_items
        print(
            f"[KB TEMPLATE] найден шаблон id={template.id}, "
            f"coverage={template_score:.2f}, заполнено={template_filled}; "
            "режим TEMPLATE_ONLY: остальные пустые поля НЕ заполняются"
        )
        # Жёстко очищаем алгоритмические подсказки у строк, которые шаблон не заполнил.
        # Иначе последующий merge мог бы принять старое/служебное значение за источник заполнения.
        for _item in algorithm_items:
            if not _item.get("template_value"):
                _item["algorithm_value"] = ""
                _item["algorithm_source"] = ""
    else:
        # Шаблон не найден — запускаем обычный первичный конвейер.
        algorithm_items = _algorithm_fill(
            session, items, tr_type, voltage,
            data_kb=data_kb, document_text=document_text
        )

        profile_filled = apply_voltage_profile(session, algorithm_items, voltage)
        if profile_filled:
            print(f"[KB PROFILE] заполнено из профиля {voltage} кВ: {profile_filled}")
        print("[KB TEMPLATE] подходящий шаблон не найден")

    # Полный снимок документа после первого (детерминированного) прохода.
    filled_context = [
        {k: item.get(k) for k in ["id", "table", "row", "num", "param_name", "required_val", "algorithm_value", "algorithm_source"]}
        for item in algorithm_items
    ]

    for item in algorithm_items:
        item["_all_items"] = algorithm_items
        target_cell = _resolve_target_cell(doc, item)
        if target_cell is None:
            print(
                f"[DOCX] пропущена строка {item.get('id')}: "
                f"не удалось определить физическую ячейку ответа "
                f"table={item.get('table')} row={item.get('row')} "
                f"target_col={item.get('target_cell_index')}"
            )
            continue
        item["target_cell"] = target_cell
        if item.get("algorithm_value"):
            _set_cell_text(target_cell, str(item["algorithm_value"]), "")

    kb_context = data_kb.build_full_context(model=data_kb.detect_model(document_text) or None, voltage=voltage, include_images=True)
    self_profile = _build_self_context(algorithm_items, extra["document"])
    if template is not None:
        ai_targets, self_filled = [], []
        # В TEMPLATE_ONLY нельзя передавать незаполненные строки дальше в AI.
    else:
        ai_targets, self_filled = _prepare_ai_targets(algorithm_items, self_profile)
    doc_context = {
        **extra["document"],
        "algorithm_pass": filled_context,
        "detected": {"tr_type": tr_type.name, "voltage": voltage},
        "self_profile": _compact_ai_profile(self_profile),
        "template": {
            "id": template.id if template else None,
            "score": template_score if template else 0.0,
            "filled_rows": template_filled,
            "profile_filled_rows": profile_filled,
        },
    }

    if __import__("os").getenv("TENDER_SKIP_AI", "0") == "1" or not ai_targets:
        print("[GigaChat] AI-проход не требуется: TENDER_SKIP_AI=1 или все пропуски закрыты контекстом/БД")
        audit = {}
    else:
        with open_client() as client:
            # get_models()/verify_connection() — отдельный сетевой запрос.
            # В рабочем контуре его не выполняем: chat-запрос сам проверит авторизацию.
            if __import__("os").getenv("TENDER_VERIFY_GIGACHAT", "0") == "1":
                verify_connection(client)
            audit = _ai_review(
                client,
                algorithm_items,
                self_profile,
                kb_context,
                doc_context,
            )

    audit_rows = []
    marked = 0
    unresolved: list[str] = []
    for item in algorithm_items:
        aid = audit.get(item["id"], {}) or {}
        if item.get("pre_ai_value") and not aid.get("value"):
            aid = {
                "value": item["pre_ai_value"],
                "source": "SELF_CONTEXT",
                "confidence": 1.0,
                "evidence": item.get("pre_ai_evidence", []),
                "reason": "Значение однозначно перенесено из связанной строки текущего документа.",
            }
        evidence = aid.get("evidence", []) or []
        if not isinstance(evidence, list):
            evidence = [str(evidence)]
        value, source, mark, confidence, reason = _merge_algorithm_and_ai(
            item, aid, session=session, tr_type=tr_type, voltage=voltage
        )

        # Уже заполненный участником ответ по умолчанию не стираем.
        if item.get("current_value") and __import__("os").getenv("TENDER_OVERWRITE_EXISTING", "0") != "1":
            value = str(item["current_value"]).strip()
            source, mark = "EXISTING_ANSWER", ""

        # Если AI промолчал, используем явное требование из соседней колонки.
        # Для значений, выведенных не из БД, ставится ** (или * при уже отмеченном *).
        if not value and template is None:
            requirement_value = _fallback_from_requirement(item.get("required_val", ""))
            if requirement_value:
                value = requirement_value
                source, mark = "REQUIREMENT", _ai_mark_for_required(item.get("required_val", ""))

        if not value and not _is_header_row(item["param_name"], item["num"], item.get("required_val", "")):
            unresolved.append(item["id"])

        # Защита от повторной генерации: не дублируем **/***.
        value = _clean_ai_value(value)
        item["final_value"] = value
        item["final_mark"] = mark

        if mark:
            marked += 1
        _set_cell_text(item["target_cell"], value, mark)
        audit_rows.append(
            {
                "id": item["id"],
                "num": item["num"],
                "param_name": item["param_name"],
                "required_val": item["required_val"],
                "algorithm_value": item.get("algorithm_value", ""),
                "algorithm_source": item["algorithm_source"],
                "ai_value": _clean_ai_value(aid.get("value", "")),
                "ai_source": str(aid.get("source", "") or ""),
                "ai_mark": str(aid.get("mark", "") or ""),
                "final_value": value,
                "source": source,
                "mark": mark,
                "confidence": confidence,
                "reason": reason,
                "evidence": evidence,
                "decision": ("DB_PRIORITY" if source == "DB" else "AI_OVERRIDE" if source == "AI_CONTEXT" else "ALGORITHM_OR_FALLBACK"),
            }
        )

    out = Path(output_path)
    doc.save(out)

    # ВАЖНО: результат генерации НЕ обучает БЗ автоматически.
    # На этом этапе файл только создается. Обучение выполняется после того,
    # как инженер реально открыл/исправил/сохранил документ и tracker вызывает
    # capture_docx_file(). Иначе ошибочный AI-ответ сразу становится шаблоном.
    kb_saved = {"saved_rows": 0, "template_created": 0, "template_id": 0}

    audit_path = out.with_suffix(out.suffix + ".audit.json")
    audit_path.write_text(
        json.dumps(
            {
                "run_id": str(uuid.uuid4()),
                "input": str(docx_path),
                "output": str(out),
                "detected_type": tr_type.name,
                "voltage": voltage,
                "rows": audit_rows,
                "ai_marked_rows": marked,
                "unresolved_rows": unresolved,
                "self_context_filled_rows": len(self_filled),
                "ai_candidate_rows": len(ai_targets),
                "rules": {
                    "DATA_TENDERS_is_constant_source": True,
                    "DB_is_constant_source": False,
                    "AI_context_mark": "**",
                    "unknown_placeholders": "never",
                    "star_in_required_value_uses_single_ai_mark": True,
                    "star_does_not_block_context_fill": True,
                    "requirement_fallback_when_ai_missing": True,
                    "ai_confidence_threshold": AI_CONFIDENCE_THRESHOLD,
                    "context_first": True,
                    "ai_only_for_unresolved_blank_fields": True,
                    "template_knowledge_priority": True,
                    "ai_not_called_for_template_or_field_kb_values": True,
                    "knowledge_saved_only_after_engineer_final_save": True,
                    "gigachat_connection_check_disabled_by_default": True,
                },
            },
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )
    print(f"[OK] DOCX сохранен: {out}")
    print(
        f"[OK] Аудит сохранен: {audit_path} "
        f"(строк с mark='**': {marked}, без значения: {len(unresolved)})"
    )
    return str(out)
