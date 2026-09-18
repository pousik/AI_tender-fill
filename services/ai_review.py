"""Контекстный AI-review тендерного документа.

Архитектура:

    текущий файл целиком
            +
    весь data_tenders по частям
            ↓
    кандидаты характеристик
            ↓
    финальный AI-review
            ↓
    item_id -> value

Ограничение размера prompt решается адаптивным разбиением только корпуса
``data_tenders``. Сам текущий документ всегда передаётся целиком.
"""
from __future__ import annotations

import hashlib
import json
import os
import time
from typing import Any

from services.ai_context import ContextBuilder, DocumentContext, TargetContext
from services.data_tenders_knowledge import get_data_tenders_knowledge, _norm as normalize
from services.gigachat_client import AI_CONFIDENCE_THRESHOLD, ask_json, ask_json_schema


_RESPONSE_CACHE: dict[str, dict[str, dict[str, Any]]] = {}


class AIReviewService:
    """Оркестратор AI-этапа без логики записи DOCX/PDF."""

    def __init__(self, client) -> None:
        self.client = client
        self.context_builder = ContextBuilder()
        self.safe_limit = self._safe_prompt_limit()

    @staticmethod
    def _needs_dimension_drawings(targets: tuple[TargetContext, ...]) -> bool:
        """Нужны ли чертежи из 1БП.769.001 РЭ для текущих целей."""
        for target in targets:
            text = normalize(f"{target.param_name} {target.required_val}").lower()
            if any(token in text for token in ("габарит", "размер", "масса")):
                return True
        return False

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------
    def review(
        self,
        items: list[dict[str, Any]],
        self_profile: dict[str, Any],
        data_kb=None,
        document_context: dict[str, Any] | None = None,
    ) -> dict[str, dict[str, Any]]:
        """Заполняет пустые поля с полным контекстом текущего файла и data_tenders."""
        pending = [
            item
            for item in items
            if self._is_blank(item)
            and item.get("pre_ai_source") != "SKIP_NON_FIELD"
            and not item.get("pre_ai_value")
        ]
        if not pending:
            return {}

        data_kb = data_kb or get_data_tenders_knowledge()
        current = self.context_builder.build_current_document(document_context, items)
        targets = self.context_builder.build_targets(pending)
        scan_focus = self._scan_target_focus(targets)
        scan_chunk_size = self._scan_chunk_size(len(current.text), len(scan_focus))
        data_context = self.context_builder.build_data_tenders(
            data_kb,
            chunk_chars=scan_chunk_size,
        )
        detected = (document_context or {}).get("detected", {})
        compact_profile = self._compact_self_profile(self_profile, max_facts=40)

        # Чертежи читаются лениво: только когда среди целей есть габариты/масса.
        # Обрабатываем каждый рисунок А.1-А.6 из 1БП.769.001 РЭ один раз и кэшируем.
        drawing_context = ""
        if self._needs_dimension_drawings(targets):
            try:
                drawing_data = data_kb.read_reference_dimension_drawings(client=self.client)
                drawing_context = json.dumps(
                    drawing_data, ensure_ascii=False, separators=(",", ":")
                )
                print(
                    f"[GigaChat][DRAWINGS] контекст габаритов подготовлен: "
                    f"figures={len(drawing_data.get('figures', []) or [])}",
                    flush=True,
                )
            except Exception as exc:
                print(f"[GigaChat][DRAWINGS] не удалось прочитать чертежи: {exc}", flush=True)

        print(
            f"[GigaChat] полный контекст подготовлен: "
            f"current_file={len(current.text)} chars, "
            f"data_tenders_raw={len(data_context.text)} chars, "
            f"data_tenders_index={len(data_context.ai_text)} chars, "
            f"records={data_context.record_count}, models={data_context.model_count}, "
            f"files={data_context.file_count}, targets={len(targets)}, "
            f"data_chunks={len(data_context.ai_chunks)}, scan_chunk={scan_chunk_size}",
            flush=True,
        )

        # Если весь контекст помещается, выполняем один запрос.
        direct_prompt = self._build_final_prompt(
            current=current,
            knowledge=data_context.ai_text,
            targets=targets,
            self_profile=compact_profile,
            detected=detected,
            drawing_context=drawing_context,
            mode="FULL_CURRENT_FILE_PLUS_FULL_DATA_TENDERS_INDEX",
        )
        if len(direct_prompt) <= self.safe_limit:
            print(
                f"[GigaChat] полный контекст помещается в один запрос: {len(direct_prompt)} chars",
                flush=True,
            )
            return self._call_review(direct_prompt, pending)

        # Иначе ВСЕ части data_tenders проходят через AI. Каждый проход видит
        # полный текущий файл, поэтому модель может связывать сведения из каталога
        # с реальным ТЗ, даже если каталог разбит на несколько HTTP-запросов.
        candidates = self._scan_all_data_tenders_chunks(
            current=current,
            data_chunks=data_context.ai_chunks,
            targets=targets,
        )

        batches = self._make_target_batches(
            current=current,
            knowledge="",
            candidates=candidates,
            targets=targets,
            self_profile=compact_profile,
            detected=detected,
            drawing_context=drawing_context,
        )
        print(
            f"[GigaChat] финальный AI-review: targets={len(targets)} "
            f"batches={len(batches)} scan_candidates={len(candidates)}",
            flush=True,
        )

        result: dict[str, dict[str, Any]] = {}
        for number, batch in enumerate(batches, start=1):
            batch_candidates = self._format_candidates_for_targets(candidates, batch)
            prompt = self._build_final_prompt(
                current=current,
                knowledge=batch_candidates,
                targets=batch,
                self_profile=compact_profile,
                detected=detected,
                mode="FULL_CURRENT_FILE_PLUS_FULL_DATA_TENDERS_SCAN",
            )
            try:
                part = self._call_review(
                    prompt,
                    [self._find_item(items, target.id) for target in batch],
                )
                result.update(part)
                print(
                    f"[GigaChat] final batch {number}/{len(batches)}: "
                    f"targets={len(batch)} answered={len(part)}",
                    flush=True,
                )
            except Exception as exc:
                print(
                    f"[GigaChat] final batch {number}/{len(batches)} не выполнен: {exc}",
                    flush=True,
                )

        return result

    # ------------------------------------------------------------------
    # Full data_tenders scan
    # ------------------------------------------------------------------
    def _scan_all_data_tenders_chunks(
        self,
        *,
        current: DocumentContext,
        data_chunks: tuple[str, ...],
        targets: tuple[TargetContext, ...],
    ) -> list[dict[str, Any]]:
        """Просматривает все части data_tenders, сохраняя кандидатов по целям."""
        if not data_chunks:
            return []

        focus = self._scan_target_focus(targets)
        result: list[dict[str, Any]] = []
        print(
            f"[GigaChat][DATA-SCAN] старт полного просмотра data_tenders: "
            f"chunks={len(data_chunks)} targets={len(targets)}",
            flush=True,
        )

        for number, chunk in enumerate(data_chunks, start=1):
            prompt = self._build_target_scan_prompt(
                current=current.text,
                data_chunk=chunk,
                target_focus=focus,
                chunk_number=number,
                total_chunks=len(data_chunks),
            )
            try:
                t0 = time.perf_counter()
                raw = self._call_schema(
                    prompt,
                    self._candidate_schema(),
                    partial_array_key="candidates",
                )
                found = raw.get("candidates", []) if isinstance(raw, dict) else []
                if not isinstance(found, list):
                    found = []

                valid_ids = {target.id for target in targets}
                added = 0
                for candidate in found:
                    if not isinstance(candidate, dict):
                        continue
                    if str(candidate.get("id", "")) not in valid_ids:
                        continue
                    value = self._clean_ai_value(candidate.get("value", ""))
                    if not value:
                        continue
                    candidate["value"] = value
                    result.append(candidate)
                    added += 1

                print(
                    f"[GigaChat][DATA-SCAN] chunk {number}/{len(data_chunks)}: "
                    f"candidates={added} time={time.perf_counter() - t0:.2f}s",
                    flush=True,
                )
            except Exception as exc:
                print(
                    f"[GigaChat][DATA-SCAN] chunk {number}/{len(data_chunks)} не выполнен: {exc}",
                    flush=True,
                )

        result = self._deduplicate_candidates(result)
        print(
            f"[GigaChat][DATA-SCAN] полный просмотр завершён: candidates={len(result)}",
            flush=True,
        )
        return result

    @staticmethod
    def _scan_target_focus(targets: tuple[TargetContext, ...]) -> str:
        lines = []
        for target in targets:
            context = target.product_context or {}
            lines.append(
                f"{target.id} | parameter={target.param_name} | "
                f"model={context.get('model', '')} | voltage={context.get('voltage', '')} | "
                f"type={context.get('product_kind', '')} | required={target.required_val}"
            )
        return "\n".join(lines)

    def _scan_chunk_size(self, current_chars: int, focus_chars: int) -> int:
        """Размер части каталога так, чтобы scan prompt гарантированно помещался."""
        available = self.safe_limit - current_chars - focus_chars - 10000
        return max(7000, min(14000, available))

    def _build_target_scan_prompt(
        self,
        *,
        current: str,
        data_chunk: str,
        target_focus: str,
        chunk_number: int,
        total_chunks: int,
    ) -> str:
        return f"""Ты анализируешь технический тендерный документ и справочный каталог data_tenders.

Это проход {chunk_number}/{total_chunks} по ПОЛНОМУ каталогу data_tenders.
ПОЛНЫЙ текущий документ передан ниже целиком.

Цель этого прохода — только НАЙТИ подтверждённые варианты характеристик.
Не выбирай случайное значение и не переноси характеристику между моделями.
Если найдено несколько значений одной характеристики, сохрани их отдельными
candidates. Возвращай только значения, которые реально есть в этой части каталога.
Не добавляй пояснений внутри JSON. Максимум 24 кандидата за один проход.

Особенно ищи характеристики, которые зависят от модели:
масса, габариты, номинальные токи, напряжения, путь утечки,
изоляция, климатическое исполнение и механические нагрузки.

ЦЕЛИ:
{target_focus}

ПОЛНЫЙ ТЕКУЩИЙ ДОКУМЕНТ:
{current}

ЧАСТЬ DATA_TENDERS {chunk_number}/{total_chunks}:
{data_chunk}

Верни только JSON по схеме candidates. Не добавляй поле evidence или длинные пояснения.
source_file достаточно для идентификации источника. Не возвращай значения, которых
нет в данном фрагменте.
"""

    @staticmethod
    def _candidate_schema() -> dict[str, Any]:
        """Компактная схема DATA-SCAN.

        DATA-SCAN не должен генерировать длинные доказательства: его задача —
        собрать кандидатов. Подробное обоснование выполняется на финальном AI-этапе.
        Чем короче ответ, тем меньше риск усечения structured output.
        """
        candidate = {
            "type": "object",
            "properties": {
                "id": {"type": "string"},
                "model": {"type": "string"},
                "parameter": {"type": "string"},
                "value": {"type": "string"},
                "unit": {"type": "string"},
                "source_file": {"type": "string"},
            },
            "required": ["id", "model", "parameter", "value", "unit", "source_file"],
            "additionalProperties": False,
        }
        return {
            "type": "object",
            "properties": {
                "candidates": {"type": "array", "items": candidate, "maxItems": 24},
            },
            "required": ["candidates"],
            "additionalProperties": False,
        }

    @staticmethod
    def _deduplicate_candidates(candidates: list[dict[str, Any]]) -> list[dict[str, Any]]:
        result = []
        seen: set[tuple[str, str, str, str, str]] = set()
        for candidate in candidates:
            key = (
                str(candidate.get("id", "")),
                normalize(str(candidate.get("model", ""))),
                normalize(str(candidate.get("parameter", ""))),
                normalize(str(candidate.get("value", ""))),
                normalize(str(candidate.get("source_file", ""))),
            )
            if key in seen:
                continue
            seen.add(key)
            result.append(candidate)
        return result

    # ------------------------------------------------------------------
    # Final answer stage
    # ------------------------------------------------------------------
    def _make_target_batches(
        self,
        *,
        current: DocumentContext,
        knowledge: str,
        candidates: list[dict[str, Any]],
        targets: tuple[TargetContext, ...],
        self_profile: dict[str, Any],
        detected: dict[str, Any],
        drawing_context: str = "",
    ) -> list[tuple[TargetContext, ...]]:
        del knowledge
        batches: list[tuple[TargetContext, ...]] = []
        start = 0
        while start < len(targets):
            end = min(len(targets), start + 12)
            best: tuple[TargetContext, ...] | None = None
            while end > start:
                batch = tuple(targets[start:end])
                batch_knowledge = self._format_candidates_for_targets(candidates, batch)
                prompt = self._build_final_prompt(
                    current=current,
                    knowledge=batch_knowledge,
                    targets=batch,
                    self_profile=self_profile,
                    detected=detected,
                    drawing_context=drawing_context,
                    mode="FULL_CURRENT_FILE_PLUS_FULL_DATA_TENDERS_SCAN",
                )
                if len(prompt) <= self.safe_limit:
                    best = batch
                    break
                end = start + max(1, (end - start) // 2)
            if best is None:
                raise ValueError("Даже одно AI-поле не помещается в безопасный размер prompt")
            batches.append(best)
            start += len(best)
        return batches

    @staticmethod
    def _format_candidates_for_targets(
        candidates: list[dict[str, Any]],
        targets: tuple[TargetContext, ...],
        max_chars: int = 14000,
    ) -> str:
        target_ids = {target.id for target in targets}
        lines = [
            "=== КАНДИДАТЫ ИЗ ПОЛНОГО ПРОСМОТРА DATA_TENDERS ===",
            "Эти кандидаты собраны после просмотра ВСЕХ частей директории. "
            "Финальный выбор должен учитывать текущий документ и изделие.",
        ]
        total = sum(len(x) + 1 for x in lines)
        for candidate in candidates:
            if str(candidate.get("id", "")) not in target_ids:
                continue
            line = " | ".join(
                (
                    f"id={candidate.get('id', '')}",
                    f"model={candidate.get('model', '')}",
                    f"parameter={candidate.get('parameter', '')}",
                    f"value={candidate.get('value', '')}",
                    f"unit={candidate.get('unit', '')}",
                    f"source={candidate.get('source_file', '')}",
                    f"confidence={candidate.get('confidence', '')}",
                    f"evidence={candidate.get('evidence', '')}",
                )
            )
            if total + len(line) + 1 > max_chars:
                break
            lines.append(line)
            total += len(line) + 1
        if len(lines) == 2:
            lines.append("Подтверждённых кандидатов в data_tenders для этих целей не найдено.")
        return "\n".join(lines)

    def _build_final_prompt(
        self,
        *,
        current: DocumentContext,
        knowledge: str,
        targets: tuple[TargetContext, ...],
        self_profile: dict[str, Any],
        detected: dict[str, Any],
        drawing_context: str = "",
        mode: str = "",
    ) -> str:
        target_json = json.dumps(
            [target.as_dict() for target in targets],
            ensure_ascii=False,
            separators=(",", ":"),
        )
        profile_json = json.dumps(self_profile, ensure_ascii=False, separators=(",", ":"))
        detected_json = json.dumps(detected, ensure_ascii=False, separators=(",", ":"))

        return f"""Ты выполняешь второй этап автозаполнения технического тендерного документа.

ГЛАВНОЕ ПРАВИЛО:
Сначала изучи ВЕСЬ текущий документ. Для каждой цели определи конкретное изделие:
раздел -> вид изделия -> марка/модель -> напряжение -> исполнение.
Только после этого сопоставляй характеристики data_tenders.

КОРПУС DATA_TENDERS:
Каталог был полностью просмотрен AI по всем его частям. Поэтому кандидаты ниже
являются результатом полного просмотра директории, а не случайной выборкой.
Режим: {mode}.

НЕ ДЕЛАЙ:
- не копируй значение только по похожему названию параметра;
- не переносить массу, габариты, токи, путь утечки и другие характеристики между моделями;
- не использовать модель соседнего раздела;
- не выбирать случайный вариант из нескольких значений;
- не выдумывать значение без подтверждения.

ДЕЛАЙ:
- учитывай весь текущий файл как единый технический контекст;
- определяй изделие до определения его характеристики;
- используй модель, напряжение, изготовителя, изоляцию, климат и другие признаки;
- используй все найденные кандидаты data_tenders и выбирай только подтверждённые для изделия;
- учитывай требования целевого поля и связанные строки текущего документа;
- выполняй преобразование единиц только если оно однозначно;
- если есть несколько вариантов, связывай их с моделью/исполнением, а не выбирай случайно;
- если подтверждения нет, возвращай пустое value.

ИСТОЧНИК ОТВЕТА:
DATA_TENDERS — значение подтверждено каталогом;
SELF_CONTEXT — значение однозначно следует из текущего документа;
AI_CONTEXT — значение получено из совокупного контекста; evidence обязательно;
NONE — значение не определено.
Не возвращай '*' или '**'.

ПОЛНЫЙ ТЕКУЩИЙ ДОКУМЕНТ:
{current.text}

КАНДИДАТЫ ИЗ ПОЛНОГО КОРПУСА DATA_TENDERS:
{knowledge}

СТРУКТУРИРОВАННЫЙ ПРОФИЛЬ ТЕКУЩЕГО ФАЙЛА:
{detected_json}

ФАКТЫ, УЖЕ ИЗВЛЕЧЁННЫЕ ИЗ ТЕКУЩЕГО ДОКУМЕНТА:
{profile_json}


REFERENCE DRAWINGS — 1БП.769.001 РЭ:
{drawing_context}

Если этот блок не пуст и цель относится к габаритам/размерам/массе, обязательно
анализируй рисунки А.1-А.6 как первичный визуальный источник. Не используй номера
позиций деталей как размеры. Связывай размеры с размерными линиями, видом и моделью.
Если на разных рисунках разные модели — выбирай только рисунок, соответствующий изделию.

ЦЕЛИ:
{target_json}

Для каждой цели верни итоговое значение именно для соответствующего изделия.
Если значение неизвестно — пустая строка. Evidence должна кратко объяснять,
на каком параметре, модели и/или файле основан выбор.
"""

    # ------------------------------------------------------------------
    # Profile helpers
    # ------------------------------------------------------------------
    @staticmethod
    def _compact_self_profile(self_profile: dict[str, Any], *, max_facts: int) -> dict[str, Any]:
        """Убирает дублирование: полный текст текущего файла уже передан отдельно."""
        facts = list(self_profile.get("document_facts", []) or [])
        global_facts = list(self_profile.get("global_facts", []) or [])
        selected: list[dict[str, Any]] = []
        seen: set[str] = set()
        for fact in global_facts + facts:
            fid = str(fact.get("id", ""))
            if fid in seen:
                continue
            seen.add(fid)
            selected.append(fact)
            if len(selected) >= max_facts:
                break
        return {
            "key_paragraphs": list(self_profile.get("key_paragraphs", []) or [])[:10],
            "document_facts": selected,
        }

    # ------------------------------------------------------------------
    # GigaChat calls
    # ------------------------------------------------------------------
    def _call_review(
        self,
        prompt: str,
        pending_items: list[dict[str, Any]],
    ) -> dict[str, dict[str, Any]]:
        cache_key = self._cache_key(prompt)
        cached = _RESPONSE_CACHE.get(cache_key)
        if cached is not None:
            return dict(cached)
        if len(prompt) > self.safe_limit:
            raise ValueError(f"AI prompt too large: {len(prompt)} > {self.safe_limit}")

        t0 = time.perf_counter()
        raw = self._call_with_backoff(prompt)
        if not isinstance(raw, dict):
            raise RuntimeError("GigaChat вернул неверный формат AI-review")

        valid_ids = {str(item.get("id", "")) for item in pending_items}
        cleaned: dict[str, dict[str, Any]] = {}
        for key, value in raw.items():
            if key not in valid_ids or not isinstance(value, dict):
                continue
            result = dict(value)
            result["value"] = self._clean_ai_value(result.get("value", ""))
            cleaned[key] = result

        _RESPONSE_CACHE[cache_key] = cleaned
        print(
            f"[GigaChat] AI-review получен: requested={len(valid_ids)} answered={len(cleaned)} "
            f"time={time.perf_counter() - t0:.2f}s",
            flush=True,
        )
        return cleaned

    def _call_schema(
        self,
        prompt: str,
        schema: dict[str, Any],
        *,
        partial_array_key: str | None = None,
    ) -> dict[str, Any]:
        if len(prompt) > self.safe_limit:
            raise ValueError(f"AI prompt too large: {len(prompt)} > {self.safe_limit}")
        retries = self._retries()
        base_delay = self._base_delay()
        for attempt in range(retries + 1):
            try:
                return ask_json_schema(
                    self.client,
                    prompt,
                    schema,
                    partial_array_key=partial_array_key,
                )
            except Exception as exc:
                if not self._is_rate_limit(exc) or attempt >= retries:
                    raise
                delay = base_delay * (2**attempt)
                print(f"[GigaChat] 429: повтор через {delay:.1f} c", flush=True)
                time.sleep(delay)
        raise RuntimeError("AI schema request failed")

    def _call_with_backoff(self, prompt: str) -> dict[str, Any]:
        retries = self._retries()
        base_delay = self._base_delay()
        for attempt in range(retries + 1):
            try:
                return ask_json(self.client, prompt)
            except Exception as exc:
                if not self._is_rate_limit(exc) or attempt >= retries:
                    raise
                delay = base_delay * (2**attempt)
                print(f"[GigaChat] 429: повтор через {delay:.1f} c", flush=True)
                time.sleep(delay)
        raise RuntimeError("AI request failed")

    # ------------------------------------------------------------------
    # Generic helpers
    # ------------------------------------------------------------------
    @staticmethod
    def _is_blank(item: dict[str, Any]) -> bool:
        value = str(item.get("current_value", "") or "").strip()
        return not value or value.strip("*") == ""

    @staticmethod
    def _find_item(items: list[dict[str, Any]], target_id: str) -> dict[str, Any]:
        for item in items:
            if str(item.get("id", "")) == str(target_id):
                return item
        return {"id": target_id}

    @staticmethod
    def _clean_ai_value(value: Any) -> str:
        text = str(value or "").strip()
        while text.endswith("**"):
            text = text[:-2].rstrip()
        return text

    @staticmethod
    def _cache_key(prompt: str) -> str:
        return hashlib.sha256(prompt.encode("utf-8")).hexdigest()

    @staticmethod
    def _safe_prompt_limit() -> int:
        try:
            configured = int(os.getenv("TENDER_AI_MAX_PROMPT_CHARS", "60000"))
        except ValueError:
            configured = 60000
        return max(20000, min(configured, 58000))

    @staticmethod
    def _retries() -> int:
        try:
            return max(0, int(os.getenv("TENDER_AI_429_RETRIES", "3")))
        except ValueError:
            return 3

    @staticmethod
    def _base_delay() -> float:
        try:
            return max(0.5, float(os.getenv("TENDER_AI_429_BASE_DELAY", "2")))
        except ValueError:
            return 2.0

    @staticmethod
    def _is_rate_limit(exc: Exception) -> bool:
        text = str(exc).lower()
        return "429" in text or "too many requests" in text


__all__ = ["AIReviewService"]
