"""Сбор и форматирование контекста для AI-этапа.

Модуль намеренно не знает, как записывать значения в DOCX/PDF. Его задача
только одна: подготовить понятный и воспроизводимый контекст для GigaChat.

Источники контекста:
1. ВЕСЬ текущий документ, включая все абзацы и все строки таблиц.
2. ВЕСЬ корпус директории ``data_tenders``.
3. Паспорт конкретного изделия для каждой цели.
4. Связанные параметры текущего документа.
"""
from __future__ import annotations

import json
import re
from dataclasses import dataclass
from typing import Any, Iterable

from services.data_tenders_knowledge import _norm as normalize
from services.data_tenders_knowledge import _param_key as parameter_key


@dataclass(frozen=True)
class DocumentContext:
    """Полный текст текущего файла в AI-friendly представлении."""

    text: str
    items: tuple[dict[str, Any], ...]


@dataclass(frozen=True)
class DataTendersContext:
    """Полный корпус ``data_tenders`` в двух представлениях.

    ``text`` — полный текстовый корпус для диагностики/совместимости.
    ``ai_text`` — полный структурированный индекс всей директории,
    оптимизированный для передачи GigaChat без потери уникальных значений.
    """

    text: str
    ai_text: str
    chunks: tuple[str, ...]
    ai_chunks: tuple[str, ...]
    record_count: int
    model_count: int
    file_count: int


@dataclass(frozen=True)
class TargetContext:
    """Контекст одной пустой строки, которую должен решить AI."""

    id: str
    num: str
    param_name: str
    required_val: str
    allowed_values: tuple[str, ...]
    product_context: dict[str, Any]
    related_facts: tuple[dict[str, Any], ...]
    data_candidates: tuple[str, ...]

    def as_dict(self) -> dict[str, Any]:
        return {
            "id": self.id,
            "num": self.num,
            "param_name": self.param_name,
            "required_val": self.required_val,
            "allowed_values": list(self.allowed_values),
            "product_context": self.product_context,
            "related_facts": list(self.related_facts),
            "data_candidates": list(self.data_candidates),
        }


class ContextBuilder:
    """Единая точка подготовки полного AI-контекста."""

    def build_current_document(
        self,
        document_context: dict[str, Any] | None,
        items: Iterable[dict[str, Any]],
    ) -> DocumentContext:
        dc = document_context or {}
        lines: list[str] = []

        lines.append("=== ТЕКУЩИЙ ФАЙЛ: ПОЛНЫЙ КОНТЕКСТ ===")
        lines.append("[АБЗАЦЫ]")
        for paragraph in dc.get("paragraphs", []) or []:
            text = self._clean(paragraph)
            if text:
                lines.append(text)

        lines.append("[ТАБЛИЦЫ]")
        for table_number, table in enumerate(dc.get("tables", []) or [], start=1):
            lines.append(f"--- ТАБЛИЦА {table_number} ---")
            for row in table or []:
                cells = row.get("cells", []) if isinstance(row, dict) else row
                values = [self._clean(cell) for cell in (cells or [])]
                values = [value for value in values if value]
                if values:
                    lines.append(" | ".join(values))

        # Весь текст таблиц уже находится выше и содержит все строки ТЗ.
        # Отдельный JSON-дамп всех 149 строк здесь не нужен и только раздувает
        # prompt; конкретные цели и их идентификаторы передаются отдельно.
        text = "\n".join(self._remove_adjacent_duplicates(lines))
        frozen_items = tuple(
            {
                key: item.get(key, "")
                for key in (
                    "id",
                    "table",
                    "row",
                    "num",
                    "param_name",
                    "required_val",
                    "current_value",
                    "algorithm_value",
                )
            }
            for item in items
        )
        return DocumentContext(text=text, items=frozen_items)

    def build_data_tenders(
        self,
        data_kb,
        *,
        chunk_chars: int = 26000,
    ) -> DataTendersContext:
        """Строит ПОЛНЫЙ текст корпуса ``data_tenders``.

        Внутри сохраняются:
        - все модели и их профили;
        - все исходные файлы и строки таблиц;
        - все закэшированные Vision-записи;
        - служебная статистика.

        Для API полный текст затем разбивается на последовательные чанки, но
        ни один исходный файл намеренно не исключается.
        """
        # Поддерживаем два корректных варианта входа:
        # 1) объект DataTendersKnowledge — строим полный контекст директории;
        # 2) уже подготовленная строка полного контекста — используем её без
        #    повторного обращения к data_tenders. Это устраняет ошибку
        #    "str object has no attribute build_full_context".
        if isinstance(data_kb, str):
            full_text = data_kb.strip()
            chunks = tuple(self._split_text(full_text, max(8000, int(chunk_chars))))
            ai_chunks = tuple(self._split_structured_index(full_text, max(8000, int(chunk_chars))))
            return DataTendersContext(
                text=full_text,
                ai_text=full_text,
                chunks=chunks,
                ai_chunks=ai_chunks,
                record_count=0,
                model_count=0,
                file_count=0,
            )

        if hasattr(data_kb, "build_full_context"):
            raw = data_kb.build_full_context(
                model=None,
                voltage=None,
                include_images=True,
                max_chars=10_000_000,
            )
        elif isinstance(data_kb, dict):
            # Допускаем raw-словарь, если контекст уже построен вызывающим кодом.
            raw = data_kb
        else:
            raise TypeError(
                "data_kb должен быть DataTendersKnowledge, dict или str с полным "
                f"контекстом data_tenders; получено: {type(data_kb).__name__}"
            )

        lines: list[str] = [
            "=== КОРПУС DATA_TENDERS: ПОЛНЫЙ КОНТЕКСТ ДИРЕКТОРИИ ===",
            f"ROOT: {raw.get('root', '')}",
            f"RECORD_COUNT: {raw.get('record_count', 0)}",
            f"MODEL_COUNT: {raw.get('model_count', 0)}",
            "",
            "=== ПАСПОРТА ВСЕХ МОДЕЛЕЙ ===",
        ]

        model_profiles = raw.get("model_profiles", {}) or {}
        for model in sorted(model_profiles, key=normalize):
            profile = model_profiles.get(model) or {}
            lines.append(f"[MODEL] {model}")
            lines.append(f"[VOLTAGE] {profile.get('voltage', '')}")
            parameters = profile.get("parameters", {}) or {}
            for key in sorted(parameters, key=normalize):
                values = parameters.get(key) or []
                cleaned_values = [self._clean(value) for value in values if self._clean(value)]
                if cleaned_values:
                    lines.append(f"PARAM: {key} => {' || '.join(cleaned_values)}")
            lines.append("")

        lines.append("=== ВСЕ ФАЙЛЫ DATA_TENDERS ===")
        for file_context in raw.get("files", []) or []:
            filename = self._clean(file_context.get("file", ""))
            lines.append(f"[FILE] {filename}")
            if file_context.get("error"):
                lines.append(f"ERROR: {self._clean(file_context['error'])}")
            for table in file_context.get("tables", []) or []:
                lines.append(f"[TABLE {table.get('table', '')}]")
                for row in table.get("rows", []) or []:
                    values = [self._clean(cell) for cell in row or []]
                    values = [value for value in values if value]
                    if values:
                        lines.append(" | ".join(values))
            lines.append("")

        images = raw.get("images", []) or []
        if images:
            lines.append("=== ВСЕ КЭШИРОВАННЫЕ VISION/OCR ДАННЫЕ ===")
            for image in images:
                lines.append(
                    json.dumps(
                        image,
                        ensure_ascii=False,
                        separators=(",", ":"),
                    )
                )

        full_text = "\n".join(self._remove_adjacent_duplicates(lines))

        # Полный структурированный AI-индекс строится из ВСЕХ записей data_tenders.
        # Это не выборка релевантных строк: поиск/фильтрация выполняется уже AI
        # с учётом модели, напряжения и других признаков текущего файла.
        if hasattr(data_kb, "build_ai_index"):
            ai_index = data_kb.build_ai_index(include_images=True)
        elif isinstance(raw, dict):
            ai_index = {
                "source": "data_tenders",
                "root": raw.get("root", ""),
                "record_count": raw.get("record_count", 0),
                "model_count": raw.get("model_count", 0),
                "file_count": len(raw.get("files", []) or []),
                "file_manifest": [{"file": f.get("file", ""), "models": []} for f in raw.get("files", []) or []],
                "models": raw.get("model_profiles", {}),
                "images": raw.get("images", []),
                "coverage": "Представление построено из полного доступного корпуса data_tenders.",
            }
        else:
            ai_index = {}
        ai_text = self._format_full_data_tenders_index(ai_index)

        chunks = tuple(self._split_text(full_text, max(8000, int(chunk_chars))))
        ai_chunks = tuple(self._split_structured_index(ai_text, max(8000, int(chunk_chars))))
        return DataTendersContext(
            text=full_text,
            ai_text=ai_text,
            chunks=chunks,
            ai_chunks=ai_chunks,
            record_count=int(raw.get("record_count", 0) or 0),
            model_count=int(raw.get("model_count", 0) or 0),
            file_count=len(raw.get("files", []) or []),
        )

    @staticmethod
    def _format_full_data_tenders_index(index: dict[str, Any]) -> str:
        """Делает полный, но компактный и читаемый корпус для GigaChat."""
        lines = [
            "=== ПОЛНЫЙ КОНТЕКСТ ДИРЕКТОРИИ DATA_TENDERS ===",
            f"record_count={index.get('record_count', 0)}; "
            f"model_count={index.get('model_count', 0)}; "
            f"file_count={index.get('file_count', 0)}",
            "",
            "Все значения ниже собраны из всех индексированных записей data_tenders. "
            "Одинаковые значения внутри одной модели и параметра объединены только для "
            "уменьшения объёма; уникальные значения и источники не удаляются.",
            "",
            "=== ФАЙЛЫ ДИРЕКТОРИИ ===",
        ]
        for file_info in index.get("file_manifest", []) or []:
            lines.append(
                f"FILE: {file_info.get('file', '')} | "
                f"models={', '.join(file_info.get('models', []) or [])}"
            )

        lines.append("")
        lines.append("=== ВСЕ МОДЕЛИ И ИХ ПАРАМЕТРЫ ===")
        for model, payload in (index.get("models", {}) or {}).items():
            lines.append(
                f"MODEL: {model} | voltage={payload.get('voltage') or ''} | "
                f"product_type={payload.get('product_type') or ''}"
            )
            for parameter in payload.get("parameters", []) or []:
                values = " || ".join(str(v) for v in parameter.get("values", []) if str(v).strip())
                sources = ", ".join(str(v) for v in parameter.get("source_files", []) if str(v).strip())
                if not values:
                    continue
                lines.append(
                    f"  PARAM: {parameter.get('parameter', '')} | "
                    f"VALUES: {values} | SOURCES: {sources}"
                )
            lines.append("")

        images = index.get("images", []) or []
        if images:
            lines.append("=== ВСЕ КЭШИРОВАННЫЕ VISION/OCR ЗАПИСИ ===")
            for image in images:
                lines.append(
                    f"IMAGE: {image.get('source_file', '')}/{image.get('image', '')} | "
                    f"models={','.join(image.get('models', []) or [])} | "
                    f"text={image.get('text', '')} | context={image.get('context', '')}"
                )

        return "\n".join(lines)

    def build_targets(
        self,
        pending_items: Iterable[dict[str, Any]],
    ) -> tuple[TargetContext, ...]:
        targets: list[TargetContext] = []
        for item in pending_items:
            product_context = dict(item.get("_product_context") or {})
            related = self._related_facts(item, pending_items)
            candidates = self._candidate_values(item)
            targets.append(
                TargetContext(
                    id=str(item.get("id", "")),
                    num=str(item.get("num", "")),
                    param_name=str(item.get("param_name", "")),
                    required_val=str(item.get("required_val", "")),
                    allowed_values=tuple(self._allowed_values(item)),
                    product_context=product_context,
                    related_facts=tuple(related),
                    data_candidates=tuple(candidates),
                )
            )
        return tuple(targets)

    @staticmethod
    def target_models(targets: Iterable[TargetContext]) -> list[str]:
        models: list[str] = []
        seen: set[str] = set()
        for target in targets:
            model = str(target.product_context.get("model") or "").strip()
            if model and normalize(model) not in seen:
                seen.add(normalize(model))
                models.append(model)
        return models

    @staticmethod
    def _clean(value: Any) -> str:
        return re.sub(r"\s+", " ", str(value or "")).strip()

    @staticmethod
    def _remove_adjacent_duplicates(lines: list[str]) -> list[str]:
        result: list[str] = []
        previous = None
        for line in lines:
            if line == previous:
                continue
            result.append(line)
            previous = line
        return result

    @classmethod
    def split_text(cls, text: str, chunk_chars: int) -> tuple[str, ...]:
        """Разбивает произвольный текст на чанки без потери строк."""
        return tuple(cls._split_text(text, max(1000, int(chunk_chars))))

    @staticmethod
    def _split_structured_index(text: str, chunk_chars: int) -> list[str]:
        """Делит AI-индекс по целым блокам моделей, не разрывая паспорт модели."""
        lines = text.splitlines()
        header: list[str] = []
        blocks: list[list[str]] = []
        current_block: list[str] = []

        for line in lines:
            if line.startswith("MODEL: "):
                if current_block:
                    blocks.append(current_block)
                current_block = [line]
            else:
                if current_block:
                    current_block.append(line)
                else:
                    header.append(line)
        if current_block:
            blocks.append(current_block)

        chunks: list[str] = []
        current = list(header)
        current_size = sum(len(x) + 1 for x in current)
        for block in blocks:
            block_size = sum(len(x) + 1 for x in block)
            if len(current) > len(header) and current_size + block_size > chunk_chars:
                chunks.append("\n".join(current).strip())
                current = list(header)
                current_size = sum(len(x) + 1 for x in current)
            current.extend(block)
            current_size += block_size
        if current:
            chunks.append("\n".join(current).strip())
        return [chunk for chunk in chunks if chunk]

    @staticmethod
    def _split_text(text: str, chunk_chars: int) -> list[str]:
        if len(text) <= chunk_chars:
            return [text]
        chunks: list[str] = []
        current: list[str] = []
        current_size = 0
        for line in text.splitlines(keepends=True):
            line_size = len(line)
            if current and current_size + line_size > chunk_chars:
                chunks.append("".join(current))
                current = []
                current_size = 0
            current.append(line)
            current_size += line_size
        if current:
            chunks.append("".join(current))
        return chunks

    def _related_facts(self, item: dict[str, Any], pending_items: Iterable[dict[str, Any]]) -> list[dict[str, Any]]:
        target_key = parameter_key(str(item.get("param_name", "")))
        result: list[tuple[float, dict[str, Any]]] = []
        for other in pending_items:
            if other.get("id") == item.get("id"):
                continue
            other_value = str(other.get("algorithm_value") or other.get("current_value") or "").strip()
            if not other_value:
                continue
            other_key = parameter_key(str(other.get("param_name", "")))
            score = self._text_similarity(target_key, other_key)
            if score >= 0.40:
                result.append(
                    (
                        score,
                        {
                            "id": other.get("id", ""),
                            "num": other.get("num", ""),
                            "param_name": other.get("param_name", ""),
                            "value": other_value,
                        },
                    )
                )
        result.sort(key=lambda pair: pair[0], reverse=True)
        return [fact for _, fact in result[:6]]

    @staticmethod
    def _candidate_values(item: dict[str, Any]) -> list[str]:
        values = []
        for value in item.get("_ai_data_candidates", []) or []:
            value = str(value).strip()
            if value and value not in values:
                values.append(value)
        return values[:12]

    @staticmethod
    def _allowed_values(item: dict[str, Any]) -> list[str]:
        values: list[str] = []
        text = f"{item.get('param_name', '')} {item.get('required_val', '')}"
        for match in re.finditer(r"\(([^()]*)\)", text):
            content = match.group(1).strip()
            if not content:
                continue
            # Скобки вида "(марка)" и "(по ГОСТ...)" — это уточнения,
            # а не перечень значений.
            if "," not in content and ";" not in content and " или " not in content.lower() and "/" not in content:
                continue
            parts = re.split(r"\s*[,;]\s*|\s+или\s+", content, flags=re.I)
            if len(parts) == 1 and "/" in content and not re.search(r"\b(?:в|кв|ка|а|мм|см|м|гц|кн)\b", content, re.I):
                parts = [part.strip() for part in content.split("/")]
            for part in parts:
                part = part.strip()
                if part and normalize(part) not in {normalize(existing) for existing in values}:
                    values.append(part)
        return values

    @staticmethod
    def _text_similarity(left: str, right: str) -> float:
        left_tokens = set(re.findall(r"[а-яa-z0-9]+", normalize(left)))
        right_tokens = set(re.findall(r"[а-яa-z0-9]+", normalize(right)))
        if not left_tokens or not right_tokens:
            return 0.0
        overlap = len(left_tokens & right_tokens) / max(1, min(len(left_tokens), len(right_tokens)))
        return overlap
