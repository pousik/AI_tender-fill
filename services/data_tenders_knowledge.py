"""Структурированный файловый источник истины по data_tenders.

Первичный проход заполнения НЕ использует transformers_legacy для констант.
Эталонные DOCX из data_tenders разбираются на:
  - модель/типоисполнение;
  - напряжение;
  - параметр -> значение;
  - размерные/массовые матрицы;
  - климатические и изоляционные исполнения;
  - изображения/чертежи с OCR;
  - связь изображения с ближайшей маркой/типоисполнением.

Далее эта структура целиком попадает в AI-контекст. AI используется вторым
этапом для сопоставления требований и тех характеристик, которые нельзя
однозначно вывести детерминированно.
"""
from __future__ import annotations

import hashlib
import io
from io import BytesIO
import json
import os
import re
import subprocess
import tempfile
import zipfile
from collections import defaultdict
from dataclasses import dataclass, asdict
from functools import lru_cache
from difflib import SequenceMatcher
from pathlib import Path
from typing import Any

from docx import Document
from docx.oxml.ns import qn

try:
    import pytesseract
    from PIL import Image, ImageOps, ImageFilter
except Exception:
    pytesseract = None
    Image = ImageOps = ImageFilter = None

from services.gigachat_client import ask_vision_json, open_client

_MODEL_RE = re.compile(
    r"трг\s*[-‑–—]?\s*уэтм\s*(?:®|\(r\))?\s*[-‑–—]?\s*(\d{1,4})",
    re.IGNORECASE,
)
_VOLT_RE = re.compile(r"(?<!\d)(\d{1,4}(?:[.,]\d+)?)\s*(?:кв|kv)\b", re.IGNORECASE)
_NUM_RE = re.compile(r"(?<!\d)(\d+(?:[.,]\d+)?)(?!\d)")

_PRODUCT_MODEL_RE = re.compile(
    r"(?:ТРГ\s*[-‑–—]?\s*УЭТМ\s*[-‑–—]?\s*(\d{1,4})|[А-ЯЁA-Z0-9]{2,}(?:[-_][А-ЯЁA-Z0-9]+)+[-_]\d{1,4})",
    re.IGNORECASE,
)
_INSULATION_TYPE_RE = re.compile(r"\b(фарфор(?:овый|овая|овое|овые)?|полимер(?:ный|ная|ное|ные)?)\b", re.IGNORECASE)
_INSULATION_COLOR_RE = re.compile(r"\b(бел(?:ый|ая|ое|ые)|коричнев(?:ый|ая|ое|ые)|сер(?:ый|ая|ое|ые)|черн(?:ый|ая|ое|ые)|черный|серый|белый|коричневый)\b", re.IGNORECASE)
_TYPE_HINTS = ("тип внешней изоляции", "тип изоляции", "марка", "заводской тип", "исполнение")
_MANUFACTURER_RE = re.compile(
    r'''ООО\s*[«\"]?Эльмаш\s*\(\s*УЭТМ\s*\)[»\"]?''',
    re.IGNORECASE,
)


_ALIAS_GROUPS = (
    ("номинальное напряжение", "ном напряжение", "номинальное напряжение uн"),
    ("наибольшее рабочее напряжение", "максимальное рабочее напряжение"),
    ("количество вторичных обмоток", "число вторичных обмоток"),
    ("количество ответвлений от вторичных обмоток", "число ответвлений от вторичных обмоток"),
    ("климатическое исполнение и категория размещения", "климатическое исполнение"),
    ("срок службы трансформатора", "срок службы трансформатора"),
    ("срок периодической поверки", "периодичность поверки", "срок поверки"),
    ("расширенный диапазон рабочих частот", "диапазон рабочих частот"),
    ("номинальный первичный ток", "номинальный ток первичной обмотки"),
    ("номинальный вторичный ток", "номинальный ток вторичной обмотки"),
    ("номинальная частота", "частота"),
    ("номинальная вторичная нагрузка", "вторичная нагрузка"),
    ("коэффициент мощности", "cosφ2", "cos2"),
    ("уровень шума при работе", "уровень шума"),
    ("габаритные размеры", "габаритные размеры трансформатора"),
    ("масса трансформатора тока", "масса трансформатора", "масса"),
    ("изготовитель", "производитель", "предприятие-изготовитель", "предприятие изготовитель", "изготовитель оборудования"),
)
_ALIAS_TO_CANON = {alias: group[0] for group in _ALIAS_GROUPS for alias in group}


@lru_cache(maxsize=32768)
def _clean_text(text: str) -> str:
    text = str(text or "").replace("\u00a0", " ")
    text = re.sub(r"\s+", " ", text).strip()
    return text


@lru_cache(maxsize=32768)
def _norm(text: str) -> str:
    text = _clean_text(text).lower().replace("ё", "е").replace("®", "")
    text = re.sub(r"[‐‑‒–—−]", "-", text)
    return re.sub(r"\s+", " ", text).strip()


@lru_cache(maxsize=32768)
def _param_key(text: str) -> str:
    raw = _clean_text(text)
    low = _norm(raw)
    # Специальные параметры должны проверяться ДО общего слова "масса".
    if "масса масла" in low or "масса элегаза" in low or "масса изолирующего газа" in low:
        return "масса масла"
    if "масса" in low and "трансформатор" in low:
        return "масса трансформатора"
    if re.search(r"\bмасса\b", low):
        return "масса трансформатора"
    if "габарит" in low and ("размер" in low or "высота" in low or "диаметр" in low):
        return "габаритные размеры"

    text = low
    text = re.sub(r"^\d+(?:\.\d+)*[.)]?\s*", "", text)
    text = re.sub(r"\([^)]*\)", " ", text)
    text = re.sub(r"\bпо\s+гост[^,;]*", " ", text, flags=re.IGNORECASE)
    text = re.sub(r"\b(?:кв|кв\.|мм|см/кв|шт\.?|лет|год(?:а|ы)?|м/с|мпа|дба|гц)\b", " ", text)
    text = re.sub(r"\b(?:не\s+менее|не\s+более)\b", " ", text)
    text = re.sub(r"\s+", " ", text).strip(" ;:,")
    for alias, canonical in sorted(_ALIAS_TO_CANON.items(), key=lambda x: len(x[0]), reverse=True):
        if alias in text:
            # Масса масла уже обработана выше.
            return canonical
    return text


@lru_cache(maxsize=8192)
def _model_name(text: str) -> str | None:
    m = _MODEL_RE.search(_norm(text))
    return f"ТРГ-УЭТМ-{int(m.group(1))}" if m else None


@lru_cache(maxsize=8192)
def _all_models(text: str) -> list[str]:
    return list(dict.fromkeys(f"ТРГ-УЭТМ-{int(x)}" for x in _MODEL_RE.findall(_norm(text))))


def _model_voltage(model: str | None) -> str | None:
    if not model:
        return None
    m = re.search(r"(\d+)$", model)
    return m.group(1) if m else None


def _same_voltage(a: str | None, b: str | None) -> bool:
    if not a or not b:
        return False
    try:
        return float(str(a).replace(",", ".")) == float(str(b).replace(",", "."))
    except (ValueError, TypeError):
        return _norm(a) == _norm(b)


def _split_numeric_values(value: str) -> list[str]:
    return [x.replace(".", ".") for x in _NUM_RE.findall(value or "")]


def _split_options(text: str) -> list[str]:
    """Разбирает инженерные списки вариантов, не разрушая значения в скобках."""
    s = _clean_text(text)
    # Убираем только пояснительные скобки, сохраняем варианты внутри них.
    out: list[str] = []
    for part in re.split(r"\s*;\s*|\s+или\s+|\s*[/|]\s*|\s*,\s*", s, flags=re.IGNORECASE):
        part = part.strip(" .")
        if part and not part.isdigit() and not part.lower().startswith("примеч"):
            out.append(part)
    return list(dict.fromkeys(out))


def _extract_product_signature(text: str) -> dict[str, str | None]:
    """Извлекает признаки исполнения один раз из текста документа."""
    text = _clean_text(text)
    model = _model_name(text)
    voltage = _model_voltage(model) if model else None
    if not voltage:
        m = _VOLT_RE.search(text)
        voltage = m.group(1).replace(",", ".") if m else None
    ins = _INSULATION_TYPE_RE.search(text)
    color = _INSULATION_COLOR_RE.search(text)
    return {
        "model": model,
        "voltage_class": voltage,
        "product_type": ("ТРГ-УЭТМ" if model else None),
        "insulation_type": ins.group(1).lower() if ins else None,
        "insulation_color": color.group(1).lower() if color else None,
    }

def _profile_traits(records: list[DataTenderRecord], model: str) -> dict[str, Any]:
    text = " ".join([model] + [f"{r.param_name} {r.value}" for r in records])
    sig = _extract_product_signature(text)
    all_text = " ".join([f"{r.param_name} {r.value}" for r in records])
    ins_types = sorted({m.group(1).lower() for m in _INSULATION_TYPE_RE.finditer(all_text)})
    colors = sorted({m.group(1).lower() for m in _INSULATION_COLOR_RE.finditer(all_text)})
    sig["insulation_types"] = ins_types
    sig["insulation_colors"] = colors
    return sig

def _profile_signature_from_records(records: list[DataTenderRecord], model: str) -> dict[str, Any]:
    text_parts = [model]
    for r in records:
        text_parts.append(f"{r.param_name} {r.value}")
    sig = _extract_product_signature(" ".join(text_parts))
    # Нормализуем тип/цвет по значениям параметров, если они встретились там.
    return sig

@dataclass(frozen=True)
class DataTenderRecord:
    param_name: str
    value: str
    model: str | None
    voltage: str | None
    source_file: str
    table_index: int
    row_index: int
    source_kind: str = "table"


class DataTendersKnowledge:
    def __init__(self, root: str | Path | None = None):
        base = Path(root) if root else Path(__file__).resolve().parents[1]
        self.root = self._find_root(base)
        self._signature: tuple[tuple[str, int, int], ...] = ()
        self._initialized = False
        self._records: list[DataTenderRecord] = []
        self._profiles: dict[str, dict[str, Any]] = {}
        self._image_records: list[dict[str, Any]] = []
        # Индексы: O(1) для основного поиска вместо полного прохода по всем records.
        self._index_exact: dict[tuple[str, str | None, str | None], list[DataTenderRecord]] = {}
        self._index_param_model: dict[tuple[str, str | None], list[DataTenderRecord]] = {}
        self._index_param: dict[str, list[DataTenderRecord]] = {}
        self._model_records: dict[str, list[DataTenderRecord]] = {}
        self._model_keys: tuple[str, ...] = ()
        self._profile_signatures: dict[str, dict[str, str | None]] = {}
        self._signature_index: dict[tuple[str | None, str | None, str | None, str | None], tuple[str, ...]] = {}
        self._context_cache_key: tuple[Any, ...] | None = None
        self._context_cache: dict[str, Any] | None = None
        self._image_ocr_cache_path = self.root / ".image_ocr_cache.json"
        self._image_ocr_cache: dict[str, Any] | None = None

    @staticmethod
    def _find_root(base: Path) -> Path:
        base = base.resolve()
        # Поддерживаем как корень проекта, так и прямую передачу самой папки.
        if base.is_dir() and base.name.lower().replace(" ", "_") == "data_tenders":
            return base
        if base.is_dir() and base.name.lower() == "tenders" and base.parent.name.lower() == "data":
            return base
        for candidate in (base / "data_tenders", base / "data" / "tenders", base / "data tenders"):
            if candidate.is_dir():
                return candidate
        return base / "data_tenders"

    def _files(self) -> list[Path]:
        if not self.root.exists():
            return []
        return sorted(self.root.rglob("*.docx"), key=lambda p: p.name.lower())

    def _get_signature(self) -> tuple[tuple[str, int, int], ...]:
        return tuple((str(p), p.stat().st_mtime_ns, p.stat().st_size) for p in self._files())

    def refresh(self, force: bool = False, *, ocr_images: bool = False) -> None:
        # data_tenders считается неизменяемым в рамках одного запуска.
        # Не делаем rglob/stat на каждом lookup. Для принудительной проверки
        # используйте refresh(force=True) или DATA_TENDERS_REFRESH_CHECK=1.
        if self._initialized and not force and os.getenv("DATA_TENDERS_REFRESH_CHECK", "0") != "1":
            if ocr_images and not self._image_records:
                if os.getenv("DATA_TENDERS_VISION_LIVE", "0") == "1":
                    self._build_image_index()
                else:
                    self._image_records = self._load_cached_image_records() or []
            return
        signature = self._get_signature()
        if not force and signature == self._signature:
            if ocr_images and not self._image_records:
                if os.getenv("DATA_TENDERS_VISION_LIVE", "0") == "1":
                    self._build_image_index()
                else:
                    self._image_records = self._load_cached_image_records() or []
            return
        records: list[DataTenderRecord] = []
        profiles: dict[str, dict[str, Any]] = defaultdict(lambda: {
            "model": None,
            "voltage": None,
            "product_type": "ТРГ-УЭТМ",
            "manufacturer": None,
            "parameters": {},
            "images": [],
            "source_files": [],
        })
        for path in self._files():
            try:
                doc = Document(path)
                file_records, file_profiles = self._parse_doc(doc, path.name)
                records.extend(file_records)
                # Важные реквизиты часто лежат в обычных абзацах, а не в таблицах.
                paragraphs_text = "\n".join(_clean_text(x.text) for x in doc.paragraphs if _clean_text(x.text))
                metadata = self._extract_document_metadata(paragraphs_text)
                # Если в основной части абзацев модель не попалась, используем текст всех таблиц.
                all_models_in_file = list(metadata.get("models") or [])
                if not all_models_in_file:
                    all_models_in_file = list(file_profiles.keys())
                manufacturer = metadata.get("manufacturer")
                if manufacturer and all_models_in_file:
                    for model in all_models_in_file:
                        key = _param_key("Изготовитель")
                        rec = DataTenderRecord("Изготовитель", manufacturer, model, _model_voltage(model), path.name, -2, -1)
                        records.append(rec)
                        file_profiles.setdefault(model, {"parameters": {}}).setdefault("parameters", {}).setdefault(key, []).append(manufacturer)
                for model, profile in file_profiles.items():
                    dst = profiles[model]
                    dst["model"] = model
                    dst["voltage"] = _model_voltage(model)
                    dst["product_type"] = "ТРГ-УЭТМ"
                    if manufacturer:
                        dst["manufacturer"] = manufacturer
                    if path.name not in dst["source_files"]:
                        dst["source_files"].append(path.name)
                    for key, values in profile.get("parameters", {}).items():
                        dst["parameters"].setdefault(key, []).extend(v for v in values if v not in dst["parameters"].get(key, []))
            except Exception as exc:
                print(f"[DATA_TENDERS] пропущен {path.name}: {exc}")
        valid_models = {m for m in profiles if _model_voltage(m) in {"35", "110", "220", "330", "500", "750"}}
        records = [r for r in records if not r.model or r.model in valid_models]
        profiles = {m: v for m, v in profiles.items() if m in valid_models}
        self._records = records
        self._profiles = dict(profiles)
        self._build_record_indexes()
        self._build_profile_signature_index()
        self._model_keys = tuple(sorted(self._profiles))
        self._signature = signature
        self._initialized = True
        self._image_records = []
        self._context_cache_key = None
        self._context_cache = None
        if ocr_images:
            if os.getenv("DATA_TENDERS_VISION_LIVE", "0") == "1":
                self._build_image_index()
            else:
                self._image_records = self._load_cached_image_records() or []
        print(f"[DATA_TENDERS] индекс: files={len(signature)} records={len(records)} models={len(self._profiles)} root={self.root}")

    def _parse_doc(self, doc: Document, filename: str) -> tuple[list[DataTenderRecord], dict[str, Any]]:
        out: list[DataTenderRecord] = []
        profiles: dict[str, dict[str, Any]] = defaultdict(lambda: {"parameters": {}})
        for ti, table in enumerate(doc.tables):
            rows = [[_clean_text(c.text) for c in row.cells] for row in table.rows]
            rows = [r for r in rows if any(r)]
            if not rows:
                continue

            # 1) Матрица «параметр | значения по моделям».
            header_idx = None
            for i, cells in enumerate(rows[:8]):
                if cells and any(_norm(x).startswith(("наименование параметра", "наименование характеристики")) for x in cells):
                    header_idx = i
                    break
            if header_idx is not None:
                header = rows[header_idx]
                model_row_idx = header_idx + 1
                candidate = rows[model_row_idx] if model_row_idx < len(rows) else []
                model_by_col: dict[int, str] = {}
                if candidate:
                    for ci, cell in enumerate(candidate):
                        model = _model_name(cell)
                        if model:
                            model_by_col[ci] = model
                start = model_row_idx + 1 if model_by_col else header_idx + 1
                for ri in range(start, len(rows)):
                    cells = rows[ri]
                    if not cells or not cells[0] or _looks_like_note(cells[0]):
                        continue
                    param_name = cells[0]
                    for ci, value in enumerate(cells[1:], start=1):
                        if not value or value.lower() == "значение":
                            continue
                        model = model_by_col.get(ci)
                        rec = DataTenderRecord(param_name, value, model, _model_voltage(model), filename, ti, ri)
                        out.append(rec)
                        if model:
                            key = _param_key(param_name)
                            profiles[model]["parameters"].setdefault(key, []).append(value)
                # Не прекращаем: следующие таблицы содержат одно-значные/general записи.

            # 2) Таблицы «параметр | значение», и строки, где модель явно указана в первой ячейке.
            for ri, cells in enumerate(rows):
                joined = " | ".join(cells)
                models_here = _all_models(joined)
                if len(cells) >= 2 and not (header_idx is not None and ri <= header_idx + 1):
                    param = cells[0]
                    if _looks_like_note(param):
                        continue
                    # Явно модельные строки (модель + характеристики).
                    explicit_models = models_here
                    if explicit_models and cells[0].strip().upper().startswith("ТРГ"):
                        for model in explicit_models:
                            for ci, value in enumerate(cells[1:], start=1):
                                if value and not value.lower().startswith("значение"):
                                    rec = DataTenderRecord(param, value, model, _model_voltage(model), filename, ti, ri)
                                    out.append(rec)
                                    profiles[model]["parameters"].setdefault(_param_key(param), []).append(value)
                    elif len(cells) == 2 and param and cells[1]:
                        rec = DataTenderRecord(param, cells[1], None, None, filename, ti, ri)
                        out.append(rec)
        return out, profiles

    @staticmethod
    def _extract_document_metadata(text: str) -> dict[str, Any]:
        """Извлекает общие реквизиты документа, которые не обязаны находиться в таблицах."""
        low = _norm(text)
        models = _all_models(text)
        product_type = "ТРГ-УЭТМ" if models or "трг-уэтм" in low else None
        manufacturer = None
        m = _MANUFACTURER_RE.search(text or "")
        if m:
            manufacturer = _clean_text(m.group(0))
        return {"product_type": product_type, "manufacturer": manufacturer, "models": models}

    def _build_record_indexes(self) -> None:
        exact: defaultdict[tuple[str, str | None, str | None], list[DataTenderRecord]] = defaultdict(list)
        by_pm: defaultdict[tuple[str, str | None], list[DataTenderRecord]] = defaultdict(list)
        by_p: defaultdict[str, list[DataTenderRecord]] = defaultdict(list)
        by_m: defaultdict[str, list[DataTenderRecord]] = defaultdict(list)
        for rec in self._records:
            pkey = _param_key(rec.param_name)
            if not pkey:
                continue
            mkey = _norm(rec.model) if rec.model else None
            vkey = _norm(rec.voltage) if rec.voltage else None
            exact[(pkey, mkey, vkey)].append(rec)
            by_pm[(pkey, mkey)].append(rec)
            by_p[pkey].append(rec)
            if mkey:
                by_m[mkey].append(rec)
        self._index_exact = dict(exact)
        self._index_param_model = dict(by_pm)
        self._index_param = dict(by_p)
        self._model_records = dict(by_m)

    def _build_profile_signature_index(self) -> None:
        signatures: dict[str, dict[str, Any]] = {}
        buckets: defaultdict[tuple[str | None, str | None, str | None, str | None], list[str]] = defaultdict(list)
        for model, profile in self._profiles.items():
            recs = [
                DataTenderRecord(k, v, model, _model_voltage(model), "", -1, -1)
                for k, vals in profile.get("parameters", {}).items()
                for v in vals[:32]
            ]
            sig = _profile_traits(recs, model)
            sig["product_type"] = sig.get("product_type") or "ТРГ-УЭТМ"
            manufacturer_values = profile.get("parameters", {}).get(_param_key("Изготовитель"), [])
            if manufacturer_values:
                sig["manufacturer"] = manufacturer_values[0]
            signatures[model] = sig
            key = (
                _norm(sig.get("product_type") or "") or None,
                _norm(sig.get("voltage_class") or "") or None,
                _norm(sig.get("insulation_type") or "") or None,
                _norm(sig.get("insulation_color") or "") or None,
            )
            buckets[key].append(model)
        self._profile_signatures = signatures
        self._signature_index = {k: tuple(v) for k, v in buckets.items()}


    def select_profile(
        self,
        text: str,
        items: list[dict] | None = None,
    ) -> tuple[str | None, dict[str, Any], list[str]]:
        """Определяет паспорт изделия по типу, классу напряжения, изоляции и цвету."""
        combined = [str(text or "")]
        for item in items or []:
            combined.extend([
                str(item.get("param_name", "")),
                str(item.get("required_val", "")),
                str(item.get("current_value", "")),
            ])
        all_text = " ".join(combined)
        sig = _extract_product_signature(all_text)
        sig["product_type"] = sig.get("product_type") or ("ТРГ-УЭТМ" if "трг" in _norm(all_text) and "уэтм" in _norm(all_text) else None)
        mm = _MANUFACTURER_RE.search(all_text)
        if mm:
            sig["manufacturer"] = _clean_text(mm.group(0))
        sig["insulation_types"] = sorted({m.group(1).lower() for m in _INSULATION_TYPE_RE.finditer(all_text)})
        sig["insulation_colors"] = sorted({m.group(1).lower() for m in _INSULATION_COLOR_RE.finditer(all_text)})
        model = sig.get("model")
        if model and model in self._profiles:
            return model, sig, [model]

        desired_voltage = _norm(str(sig.get("voltage_class") or "")) or None
        desired_type = _norm(str(sig.get("product_type") or "")) or ("трг-уэтм" if "трг" in _norm(all_text) and "уэтм" in _norm(all_text) else None)
        desired_ins = set(sig.get("insulation_types") or [])
        desired_color = set(sig.get("insulation_colors") or [])

        scored: list[tuple[int, str]] = []
        if desired_voltage:
            voltage_models = [m for m in self._model_keys if _same_voltage(desired_voltage, _model_voltage(m))]
            if len(voltage_models) == 1:
                sig["model"] = voltage_models[0]
                sig["product_type"] = desired_type or "ТРГ-УЭТМ"
                return voltage_models[0], sig, voltage_models
        for m in self._model_keys:
            ps = self._profile_signatures.get(m, {})
            score = 0
            ptype = _norm(str(ps.get("product_type") or "")) or None
            pvoltage = _norm(str(ps.get("voltage_class") or "")) or None
            pins = set(ps.get("insulation_types") or [])
            pcolors = set(ps.get("insulation_colors") or [])
            if desired_type and ptype == desired_type:
                score += 3
            elif desired_type and ptype and ptype != desired_type:
                continue
            if desired_voltage:
                if pvoltage == desired_voltage:
                    score += 6
                else:
                    continue
            if desired_ins:
                if pins and desired_ins.isdisjoint(pins):
                    continue
                if pins & desired_ins:
                    score += 4
            if desired_color:
                if pcolors and desired_color.isdisjoint(pcolors):
                    continue
                if pcolors & desired_color:
                    score += 5
            scored.append((score, m))
        scored.sort(key=lambda x: (-x[0], x[1]))
        candidates = [m for _, m in scored]
        if not candidates:
            return None, sig, []
        best_score = scored[0][0]
        best = [m for sc, m in scored if sc == best_score]
        return (best[0] if len(best) == 1 else None), sig, candidates[:32]

    def detect_model(self, text: str, items: list[dict] | None = None) -> str | None:
        self.refresh()
        model = _model_name(text)
        if model:
            return model
        selected, _sig, _candidates = self.select_profile(text, items)
        if selected:
            return selected
        votes: dict[str, int] = defaultdict(int)
        for item in items or []:
            for m in _all_models(str(item.get("param_name", "")) + " " + str(item.get("required_val", ""))):
                votes[m] += 1
            for m in _all_models(str(item.get("current_value", ""))):
                votes[m] += 2
        return max(votes, key=votes.get) if votes else None

    def detect_voltage(self, text: str, items: list[dict] | None = None) -> str | None:
        model = self.detect_model(text, items)
        if model:
            return _model_voltage(model)
        for item in items or []:
            if "номинальн" in _norm(item.get("param_name", "")) and "напряж" in _norm(item.get("param_name", "")):
                m = _NUM_RE.search(str(item.get("required_val", "")))
                if m:
                    return m.group(1).replace(",", ".")
        m = _VOLT_RE.search(text or "")
        return m.group(1).replace(",", ".") if m else None

    @staticmethod
    def _score_param(query: str, candidate: str) -> float:
        q = _param_key(query)
        c = _param_key(candidate)
        if not q or not c:
            return 0.0
        if q == c:
            return 1.0
        if q in c or c in q:
            return 0.94
        q_tokens, c_tokens = set(q.split()), set(c.split())
        overlap = len(q_tokens & c_tokens) / max(1, min(len(q_tokens), len(c_tokens)))
        return max(SequenceMatcher(None, q, c).ratio(), overlap * 0.92)

    @staticmethod
    def _candidate_allowed_by_requirement(required_val: str, value: str) -> bool:
        req = _clean_text(required_val)
        val = _clean_text(value)
        if not req or req == "*":
            return True
        # Служебные инструкции ТЗ («Указать», «Определить проектом»,
        # «Согласно ...») не являются ограничением значения. В этих случаях
        # берём подтверждённый факт из выбранного профиля data_tenders.
        req_low = _norm(req)
        if any(token in req_low for token in (
            "указать", "определить проектом", "определяется проектом",
            "согласно", "в соответствии", "по согласованию",
        )):
            return True
        nr, nv = _norm(req), _norm(val)
        if nr == nv or nr in nv:
            # Требование не должно совпасть только с пояснительным текстом в скобках.
            return True
        req_nums = _NUM_RE.findall(req)
        if req_nums:
            value_nums = _NUM_RE.findall(val)
            if len(req_nums) == 1 and req_nums[0] in value_nums:
                return True
        return False

    @staticmethod
    def _select_value(required_val: str, candidate: str) -> str:
        req = _clean_text(required_val)
        if not req or req == "*":
            return candidate
        # Сначала точное совпадение по токену/варианту.
        # Категория размещения выводится из климатической записи: Т1/У1/УХЛ1/ХЛ1.
        if _norm(req).isdigit():
            cat = re.search(r"(?:т|у|ухл|хл)\s*(\d+)", candidate, flags=re.IGNORECASE)
            if cat and _norm(req) == cat.group(1):
                return cat.group(1)

        pieces = re.split(r"\s*;\s*|\s+или\s+|\s*[/|]\s*|\s*,\s*", candidate, flags=re.IGNORECASE)
        for piece in pieces:
            piece = piece.strip()
            if piece and _norm(piece) == _norm(req):
                return piece
        # Для «УХЛ» в «УХЛ1 (УХЛ1*)» — возвращаем конкретный классификационный вариант.
        m = re.search(rf"\b{re.escape(req)}\d?(?:\*)?\b", candidate, flags=re.IGNORECASE)
        if m and req.upper() in {"Т", "У", "УХЛ", "ХЛ"}:
            return m.group(0)
        return candidate

    def lookup(
        self,
        param_name: str,
        *,
        model: str | None = None,
        voltage: str | None = None,
        required_val: str = "",
        min_score: float = 0.91,
        profile_signature: dict[str, str | None] | None = None,
    ) -> tuple[str | None, str | None, float]:
        self.refresh()
        if not param_name.strip() or not _clean_text(required_val) and _param_key(param_name) not in {"номинальное напряжение"}:
            return None, None, 0.0
        # Пустое требование: не брать справочную строку наугад.
        if not _clean_text(required_val) and _param_key(param_name) not in {"номинальное напряжение"}:
            return None, None, 0.0
        model_norm = _norm(model or "") or None
        voltage_norm = _norm(voltage) if voltage else None
        pkey = _param_key(param_name)

        # Новый основной фильтр: участвуют только записи выбранного паспорта изделия.
        candidates_by_profile = []
        if profile_signature:
            desired_model = profile_signature.get("model")
            desired_voltage = _norm(str(profile_signature.get("voltage_class") or "")) or None
            desired_type = _norm(str(profile_signature.get("product_type") or "")) or None
            desired_ins = set(profile_signature.get("insulation_types") or [])
            desired_color = set(profile_signature.get("insulation_colors") or [])
            for m in self._model_keys:
                ps = self._profile_signatures.get(m, {})
                if desired_model and m != desired_model:
                    continue
                if desired_voltage and _norm(str(ps.get("voltage_class") or "")) != desired_voltage:
                    continue
                if desired_type and ps.get("product_type") and _norm(str(ps.get("product_type"))) != desired_type:
                    continue
                pins = set(ps.get("insulation_types") or [])
                pcolors = set(ps.get("insulation_colors") or [])
                if desired_ins and pins and desired_ins.isdisjoint(pins):
                    continue
                if desired_color and pcolors and desired_color.isdisjoint(pcolors):
                    continue
                candidates_by_profile.extend(self._model_records.get(_norm(m), ()))

        # Основной путь: O(1) хеш-поиск по нормализованному параметру,
        # модели и напряжению. Сначала самый строгий ключ.
        raw_candidates: list[DataTenderRecord] = []
        if model_norm and voltage_norm:
            raw_candidates = list(self._index_exact.get((pkey, model_norm, voltage_norm), ()))
        if not raw_candidates and model_norm:
            raw_candidates = list(self._index_param_model.get((pkey, model_norm), ()))
        if not raw_candidates:
            raw_candidates = list(self._index_param.get(pkey, ()))
        if candidates_by_profile:
            allowed_ids = {id(r) for r in candidates_by_profile}
            raw_candidates = [r for r in raw_candidates if id(r) in allowed_ids]
        elif profile_signature:
            raw_candidates = []

        candidates: list[tuple[float, DataTenderRecord, str]] = []
        # Быстрый путь: после строгого O(1)-ключа достаточно одного кандидата.
        if len(raw_candidates) == 1:
            rec = raw_candidates[0]
            if (not model_norm or not rec.model or _norm(rec.model) == model_norm) and (
                not voltage_norm or not rec.voltage or _same_voltage(voltage_norm, rec.voltage)
            ) and self._candidate_allowed_by_requirement(required_val, rec.value):
                selected = self._select_value(required_val, rec.value)
                return selected, f"DATA_TENDERS:{rec.source_file}", 1.20
        for rec in raw_candidates:
            score = 1.0 if _param_key(rec.param_name) == pkey else self._score_param(param_name, rec.param_name)
            if score < min_score:
                continue
            if model_norm and rec.model:
                if _norm(rec.model) != model_norm:
                    continue
                score += 0.20
            elif model_norm and rec.model is None:
                score += 0.03
            elif not model_norm and rec.model:
                continue
            if voltage_norm and rec.voltage and not _same_voltage(voltage_norm, rec.voltage):
                continue
            if not self._candidate_allowed_by_requirement(required_val, rec.value):
                continue
            selected = self._select_value(required_val, rec.value)
            candidates.append((score, rec, selected))
        if not candidates:
            return None, None, 0.0
        candidates.sort(key=lambda x: x[0], reverse=True)
        best_score = candidates[0][0]
        top = [x for x in candidates if x[0] >= best_score - 0.04]
        counts: dict[str, int] = defaultdict(int)
        for _, _, selected in top:
            counts[_norm(selected)] += 1
        winner_norm = max(counts, key=counts.get)
        winner = next(x for x in top if _norm(x[2]) == winner_norm)
        return winner[2], f"DATA_TENDERS:{winner[1].source_file}", best_score

    def get_constant(self, param_name: str, *, text_context: str = "", model: str | None = None, voltage: str | None = None, required_val: str = ""):
        # Обычно model/voltage уже вычислены один раз на документе.
        if not model and text_context:
            model = self.detect_model(text_context)
        if model and not _model_name(model):
            model = _model_name(str(model))
        if not voltage:
            voltage = _model_voltage(model)
            if not voltage and text_context:
                voltage = self.detect_voltage(text_context)
        return self.lookup(param_name, model=model, voltage=voltage, required_val=required_val)

    def model_profile(self, model: str) -> dict[str, Any]:
        self.refresh()
        profile = dict(self._profiles.get(model, {}))
        profile["model"] = model
        profile["voltage"] = _model_voltage(model)
        image_records = self._image_records if isinstance(self._image_records, list) else []
        profile["images"] = [
            x for x in image_records
            if isinstance(x, dict) and model in (x.get("models") or [])
        ]
        return profile

    def build_context(self, *, model: str | None = None, voltage: str | None = None, limit: int = 120) -> dict[str, Any]:
        # Совместимый компактный API для старого кода.
        self.refresh()
        rows = []
        for rec in self._records:
            if model and rec.model and _norm(rec.model) != _norm(model):
                continue
            if voltage and rec.voltage and not _same_voltage(voltage, rec.voltage):
                continue
            rows.append({**asdict(rec)})
            if len(rows) >= limit:
                break
        return {"source": "data_tenders", "root": str(self.root), "model": model, "voltage": voltage, "constants": rows}

    # ------------------------------------------------------------------
    # Изображения: связь image -> ближайший текст -> модель/марка -> OCR.
    # ------------------------------------------------------------------
    def _relationship_map(self, doc: Document) -> dict[str, str]:
        mapping: dict[str, str] = {}
        for rel_id, rel in doc.part.rels.items():
            if getattr(rel, "reltype", "") == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image":
                target = str(getattr(rel, "target_ref", ""))
                if "media/" in target:
                    mapping[rel_id] = Path(target).name
        return mapping

    def _element_images(self, element, relmap: dict[str, str]) -> list[str]:
        found: list[str] = []
        for blip in element.iter(qn("a:blip")):
            rid = blip.get(qn("r:embed"))
            if rid and rid in relmap:
                found.append(relmap[rid])
        for data in element.iter("{urn:schemas-microsoft-com:vml}imagedata"):
            rid = data.get(qn("r:id"))
            if rid and rid in relmap:
                found.append(relmap[rid])
        return list(dict.fromkeys(found))

    def _block_sequence(self, doc: Document, relmap: dict[str, str]) -> list[dict[str, Any]]:
        seq: list[dict[str, Any]] = []
        body = doc.element.body
        for child in body.iterchildren():
            tag = child.tag.rsplit("}", 1)[-1]
            if tag == "p":
                text = _clean_text(" ".join(child.itertext()))
                imgs = self._element_images(child, relmap)
                if text or imgs:
                    seq.append({"text": text, "images": imgs})
            elif tag == "tbl":
                text = _clean_text(" ".join(child.itertext()))
                imgs = self._element_images(child, relmap)
                if text or imgs:
                    seq.append({"text": text, "images": imgs})
        return seq

    def _prepare_image_for_vision(self, raw: bytes, suffix: str) -> tuple[bytes, str]:
        """Нормализует изображение к PNG/JPEG, пригодному для Vision."""
        mime = {
            ".png": "image/png",
            ".jpg": "image/jpeg",
            ".jpeg": "image/jpeg",
            ".webp": "image/webp",
        }.get(suffix.lower(), "")
        if mime:
            return raw, mime
        try:
            image = Image.open(io.BytesIO(raw)) if Image is not None else None
        except Exception:
            image = None
        if image is None and suffix.lower() in {".wmf", ".emf"}:
            magick = os.getenv("MAGICK_BINARY", "magick")
            try:
                proc = subprocess.run(
                    [magick, f"{suffix.lower().lstrip('.') }:-", "png:-"],
                    input=raw, stdout=subprocess.PIPE, stderr=subprocess.DEVNULL,
                    timeout=30, check=True,
                )
                return proc.stdout, "image/png"
            except Exception:
                pass
        if image is None:
            raise ValueError(f"Неподдерживаемое изображение {suffix}")
        out = io.BytesIO()
        image.convert("RGB").save(out, format="PNG", optimize=True)
        return out.getvalue(), "image/png"

    def _vision_image(self, raw: bytes, suffix: str, *, client=None) -> dict[str, Any]:
        image_bytes, mime = self._prepare_image_for_vision(raw, suffix)
        prompt = os.getenv(
            "DATA_TENDERS_VISION_PROMPT",
            "Извлеки все параметры оборудования и их значения со схемы/изображения. "
            "Особенно внимательно прочитай таблицы, габаритные размеры, массу, "
            "марку/тип изделия, класс напряжения, тип и цвет изоляции. "
            "Сохраняй числа, единицы измерения и обозначения дословно. Не додумывай "
            "отсутствующие данные. Верни только данные, которые действительно видны на изображении."
        )
        if client is not None:
            return ask_vision_json(client, image_bytes, prompt, mime_type=mime)
        with open_client() as local_client:
            return ask_vision_json(local_client, image_bytes, prompt, mime_type=mime)

    def _ocr_image_fallback(self, raw: bytes, suffix: str) -> str:
        if pytesseract is None or Image is None:
            return ""
        try:
            image = Image.open(io.BytesIO(raw))
        except Exception:
            return ""
        try:
            image = image.convert("L")
            max_side = int(os.getenv("DATA_TENDERS_IMAGE_MAX_SIDE", "3200"))
            if max(image.size) > max_side:
                scale = max_side / max(image.size)
                image = image.resize((max(1, int(image.width * scale)), max(1, int(image.height * scale))))
            image = ImageOps.autocontrast(image, cutoff=1).filter(ImageFilter.SHARPEN)
            lang = os.getenv("TESSERACT_LANG", "rus+eng")
            psm = os.getenv("DATA_TENDERS_IMAGE_PSM", "11")
            return _clean_text(pytesseract.image_to_string(image, lang=lang, config=f"--psm {psm}"))
        except Exception:
            return ""

    def _load_image_ocr_cache(self) -> dict[str, Any]:
        if self._image_ocr_cache is not None:
            return self._image_ocr_cache
        try:
            data = json.loads(self._image_ocr_cache_path.read_text(encoding="utf-8"))
            self._image_ocr_cache = data if isinstance(data, dict) else {}
        except Exception:
            self._image_ocr_cache = {}
        return self._image_ocr_cache

    def _save_image_ocr_cache(self) -> None:
        try:
            self.root.mkdir(parents=True, exist_ok=True)
            self._image_ocr_cache_path.write_text(json.dumps(self._image_ocr_cache or {}, ensure_ascii=False, indent=2), encoding="utf-8")
        except Exception:
            pass

    def _load_cached_image_records(self) -> list[dict[str, Any]]:
        """Загружает только ранее распознанные Vision-результаты без сетевых вызовов."""
        cache = self._load_image_ocr_cache()
        records: list[dict[str, Any]] = []
        if not cache:
            return records
        for payload in cache.values():
            if not isinstance(payload, dict):
                continue
            vision = payload.get("vision")
            if not isinstance(vision, dict):
                continue
            source_file = payload.get("source_file", "")
            image = payload.get("image", "")
            text = _clean_text(payload.get("text", ""))
            models = [m for m in (payload.get("models") or []) if m in self._profiles]
            records.append({
                "source_file": source_file,
                "image": image,
                "models": models,
                "text": text,
                "vision": vision,
                "context": text[:4000],
                "source": "DATA_TENDERS_IMAGE_VISION_CACHE",
            })
        return records

    def _build_image_index(self) -> None:
        if Image is None:
            self._image_records = []
            return
        cache = self._load_image_ocr_cache()
        records: list[dict[str, Any]] = []
        self._vision_record_keys: set[tuple[Any, ...]] = {
            (r.source_file, r.model, _norm(r.param_name), _norm(r.value))
            for r in self._records if r.source_kind == "image_vision"
        }
        max_images_env = os.getenv("DATA_TENDERS_MAX_IMAGE_OCR", "0")
        try:
            max_images = int(max_images_env)
        except ValueError:
            max_images = 0

        vision_cm = None
        vision_client = None
        vision_ready = False
        try:
            for path in self._files():
                try:
                    doc = Document(path)
                    relmap = self._relationship_map(doc)
                    blocks = self._block_sequence(doc, relmap)
                    known_models = set(self._profiles)
                    all_models = [m for m in _all_models(" ".join(b.get("text", "") for b in blocks)) if m in known_models]
                    image_counter = 0
                    with zipfile.ZipFile(path, "r") as zf:
                        media = {Path(n).name: n for n in zf.namelist() if n.startswith("word/media/")}
                        for bi, block in enumerate(blocks):
                            if not block.get("images"):
                                continue
                            near_text = " ".join(b.get("text", "") for b in blocks[max(0, bi - 4):min(len(blocks), bi + 5)])
                            models = [m for m in _all_models(near_text) if m in known_models] or all_models
                            for image_name in block["images"]:
                                if image_name not in media:
                                    continue
                                image_counter += 1
                                if max_images > 0 and image_counter > max_images:
                                    continue
                                raw = zf.read(media[image_name])
                                digest = hashlib.sha1(raw).hexdigest()
                                cache_key = f"{path.resolve()}::{image_name}::{digest}"
                                old = cache.get(cache_key)
                                vision = old.get("vision") if isinstance(old, dict) else None
                                text = old.get("text", "") if isinstance(old, dict) else ""

                                needs_vision = not isinstance(vision, dict) or (not vision.get("parameters") and not vision.get("raw_text"))
                                if needs_vision:
                                    try:
                                        if vision_client is None:
                                            vision_cm = open_client()
                                            vision_client = vision_cm.__enter__()
                                            vision_ready = True
                                        vision = self._vision_image(raw, Path(image_name).suffix, client=vision_client)
                                        text = _clean_text(
                                            vision.get("raw_text", "")
                                            or " ".join(
                                                f"{p.get('name', '')}: {p.get('value', '')} {p.get('unit', '')}"
                                                for p in (vision.get("parameters") or [])
                                                if isinstance(p, dict)
                                            )
                                        )
                                        cache[cache_key] = {
                                            "vision": vision,
                                            "text": text,
                                            "models": models,
                                            "source_file": path.name,
                                            "image": image_name,
                                        }
                                    except Exception as exc:
                                        print(f"[DATA_TENDERS][VISION] {path.name}/{image_name}: {exc}")
                                        text = self._ocr_image_fallback(raw, Path(image_name).suffix)
                                        vision = {
                                            "model_mentions": [],
                                            "parameters": [],
                                            "raw_text": text,
                                            "notes": [f"vision_failed: {exc}"],
                                        }
                                        cache[cache_key] = {
                                            "vision": vision,
                                            "text": text,
                                            "models": models,
                                            "source_file": path.name,
                                            "image": image_name,
                                        }

                                image_models = [
                                    m for m in _all_models(_clean_text(" ".join(map(str, vision.get("model_mentions", [])))))
                                    if m in known_models
                                ]
                                if image_models:
                                    models = image_models
                                records.append({
                                    "source_file": path.name,
                                    "image": image_name,
                                    "models": models,
                                    "text": _clean_text(text),
                                    "vision": vision,
                                    "context": _clean_text(near_text[:4000]),
                                    "source": "DATA_TENDERS_IMAGE_VISION",
                                })
                except Exception as exc:
                    print(f"[DATA_TENDERS][IMAGE] {path.name}: {exc}")
        finally:
            if vision_cm is not None and vision_ready:
                try:
                    vision_cm.__exit__(None, None, None)
                except Exception:
                    pass

        # Данные, извлечённые Vision с изображений, становятся частью того же
        # первичного индекса, что и табличные данные. Благодаря этому get_constant()
        # может найти габариты/массу/климат прямо на чертеже, а не только передать
        # их второму уровню AI.
        vision_added = 0
        for image_rec in records:
            vision = image_rec.get("vision") or {}
            image_models = list(image_rec.get("models") or [])
            params = vision.get("parameters") or []
            for par in params:
                if not isinstance(par, dict):
                    continue
                name = _clean_text(par.get("name", ""))
                value = _clean_text(par.get("value", ""))
                unit = _clean_text(par.get("unit", ""))
                if not name or not value:
                    continue
                models_from_param = [
                    m for m in _all_models(str(par.get("model", "")))
                    if m in self._profiles
                ]
                target_models = models_from_param or image_models or [None]
                for model in target_models:
                    rec = DataTenderRecord(
                        param_name=f"{name} ({unit})" if unit else name,
                        value=value,
                        model=model,
                        voltage=_model_voltage(model),
                        source_file=image_rec.get("source_file", ""),
                        table_index=-1,
                        row_index=-1,
                        source_kind="image_vision",
                    )
                    # Не дублируем одну и ту же картинку/параметр/модель/значение.
                    duplicate = (
                        rec.source_file, rec.model, _norm(rec.param_name), _norm(rec.value)
                    ) in self._vision_record_keys
                    if not duplicate:
                        self._vision_record_keys.add(
                            (rec.source_file, rec.model, _norm(rec.param_name), _norm(rec.value))
                        )
                        self._records.append(rec)
                        if model:
                            profile = self._profiles.setdefault(model, {
                                "model": model, "voltage": _model_voltage(model),
                                "parameters": {}, "images": [], "source_files": [],
                            })
                            key = _param_key(name)
                            profile.setdefault("parameters", {}).setdefault(key, [])
                            if value not in profile["parameters"][key]:
                                profile["parameters"][key].append(value)
                        vision_added += 1

        self._image_records = records
        self._build_record_indexes()
        self._context_cache_key = None
        self._context_cache = None
        self._save_image_ocr_cache()
        print(f"[DATA_TENDERS][IMAGE] image_records={len(records)} vision_records_added={vision_added} cache={self._image_ocr_cache_path}")

    def _load_cached_image_records(self) -> None:
        """Быстро загружает только уже распознанные Vision-картинки.

        Никаких сетевых запросов здесь нет. Полный Vision-прогон выполняется
        отдельно через preprocess_data_tenders_images().
        """
        cache = self._load_image_ocr_cache()
        if not cache:
            return
        records: list[dict[str, Any]] = []
        known_models = set(self._profiles)
        for entry in cache.values():
            if not isinstance(entry, dict):
                continue
            vision = entry.get("vision")
            if not isinstance(vision, dict):
                continue
            source_file = str(entry.get("source_file", ""))
            image = str(entry.get("image", ""))
            models = [m for m in (entry.get("models") or []) if m in known_models]
            records.append({
                "source_file": source_file,
                "image": image,
                "models": models,
                "text": _clean_text(entry.get("text", "") or ""),
                "vision": vision,
                "context": _clean_text(entry.get("context", "") or ""),
                "source": "DATA_TENDERS_IMAGE_VISION_CACHE",
            })
        self._image_records = records

    def ensure_vision_for_model(self, model: str, max_images: int | None = None) -> int:
        """Однократно распознаёт Vision только изображения выбранного изделия.

        В обычной обработке не гоняем все изображения всех data_tenders. Берём
        только картинки, связанные с выбранной моделью, и сохраняем результат в кэш.
        Повторный запуск сетевых запросов не делает.
        """
        self.refresh(ocr_images=False)
        cached = self._load_cached_image_records()
        cached_keys = {(x.get("source_file"), x.get("image")) for x in cached}
        if max_images is None:
            try:
                max_images = int(os.getenv("DATA_TENDERS_VISION_MODEL_MAX", "6"))
            except ValueError:
                max_images = 6
        max_images = max(0, max_images)
        if max_images == 0 or os.getenv("DATA_TENDERS_VISION_NO_NETWORK", "0") == "1":
            self._image_records = cached
            return 0

        target: list[tuple[Path, str, bytes, list[str], str]] = []
        for path in self._files():
            try:
                doc = Document(path)
                relmap = self._relationship_map(doc)
                blocks = self._block_sequence(doc, relmap)
                known_models = set(self._profiles)
                file_text = _clean_text(" ".join(b.get("text", "") for b in blocks))
                file_has_model = model in _all_models(file_text)
                with zipfile.ZipFile(path, "r") as zf:
                    media = {Path(n).name: n for n in zf.namelist() if n.startswith("word/media/")}
                    local_candidates = []
                    for bi, block in enumerate(blocks):
                        if not block.get("images"):
                            continue
                        near_text = _clean_text(" ".join(b.get("text", "") for b in blocks[max(0, bi - 4):min(len(blocks), bi + 5)]))
                        near_models = [m for m in _all_models(near_text) if m in known_models]
                        priority = 2 if model in near_models else (1 if file_has_model else 0)
                        for image_name in block["images"]:
                            if image_name not in media or (path.name, image_name) in cached_keys:
                                continue
                            local_candidates.append((priority, bi, image_name, near_text[:4000]))
                    local_candidates.sort(key=lambda x: (-x[0], x[1]))
                    for priority, bi, image_name, near_text in local_candidates:
                        raw = zf.read(media[image_name])
                        target.append((path, image_name, raw, [model], near_text))
                        if len(target) >= max_images:
                            break
            except Exception as exc:
                print(f"[DATA_TENDERS][VISION_SELECT] {path.name}: {exc}")
            if len(target) >= max_images:
                break

        if not target:
            self._image_records = cached
            return 0

        cache = self._load_image_ocr_cache()
        added = 0
        with open_client() as client:
            for path, image_name, raw, models, near_text in target:
                digest = hashlib.sha1(raw).hexdigest()
                cache_key = f"{path.resolve()}::{image_name}::{digest}"
                try:
                    image_bytes, mime = self._prepare_image_for_vision(raw, Path(image_name).suffix)
                    vision = ask_vision_json(
                        client=client,
                        image_bytes=image_bytes,
                        user_prompt=os.getenv(
                            "DATA_TENDERS_VISION_PROMPT",
                            "Извлеки все параметры оборудования и их значения со схемы/изображения."
                        ),
                        mime_type=mime,
                    )
                    # ask_vision_json() в зависимости от версии SDK может вернуть
                    # None/пустое значение. Это не должно ломать весь pipeline.
                    if not isinstance(vision, dict):
                        vision = {"parameters": [], "raw_text": str(vision or ""), "notes": ["empty_or_non_dict_response"]}
                    params = vision.get("parameters") or []
                    if not isinstance(params, list):
                        params = []
                        vision["parameters"] = params
                    text = _clean_text(
                        vision.get("raw_text", "") or " ".join(
                            f"{p.get('name','')}: {p.get('value','')} {p.get('unit','')}"
                            for p in params if isinstance(p, dict)
                        )
                    )
                    cache[cache_key] = {"vision": vision, "text": text, "models": models, "source_file": path.name, "image": image_name, "context": near_text}
                    added += 1
                except Exception as exc:
                    print(f"[DATA_TENDERS][VISION] {path.name}/{image_name}: {exc}")
        self._save_image_ocr_cache()
        self._image_records = self._load_cached_image_records() or []
        self._context_cache_key = None
        self._context_cache = None
        return added

    def build_full_context(
        self,
        *,
        model: str | None = None,
        voltage: str | None = None,
        include_images: bool = True,
        max_chars: int | None = None,
    ) -> dict[str, Any]:
        """Полный корпус data_tenders + структурированные профили + OCR изображений."""
        # В обычном запуске НЕ вызываем Vision для всех изображений.
        # include_images означает "подключить уже закэшированные Vision-данные".
        self.refresh(ocr_images=False)
        if include_images:
            self._load_cached_image_records()
        if max_chars is None:
            try:
                max_chars = int(os.getenv("TENDER_AI_DATA_TENDERS_MAX_CHARS", "180000"))
            except ValueError:
                max_chars = 180000

        cache_key = (self._signature, _norm(model) if model else None, _norm(voltage) if voltage else None, bool(include_images), int(max_chars))
        if self._context_cache_key == cache_key and self._context_cache is not None:
            return self._context_cache

        files_ctx: list[dict[str, Any]] = []
        models = list(self._profiles)
        for path in self._files():
            try:
                doc = Document(path)
                tables = []
                for ti, table in enumerate(doc.tables):
                    rows = []
                    for row in table.rows:
                        vals = []
                        seen = set()
                        for cell in row.cells:
                            val = _clean_text(cell.text)
                            # Удаляем только повтор merged-cell, сохраняя реальную строку.
                            if val and val not in seen:
                                vals.append(val)
                                seen.add(val)
                        if vals:
                            rows.append(vals)
                    if rows:
                        tables.append({"table": ti, "rows": rows})
                imgs = [x for x in self._image_records if x["source_file"] == path.name]
                files_ctx.append({"file": path.name, "tables": tables, "images": imgs})
            except Exception as exc:
                files_ctx.append({"file": path.name, "error": str(exc)})

        model_profiles = {m: self._profiles[m] for m in models}
        context = {
            "source": "data_tenders",
            "root": str(self.root),
            "selection": {"model": model, "voltage": voltage},
            "model_profiles": model_profiles,
            "files": files_ctx,
            "images": self._image_records,
            "record_count": len(self._records),
            "model_count": len(models),
            "selection_rule": "model/voltage are filters for decisions, not permissions to copy unrelated values",
        }
        serialized = json.dumps(context, ensure_ascii=False)
        if len(serialized) <= max_chars:
            self._context_cache_key = cache_key
            self._context_cache = context
            return context
        # Сначала ужимаем OCR, затем повторяющиеся длинные таблицы; сами structured profiles сохраняем.
        for image in context["images"]:
            image["text"] = image.get("text", "")[:1000]
            image["context"] = image.get("context", "")[:1800]
        for file_ctx in context["files"]:
            for image in file_ctx.get("images", []):
                image["text"] = image.get("text", "")[:1000]
                image["context"] = image.get("context", "")[:1800]
        serialized = json.dumps(context, ensure_ascii=False)
        if len(serialized) <= max_chars:
            self._context_cache_key = cache_key
            self._context_cache = context
            return context
        # Не дублируем images на уровне files, но сохраняем отдельный полный image-index.
        for file_ctx in context["files"]:
            file_ctx["images"] = [{"image": x.get("image"), "models": x.get("models", []), "text": x.get("text", "")} for x in file_ctx.get("images", [])]
        self._context_cache_key = cache_key
        self._context_cache = context
        return context


def preprocess_data_tenders_images(root: str | Path | None = None) -> dict[str, Any]:
    """Полностью построить кэш OCR изображений data_tenders один раз."""
    kb = DataTendersKnowledge(root)
    kb.refresh(force=True, ocr_images=True)
    return {"root": str(kb.root), "files": len(kb._files()), "records": len(kb._records), "models": sorted(kb._profiles), "image_records": len(kb._image_records), "cache": str(kb._image_ocr_cache_path)}


def _looks_like_note(text: str) -> bool:
    low = _norm(text)
    return low.startswith(("примечание", "1)", "2)", "3)", "4)", "5)")) or (len(low) > 300 and ("гост" in low or "в соответствии" in low))


_default = DataTendersKnowledge()


def get_data_tenders_knowledge() -> DataTendersKnowledge:
    return _default


__all__ = ["DataTenderRecord", "DataTendersKnowledge", "get_data_tenders_knowledge", "preprocess_data_tenders_images"]
