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
        self._records: list[DataTenderRecord] = []
        self._profiles: dict[str, dict[str, Any]] = {}
        self._image_records: list[dict[str, Any]] = []
        # Индексы: O(1) для основного поиска вместо полного прохода по всем records.
        self._index_exact: dict[tuple[str, str | None, str | None], list[DataTenderRecord]] = {}
        self._index_param_model: dict[tuple[str, str | None], list[DataTenderRecord]] = {}
        self._index_param: dict[str, list[DataTenderRecord]] = {}
        self._model_records: dict[str, list[DataTenderRecord]] = {}
        self._model_keys: tuple[str, ...] = ()
        self._context_cache_key: tuple[Any, ...] | None = None
        self._context_cache: dict[str, Any] | None = None
        self._image_ocr_cache_path = self.root / ".image_ocr_cache.json"
        self._image_ocr_cache: dict[str, Any] | None = None

    @staticmethod
    def _find_root(base: Path) -> Path:
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
        signature = self._get_signature()
        if not force and signature == self._signature:
            if ocr_images and not self._image_records:
                if os.getenv("DATA_TENDERS_VISION_LIVE", "0") == "1":
                    self._build_image_index()
                else:
                    self._image_records = self._load_cached_image_records()
            return
        records: list[DataTenderRecord] = []
        profiles: dict[str, dict[str, Any]] = defaultdict(lambda: {
            "model": None,
            "voltage": None,
            "parameters": {},
            "images": [],
            "source_files": [],
        })
        for path in self._files():
            try:
                doc = Document(path)
                file_records, file_profiles = self._parse_doc(doc, path.name)
                records.extend(file_records)
                for model, profile in file_profiles.items():
                    dst = profiles[model]
                    dst["model"] = model
                    dst["voltage"] = _model_voltage(model)
                    if path.name not in dst["source_files"]:
                        dst["source_files"].append(path.name)
                    for key, values in profile.get("parameters", {}).items():
                        dst["parameters"].setdefault(key, []).extend(v for v in values if v not in dst["parameters"].get(key, []))
            except Exception as exc:
                print(f"[DATA_TENDERS] пропущен {path.name}: {exc}")
        self._records = records
        self._profiles = dict(profiles)
        self._build_record_indexes()
        self._model_keys = tuple(sorted(self._profiles))
        self._signature = signature
        self._image_records = []
        self._context_cache_key = None
        self._context_cache = None
        if ocr_images:
            if os.getenv("DATA_TENDERS_VISION_LIVE", "0") == "1":
                self._build_image_index()
            else:
                self._image_records = self._load_cached_image_records()
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

    def detect_model(self, text: str, items: list[dict] | None = None) -> str | None:
        self.refresh()
        model = _model_name(text)
        if model:
            return model
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

        # Основной путь: O(1) хеш-поиск по нормализованному параметру,
        # модели и напряжению. Сначала самый строгий ключ.
        raw_candidates: list[DataTenderRecord] = []
        if model_norm and voltage_norm:
            raw_candidates = list(self._index_exact.get((pkey, model_norm, voltage_norm), ()))
        if not raw_candidates and model_norm:
            raw_candidates = list(self._index_param_model.get((pkey, model_norm), ()))
        if not raw_candidates:
            raw_candidates = list(self._index_param.get(pkey, ()))

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
        # В обычном заполнении model/voltage уже определены один раз на документ.
        # Не запускаем повторный поиск по всему тексту для каждой строки.
        if not model:
            model = self.detect_model(text_context)
        elif not _model_name(model):
            model = None
        if not voltage:
            voltage = _model_voltage(model) or self.detect_voltage(text_context)
        return self.lookup(param_name, model=model, voltage=voltage, required_val=required_val)

    def model_profile(self, model: str) -> dict[str, Any]:
        self.refresh()
        profile = dict(self._profiles.get(model, {}))
        profile["model"] = model
        profile["voltage"] = _model_voltage(model)
        profile["images"] = [x for x in self._image_records if model in x.get("models", [])]
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
            """Ты анализируешь технический документ трансформатора тока.
Извлеки ВСЕ читаемые технические данные с изображения, включая таблицы, схемы,
чертежные размеры, массы, климатические исполнения, изоляцию, марку/тип изделия
и подписи. Сохраняй числа, единицы и обозначения максимально дословно.
Особенно внимательно распознавай многозначные строки и значения, связанные с
конкретной моделью/маркой. Не додумывай отсутствующие данные.
Верни строго JSON: {
  "model_mentions": ["..."],
  "parameters": [{"name": "...", "value": "...", "unit": "...", "model": "..."}],
  "raw_text": "полный связный текст изображения",
  "notes": ["..."]
}"""
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
