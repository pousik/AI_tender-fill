"""Совместимый фасад локальной БЗ data_tenders."""
from __future__ import annotations

from pathlib import Path

from services.tender_engine.data_source import DataTendersRepository, get_data_tenders_knowledge
from services.tender_engine.rules import field_key as _param_key
from services.tender_engine.text import norm as _norm


class DataTendersKnowledge(DataTendersRepository):
    def detect_model(self, text: str, items=None):
        import re
        match = re.search(r"ТРГ\s*[-–]?\s*УЭТМ\s*[®™]?\s*[-–]?\s*(35|110|220|330|500|750)", text or "", re.I)
        return f"ТРГ-УЭТМ-{match.group(1)}" if match else None

    def detect_voltage(self, text: str, items=None):
        model = self.detect_model(text, items)
        if model:
            return model.rsplit("-", 1)[-1]
        return None

    def select_profile(self, text: str, items=None):
        model = self.detect_model(text, items)
        voltage = self.detect_voltage(text, items)
        signature = {"model": model, "voltage_class": voltage, "product_type": "ТРГ-УЭТМ" if model else None}
        return model, signature, [model] if model else []

    def _model_voltage(self, model):
        return model.rsplit("-", 1)[-1] if model else None

    def lookup(self, parameter, *, model=None, voltage=None, required_val="", profile_signature=None, min_score=0.0, **kwargs):
        candidates = super().lookup(parameter, model=model, voltage=voltage, requirement=required_val, parent_context=kwargs.get("parent_context", ""))
        if not candidates:
            return "", "NONE", 0.0
        best = candidates[0]
        return best.value, best.source, best.score

    def model_profile(self, model):
        return super().model_profile(model)

    def build_context(self, *, model=None, voltage=None, limit=120):
        return self.model_profile(model) if model else {"documents": len(self._docs)}

    def build_full_context(self, *, model=None, voltage=None, max_chars=120000, **kwargs):
        return self.all_text(model=model, limit=max_chars)

    def build_ai_index(self, *, include_images=True):
        return {"documents": len(self._docs), "include_images": include_images}

    def ensure_vision_for_model(self, model, max_images=None):
        return 0

    def refresh(self, force=False, *, ocr_images=False):
        super().refresh()
        return None


def preprocess_data_tenders_images(root: str | Path | None = None) -> dict:
    repo = DataTendersRepository(Path(root) if root else Path(__file__).resolve().parents[1] / "data_tenders")
    return {"documents": len(repo._docs), "vision_added": 0}


__all__ = ["DataTendersKnowledge", "get_data_tenders_knowledge", "_norm", "_param_key", "preprocess_data_tenders_images"]
