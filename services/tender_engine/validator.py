from __future__ import annotations

import re

from .schema import TenderRow
from .text import clean, norm, numbers, strip_marks


class TenderValidator:
    """Проверяет ответ уже после выбора источника/AI."""

    def validate(self, row: TenderRow, value: str, rows: list[TenderRow]) -> tuple[bool, str]:
        value = strip_marks(value)
        if not value:
            return True, ""
        if not self._type_ok(row, value):
            return False, "Значение не соответствует типу строки ТЗ."
        if not self._allowed_text_ok(row, value):
            return False, "Значение не входит в допустимое текстовое множество требования."
        if not self._cross_fields_ok(row, value, rows):
            return False, "Значение противоречит связанным параметрам документа."
        if not self._dangerous_mismatch_ok(row, value):
            return False, "Обнаружен перенос значения из другого параметра/единицы."
        return True, ""

    def _type_ok(self, row: TenderRow, value: str) -> bool:
        req = norm(row.requirement)
        val = norm(value)
        if req in {"да", "нет"}:
            # Сертификаты могут содержать подтверждающие реквизиты вместо одного «Да».
            if row.field_key in {"certificate", "type_approval", "meter_certificates"}:
                return val in {"да", "нет"} or bool(re.search(r"\b(?:ru|росс)\b|\b№\b|действует до", val, flags=re.I))
            return val in {"да", "нет"} and val == req if row.field_key not in {"chief_engineer", "certified_staff"} else val in {"да", "нет"}
        if row.field_key in {"upper_temperature", "lower_temperature", "nominal_voltage", "max_working_voltage", "frequency", "gas_pressure", "service_life", "warranty", "verification_interval", "mechanical_load", "partial_discharge", "impulse_voltage", "ac_test_voltage"}:
            return bool(numbers(value))
        if row.field_key == "dimensions":
            return len(numbers(value)) >= 2
        if row.field_key == "mass":
            nums = numbers(value)
            if "*/*" in req or "*/" in req or "/*" in req:
                return len(nums) >= 2
            return len(nums) >= 2 or "*" in row.requirement
        if row.field_key == "gas_mass":
            return any(token in val for token in ("sf6", "n2", "азот", "элегаз"))
        return True

    def _allowed_text_ok(self, row: TenderRow, value: str) -> bool:
        req = norm(row.requirement)
        val = norm(value)
        pairs = {
            "climate": ("ухл", "ухл1", "хл1", "у1", "т1"),
            "external_insulation": ("фарфор", "полимер"),
            "internal_insulation": ("элегаз", "газ", "изолирующий газ"),
        }
        allowed = pairs.get(row.field_key)
        if allowed:
            return any(x in val for x in allowed)
        if row.field_key == "transport" and req and req in val:
            return True
        return True

    def _cross_fields_ok(self, row: TenderRow, value: str, rows: list[TenderRow]) -> bool:
        def first_num(key: str) -> float | None:
            target = next((r for r in rows if r.field_key == key), None)
            vals = numbers(target.proposed_value if target else "")
            return vals[0] if vals else None

        if row.field_key == "max_working_voltage":
            nominal = first_num("nominal_voltage")
            nums = numbers(value)
            return nominal is None or not nums or nums[0] >= nominal
        if row.field_key == "upper_temperature":
            low = first_num("lower_temperature")
            nums = numbers(value)
            return low is None or not nums or nums[0] >= low
        if row.field_key == "lower_temperature":
            high = first_num("upper_temperature")
            nums = numbers(value)
            return high is None or not nums or nums[0] <= high
        return True

    def _dangerous_mismatch_ok(self, row: TenderRow, value: str) -> bool:
        val = norm(value)
        p = norm(row.parameter)
        if row.field_key in {"transport", "chief_engineer", "certified_staff"} and any(token in val for token in ("0,7", "0.11", "0.15", "297", "302", "317", "252")):
            return False
        if row.field_key not in {"mass", "dimensions", "gas_mass", "certificate", "manufacturer", "brand"} and re.search(r"\b(?:297|302|317|252)\b", val):
            return False
        if "скорость ветра" in p and len(numbers(value)) > 1:
            return False
        return True
