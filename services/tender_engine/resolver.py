from __future__ import annotations

from dataclasses import dataclass

from .data_source import DataTendersRepository
from .requirements import RequirementResolver
from .schema import Decision, ProductContext, SourceCandidate, TenderRow
from .text import clean, norm, strip_marks


@dataclass(frozen=True)
class ResolverConfig:
    overwrite_existing: bool = False


class ValueResolver:
    """Выбирает источник ответа без генерации текста нейросетью."""

    # Эти поля напрямую описывают реальные возможности/исполнение изделия.
    PRODUCT_FIRST = {
        "manufacturer", "brand", "internal_insulation", "external_insulation", "external_color",
        "max_working_voltage", "gas_pressure", "dimensions", "mass", "gas_mass", "impulse_voltage",
        "ac_test_voltage", "service_life",
    }
    # Эти поля в конкурсной таблице обычно должны повторять конкретно заказанную
    # конфигурацию/условие, если оно задано заказчиком.
    REQUIREMENT_FIRST = {
        "nominal_voltage", "frequency", "primary_current", "max_primary_current", "thermal_current",
        "thermal_time", "secondary_current", "secondary_count", "accuracy_class", "secondary_load",
        "instrument_security", "protection_limit_multiple", "verification_capability", "mechanical_load",
        "tightness", "gas_alarm", "gas_gauge", "gas_safety_valve", "gas_leak", "placement_category",
        "climate", "upper_temperature", "lower_temperature", "wind_no_ice", "wind_ice", "ice_thickness",
        "installation_altitude", "seismicity", "overvoltage", "partial_discharge", "secondary_ac_test",
        "interturn_test", "repair_free", "explosion_safety", "verification_interval", "warranty", "radio_noise",
        "documentation", "terminal_contacts", "support_structures", "service_kit", "anticorrosion", "initial_verification",
        "marking_packaging", "transport", "delivery", "factory_tests", "commissioning", "training", "arrival_72h",
        "spare_parts_20y", "spare_parts_6m",
    }

    def __init__(self, repository: DataTendersRepository, config: ResolverConfig | None = None):
        self.repository = repository
        self.config = config or ResolverConfig()
        self.requirements = RequirementResolver()

    def resolve_initial(self, row: TenderRow, context: ProductContext, *, template_value: str = "") -> Decision:
        if row.current_value and not self.config.overwrite_existing:
            return Decision(row.current_value, "EXISTING_ANSWER", "", 1.0, "Существующее значение сохранено.")

        if template_value:
            return Decision(strip_marks(template_value), "DB_TEMPLATE", "", 1.0, "Значение взято из ранее исправленного шаблона инженера.")

        candidates = self.repository.lookup(
            row.parameter,
            model=context.model,
            voltage=context.voltage,
            requirement=row.requirement,
            parent_context=row.parent_context,
        )
        row.candidate_values = [c.value for c in candidates]
        row.candidate_sources = [c.evidence for c in candidates]

        requirement_value = self.requirements.resolve(row)
        product = self._best_candidate(candidates)
        preferred = self._preferred(row.field_key, requirement_value, product)
        if preferred:
            return Decision(preferred.value, preferred.source, "", preferred.score, preferred.evidence, (preferred.evidence,))
        if requirement_value and row.field_key not in self.PRODUCT_FIRST:
            return Decision(requirement_value, "REQUIREMENT", "", 1.0, "Явное значение взято из требования ТЗ.", (f"Требование: {row.requirement}",))
        if product and row.field_key not in self.PRODUCT_FIRST:
            return Decision(product.value, product.source, "", product.score, product.evidence, (product.evidence,))
        return Decision("", "UNRESOLVED", "", 0.0, "Для строки нет надежного детерминированного источника.")

    def _preferred(self, key: str, requirement_value: str, product: SourceCandidate | None) -> SourceCandidate | None:
        if key in self.PRODUCT_FIRST and product:
            return product
        if key in self.REQUIREMENT_FIRST and requirement_value:
            return SourceCandidate(requirement_value, "REQUIREMENT", 1.0, "Явное значение требования ТЗ", None, key)
        if product and requirement_value and self._same(requirement_value, product.value):
            return product
        if product and not requirement_value:
            return product
        return None

    @staticmethod
    def _best_candidate(candidates: list[SourceCandidate]) -> SourceCandidate | None:
        return max(candidates, key=lambda x: x.score, default=None)

    @staticmethod
    def _same(left: str, right: str) -> bool:
        a, b = norm(strip_marks(left)), norm(strip_marks(right))
        return a == b or (a and b and (a in b or b in a))

    @staticmethod
    def _mark_requirement(row: TenderRow) -> str:
        return "*" if "*" in row.requirement else "**"
