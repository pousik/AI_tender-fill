from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any


@dataclass
class TenderRow:
    row_id: str
    table_index: int
    row_index: int
    number: str
    parameter: str
    requirement: str
    current_value: str
    target_cell_index: int | None
    parent_context: str = ""
    section: str = ""
    field_key: str = ""
    is_header: bool = False
    evidence: list[str] = field(default_factory=list)
    candidate_values: list[str] = field(default_factory=list)
    candidate_sources: list[str] = field(default_factory=list)
    proposed_value: str = ""
    source: str = ""
    mark: str = ""
    confidence: float = 0.0
    reason: str = ""

    @property
    def identity(self) -> tuple[str, str, str]:
        return (
            self.number.rstrip("."),
            self.field_key,
            self.parent_context.strip().lower(),
        )

    def to_dict(self) -> dict[str, Any]:
        return {
            "id": self.row_id,
            "table": self.table_index,
            "row": self.row_index,
            "num": self.number,
            "param_name": self.parameter,
            "required_val": self.requirement,
            "current_value": self.current_value,
            "target_cell_index": self.target_cell_index,
            "parent_context": self.parent_context,
            "section": self.section,
            "field_key": self.field_key,
            "algorithm_value": self.proposed_value if self.source.startswith("DB") else "",
            "algorithm_source": self.source,
            "final_value": self.proposed_value,
            "final_mark": self.mark,
            "source": self.source,
            "confidence": self.confidence,
            "evidence": list(self.evidence),
            "reason": self.reason,
        }


@dataclass(frozen=True)
class ProductContext:
    product_kind: str
    model: str | None
    voltage: str | None
    manufacturer: str | None
    signature: dict[str, Any] = field(default_factory=dict)


@dataclass(frozen=True)
class SourceCandidate:
    value: str
    source: str
    score: float
    evidence: str = ""
    model: str | None = None
    field_key: str = ""


@dataclass(frozen=True)
class Decision:
    value: str
    source: str
    mark: str
    confidence: float
    reason: str
    evidence: tuple[str, ...] = ()
