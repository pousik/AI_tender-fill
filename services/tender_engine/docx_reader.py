from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Any

from docx.document import Document as _DocumentType
from docx.table import _Cell, Table

from .rules import field_key
from .schema import TenderRow
from .text import clean, is_numbered, norm, strip_marks


@dataclass
class DocumentSnapshot:
    paragraphs: list[str]
    tables: list[list[list[str]]]
    answer_columns: dict[int, int]
    rows: list[TenderRow]


ANSWER_HEADERS = (
    "предлагаемое участником конкурса",
    "предлагаемое участником",
    "предложение участника",
    "ответ участника",
    "данные участника",
    "заполняется участником",
)


class DocxTableReader:
    """Извлекает тендерные строки и запоминает физическую Word-ячейку."""

    def read(self, doc: _DocumentType) -> DocumentSnapshot:
        paragraphs = [clean(p.text) for p in doc.paragraphs if clean(p.text)]
        tables: list[list[list[str]]] = []
        answer_columns: dict[int, int] = {}
        rows: list[TenderRow] = []
        for table_index, table in enumerate(doc.tables):
            table_text = self._table_text(table)
            tables.append(table_text)
            answer_col = self._detect_answer_column(table)
            if answer_col is not None:
                answer_columns[table_index] = answer_col
            rows.extend(self._read_table(table, table_index, answer_col))
        return DocumentSnapshot(paragraphs, tables, answer_columns, rows)

    def _table_text(self, table: Table) -> list[list[str]]:
        return [[self._cell_text(cell) for cell in row.cells] for row in table.rows]

    @staticmethod
    def _cell_text(cell: _Cell) -> str:
        return re.sub(r"[ \t]+", " ", str(cell.text or "")).strip()

    @staticmethod
    def _physical_cells(row) -> list[_Cell]:
        cells: list[_Cell] = []
        seen: set[int] = set()
        for cell in row.cells:
            marker = id(cell._tc)
            if marker in seen:
                continue
            seen.add(marker)
            cells.append(cell)
        return cells

    def _detect_answer_column(self, table: Table) -> int | None:
        for row in table.rows[:10]:
            for idx, cell in enumerate(self._physical_cells(row)):
                text = norm(cell.text)
                if any(header in text for header in ANSWER_HEADERS):
                    return idx
        return None

    def _read_table(self, table: Table, table_index: int, answer_col: int | None) -> list[TenderRow]:
        result: list[TenderRow] = []
        section = ""
        parent_context = ""
        counter = 0
        for row_index, row in enumerate(table.rows):
            cells = [self._cell_text(c) for c in self._physical_cells(row)]
            if len(cells) < 3 or not any(cells):
                continue

            number = cells[0]
            # В реальном Word-документе заголовок ответа может занимать объединённую
            # ячейку, поэтому его физический индекс нельзя переносить 1:1 на все строки.
            # Для тендерной таблицы ответ всегда находится в последней физической
            # ячейке строки, а требование — непосредственно перед ним.
            target_idx = len(cells) - 1
            req_idx = len(cells) - 2

            if len(cells) >= 5:
                # [номер, родитель/обмотка, параметр, требование, ответ]
                row_parent = cells[1]
                parameter = cells[2]
                local_parent = row_parent if self._looks_like_winding(row_parent) else parent_context
            else:
                # [номер, параметр, требование, ответ]
                parameter = cells[1]
                local_parent = parent_context

            required = cells[req_idx] if req_idx >= 0 else ""
            current = cells[target_idx] if target_idx >= 0 else ""

            if norm(parameter) in {"технические требования к оборудованию (наименование параметра)", "наименование параметра"}:
                continue

            section_candidate = self._section_number(number, parameter, required)
            is_section = self._looks_like_section(parameter, number, required)
            if self._looks_like_group_header(parameter, required, current):
                parent_context = parameter
                continue
            if is_section:
                if section_candidate:
                    section = section_candidate
                parent_context = parameter
                continue

            if self._looks_like_winding(parameter):
                parent_context = parameter
                continue

            # Защита от сдвига: если колонка ответа почему-то указывает на требование,
            # последняя физическая колонка является безопасным резервом.
            if current == strip_marks(required) and required not in {"", "*"} and len(cells) >= 4:
                target_idx, req_idx = len(cells) - 1, len(cells) - 2
                current, required = cells[target_idx], cells[req_idx]

            if re.fullmatch(r"\*+", strip_marks(current or "")):
                current = ""

            key = field_key(parameter, local_parent)
            row_id = f"t{table_index}_r{row_index}_{counter}"
            counter += 1
            result.append(TenderRow(
                row_id=row_id,
                table_index=table_index,
                row_index=row_index,
                number=number,
                parameter=parameter,
                requirement=required,
                current_value=current,
                target_cell_index=target_idx,
                parent_context=local_parent,
                section=section,
                field_key=key,
            ))
        return result

    @staticmethod
    def _section_number(number: str, parameter: str, requirement: str) -> str:
        if is_numbered(number):
            return number.rstrip(".")
        return ""

    @staticmethod
    def _looks_like_winding(parameter: str) -> bool:
        return bool(re.search(r"обмотка\s+\d+(?:\s*[-–]\s*\d+)?", norm(parameter)))

    @staticmethod
    def _looks_like_group_header(parameter: str, requirement: str, current: str) -> bool:
        p = norm(parameter)
        return (
            not requirement
            and not current
            and (p.startswith("параметры ") or p.endswith(":"))
        )

    @staticmethod
    def _looks_like_section(parameter: str, number: str, requirement: str) -> bool:
        p = norm(parameter)
        n = clean(number)
        r = clean(requirement)
        if not p:
            return False
        if p.endswith(":") and not r:
            return True
        if is_numbered(n) and re.fullmatch(r"\d+\.?", n) and not r:
            return True
        if not r and any(p.startswith(prefix) for prefix in (
            "технические требования", "основные ", "требования по ", "требования к ",
            "для тт ", "массо-габаритные", "комплектность", "маркировка",
        )):
            return True
        return False


def snapshot_context(snapshot: DocumentSnapshot) -> dict[str, Any]:
    return {
        "paragraphs": list(snapshot.paragraphs),
        "tables": snapshot.tables,
        "rows": [row.to_dict() for row in snapshot.rows],
        "answer_columns": dict(snapshot.answer_columns),
        "has_answer_column": bool(snapshot.answer_columns),
    }
