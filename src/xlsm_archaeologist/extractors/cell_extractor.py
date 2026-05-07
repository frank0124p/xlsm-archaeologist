"""Extract 'meaningful' cells from all worksheets into CellRecord instances."""

from __future__ import annotations

from typing import Iterator

from openpyxl.workbook.workbook import Workbook

from xlsm_archaeologist.models.cell import CellRecord, ValueType
from xlsm_archaeologist.utils.logging import get_logger

logger = get_logger(__name__)


def _value_type(cell: object) -> ValueType:
    """Map an openpyxl cell to our ValueType enum."""
    from openpyxl.cell.cell import Cell, TYPE_BOOL, TYPE_ERROR, TYPE_NUMERIC, TYPE_STRING

    if not isinstance(cell, Cell):
        return "empty"
    if cell.value is None:
        return "empty"
    if cell.data_type == TYPE_BOOL:
        return "boolean"
    if cell.data_type == TYPE_ERROR:
        return "error"
    if cell.data_type == TYPE_NUMERIC:
        # openpyxl stores dates as numbers with a date format
        if cell.is_date:
            return "date"
        return "number"
    if cell.data_type == TYPE_STRING or cell.data_type == "s":
        return "string"
    # formula cell — report the cached value type if possible
    if cell.data_type == "f":
        # value may be the cached display value (when data_only=False it's the formula string)
        return "string"
    return "string"


def _raw_value(cell: object) -> str:
    """Return the string representation of a cell's stored value."""
    from openpyxl.cell.cell import Cell

    if not isinstance(cell, Cell) or cell.value is None:
        return ""
    return str(cell.value)


def extract_cells(
    wb: Workbook,
    named_addresses: set[str],
    validation_addresses: set[str],
) -> Iterator[CellRecord]:
    """Yield CellRecord for every 'meaningful' cell across all sheets.

    A cell is meaningful if it satisfies at least one of:
    - has a formula (data_type == 'f')
    - its qualified address is in *validation_addresses*
    - its qualified address is in *named_addresses*

    ``is_referenced`` is always False here; Phase 5 back-fills it.

    Args:
        wb: Open openpyxl Workbook (read_only=False, data_only=False).
        named_addresses: Set of ``"SheetName!A1"`` strings from named ranges.
        validation_addresses: Set of ``"SheetName!A1"`` strings from validations.

    Yields:
        CellRecord instances (unsorted; caller sorts by qualified_address).
    """
    seen_id = 0
    for sheet in wb.worksheets:
        sheet_name = sheet.title
        for row in sheet.iter_rows():
            for cell in row:
                has_formula = cell.data_type == "f"
                qa = f"{sheet_name}!{cell.coordinate}"
                has_validation = qa in validation_addresses
                is_named = qa in named_addresses

                if not (has_formula or has_validation or is_named):
                    continue

                seen_id += 1
                yield CellRecord(
                    cell_id=seen_id,
                    sheet_name=sheet_name,
                    cell_address=cell.coordinate,
                    qualified_address=qa,
                    cell_row=cell.row,
                    cell_col=cell.column,
                    has_formula=has_formula,
                    has_validation=has_validation,
                    is_named=is_named,
                    is_referenced=False,
                    value_type=_value_type(cell),
                    raw_value=_raw_value(cell),
                )
