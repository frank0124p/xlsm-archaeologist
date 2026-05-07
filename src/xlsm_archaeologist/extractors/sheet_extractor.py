"""Extract sheet-level metadata into SheetRecord instances."""

from __future__ import annotations

from typing import Iterator

from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from xlsm_archaeologist.models.workbook import SheetRecord
from xlsm_archaeologist.utils.logging import get_logger

logger = get_logger(__name__)

# openpyxl sheet state constants
_STATE_HIDDEN = "hidden"
_STATE_VERY_HIDDEN = "veryHidden"


def _parse_dimension(dim: str | None) -> tuple[int, int]:
    """Parse 'A1:Z100' into (row_count, col_count).

    Returns (0, 0) when the dimension string is absent or malformed.
    """
    if not dim or ":" not in dim:
        return 0, 0
    try:
        from openpyxl.utils.cell import range_boundaries

        min_col, min_row, max_col, max_row = range_boundaries(dim)
        return max_row - min_row + 1, max_col - min_col + 1
    except Exception:
        return 0, 0


def _count_cells(sheet: Worksheet) -> tuple[int, int]:
    """Count total non-empty cells and formula cells in the sheet.

    Returns (cell_count, formula_cell_count).
    """
    total = 0
    formulas = 0
    for row in sheet.iter_rows():
        for cell in row:
            if cell.value is not None:
                total += 1
                if cell.data_type == "f":
                    formulas += 1
    return total, formulas


def extract_sheets(wb: Workbook) -> Iterator[SheetRecord]:
    """Yield one SheetRecord per worksheet in the workbook.

    Args:
        wb: Open openpyxl Workbook (read_only=False).

    Yields:
        SheetRecord sorted by sheet_index (which matches wb.worksheets order).
    """
    for idx, sheet in enumerate(wb.worksheets):
        state = sheet.sheet_state or ""
        is_hidden = state == _STATE_HIDDEN
        is_very_hidden = state == _STATE_VERY_HIDDEN

        dim = sheet.calculate_dimension()
        used_range = dim if dim else "A1:A1"
        row_count, col_count = _parse_dimension(used_range)
        cell_count, formula_cell_count = _count_cells(sheet)

        logger.debug("Sheet[%d] %r: %s (%d cells)", idx, sheet.title, used_range, cell_count)
        yield SheetRecord(
            sheet_id=idx + 1,
            sheet_name=sheet.title,
            sheet_index=idx,
            is_hidden=is_hidden,
            is_very_hidden=is_very_hidden,
            used_range=used_range,
            row_count=row_count,
            col_count=col_count,
            cell_count=cell_count,
            formula_cell_count=formula_cell_count,
        )
