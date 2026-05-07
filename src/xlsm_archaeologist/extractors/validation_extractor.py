"""Extract data validation rules from all worksheets."""

from __future__ import annotations

from typing import Iterator

from openpyxl.workbook.workbook import Workbook

from xlsm_archaeologist.models.cell import ValidationRecord, ValidationType
from xlsm_archaeologist.utils.logging import get_logger

logger = get_logger(__name__)

_VALID_TYPES: set[str] = {"list", "whole", "decimal", "date", "time", "length", "custom"}


def _normalise_type(raw: str | None) -> ValidationType:
    """Map openpyxl validation type string to our ValidationType enum value."""
    if raw and raw.lower() in _VALID_TYPES:
        return raw.lower()  # type: ignore[return-value]
    return "custom"


def _parse_enum_values(formula1: str, sheet_name: str, wb: Workbook) -> str:
    """Parse list-type formula1 into a pipe-separated string of options.

    Supports two formats:
    - Literal list: ``'"A,B,C"'`` or ``'"A","B","C"'``
    - Range reference: ``'=Sheet!$A$1:$A$5'`` → reads cell values

    Returns empty string on any failure (non-fatal).
    """
    if not formula1:
        return ""

    # Literal list: formula1 is wrapped in double-quotes e.g. '"Yes,No"'
    stripped = formula1.strip()
    if stripped.startswith('"') and stripped.endswith('"'):
        inner = stripped[1:-1]
        # Items may be comma or semicolon separated; strip whitespace
        parts = [p.strip() for p in inner.replace(";", ",").split(",")]
        return "|".join(p for p in parts if p)

    # Range reference: e.g. '=Params!$A$2:$A$10' or 'Sheet1!A1:A5'
    ref = stripped.lstrip("=")
    try:
        from openpyxl.utils.cell import range_boundaries

        if "!" in ref:
            sheet_part, range_part = ref.split("!", 1)
            target_sheet_name = sheet_part.strip("'$")
        else:
            target_sheet_name = sheet_name
            range_part = ref

        range_part = range_part.replace("$", "")
        if target_sheet_name not in wb.sheetnames:
            return ""

        target_ws = wb[target_sheet_name]
        values: list[str] = []
        for row in target_ws[range_part]:
            for cell in row if hasattr(row, "__iter__") else [row]:
                if cell.value is not None:
                    values.append(str(cell.value))
        return "|".join(values)
    except Exception as exc:
        logger.debug("Could not parse enum_values from %r: %s", formula1, exc)
        return ""


def extract_validations(wb: Workbook) -> Iterator[ValidationRecord]:
    """Yield one ValidationRecord per data-validation rule per sheet.

    Args:
        wb: Open openpyxl Workbook (read_only=False).

    Yields:
        ValidationRecord instances (unsorted; caller sorts by qualified_address).
    """
    seen_id = 0
    for sheet in wb.worksheets:
        for dv in sheet.data_validations.dataValidation:
            raw_type = getattr(dv, "type", None)
            val_type = _normalise_type(raw_type)
            f1: str = str(dv.formula1) if dv.formula1 else ""
            f2: str = str(dv.formula2) if dv.formula2 else ""

            enum_values = ""
            if val_type == "list":
                enum_values = _parse_enum_values(f1, sheet.title, wb)

            # sqref can be a CellRange object or a string with multiple tokens
            sqref = str(dv.sqref) if dv.sqref else ""
            # Split on whitespace to get individual range tokens
            tokens = sqref.split() if sqref else []
            if not tokens:
                tokens = [""]

            # First token gives us the primary address; we emit one row per rule
            # (not per cell — that would explode the CSV for large ranges)
            first_range = tokens[0] if tokens else ""
            first_cell = first_range.split(":")[0] if first_range else ""
            qualified_address = f"{sheet.title}!{first_cell}" if first_cell else sheet.title

            seen_id += 1
            yield ValidationRecord(
                validation_id=seen_id,
                qualified_address=qualified_address,
                range_text=sqref,
                validation_type=val_type,
                formula1=f1,
                formula2=f2,
                enum_values=enum_values,
                allow_blank=bool(dv.allow_blank),
                error_title=str(dv.error_title) if dv.error_title else "",
                error_message=str(dv.error) if dv.error else "",
            )
