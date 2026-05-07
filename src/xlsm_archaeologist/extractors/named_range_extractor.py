"""Extract named range (defined name) records from the workbook."""

from __future__ import annotations

from collections.abc import Iterator

from openpyxl.workbook.workbook import Workbook

from xlsm_archaeologist.models.named_range import NamedRangeRecord
from xlsm_archaeologist.utils.logging import get_logger

logger = get_logger(__name__)

_DYNAMIC_FUNCS = {"OFFSET(", "INDIRECT(", "INDEX("}


def _is_dynamic(refers_to: str) -> bool:
    """Return True when refers_to contains a volatile/dynamic function call."""
    upper = refers_to.upper()
    return any(fn in upper for fn in _DYNAMIC_FUNCS)


def _is_valid_ref(refers_to: str) -> bool:
    """Return False when refers_to contains a broken reference."""
    return "#REF!" not in refers_to.upper()


def extract_named_ranges(wb: Workbook) -> Iterator[NamedRangeRecord]:
    """Yield one NamedRangeRecord per defined name in the workbook.

    Skips built-in / print-area names that start with '_xlnm.'.

    Args:
        wb: Open openpyxl Workbook (read_only=False).

    Yields:
        NamedRangeRecord instances sorted by range_name (caller is responsible
        for final sort before writing to CSV).
    """
    seen_id = 0
    for name in sorted(wb.defined_names):
        if name.startswith("_xlnm."):
            continue

        dn = wb.defined_names[name]

        # Resolve refers_to text and scope
        raw_refers: str = getattr(dn, "attr_text", None) or getattr(dn, "value", "") or ""
        refers_to = raw_refers if raw_refers.startswith("=") else f"={raw_refers}"

        # localSheetId present → sheet-scoped name
        local_id = getattr(dn, "localSheetId", None)
        if local_id is not None:
            try:
                scope = wb.worksheets[local_id].title
            except IndexError:
                scope = str(local_id)
        else:
            scope = "workbook"

        seen_id += 1
        yield NamedRangeRecord(
            named_range_id=seen_id,
            range_name=name,
            scope=scope,
            refers_to=refers_to,
            has_dynamic_formula=_is_dynamic(refers_to),
            is_valid=_is_valid_ref(refers_to),
        )
