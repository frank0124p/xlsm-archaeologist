"""Detect cell/range read and write accesses in VBA procedure source code."""

from __future__ import annotations

import re

from xlsm_archaeologist.models.vba import EventTrigger, RangeAccess, RangeAccessVia

# Patterns for explicit Range("...") or Cells(r,c) references
_EXPLICIT_RANGE = re.compile(
    r'(?:(?P<sheet>["\w]+)\.)?Range\(\s*"(?P<addr>[A-Za-z0-9$:]+)"\s*\)',
    re.IGNORECASE,
)
_NAMED_RANGE = re.compile(
    r'(?:(?P<sheet>["\w]+)\.)?Range\(\s*"(?P<name>[A-Za-z_]\w*[A-Za-z_]\w*)"\s*\)',
    re.IGNORECASE,
)
_CELLS_METHOD = re.compile(r"\bCells\s*\(", re.IGNORECASE)
_DYNAMIC_PATTERNS = [
    re.compile(r'Range\s*\(\s*(?:"[^"]*"|[A-Za-z_]\w*)\s*&', re.IGNORECASE),  # Range("A" & n)
    re.compile(r'\bIndirect\s*\(', re.IGNORECASE),
    re.compile(r'\bOffset\s*\(', re.IGNORECASE),
    re.compile(r'Range\s*\(\s*[A-Za-z_]\w*\s*\)', re.IGNORECASE),  # Range(var)
]
_WRITE_PATTERN = re.compile(
    r'(?:(?P<sheet>["\w]+)\.)?Range\(\s*"(?P<addr>[^"]+)"\s*\)\s*(?:\.Value\s*)?=(?!=)',
    re.IGNORECASE,
)
_CELLS_WRITE = re.compile(r'\bCells\s*\([^)]+\)\s*(?:\.Value\s*)?=(?!=)', re.IGNORECASE)

_EVENT_NAMES: dict[str, str] = {
    "worksheet_change": "Worksheet_Change",
    "worksheet_selectionchange": "Worksheet_SelectionChange",
    "worksheet_beforedoubleclick": "Worksheet_BeforeDoubleClick",
    "worksheet_beforerightclick": "Worksheet_BeforeRightClick",
    "worksheet_activate": "Worksheet_Activate",
    "worksheet_deactivate": "Worksheet_Deactivate",
    "workbook_open": "Workbook_Open",
    "workbook_beforesave": "Workbook_BeforeSave",
    "workbook_beforeclose": "Workbook_BeforeClose",
    "workbook_sheetchange": "Workbook_SheetChange",
}

_INTERSECT_RE = re.compile(
    r'Intersect\s*\(Target\s*,\s*Range\s*\(\s*"([^"]+)"',
    re.IGNORECASE,
)


def _strip_comments(line: str) -> str:
    """Remove VBA line comment starting with ' (single quote)."""
    in_str = False
    for i, ch in enumerate(line):
        if ch == '"':
            in_str = not in_str
        elif ch == "'" and not in_str:
            return line[:i]
    return line


def _is_address(value: str) -> bool:
    """Return True if value looks like a cell address rather than a named range."""
    if not value:
        return False
    v = value.replace("$", "").upper()
    if ":" in v:
        return True
    # column-only like A:A
    if re.fullmatch(r"[A-Z]+:[A-Z]+", v):
        return True
    # single cell A1
    m = re.fullmatch(r"[A-Z]+\d+", v)
    return m is not None


def detect_range_accesses(
    source_code: str,
) -> tuple[list[RangeAccess], list[RangeAccess], list[str]]:
    """Scan VBA source code for range read and write accesses.

    Args:
        source_code: Procedure or module source.

    Returns:
        Tuple of (reads, writes, dynamic_notes).
        - reads: RangeAccess list for cells read
        - writes: RangeAccess list for cells written
        - dynamic_notes: Human-readable descriptions of dynamic range patterns
    """
    reads: list[RangeAccess] = []
    writes: list[RangeAccess] = []
    dynamic_notes: list[str] = []
    write_addresses: set[str] = set()

    for raw_line in source_code.splitlines():
        line = _strip_comments(raw_line)

        # Write accesses
        for m in _WRITE_PATTERN.finditer(line):
            addr = m.group("addr")
            sheet = m.group("sheet")
            via: RangeAccessVia = "explicit_range" if _is_address(addr) else "named_range"
            write_addresses.add(addr.upper())
            writes.append(RangeAccess(sheet=sheet, range_ref=addr, via=via))

        if _CELLS_WRITE.search(line):
            writes.append(RangeAccess(sheet=None, range_ref="(cells)", via="cells_method"))

        # Dynamic patterns
        for pattern in _DYNAMIC_PATTERNS:
            if pattern.search(line):
                note = f"Dynamic range: {raw_line.strip()}"
                if note not in dynamic_notes:
                    dynamic_notes.append(note)

        # All explicit Range reads (not writes)
        for m in _EXPLICIT_RANGE.finditer(line):
            addr = m.group("addr")
            sheet = m.group("sheet")
            if addr.upper() not in write_addresses:
                via = "explicit_range" if _is_address(addr) else "named_range"
                reads.append(RangeAccess(sheet=sheet, range_ref=addr, via=via))

        if _CELLS_METHOD.search(line) and not _CELLS_WRITE.search(line):
            reads.append(RangeAccess(sheet=None, range_ref="(cells)", via="cells_method"))

    return reads, writes, dynamic_notes


def detect_triggers(procedure_name: str, source_code: str) -> list[EventTrigger]:
    """Detect Excel event triggers from the procedure name convention.

    Args:
        procedure_name: Name of the procedure (e.g. 'Worksheet_Change').
        source_code: Procedure body for Intersect heuristic.

    Returns:
        List of EventTrigger, typically 0 or 1 entry.
    """
    canonical = _EVENT_NAMES.get(procedure_name.lower())
    if canonical is None:
        return []

    # Try to extract target range from Intersect(Target, Range("..."))
    target: str | None = None
    m = _INTERSECT_RE.search(source_code)
    if m:
        target = m.group(1)

    return [EventTrigger(event=canonical, target=target)]
