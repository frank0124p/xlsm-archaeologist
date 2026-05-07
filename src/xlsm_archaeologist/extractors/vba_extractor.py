"""Extract VBA modules from a .xlsm file using oletools."""

from __future__ import annotations

from collections.abc import Iterator
from pathlib import Path

from xlsm_archaeologist.models.vba import VbaModuleRecord, VbaModuleType
from xlsm_archaeologist.utils.logging import get_logger

logger = get_logger(__name__)

_SHEET_KEYWORDS = ("sheet", "feuille", "tabelle", "hoja")  # common multi-lang patterns


def _detect_module_type(module_name: str, code_path: str) -> VbaModuleType:
    """Heuristically classify a VBA module type from its name and internal path."""
    name_lower = module_name.lower()
    path_lower = (code_path or "").lower()

    if "thisworkbook" in name_lower or "thisworkbook" in path_lower:
        return "workbook"
    if any(k in name_lower for k in _SHEET_KEYWORDS):
        return "sheet"
    if path_lower.endswith(".cls"):
        return "class"
    if path_lower.endswith(".frm"):
        return "form"
    # Check codepath extensions for standard
    if path_lower.endswith(".bas"):
        return "standard"
    return "standard"


def extract_vba_modules(
    path: Path,
    warnings: list[str],
) -> Iterator[VbaModuleRecord]:
    """Extract VBA modules from a .xlsm / .xlsb file using oletools.

    Args:
        path: Absolute path to the workbook file.
        warnings: Mutable list to append warning strings into.

    Yields:
        VbaModuleRecord per discovered module.
    """
    try:
        from oletools.olevba import VBA_Parser
    except ImportError:
        warnings.append("oletools not available — VBA extraction skipped")
        return

    module_id = 0
    try:
        vba_parser = VBA_Parser(str(path))
        if not vba_parser.detect_vba_macros():
            logger.debug("No VBA macros detected in %s", path)
            return

        for filename, stream_path, vba_filename, source_code in vba_parser.extract_macros():
            module_id += 1
            # Derive module name from vba_filename (e.g. 'Module1.bas' → 'Module1')
            raw_name = vba_filename or filename or f"module_{module_id}"
            module_name = raw_name.rsplit(".", 1)[0] if "." in raw_name else raw_name
            code_path = stream_path or vba_filename or ""
            module_type = _detect_module_type(module_name, code_path)
            lines = source_code.splitlines() if source_code else []

            # Count procedures
            import re

            proc_count = len(
                re.findall(
                    r"^\s*(public\s+|private\s+|friend\s+)?(sub|function|property\s+get"
                    r"|property\s+let|property\s+set)\s+\w",
                    source_code or "",
                    re.IGNORECASE | re.MULTILINE,
                )
            )

            yield VbaModuleRecord(
                vba_module_id=module_id,
                module_name=module_name,
                module_type=module_type,
                line_count=len(lines),
                procedure_count=proc_count,
                source_code=source_code or "",
            )

    except Exception as exc:  # noqa: BLE001
        warnings.append(f"VBA extraction failed for {path.name}: {exc}")
        logger.warning("VBA extraction error: %s", exc)
