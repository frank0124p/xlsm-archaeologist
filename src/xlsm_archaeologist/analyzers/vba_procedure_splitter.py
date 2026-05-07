"""Split a VBA module's source code into individual procedure chunks."""

from __future__ import annotations

import re
from dataclasses import dataclass

from xlsm_archaeologist.models.vba import Parameter, ProcedureType

_PROC_START = re.compile(
    r"^(?P<scope>Public|Private|Friend|Static)?\s*"
    r"(?P<ptype>Sub|Function|Property\s+Get|Property\s+Let|Property\s+Set)\s+"
    r"(?P<name>\w+)\s*(?P<params>\([^)]*\))?",
    re.IGNORECASE,
)

_PROC_END = re.compile(
    r"^\s*End\s+(Sub|Function|Property)\b",
    re.IGNORECASE,
)

_PARAM_RE = re.compile(
    r"(?P<opt>Optional\s+)?(?:ByVal\s+|ByRef\s+)?(?P<name>\w+)"
    r"(?:\s+As\s+(?P<type>\w+(?:\.\w+)*))?",
    re.IGNORECASE,
)


def _parse_params(params_str: str) -> list[Parameter]:
    """Parse a parameter list string '(a As Integer, Optional b As String)' into Parameters."""
    inner = params_str.strip().strip("()")
    if not inner.strip():
        return []
    result: list[Parameter] = []
    for part in inner.split(","):
        part = part.strip()
        if not part:
            continue
        m = _PARAM_RE.match(part)
        if m:
            result.append(
                Parameter(
                    name=m.group("name"),
                    type_hint=m.group("type"),
                    is_optional=bool(m.group("opt")),
                )
            )
    return result


def _normalize_type(ptype_str: str) -> ProcedureType:
    s = ptype_str.lower().replace(" ", "_")
    mapping: dict[str, ProcedureType] = {
        "sub": "sub",
        "function": "function",
        "property_get": "property_get",
        "property_let": "property_let",
        "property_set": "property_set",
    }
    return mapping.get(s, "sub")


@dataclass
class ProcedureChunk:
    """Raw data for one procedure extracted from a module."""

    name: str
    procedure_type: ProcedureType
    is_public: bool
    parameters: list[Parameter]
    source_lines: list[str]


def split_procedures(source_code: str) -> list[ProcedureChunk]:
    """Split VBA module source into individual procedure chunks.

    Args:
        source_code: Full module source code string.

    Returns:
        List of ProcedureChunk, one per detected Sub/Function/Property.
    """
    # Join continuation lines (ending with ' _')
    joined_lines: list[str] = []
    pending = ""
    for raw_line in source_code.splitlines():
        stripped = raw_line.rstrip()
        if stripped.endswith(" _"):
            pending += stripped[:-2]
        else:
            joined_lines.append(pending + stripped)
            pending = ""
    if pending:
        joined_lines.append(pending)

    chunks: list[ProcedureChunk] = []
    in_proc = False
    current_lines: list[str] = []
    current_name = ""
    current_type: ProcedureType = "sub"
    current_public = True
    current_params: list[Parameter] = []

    for line in joined_lines:
        stripped = line.strip()

        # Skip pure comment lines outside a procedure
        if not in_proc and stripped.startswith("'"):
            continue

        if not in_proc:
            m = _PROC_START.match(stripped)
            if m:
                scope = (m.group("scope") or "").lower()
                current_public = scope not in ("private", "friend")
                current_name = m.group("name")
                current_type = _normalize_type(m.group("ptype"))
                params_str = m.group("params") or "()"
                current_params = _parse_params(params_str)
                current_lines = [line]
                in_proc = True
        else:
            current_lines.append(line)
            if _PROC_END.match(stripped):
                chunks.append(
                    ProcedureChunk(
                        name=current_name,
                        procedure_type=current_type,
                        is_public=current_public,
                        parameters=current_params,
                        source_lines=current_lines,
                    )
                )
                in_proc = False
                current_lines = []

    return chunks
