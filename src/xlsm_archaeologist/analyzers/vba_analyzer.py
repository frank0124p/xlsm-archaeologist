"""Main VBA analysis pipeline: extract modules → split → analyze procedures."""

from __future__ import annotations

import re
from collections.abc import Iterator
from pathlib import Path

from xlsm_archaeologist.analyzers.vba_call_graph import extract_calls
from xlsm_archaeologist.analyzers.vba_procedure_splitter import ProcedureChunk, split_procedures
from xlsm_archaeologist.analyzers.vba_range_detector import detect_range_accesses, detect_triggers
from xlsm_archaeologist.extractors.vba_extractor import extract_vba_modules
from xlsm_archaeologist.models.vba import RangeAccess, VbaModuleRecord, VbaProcedureRecord
from xlsm_archaeologist.utils.logging import get_logger

logger = get_logger(__name__)

_BRANCH_KEYWORDS = re.compile(
    r"\b(If|ElseIf|Select\s+Case|Case\s+(?!Else))\b",
    re.IGNORECASE,
)


def _complexity(chunk: ProcedureChunk, reads: list[RangeAccess], writes: list[RangeAccess]) -> int:
    """Compute complexity = line_count + read_count + write_count + branch_count."""
    code = "\n".join(chunk.source_lines)
    branches = len(_BRANCH_KEYWORDS.findall(code))
    return len(chunk.source_lines) + len(reads) + len(writes) + branches


def _analyze_procedure(
    chunk: ProcedureChunk,
    module_id: int,
    proc_id: int,
    all_proc_names: set[str],
    module_type: str,
) -> VbaProcedureRecord:
    code = "\n".join(chunk.source_lines)
    reads, writes, dynamic_notes = detect_range_accesses(code)
    triggers = detect_triggers(chunk.name, code) if module_type in ("sheet", "workbook") else []
    calls = extract_calls(code, all_proc_names - {chunk.name})
    has_dynamic = bool(dynamic_notes)
    score = _complexity(chunk, reads, writes)

    return VbaProcedureRecord(
        vba_procedure_id=proc_id,
        vba_module_id=module_id,
        procedure_name=chunk.name,
        procedure_type=chunk.procedure_type,
        is_public=chunk.is_public,
        parameters=chunk.parameters,
        line_count=len(chunk.source_lines),
        reads=reads,
        writes=writes,
        calls=calls,
        triggers=triggers,
        has_dynamic_range=has_dynamic,
        dynamic_range_notes=dynamic_notes,
        complexity_score=score,
        source_code=code,
    )


def analyze_vba(
    input_path: Path,
    warnings: list[str],
) -> tuple[list[VbaModuleRecord], list[VbaProcedureRecord]]:
    """Run full VBA analysis pipeline.

    Two-pass approach:
    1. Extract modules + split procedures (collects all_procedure_names)
    2. Analyze each procedure for ranges/calls/triggers

    Args:
        input_path: Path to the .xlsm file.
        warnings: Mutable list to collect warning strings.

    Returns:
        Tuple of (modules, procedures) sorted by id.
    """
    modules = list(extract_vba_modules(input_path, warnings))

    # Pass 1: collect all procedure names
    module_chunks: list[tuple[VbaModuleRecord, list[ProcedureChunk]]] = []
    all_proc_names: set[str] = set()
    for module in modules:
        chunks = split_procedures(module.source_code)
        module_chunks.append((module, chunks))
        for c in chunks:
            all_proc_names.add(c.name)

    # Pass 2: analyze
    procedures: list[VbaProcedureRecord] = []
    proc_id = 0
    for module, chunks in module_chunks:
        for chunk in chunks:
            proc_id += 1
            try:
                rec = _analyze_procedure(
                    chunk, module.vba_module_id, proc_id, all_proc_names, module.module_type
                )
                procedures.append(rec)
            except Exception as exc:  # noqa: BLE001
                warnings.append(
                    f"VBA procedure analysis failed for {module.module_name}.{chunk.name}: {exc}"
                )
                logger.warning("Procedure analysis error: %s", exc)

    return modules, procedures


def iter_modules(
    input_path: Path,
    warnings: list[str],
) -> Iterator[VbaModuleRecord]:
    """Yield VbaModuleRecords from *input_path* (convenience wrapper)."""
    yield from extract_vba_modules(input_path, warnings)
