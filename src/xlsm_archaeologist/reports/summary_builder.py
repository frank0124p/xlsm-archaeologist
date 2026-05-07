"""Build the SummaryRecord from all phase results."""

from __future__ import annotations

import datetime
from pathlib import Path
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from xlsm_archaeologist.models.cell import ValidationRecord
    from xlsm_archaeologist.models.dependency import CycleRecord, DependencyEdge
    from xlsm_archaeologist.models.formula import FormulaRecord
    from xlsm_archaeologist.models.named_range import NamedRangeRecord
    from xlsm_archaeologist.models.vba import VbaModuleRecord, VbaProcedureRecord
    from xlsm_archaeologist.models.workbook import SheetRecord

from xlsm_archaeologist.analyzers.summary_analyzer import (
    compute_complexity_score,
    compute_risk_indicators,
    compute_stats,
    derive_migration_difficulty,
)
from xlsm_archaeologist.models.summary import SummaryRecord, SummaryWarning


def build_summary(
    input_path: Path,
    tool_version: str,
    sheets: list[SheetRecord],
    named_ranges: list[NamedRangeRecord],
    formulas: list[FormulaRecord],
    validations: list[ValidationRecord],
    vba_modules: list[VbaModuleRecord],
    vba_procedures: list[VbaProcedureRecord],
    dep_edges: list[DependencyEdge],
    cycles: list[CycleRecord],
    orphan_ids: list[str],
    raw_warnings: list[str],
) -> SummaryRecord:
    """Assemble the complete SummaryRecord from all phase outputs.

    Args:
        input_path: Source .xlsm path.
        tool_version: xlsm-archaeologist.__version__.
        *: Per-phase results.
        raw_warnings: Plain string warnings collected across all phases.

    Returns:
        Fully populated SummaryRecord.
    """
    stats = compute_stats(
        sheets, named_ranges, formulas, validations,
        vba_modules, vba_procedures, dep_edges,
    )
    risk = compute_risk_indicators(formulas, vba_procedures, cycles, orphan_ids, dep_edges)
    score = compute_complexity_score(risk, stats)
    difficulty = derive_migration_difficulty(score)

    # Convert raw warning strings to SummaryWarning objects
    from xlsm_archaeologist.models.summary import WarningLevel

    sw_list: list[SummaryWarning] = []
    for msg in raw_warnings:
        # Heuristically categorise from prefix
        level: WarningLevel = "warning"
        category: str = "general"
        location: str = ""

        if msg.startswith("VBA"):
            category = "vba"
        elif "formula" in msg.lower() or "parse" in msg.lower():
            category = "formula"
        elif "cycle" in msg.lower() or "circular" in msg.lower():
            category = "dependency"
            level = "error"
        elif "orphan" in msg.lower():
            category = "dependency"

        sw_list.append(
            SummaryWarning(level=level, category=category, location=location, message=msg)
        )

    # Sort: error first, then warning, then info
    _order = {"error": 0, "warning": 1, "info": 2}
    sw_list.sort(key=lambda w: (_order.get(w.level, 9), w.category, w.message))

    return SummaryRecord(
        schema_version="1.0",
        tool_version=tool_version,
        analyzed_at=datetime.datetime.now(tz=datetime.UTC).isoformat(),
        input_file=str(input_path),
        stats=stats,
        risk_indicators=risk,
        complexity_score=score,
        migration_difficulty=difficulty,
        warnings=sw_list,
    )
