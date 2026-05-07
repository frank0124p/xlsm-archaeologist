"""Compute complexity score, risk indicators, and migration difficulty."""

from __future__ import annotations

from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from xlsm_archaeologist.models.cell import ValidationRecord
    from xlsm_archaeologist.models.dependency import CycleRecord, DependencyEdge
    from xlsm_archaeologist.models.formula import FormulaRecord
    from xlsm_archaeologist.models.named_range import NamedRangeRecord
    from xlsm_archaeologist.models.vba import VbaModuleRecord, VbaProcedureRecord
    from xlsm_archaeologist.models.workbook import SheetRecord

from xlsm_archaeologist.models.summary import MigrationDifficulty, RiskIndicators, Stats

_DEEP_NESTING_THRESHOLD = 5


def compute_stats(
    sheets: list[SheetRecord],
    named_ranges: list[NamedRangeRecord],
    formulas: list[FormulaRecord],
    validations: list[ValidationRecord],
    vba_modules: list[VbaModuleRecord],
    vba_procedures: list[VbaProcedureRecord],
    dep_edges: list[DependencyEdge],
) -> Stats:
    """Aggregate raw counts from all phases."""
    return Stats(
        sheet_count=len(sheets),
        named_range_count=len(named_ranges),
        formula_count=len(formulas),
        validation_count=len(validations),
        vba_module_count=len(vba_modules),
        vba_procedure_count=len(vba_procedures),
        dependency_edge_count=len(dep_edges),
    )


def compute_risk_indicators(
    formulas: list[FormulaRecord],
    vba_procedures: list[VbaProcedureRecord],
    cycles: list[CycleRecord],
    orphan_ids: list[str],
    dep_edges: list[DependencyEdge],
) -> RiskIndicators:
    """Derive risk indicators from analysis results."""
    external_refs = sum(1 for f in formulas if f.has_external_reference)
    volatile_count = sum(1 for f in formulas if f.is_volatile)
    dynamic_vba = sum(1 for p in vba_procedures if p.has_dynamic_range)
    deep_nested = sum(1 for f in formulas if f.nesting_depth > _DEEP_NESTING_THRESHOLD)
    cross_sheet = sum(1 for e in dep_edges if e.is_cross_sheet)

    return RiskIndicators(
        circular_reference_count=len(cycles),
        external_reference_count=external_refs,
        volatile_function_count=volatile_count,
        dynamic_vba_range_count=dynamic_vba,
        deeply_nested_formula_count=deep_nested,
        orphan_formula_count=len(orphan_ids),
        cross_sheet_dependency_count=cross_sheet,
    )


def compute_complexity_score(
    risk: RiskIndicators,
    stats: Stats,
) -> int:
    """Derive a composite complexity score from risk indicators and stats.

    Higher score = more complex migration.
    """
    score = (
        stats.formula_count // 10
        + stats.vba_procedure_count * 5
        + risk.circular_reference_count * 50
        + risk.external_reference_count * 20
        + risk.volatile_function_count * 3
        + risk.dynamic_vba_range_count * 15
        + risk.deeply_nested_formula_count * 5
        + risk.orphan_formula_count * 2
        + risk.cross_sheet_dependency_count // 5
    )
    return score


def derive_migration_difficulty(score: int) -> MigrationDifficulty:
    """Map a complexity score to a migration difficulty label.

    Thresholds (chosen to match the reference example with score=847):
    - < 50   → low
    - < 200  → medium
    - < 500  → high
    - >= 500 → very_high
    """
    if score < 50:
        return "low"
    if score < 200:
        return "medium"
    if score < 500:
        return "high"
    return "very_high"
