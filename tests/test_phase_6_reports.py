"""Phase 6 reports and summary tests."""

from __future__ import annotations

from pathlib import Path

import pytest

from xlsm_archaeologist.analyzers.summary_analyzer import (
    compute_complexity_score,
    compute_risk_indicators,
    compute_stats,
    derive_migration_difficulty,
)
from xlsm_archaeologist.models.cell import CellRecord, ValidationRecord, ValueType
from xlsm_archaeologist.models.dependency import CycleRecord, DependencyEdge
from xlsm_archaeologist.models.formula import CellRef, FormulaRecord
from xlsm_archaeologist.models.named_range import NamedRangeRecord
from xlsm_archaeologist.models.summary import RiskIndicators, Stats
from xlsm_archaeologist.models.vba import VbaModuleRecord, VbaProcedureRecord
from xlsm_archaeologist.models.workbook import SheetRecord
from xlsm_archaeologist.reports.cross_sheet_refs_report import build_cross_sheet_refs
from xlsm_archaeologist.reports.formula_categories_report import build_categories_report
from xlsm_archaeologist.reports.hotspot_cells_report import build_hotspot_cells
from xlsm_archaeologist.reports.top_complex_formulas_report import build_top_complex_formulas
from xlsm_archaeologist.reports.vba_behavior_report import build_vba_behavior


@pytest.fixture
def fixtures_dir() -> Path:
    return Path(__file__).parent / "fixtures"


def _make_sheet(idx: int) -> SheetRecord:
    return SheetRecord(
        sheet_id=idx,
        sheet_name=f"Sheet{idx}",
        sheet_index=idx - 1,
        is_hidden=False,
        is_very_hidden=False,
        used_range="A1:B10",
        row_count=10,
        col_count=2,
        cell_count=5,
        formula_cell_count=1,
    )


def _make_formula(fid: int, qa: str, cat: str = "compute", depth: int = 0, score: int = 0) -> FormulaRecord:
    return FormulaRecord(
        formula_id=fid,
        qualified_address=qa,
        formula_text="=A1",
        formula_category=cat,  # type: ignore[arg-type]
        function_list=[],
        referenced_cells=[],
        referenced_named_ranges=[],
        has_external_reference=False,
        is_volatile=False,
        is_array_formula=False,
        nesting_depth=depth,
        function_count=0,
        complexity_score=score,
        ast=None,
        is_parsable=True,
        parse_error=None,
    )


def _make_risk(**kwargs: int) -> RiskIndicators:
    defaults = dict(
        circular_reference_count=0,
        external_reference_count=0,
        volatile_function_count=0,
        dynamic_vba_range_count=0,
        deeply_nested_formula_count=0,
        orphan_formula_count=0,
        cross_sheet_dependency_count=0,
    )
    defaults.update(kwargs)
    return RiskIndicators(**defaults)


# ---------------------------------------------------------------------------
# Complexity score tests
# ---------------------------------------------------------------------------


def test_complexity_score_calculation() -> None:
    stats = Stats(
        sheet_count=2,
        named_range_count=5,
        formula_count=100,
        validation_count=10,
        vba_module_count=2,
        vba_procedure_count=10,
        dependency_edge_count=50,
    )
    risk = _make_risk(circular_reference_count=1)
    score = compute_complexity_score(risk, stats)
    # 100//10 + 10*5 + 1*50 = 10 + 50 + 50 = 110
    assert score == 110


def test_complexity_score_zero_for_empty() -> None:
    stats = Stats(
        sheet_count=1, named_range_count=0, formula_count=0,
        validation_count=0, vba_module_count=0, vba_procedure_count=0,
        dependency_edge_count=0,
    )
    risk = _make_risk()
    assert compute_complexity_score(risk, stats) == 0


# ---------------------------------------------------------------------------
# Migration difficulty thresholds
# ---------------------------------------------------------------------------


def test_migration_difficulty_thresholds() -> None:
    assert derive_migration_difficulty(0) == "low"
    assert derive_migration_difficulty(49) == "low"
    assert derive_migration_difficulty(50) == "medium"
    assert derive_migration_difficulty(199) == "medium"
    assert derive_migration_difficulty(200) == "high"
    assert derive_migration_difficulty(499) == "high"
    assert derive_migration_difficulty(500) == "very_high"
    assert derive_migration_difficulty(9999) == "very_high"


# ---------------------------------------------------------------------------
# Formula categories report
# ---------------------------------------------------------------------------


def test_formula_categories_aggregation() -> None:
    formulas = [
        _make_formula(1, "Sheet1!A1", "compute", score=3),
        _make_formula(2, "Sheet1!A2", "compute", score=5),
        _make_formula(3, "Sheet1!A3", "lookup", score=10),
    ]
    rows = build_categories_report(formulas)
    cats = {r["category"]: r for r in rows}
    assert cats["compute"]["formula_count"] == 2
    assert cats["compute"]["total_complexity"] == 8
    assert cats["lookup"]["formula_count"] == 1
    pct = cats["compute"]["pct_of_total"]
    assert abs(float(pct) - 66.67) < 0.1  # type: ignore[arg-type]


def test_formula_categories_empty() -> None:
    assert build_categories_report([]) == []


# ---------------------------------------------------------------------------
# Top complex formulas
# ---------------------------------------------------------------------------


def test_top_complex_top50_limit() -> None:
    formulas = [_make_formula(i, f"Sheet1!A{i}", score=i) for i in range(1, 101)]
    rows = build_top_complex_formulas(formulas, top_n=50)
    assert len(rows) == 50
    # Top entry should be the most complex (i=100)
    assert rows[0]["complexity_score"] == 100
    assert rows[0]["rank"] == 1


def test_top_complex_sorted_desc() -> None:
    formulas = [
        _make_formula(1, "Sheet1!A1", score=5),
        _make_formula(2, "Sheet1!A2", score=20),
        _make_formula(3, "Sheet1!A3", score=1),
    ]
    rows = build_top_complex_formulas(formulas)
    scores = [r["complexity_score"] for r in rows]
    assert scores == sorted(scores, reverse=True)


# ---------------------------------------------------------------------------
# Cross-sheet refs report
# ---------------------------------------------------------------------------


def test_cross_sheet_refs_only_cross() -> None:
    edges = [
        DependencyEdge(
            dependency_id=1,
            source_qualified_address="Sheet1!A1",
            target_qualified_address="Sheet2!B1",
            via="formula",
            via_detail="1",
            is_cross_sheet=True,
        ),
        DependencyEdge(
            dependency_id=2,
            source_qualified_address="Sheet1!A1",
            target_qualified_address="Sheet1!C1",
            via="formula",
            via_detail="2",
            is_cross_sheet=False,
        ),
    ]
    rows = build_cross_sheet_refs(edges)
    assert len(rows) == 1
    assert rows[0]["source_sheet"] == "Sheet1"
    assert rows[0]["target_sheet"] == "Sheet2"


# ---------------------------------------------------------------------------
# Hotspot cells
# ---------------------------------------------------------------------------


def test_hotspot_in_degree_correctness() -> None:
    import networkx as nx

    graph: nx.DiGraph = nx.DiGraph()
    graph.add_node("Sheet1!A1", node_type="input_cell", value_type="number")
    graph.add_node("Sheet1!B1", node_type="formula_cell", value_type="number")
    graph.add_node("Sheet1!C1", node_type="formula_cell", value_type="number")
    # A1 is referenced by both B1 and C1
    graph.add_edge("Sheet1!A1", "Sheet1!B1", via="formula", via_detail="1")
    graph.add_edge("Sheet1!A1", "Sheet1!C1", via="formula", via_detail="2")

    vtype: ValueType = "number"
    cells = [
        CellRecord(
            cell_id=1, sheet_name="Sheet1", cell_address="A1",
            qualified_address="Sheet1!A1", cell_row=1, cell_col=1,
            has_formula=False, has_validation=False, is_named=False,
            is_referenced=True, value_type=vtype, raw_value="5",
        ),
    ]
    rows = build_hotspot_cells(graph, cells)
    top = rows[0]
    assert top["qualified_address"] == "Sheet1!A1"
    assert top["in_degree"] == 2


# ---------------------------------------------------------------------------
# Summary warnings sorted by level
# ---------------------------------------------------------------------------


def test_warnings_sorted_by_level() -> None:
    from xlsm_archaeologist.reports.summary_builder import build_summary

    summary = build_summary(
        input_path=Path("test.xlsm"),
        tool_version="0.1.0",
        sheets=[_make_sheet(1)],
        named_ranges=[],
        formulas=[],
        validations=[],
        vba_modules=[],
        vba_procedures=[],
        dep_edges=[],
        cycles=[CycleRecord(cycle_id=1, length=2, nodes=["A", "B"], edges_via=["formula"])],
        orphan_ids=[],
        raw_warnings=["Detected 1 circular reference cycle(s)", "Some parse warning"],
    )
    levels = [w.level for w in summary.warnings]
    _order = {"error": 0, "warning": 1, "info": 2}
    sorted_levels = sorted(levels, key=lambda lv: _order.get(lv, 9))
    assert levels == sorted_levels


def test_summary_schema_version() -> None:
    from xlsm_archaeologist.reports.summary_builder import build_summary

    summary = build_summary(
        input_path=Path("test.xlsm"),
        tool_version="0.1.0",
        sheets=[],
        named_ranges=[],
        formulas=[],
        validations=[],
        vba_modules=[],
        vba_procedures=[],
        dep_edges=[],
        cycles=[],
        orphan_ids=[],
        raw_warnings=[],
    )
    assert summary.schema_version == "1.0"
    assert summary.stats.sheet_count == 0
