"""End-to-end test: full analyze pipeline on formulas_complex.xlsm."""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from xlsm_archaeologist.runner import run_extraction


@pytest.fixture
def fixtures_dir() -> Path:
    return Path(__file__).parent / "fixtures"


@pytest.fixture
def output_dir(tmp_path: Path) -> Path:
    return tmp_path / "output"


def test_e2e_all_output_files_exist(fixtures_dir: Path, output_dir: Path) -> None:
    """Run analyze on formulas_complex.xlsm and verify all 11+ output files exist."""
    run_extraction(
        input_path=fixtures_dir / "formulas_complex.xlsm",
        output_dir=output_dir,
        quiet=True,
    )

    expected = [
        "00_summary.json",
        "01_workbook.json",
        "02_sheets.csv",
        "03_named_ranges.csv",
        "04_cells.csv",
        "05_formulas.json",
        "06_validations.csv",
        "07_vba_modules.json",
        "08_vba_procedures.json",
        "09_dependencies.csv",
        "10_dependency_graph.json",
        "reports/cycles.json",
        "reports/formula_categories.csv",
        "reports/top_complex_formulas.csv",
        "reports/hotspot_cells.csv",
        "reports/cross_sheet_refs.csv",
    ]
    for rel_path in expected:
        assert (output_dir / rel_path).exists(), f"Missing output file: {rel_path}"


def test_e2e_summary_valid_structure(fixtures_dir: Path, output_dir: Path) -> None:
    """Verify 00_summary.json has the correct top-level keys and schema_version."""
    run_extraction(
        input_path=fixtures_dir / "formulas_complex.xlsm",
        output_dir=output_dir,
        quiet=True,
    )
    with open(output_dir / "00_summary.json") as fh:
        data = json.load(fh)

    assert data["schema_version"] == "1.0"
    assert "stats" in data
    assert "risk_indicators" in data
    assert "complexity_score" in data
    assert "migration_difficulty" in data
    assert data["migration_difficulty"] in ("low", "medium", "high", "very_high")


def test_e2e_deterministic(fixtures_dir: Path, tmp_path: Path) -> None:
    """Running analyze twice on same input produces identical output."""
    out1 = tmp_path / "run1"
    out2 = tmp_path / "run2"

    run_extraction(
        input_path=fixtures_dir / "formulas_complex.xlsm",
        output_dir=out1,
        quiet=True,
    )
    run_extraction(
        input_path=fixtures_dir / "formulas_complex.xlsm",
        output_dir=out2,
        quiet=True,
    )

    for filename in ["02_sheets.csv", "03_named_ranges.csv", "04_cells.csv", "09_dependencies.csv"]:
        content1 = (out1 / filename).read_text()
        content2 = (out2 / filename).read_text()
        assert content1 == content2, f"Non-deterministic output: {filename}"

    # Compare formula counts (ignore analyzed_at timestamp)
    s1 = json.loads((out1 / "00_summary.json").read_text())
    s2 = json.loads((out2 / "00_summary.json").read_text())
    assert s1["stats"] == s2["stats"]
    assert s1["complexity_score"] == s2["complexity_score"]
