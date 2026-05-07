"""Tests for Phase 2: Structure Extraction."""

from __future__ import annotations

import csv
import json
from pathlib import Path

import pytest
from typer.testing import CliRunner

from xlsm_archaeologist.cli import app
from xlsm_archaeologist.extractors.cell_extractor import extract_cells
from xlsm_archaeologist.extractors.named_range_extractor import extract_named_ranges
from xlsm_archaeologist.extractors.sheet_extractor import extract_sheets
from xlsm_archaeologist.extractors.validation_extractor import extract_validations
from xlsm_archaeologist.extractors.workbook_extractor import extract_workbook

runner = CliRunner()


@pytest.fixture()
def simple_xlsm(fixtures_dir: Path) -> Path:
    return fixtures_dir / "simple.xlsm"


@pytest.fixture()
def formulas_basic_xlsm(fixtures_dir: Path) -> Path:
    return fixtures_dir / "formulas_basic.xlsm"


@pytest.fixture()
def with_named_range_xlsm(fixtures_dir: Path) -> Path:
    return fixtures_dir / "with_named_range.xlsm"


@pytest.fixture()
def with_validation_xlsm(fixtures_dir: Path) -> Path:
    return fixtures_dir / "with_validation.xlsm"


@pytest.fixture()
def hidden_sheets_xlsm(fixtures_dir: Path) -> Path:
    return fixtures_dir / "hidden_sheets.xlsm"


# ─── Workbook metadata ────────────────────────────────────────────────────────


def test_workbook_extracts_sha256(simple_xlsm: Path) -> None:
    record, wb = extract_workbook(simple_xlsm)
    wb.close()
    assert len(record.file_sha256) == 64
    assert all(c in "0123456789abcdef" for c in record.file_sha256)


def test_workbook_detects_has_vba_false(simple_xlsm: Path) -> None:
    record, wb = extract_workbook(simple_xlsm)
    wb.close()
    assert record.has_vba is False


def test_workbook_size_positive(simple_xlsm: Path) -> None:
    record, wb = extract_workbook(simple_xlsm)
    wb.close()
    assert record.size_bytes > 0


# ─── Sheet extraction ─────────────────────────────────────────────────────────


def test_sheets_extracted_count(simple_xlsm: Path) -> None:
    _, wb = extract_workbook(simple_xlsm)
    sheets = list(extract_sheets(wb))
    wb.close()
    assert len(sheets) == 1
    assert sheets[0].sheet_name == "Data"


def test_used_range_calculated(formulas_basic_xlsm: Path) -> None:
    _, wb = extract_workbook(formulas_basic_xlsm)
    sheets = list(extract_sheets(wb))
    wb.close()
    assert sheets[0].used_range != ""
    assert ":" in sheets[0].used_range


def test_formula_cell_count(formulas_basic_xlsm: Path) -> None:
    _, wb = extract_workbook(formulas_basic_xlsm)
    sheets = list(extract_sheets(wb))
    wb.close()
    assert sheets[0].formula_cell_count >= 4


def test_hidden_sheet_marked(hidden_sheets_xlsm: Path) -> None:
    _, wb = extract_workbook(hidden_sheets_xlsm)
    sheets = list(extract_sheets(wb))
    wb.close()
    hidden = [s for s in sheets if s.sheet_name == "HiddenSheet"]
    assert len(hidden) == 1
    assert hidden[0].is_hidden is True
    assert hidden[0].is_very_hidden is False


def test_very_hidden_sheet_marked(hidden_sheets_xlsm: Path) -> None:
    _, wb = extract_workbook(hidden_sheets_xlsm)
    sheets = list(extract_sheets(wb))
    wb.close()
    very = [s for s in sheets if s.sheet_name == "VeryHiddenSheet"]
    assert len(very) == 1
    assert very[0].is_very_hidden is True
    assert very[0].is_hidden is False


def test_visible_sheet_not_hidden(hidden_sheets_xlsm: Path) -> None:
    _, wb = extract_workbook(hidden_sheets_xlsm)
    sheets = list(extract_sheets(wb))
    wb.close()
    vis = [s for s in sheets if s.sheet_name == "Visible"]
    assert len(vis) == 1
    assert vis[0].is_hidden is False
    assert vis[0].is_very_hidden is False


# ─── Named ranges ─────────────────────────────────────────────────────────────


def test_named_range_basic(with_named_range_xlsm: Path) -> None:
    _, wb = extract_workbook(with_named_range_xlsm)
    nrs = list(extract_named_ranges(wb))
    wb.close()
    names = [nr.range_name for nr in nrs]
    assert "TaxRate" in names


def test_named_range_dynamic_detection(with_named_range_xlsm: Path) -> None:
    _, wb = extract_workbook(with_named_range_xlsm)
    nrs = list(extract_named_ranges(wb))
    wb.close()
    dynamic = [nr for nr in nrs if nr.range_name == "DynamicRange"]
    assert len(dynamic) == 1
    assert dynamic[0].has_dynamic_formula is True


def test_named_range_static_not_dynamic(with_named_range_xlsm: Path) -> None:
    _, wb = extract_workbook(with_named_range_xlsm)
    nrs = list(extract_named_ranges(wb))
    wb.close()
    tax = [nr for nr in nrs if nr.range_name == "TaxRate"]
    assert len(tax) == 1
    assert tax[0].has_dynamic_formula is False
    assert tax[0].is_valid is True


# ─── Validations ──────────────────────────────────────────────────────────────


def test_validation_list_literal_parsed(with_validation_xlsm: Path) -> None:
    _, wb = extract_workbook(with_validation_xlsm)
    vals = list(extract_validations(wb))
    wb.close()
    literal_vals = [v for v in vals if "A2" in v.qualified_address]
    assert len(literal_vals) >= 1
    enum_vals = literal_vals[0].enum_values.split("|")
    assert "Active" in enum_vals
    assert "Inactive" in enum_vals
    assert "Pending" in enum_vals


def test_validation_list_range_reference(with_validation_xlsm: Path) -> None:
    _, wb = extract_workbook(with_validation_xlsm)
    vals = list(extract_validations(wb))
    wb.close()
    range_vals = [v for v in vals if "B2" in v.qualified_address]
    assert len(range_vals) >= 1
    enum_vals = range_vals[0].enum_values.split("|")
    assert "Cat A" in enum_vals
    assert "Cat B" in enum_vals
    assert "Cat C" in enum_vals


def test_validation_has_error_message(with_validation_xlsm: Path) -> None:
    _, wb = extract_workbook(with_validation_xlsm)
    vals = list(extract_validations(wb))
    wb.close()
    with_msg = [v for v in vals if v.error_message]
    assert len(with_msg) >= 1


# ─── Cell filter ──────────────────────────────────────────────────────────────


def test_cell_filter_meaningful_only_empty_for_simple(simple_xlsm: Path) -> None:
    _, wb = extract_workbook(simple_xlsm)
    cells = list(extract_cells(wb, set(), set()))
    wb.close()
    assert len(cells) == 0  # pure values, no formula/validation/named


def test_formula_cells_recorded(formulas_basic_xlsm: Path) -> None:
    _, wb = extract_workbook(formulas_basic_xlsm)
    cells = list(extract_cells(wb, set(), set()))
    wb.close()
    formula_cells = [c for c in cells if c.has_formula]
    assert len(formula_cells) >= 4


def test_is_referenced_false_in_phase2(formulas_basic_xlsm: Path) -> None:
    _, wb = extract_workbook(formulas_basic_xlsm)
    cells = list(extract_cells(wb, set(), set()))
    wb.close()
    assert all(c.is_referenced is False for c in cells)


# ─── CLI integration & determinism ───────────────────────────────────────────


def test_analyze_writes_expected_files(simple_xlsm: Path, tmp_path: Path) -> None:
    result = runner.invoke(app, ["analyze", str(simple_xlsm), "-o", str(tmp_path), "-q"])
    assert result.exit_code == 0
    assert (tmp_path / "01_workbook.json").exists()
    assert (tmp_path / "02_sheets.csv").exists()
    assert (tmp_path / "03_named_ranges.csv").exists()
    assert (tmp_path / "04_cells.csv").exists()
    assert (tmp_path / "06_validations.csv").exists()


def test_workbook_json_has_schema_version(simple_xlsm: Path, tmp_path: Path) -> None:
    runner.invoke(app, ["analyze", str(simple_xlsm), "-o", str(tmp_path), "-q"])
    data = json.loads((tmp_path / "01_workbook.json").read_text(encoding="utf-8"))
    assert data["schema_version"] == "1.0"


def test_sheets_csv_header_correct(simple_xlsm: Path, tmp_path: Path) -> None:
    runner.invoke(app, ["analyze", str(simple_xlsm), "-o", str(tmp_path), "-q"])
    with (tmp_path / "02_sheets.csv").open(encoding="utf-8-sig") as fh:
        reader = csv.DictReader(fh)
        assert reader.fieldnames is not None
        expected = [
            "sheet_id",
            "sheet_name",
            "sheet_index",
            "is_hidden",
            "is_very_hidden",
            "used_range",
            "row_count",
            "col_count",
            "cell_count",
            "formula_cell_count",
        ]
        assert list(reader.fieldnames) == expected


def test_extraction_deterministic(simple_xlsm: Path, tmp_path: Path) -> None:
    out1 = tmp_path / "run1"
    out2 = tmp_path / "run2"
    runner.invoke(app, ["analyze", str(simple_xlsm), "-o", str(out1), "-q"])
    runner.invoke(app, ["analyze", str(simple_xlsm), "-o", str(out2), "-q"])
    for fname in [
        "01_workbook.json",
        "02_sheets.csv",
        "03_named_ranges.csv",
        "04_cells.csv",
        "06_validations.csv",
    ]:
        assert (out1 / fname).read_bytes() == (out2 / fname).read_bytes(), (
            f"{fname} differs between runs"
        )


def test_csv_row_order_stable(formulas_basic_xlsm: Path, tmp_path: Path) -> None:
    runner.invoke(app, ["analyze", str(formulas_basic_xlsm), "-o", str(tmp_path), "-q"])
    with (tmp_path / "04_cells.csv").open(encoding="utf-8-sig") as fh:
        rows = list(csv.DictReader(fh))
    addresses = [r["qualified_address"] for r in rows]
    assert addresses == sorted(addresses)
