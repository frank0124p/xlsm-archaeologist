"""Smoke tests for Phase 1: CLI skeleton."""

from __future__ import annotations

from typer.testing import CliRunner

from xlsm_archaeologist.cli import app

runner = CliRunner()


def test_version_command_works() -> None:
    """version command exits 0 and prints at least 4 lines."""
    result = runner.invoke(app, ["version"])
    assert result.exit_code == 0
    lines = [ln for ln in result.output.splitlines() if ln.strip()]
    assert len(lines) >= 4
    assert "xlsm-archaeologist" in lines[0]
    assert "schema_version" in lines[1]
    assert "python" in lines[2]
    assert "openpyxl" in lines[3]


def test_analyze_command_callable() -> None:
    """analyze command is callable; returns exit 1 for a missing file (Phase 2+)."""
    result = runner.invoke(app, ["analyze", "nonexistent_file.xlsm"])
    # Phase 2 implemented: missing file → exit 1 with error message
    assert result.exit_code == 1


def test_inspect_command_callable() -> None:
    """inspect command exits 0 and prints stub message."""
    result = runner.invoke(app, ["inspect", "dummy.xlsm"])
    assert result.exit_code == 0
    assert "not implemented in phase 1" in result.output


def test_help_lists_three_commands() -> None:
    """--help output includes version, analyze, and inspect commands."""
    result = runner.invoke(app, ["--help"])
    assert result.exit_code == 0
    assert "version" in result.output
    assert "analyze" in result.output
    assert "inspect" in result.output


def test_logging_writes_to_stderr() -> None:
    """Standard output is not polluted by log messages from version command."""
    result = runner.invoke(app, ["version"])
    # All log noise goes to stderr; stdout should only contain the version lines
    assert result.exit_code == 0
    for line in result.output.splitlines():
        if line.strip():
            assert any(
                kw in line
                for kw in ("xlsm-archaeologist", "schema_version", "python", "openpyxl", "oletools")
            ), f"Unexpected line in stdout: {line!r}"
