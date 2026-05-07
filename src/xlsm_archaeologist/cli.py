"""Typer CLI entry point: version / analyze / inspect commands."""

from __future__ import annotations

import sys
from importlib.metadata import PackageNotFoundError
from importlib.metadata import version as pkg_version
from pathlib import Path
from typing import Annotated

import typer

from xlsm_archaeologist import __version__

app = typer.Typer(
    name="xlsm-archaeologist",
    help="Archaeologize complex .xlsm files into structured JSON/CSV data.",
    add_completion=False,
)


def _get_pkg_version(name: str) -> str:
    """Return installed package version or 'unknown'."""
    try:
        return pkg_version(name)
    except PackageNotFoundError:
        return "unknown"


@app.command()
def version() -> None:
    """Show tool and dependency versions."""
    py = f"{sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro}"
    lines = [
        f"xlsm-archaeologist {__version__}",
        "schema_version: 1.0",
        f"python: {py}",
        f"openpyxl: {_get_pkg_version('openpyxl')}",
        f"oletools: {_get_pkg_version('oletools')}",
    ]
    typer.echo("\n".join(lines))


@app.command()
def analyze(
    input_path: Annotated[Path, typer.Argument(help="Path to the .xlsm file to analyze")],
    output: Annotated[Path, typer.Option("--output", "-o", help="Output directory")] = Path(
        "./archaeology_output"
    ),
    phases: Annotated[str, typer.Option(help="Phases to run, comma-separated or 'all'")] = "all",
    no_vba: Annotated[bool, typer.Option("--no-vba", help="Skip VBA analysis")] = False,
    no_graph: Annotated[bool, typer.Option("--no-graph", help="Skip dependency graph")] = False,
    no_reports: Annotated[
        bool, typer.Option("--no-reports", help="Skip report generation")
    ] = False,  # noqa: E501
    max_formula_depth: Annotated[int, typer.Option(help="Max formula AST nesting depth")] = 20,
    log_level: Annotated[str, typer.Option(help="Log level: debug/info/warning/error")] = "info",
    quiet: Annotated[bool, typer.Option("--quiet", "-q", help="Suppress progress bar")] = False,
    force: Annotated[bool, typer.Option("--force", help="Overwrite non-empty output dir")] = False,
) -> None:
    """Perform full archaeological analysis of an .xlsm file."""
    typer.echo("not implemented in phase 1")


@app.command()
def inspect(
    input_path: Annotated[Path, typer.Argument(help="Path to the .xlsm file to inspect")],
) -> None:
    """Quickly display .xlsm overview without writing files."""
    typer.echo("not implemented in phase 1")
