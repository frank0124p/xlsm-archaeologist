"""Pydantic models for the run summary (Phase 6)."""

from __future__ import annotations

from typing import Literal

from pydantic import BaseModel, ConfigDict, Field

MigrationDifficulty = Literal["low", "medium", "high", "very_high"]
WarningLevel = Literal["info", "warning", "error"]


class Stats(BaseModel):
    """Raw counts from the extraction phases."""

    model_config = ConfigDict(frozen=True)

    sheet_count: int = Field(description="Number of worksheets")
    named_range_count: int = Field(description="Number of defined names")
    formula_count: int = Field(description="Number of formula cells")
    validation_count: int = Field(description="Number of data-validation rules")
    vba_module_count: int = Field(description="Number of VBA modules")
    vba_procedure_count: int = Field(description="Number of VBA procedures")
    dependency_edge_count: int = Field(description="Total directed dependency edges")


class RiskIndicators(BaseModel):
    """Indicators that raise migration complexity."""

    model_config = ConfigDict(frozen=True)

    circular_reference_count: int = Field(description="Number of detected cycles")
    external_reference_count: int = Field(
        description="Formulas referencing external workbooks"
    )
    volatile_function_count: int = Field(
        description="Formulas using NOW/RAND/OFFSET/INDIRECT etc."
    )
    dynamic_vba_range_count: int = Field(
        description="VBA procedures with dynamic range access"
    )
    deeply_nested_formula_count: int = Field(
        description="Formulas with nesting_depth > 5"
    )
    orphan_formula_count: int = Field(description="Formula cells with no dependents")
    cross_sheet_dependency_count: int = Field(
        description="Dependency edges crossing sheet boundaries"
    )


class SummaryWarning(BaseModel):
    """A single warning entry collected during the run."""

    model_config = ConfigDict(frozen=True)

    level: WarningLevel = Field(description="Severity level")
    category: str = Field(description="Phase or component that raised the warning")
    location: str = Field(default="", description="Cell address or module.procedure")
    message: str = Field(description="Human-readable warning description")


class SummaryRecord(BaseModel):
    """Top-level run summary written to 00_summary.json."""

    model_config = ConfigDict(frozen=True)

    schema_version: str = Field(default="1.0", description="Output schema version")
    tool_version: str = Field(description="xlsm-archaeologist version string")
    analyzed_at: str = Field(description="ISO-8601 timestamp when analysis completed")
    input_file: str = Field(description="Absolute path of the analysed file")
    stats: Stats = Field(description="Raw extraction counts")
    risk_indicators: RiskIndicators = Field(description="Risk and complexity indicators")
    complexity_score: int = Field(description="Composite complexity score")
    migration_difficulty: MigrationDifficulty = Field(
        description="Derived migration difficulty label"
    )
    warnings: list[SummaryWarning] = Field(
        default_factory=list, description="Aggregated warnings from all phases"
    )
