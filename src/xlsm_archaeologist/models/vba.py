"""Pydantic models for VBA analysis (Phase 4)."""

from __future__ import annotations

from typing import Literal

from pydantic import BaseModel, ConfigDict, Field

VbaModuleType = Literal["standard", "class", "form", "sheet", "workbook", "unknown"]
ProcedureType = Literal["sub", "function", "property_get", "property_let", "property_set"]
RangeAccessVia = Literal[
    "explicit_range", "cells_method", "named_range", "dynamic", "unknown"
]


class RangeAccess(BaseModel):
    """A single cell/range read or write access detected in VBA code."""

    model_config = ConfigDict(frozen=True)

    sheet: str | None = Field(
        default=None, description="Sheet name if determinable; None = unknown or ActiveSheet"
    )
    range_ref: str = Field(description="Range address string, e.g. 'A1' or 'A1:B10'")
    via: RangeAccessVia = Field(description="How the range was referenced")


class EventTrigger(BaseModel):
    """An Excel event that triggers a VBA procedure."""

    model_config = ConfigDict(frozen=True)

    event: str = Field(description="Event name, e.g. 'Worksheet_Change'")
    target: str | None = Field(
        default=None, description="Target range from Intersect heuristic, if detectable"
    )


class Parameter(BaseModel):
    """A single VBA procedure parameter."""

    model_config = ConfigDict(frozen=True)

    name: str = Field(description="Parameter name")
    type_hint: str | None = Field(
        default=None, description="Declared VBA type, e.g. 'Integer', 'String'"
    )
    is_optional: bool = Field(default=False, description="True if preceded by Optional keyword")


class VbaModuleRecord(BaseModel):
    """Metadata and source for one VBA module extracted from the workbook."""

    model_config = ConfigDict(frozen=True)

    vba_module_id: int = Field(description="1-based unique module identifier")
    module_name: str = Field(description="Module name as declared in the VBA project")
    module_type: VbaModuleType = Field(description="Module classification")
    line_count: int = Field(description="Total source lines including blanks and comments")
    procedure_count: int = Field(description="Number of Sub/Function/Property procedures found")
    source_code: str = Field(description="Full module source code")


class VbaProcedureRecord(BaseModel):
    """Analysis record for one Sub, Function, or Property procedure."""

    model_config = ConfigDict(frozen=True)

    vba_procedure_id: int = Field(description="1-based unique procedure identifier")
    vba_module_id: int = Field(description="Parent VbaModuleRecord.vba_module_id")
    procedure_name: str = Field(description="Procedure name without module prefix")
    procedure_type: ProcedureType = Field(description="Sub / Function / Property variant")
    is_public: bool = Field(description="True if Public (or no scope keyword = default Public)")
    parameters: list[Parameter] = Field(
        default_factory=list, description="Declared parameters in order"
    )
    line_count: int = Field(description="Lines from first declaration line to End")
    reads: list[RangeAccess] = Field(
        default_factory=list, description="Cell/range read accesses detected"
    )
    writes: list[RangeAccess] = Field(
        default_factory=list, description="Cell/range write accesses detected"
    )
    calls: list[str] = Field(
        default_factory=list,
        description="Names of other known procedures called; sorted alphabetically",
    )
    triggers: list[EventTrigger] = Field(
        default_factory=list,
        description="Excel events this procedure handles (from naming convention)",
    )
    has_dynamic_range: bool = Field(
        description="True if any range access uses a dynamic/indeterminate reference"
    )
    dynamic_range_notes: list[str] = Field(
        default_factory=list,
        description="Human-readable notes for each dynamic range pattern found",
    )
    complexity_score: int = Field(
        description="line_count + len(reads) + len(writes) + branch_count"
    )
    source_code: str = Field(description="Extracted procedure source code")
