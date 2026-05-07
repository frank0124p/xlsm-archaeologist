"""Pydantic models for the dependency graph (Phase 5)."""

from __future__ import annotations

from typing import Literal

from pydantic import BaseModel, ConfigDict, Field

NodeType = Literal[
    "input_cell", "formula_cell", "output_cell", "named_range", "vba_procedure"
]
DependencyVia = Literal["formula", "vba_read_write", "validation", "named_range"]


class DependencyEdge(BaseModel):
    """One directed dependency edge in the workbook graph."""

    model_config = ConfigDict(frozen=True)

    dependency_id: int = Field(description="1-based unique edge identifier")
    source_qualified_address: str = Field(
        description="Source node id, e.g. 'Sheet1!A1' or '_named:TaxRate'"
    )
    target_qualified_address: str = Field(
        description="Target node id that depends on the source"
    )
    via: DependencyVia = Field(description="How this dependency was established")
    via_detail: str = Field(
        default="", description="Extra context, e.g. formula_id or procedure name"
    )
    is_cross_sheet: bool = Field(
        description="True when source and target belong to different sheets"
    )


class CycleRecord(BaseModel):
    """A detected circular reference cycle in the dependency graph."""

    model_config = ConfigDict(frozen=True)

    cycle_id: int = Field(description="1-based cycle identifier")
    length: int = Field(description="Number of nodes in the cycle")
    nodes: list[str] = Field(description="Ordered list of node ids forming the cycle")
    edges_via: list[str] = Field(
        description="Edge 'via' values corresponding to each node-to-node step"
    )
