"""Pydantic model for named range records."""

from __future__ import annotations

from pydantic import BaseModel, ConfigDict, Field


class NamedRangeRecord(BaseModel):
    """A single named range (defined name) in the workbook.

    One row per named range in 03_named_ranges.csv.
    """

    model_config = ConfigDict(frozen=True)

    named_range_id: int = Field(description="1-based sequential ID")
    range_name: str = Field(description="Defined name as stored (original casing)")
    scope: str = Field(
        description="'workbook' for workbook-scoped names, or the sheet name for local names"
    )
    refers_to: str = Field(
        description="Full refers_to string including leading '=' and absolute refs, e.g. '=Params!$B$2'"
    )
    has_dynamic_formula: bool = Field(
        description="True when refers_to contains OFFSET, INDIRECT, or INDEX (volatile/dynamic)"
    )
    is_valid: bool = Field(
        description="False when refers_to contains '#REF!' or cannot be resolved"
    )
