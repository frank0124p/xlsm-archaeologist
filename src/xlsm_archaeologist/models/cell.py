"""Pydantic models for cell and data-validation records."""

from __future__ import annotations

from typing import Literal

from pydantic import BaseModel, ConfigDict, Field

ValueType = Literal["number", "string", "boolean", "date", "error", "empty"]
ValidationType = Literal["list", "whole", "decimal", "date", "time", "length", "custom"]


class CellRecord(BaseModel):
    """A single 'meaningful' cell extracted from a worksheet.

    Only cells that have a formula, validation, or are targeted by a named
    range are recorded here.  ``is_referenced`` is always False in Phase 2;
    Phase 5 back-fills it after the dependency graph is built.

    One row per cell in 04_cells.csv.
    """

    model_config = ConfigDict(frozen=True)

    cell_id: int = Field(description="1-based sequential ID across all sheets")
    sheet_name: str = Field(description="Name of the containing sheet")
    cell_address: str = Field(description="A1-notation without sheet prefix, e.g. 'B7'")
    qualified_address: str = Field(description="Sheet-qualified unique address, e.g. 'Calc!B7'")
    cell_row: int = Field(description="1-based row number")
    cell_col: int = Field(description="1-based column number (A=1)")
    has_formula: bool = Field(description="True when the cell contains a formula")
    has_validation: bool = Field(description="True when the cell is covered by a data validation")
    is_named: bool = Field(description="True when the cell is the target of a named range")
    is_referenced: bool = Field(
        description="True when referenced by another cell or VBA — filled in Phase 5"
    )
    value_type: ValueType = Field(description="Type of the stored value")
    raw_value: str = Field(
        description="String representation of the stored value; empty string for empty cells"
    )


class ValidationRecord(BaseModel):
    """A data validation rule attached to one or more cells.

    One row per (rule, sqref-token) pair in 06_validations.csv.
    """

    model_config = ConfigDict(frozen=True)

    validation_id: int = Field(description="1-based sequential ID")
    qualified_address: str = Field(
        description="Sheet-qualified address of the first cell in the validation range"
    )
    range_text: str = Field(description="Full sqref string as stored by openpyxl, e.g. 'A2:A100'")
    validation_type: ValidationType = Field(
        description="Validation rule type from the openpyxl DataValidation.type"
    )
    formula1: str = Field(
        default="",
        description="First condition expression (list source, upper bound, etc.)",
    )
    formula2: str = Field(
        default="",
        description="Second condition expression (lower bound); empty when not applicable",
    )
    enum_values: str = Field(
        default="",
        description=(
            "Pipe-separated parsed options for 'list' type validations. Empty for non-list types."
        ),
    )
    allow_blank: bool = Field(description="True when blank cells pass validation")
    error_title: str = Field(default="", description="Title of the validation error dialog")
    error_message: str = Field(default="", description="Body text of the validation error dialog")
