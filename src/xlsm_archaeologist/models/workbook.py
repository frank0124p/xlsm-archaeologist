"""Pydantic models for workbook-level and sheet-level data."""

from __future__ import annotations

from pydantic import BaseModel, ConfigDict, Field


class WorkbookRecord(BaseModel):
    """Top-level metadata for a single .xlsm / .xlsx file.

    Produced by WorkbookExtractor and serialized into 01_workbook.json.
    """

    model_config = ConfigDict(frozen=True)

    file_path: str = Field(description="Original input path as provided to the CLI")
    file_sha256: str = Field(description="SHA-256 hex digest of the file bytes")
    size_bytes: int = Field(description="File size in bytes")
    has_vba: bool = Field(description="True when a VBA project is present in the archive")
    has_external_links: bool = Field(description="True when workbook contains [Book.xlsx] refs")
    default_sheet: str | None = Field(
        default=None, description="Name of the active sheet at save time"
    )
    created: str | None = Field(default=None, description="ISO-8601 creation timestamp if available")
    modified: str | None = Field(
        default=None, description="ISO-8601 last-modified timestamp if available"
    )
    author: str | None = Field(default=None, description="Workbook author from core properties")
    last_modified_by: str | None = Field(
        default=None, description="Last modifier from core properties"
    )


class SheetRecord(BaseModel):
    """Metadata for a single worksheet.

    One row per sheet in 02_sheets.csv.
    """

    model_config = ConfigDict(frozen=True)

    sheet_id: int = Field(description="1-based sequential ID")
    sheet_name: str = Field(description="Sheet name as stored in the workbook")
    sheet_index: int = Field(description="0-based position in workbook tab order")
    is_hidden: bool = Field(description="True for normally hidden sheets (xlSheetHidden)")
    is_very_hidden: bool = Field(
        description="True for sheets hidden via xlSheetVeryHidden (only VBA can reveal)"
    )
    used_range: str = Field(
        description="A1-notation bounding box of non-empty cells, e.g. 'A1:Z100'"
    )
    row_count: int = Field(description="Number of rows in used_range")
    col_count: int = Field(description="Number of columns in used_range")
    cell_count: int = Field(description="Total non-empty cells in the sheet (not just meaningful)")
    formula_cell_count: int = Field(description="Number of cells whose data_type == 'f' (formula)")
