"""Pydantic models for formula analysis (Phase 3)."""

from __future__ import annotations

from typing import Annotated, Any, Literal

from pydantic import BaseModel, ConfigDict, Field

FormulaCategory = Literal[
    "lookup", "branch", "compute", "aggregate", "text", "reference", "mixed"
]


class AstNode(BaseModel):
    """Base class for all AST node types."""

    model_config = ConfigDict(frozen=True)
    type: str


class FunctionNode(AstNode):
    """AST node representing a function call, e.g. IF(A1>0, 1, 0)."""

    type: Literal["function"] = "function"
    name: str = Field(description="Uppercase function name, e.g. 'IF'")
    args: list[AstNode] = Field(default_factory=list, description="Ordered argument nodes")


class OperandNode(AstNode):
    """AST node for a literal value (number, text, bool, error)."""

    type: Literal["operand"] = "operand"
    operand_type: Literal["number", "text", "boolean", "error"] = Field(
        description="Literal value type"
    )
    value: str = Field(description="Raw string representation of the literal")


class OperatorNode(AstNode):
    """AST node for a binary or unary operator."""

    type: Literal["operator"] = "operator"
    op: str = Field(description="Operator symbol, e.g. '+', '>', '&'")
    left: AstNode | None = Field(default=None, description="Left operand (None for unary)")
    right: AstNode = Field(description="Right operand")


class RangeNode(AstNode):
    """AST node for a cell or range reference, e.g. A1 or Sheet2!A1:B10."""

    type: Literal["range"] = "range"
    sheet: str | None = Field(default=None, description="Sheet name if cross-sheet; else None")
    address: str = Field(description="Cell or range address without sheet prefix")


class NamedRangeNode(AstNode):
    """AST node for a named range reference, e.g. TaxRate."""

    type: Literal["named_range"] = "named_range"
    name: str = Field(description="Named range identifier")


class UnparsableNode(AstNode):
    """Sentinel node when parsing fails or encounters unsupported constructs."""

    type: Literal["unparsable"] = "unparsable"
    raw: str = Field(description="Original formula text that could not be parsed")


# Union type for AST nodes used in annotations
AstNodeUnion = Annotated[
    FunctionNode | OperandNode | OperatorNode | RangeNode | NamedRangeNode | UnparsableNode,
    Field(discriminator="type"),
]


class CellRef(BaseModel):
    """A single cell or range reference extracted from a formula."""

    model_config = ConfigDict(frozen=True)

    sheet: str | None = Field(default=None, description="Sheet name; None = same sheet")
    address: str = Field(description="Cell or range address, e.g. 'A1' or 'A1:B10'")


class FormulaRecord(BaseModel):
    """Full analysis record for one formula cell."""

    model_config = ConfigDict(frozen=True)

    formula_id: int = Field(description="1-based unique formula identifier")
    qualified_address: str = Field(description="'SheetName!A1' cell address")
    formula_text: str = Field(description="Raw formula string including leading '='")
    formula_category: FormulaCategory = Field(description="Classifier output category")
    function_list: list[str] = Field(
        description="Sorted, deduplicated uppercase function names used"
    )
    referenced_cells: list[CellRef] = Field(
        description="Sorted list of cell/range references in the formula"
    )
    referenced_named_ranges: list[str] = Field(
        description="Sorted list of named range identifiers used"
    )
    has_external_reference: bool = Field(
        description="True if any range reference contains '[...]' external workbook syntax"
    )
    is_volatile: bool = Field(
        description="True if formula contains a volatile function (NOW, RAND, OFFSET, etc.)"
    )
    is_array_formula: bool = Field(
        description="True if cell uses array formula syntax (entered with Ctrl+Shift+Enter)"
    )
    nesting_depth: int = Field(description="Maximum function call nesting depth")
    function_count: int = Field(description="Total count of function invocations (with dupes)")
    complexity_score: int = Field(
        description="nesting_depth * 2 + function_count + len(referenced_cells)"
    )
    ast: Any = Field(
        default=None,
        description="Parsed AST root node; None when is_parsable=False",
    )
    is_parsable: bool = Field(description="False when tokenizer/parser could not parse formula")
    parse_error: str | None = Field(
        default=None, description="Error message when is_parsable=False"
    )
