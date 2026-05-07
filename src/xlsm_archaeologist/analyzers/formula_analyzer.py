"""Main formula analysis pipeline: tokenize → parse → classify → complexity + metadata."""

from __future__ import annotations

from collections.abc import Iterator

from xlsm_archaeologist.analyzers.formula_classifier import classify
from xlsm_archaeologist.analyzers.formula_complexity import compute_complexity
from xlsm_archaeologist.analyzers.formula_parser import parse
from xlsm_archaeologist.analyzers.formula_tokenizer import tokenize
from xlsm_archaeologist.models.cell import CellRecord
from xlsm_archaeologist.models.formula import (
    AstNode,
    CellRef,
    FormulaRecord,
    FunctionNode,
    NamedRangeNode,
    OperatorNode,
    RangeNode,
    UnparsableNode,
)
from xlsm_archaeologist.utils.logging import get_logger

logger = get_logger(__name__)

VOLATILE_FUNCS: frozenset[str] = frozenset(
    {
        "NOW", "TODAY",
        "RAND", "RANDBETWEEN", "RANDARRAY",
        "OFFSET", "INDIRECT",
        "INFO", "CELL",
    }
)

UNPARSABLE_FUNCS: frozenset[str] = frozenset(
    {
        "LAMBDA", "LET",
        "BYROW", "BYCOL",
        "REDUCE", "MAP", "SCAN",
        "MAKEARRAY",
    }
)


def _collect_functions_set(node: AstNode) -> set[str]:
    """Recursively collect unique uppercase function names."""
    names: set[str] = set()
    if isinstance(node, FunctionNode):
        names.add(node.name.upper())
        for arg in node.args:
            names |= _collect_functions_set(arg)
    elif isinstance(node, OperatorNode):
        if node.left is not None:
            names |= _collect_functions_set(node.left)
        names |= _collect_functions_set(node.right)
    return names


def _extract_references(node: AstNode) -> tuple[list[CellRef], list[str]]:
    """Recursively collect all RangeNode and NamedRangeNode references.

    Returns:
        Tuple of (sorted unique CellRef list, sorted unique named range name list).
    """
    cell_refs: list[CellRef] = []
    named_refs: list[str] = []

    def _walk(n: AstNode) -> None:
        if isinstance(n, RangeNode):
            cell_refs.append(CellRef(sheet=n.sheet, address=n.address))
        elif isinstance(n, NamedRangeNode):
            named_refs.append(n.name)
        elif isinstance(n, FunctionNode):
            for arg in n.args:
                _walk(arg)
        elif isinstance(n, OperatorNode):
            if n.left is not None:
                _walk(n.left)
            _walk(n.right)

    _walk(node)

    # deduplicate preserving sort order
    seen_cells: set[tuple[str | None, str]] = set()
    unique_cells: list[CellRef] = []
    for ref in sorted(cell_refs, key=lambda r: (r.sheet or "", r.address)):
        key = (ref.sheet, ref.address)
        if key not in seen_cells:
            seen_cells.add(key)
            unique_cells.append(ref)

    unique_named = sorted(set(named_refs))
    return unique_cells, unique_named


def _has_external_ref(cell_refs: list[CellRef]) -> bool:
    """Return True if any reference address contains '[...]' external workbook syntax."""
    return any("[" in (ref.address or "") or "[" in (ref.sheet or "") for ref in cell_refs)


def _is_volatile(funcs: set[str]) -> bool:
    return bool(funcs & VOLATILE_FUNCS)


def _has_unparsable_func(funcs: set[str]) -> bool:
    return bool(funcs & UNPARSABLE_FUNCS)


def analyze_formulas(
    cells: list[CellRecord],
    warnings: list[str],
) -> Iterator[FormulaRecord]:
    """Run the full formula analysis pipeline for all formula cells.

    Args:
        cells: All CellRecords from Phase 2 (only those with has_formula=True are processed).
        warnings: Mutable list to append parse-error warning strings into.

    Yields:
        FormulaRecord instances (unsorted; caller sorts if needed).
    """
    formula_id = 0
    for cell in cells:
        if not cell.has_formula:
            continue
        formula_id += 1
        formula_text = cell.raw_value

        # Ensure leading '='
        if not formula_text.startswith("="):
            formula_text = f"={formula_text}"

        tokens = tokenize(formula_text)
        if not tokens:
            warnings.append(
                f"{cell.qualified_address}: tokenization returned empty — formula: {formula_text!r}"
            )
            yield FormulaRecord(
                formula_id=formula_id,
                qualified_address=cell.qualified_address,
                formula_text=formula_text,
                formula_category="compute",
                function_list=[],
                referenced_cells=[],
                referenced_named_ranges=[],
                has_external_reference=False,
                is_volatile=False,
                is_array_formula=False,
                nesting_depth=0,
                function_count=0,
                complexity_score=0,
                ast=None,
                is_parsable=False,
                parse_error="Tokenization returned empty",
            )
            continue

        ast = parse(tokens)
        is_parsable = not isinstance(ast, UnparsableNode)

        if not is_parsable:
            err = f"Parse failed for {cell.qualified_address}: {formula_text!r}"
            warnings.append(err)
            yield FormulaRecord(
                formula_id=formula_id,
                qualified_address=cell.qualified_address,
                formula_text=formula_text,
                formula_category="compute",
                function_list=[],
                referenced_cells=[],
                referenced_named_ranges=[],
                has_external_reference=False,
                is_volatile=False,
                is_array_formula=False,
                nesting_depth=0,
                function_count=0,
                complexity_score=0,
                ast=None,
                is_parsable=False,
                parse_error=str(getattr(ast, "raw", "unknown")),
            )
            continue

        funcs = _collect_functions_set(ast)

        if _has_unparsable_func(funcs):
            warnings.append(
                f"{cell.qualified_address}: contains unsupported function(s) "
                f"{funcs & UNPARSABLE_FUNCS} — marked is_parsable=false"
            )
            yield FormulaRecord(
                formula_id=formula_id,
                qualified_address=cell.qualified_address,
                formula_text=formula_text,
                formula_category="compute",
                function_list=sorted(funcs),
                referenced_cells=[],
                referenced_named_ranges=[],
                has_external_reference=False,
                is_volatile=_is_volatile(funcs),
                is_array_formula=False,
                nesting_depth=0,
                function_count=len(funcs),
                complexity_score=0,
                ast=None,
                is_parsable=False,
                parse_error=f"Contains unsupported functions: {sorted(funcs & UNPARSABLE_FUNCS)}",
            )
            continue

        cell_refs, named_refs = _extract_references(ast)
        category = classify(ast)
        depth, func_count, score = compute_complexity(ast, cell_refs)

        yield FormulaRecord(
            formula_id=formula_id,
            qualified_address=cell.qualified_address,
            formula_text=formula_text,
            formula_category=category,
            function_list=sorted(funcs),
            referenced_cells=cell_refs,
            referenced_named_ranges=named_refs,
            has_external_reference=_has_external_ref(cell_refs),
            is_volatile=_is_volatile(funcs),
            is_array_formula=False,
            nesting_depth=depth,
            function_count=func_count,
            complexity_score=score,
            ast=ast.model_dump(),
            is_parsable=True,
            parse_error=None,
        )
