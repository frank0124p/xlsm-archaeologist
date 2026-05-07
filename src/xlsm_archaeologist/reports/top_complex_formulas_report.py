"""Top N most complex formulas report."""

from __future__ import annotations

from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from xlsm_archaeologist.models.formula import FormulaRecord

_DEFAULT_TOP_N = 50


def build_top_complex_formulas(
    formulas: list[FormulaRecord],
    top_n: int = _DEFAULT_TOP_N,
) -> list[dict[str, object]]:
    """Return the top N formula rows sorted by complexity_score descending.

    Returns:
        List of row dicts with keys:
        rank, qualified_address, formula_text, formula_category,
        nesting_depth, function_count, referenced_cell_count, complexity_score.
    """
    sorted_formulas = sorted(formulas, key=lambda f: f.complexity_score, reverse=True)[:top_n]
    rows: list[dict[str, object]] = []
    for rank, formula in enumerate(sorted_formulas, start=1):
        rows.append(
            {
                "rank": rank,
                "qualified_address": formula.qualified_address,
                "formula_text": formula.formula_text,
                "formula_category": formula.formula_category,
                "nesting_depth": formula.nesting_depth,
                "function_count": formula.function_count,
                "referenced_cell_count": len(formula.referenced_cells),
                "complexity_score": formula.complexity_score,
            }
        )
    return rows
