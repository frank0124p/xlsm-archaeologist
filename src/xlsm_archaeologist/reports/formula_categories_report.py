"""Aggregate formula counts and complexity by category."""

from __future__ import annotations

from collections import defaultdict
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from xlsm_archaeologist.models.formula import FormulaRecord


def build_categories_report(formulas: list[FormulaRecord]) -> list[dict[str, object]]:
    """Build per-category aggregation rows.

    Returns:
        List of row dicts with keys:
        category, formula_count, total_complexity, avg_complexity,
        max_complexity, pct_of_total.
    """
    total = len(formulas)
    buckets: dict[str, list[int]] = defaultdict(list)
    for f in formulas:
        buckets[f.formula_category].append(f.complexity_score)

    rows: list[dict[str, object]] = []
    for cat in sorted(buckets):
        scores = buckets[cat]
        count = len(scores)
        total_complexity = sum(scores)
        avg = round(total_complexity / count, 2) if count else 0.0
        max_c = max(scores) if scores else 0
        pct = round(count / total * 100, 2) if total else 0.0
        rows.append(
            {
                "category": cat,
                "formula_count": count,
                "total_complexity": total_complexity,
                "avg_complexity": avg,
                "max_complexity": max_c,
                "pct_of_total": pct,
            }
        )
    return rows
