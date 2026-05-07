"""Cross-sheet reference edges report."""

from __future__ import annotations

from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from xlsm_archaeologist.models.dependency import DependencyEdge


def build_cross_sheet_refs(edges: list[DependencyEdge]) -> list[dict[str, object]]:
    """Return rows for all cross-sheet dependency edges.

    Returns:
        List of row dicts matching cross_sheet_refs.csv schema.
    """
    rows: list[dict[str, object]] = []
    for edge in sorted(
        (e for e in edges if e.is_cross_sheet),
        key=lambda e: (e.source_qualified_address, e.target_qualified_address),
    ):
        src = edge.source_qualified_address
        tgt = edge.target_qualified_address
        src_sheet = src.split("!", 1)[0] if "!" in src else src
        tgt_sheet = tgt.split("!", 1)[0] if "!" in tgt else tgt
        rows.append(
            {
                "source_qualified_address": src,
                "source_sheet": src_sheet,
                "target_qualified_address": tgt,
                "target_sheet": tgt_sheet,
                "via": edge.via,
                "via_detail": edge.via_detail,
            }
        )
    return rows
