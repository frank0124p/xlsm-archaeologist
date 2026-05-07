"""Hotspot cells: top N by in-degree (most depended-upon cells)."""

from __future__ import annotations

from typing import TYPE_CHECKING

import networkx as nx

if TYPE_CHECKING:
    from xlsm_archaeologist.models.cell import CellRecord

_DEFAULT_TOP_N = 50


def build_hotspot_cells(
    graph: nx.DiGraph,
    cells: list[CellRecord],
    top_n: int = _DEFAULT_TOP_N,
) -> list[dict[str, object]]:
    """Return top N cells by in-degree (how many things depend on them).

    Returns:
        List of row dicts matching hotspot_cells.csv schema.
    """
    cell_map = {c.qualified_address: c for c in cells}

    # "Hotspot" = a cell/node that many other cells depend on = high out_degree
    # (edge src→tgt means tgt depends on src; so out_degree = how many things reference src)
    candidates = [
        (node, graph.out_degree(node))
        for node in graph.nodes()
        if graph.out_degree(node) > 0
    ]
    candidates.sort(key=lambda x: x[1], reverse=True)

    rows: list[dict[str, object]] = []
    for rank, (node_id, out_deg) in enumerate(candidates[:top_n], start=1):
        cell = cell_map.get(node_id)
        attrs = graph.nodes[node_id]
        node_type = attrs.get("node_type", "input_cell")

        # Count formula vs vba out-edges (things that reference this node)
        formula_refs = sum(
            1 for _, _, d in graph.out_edges(node_id, data=True) if d.get("via") == "formula"
        )
        vba_refs = sum(
            1
            for _, _, d in graph.out_edges(node_id, data=True)
            if d.get("via") == "vba_read_write"
        )

        rows.append(
            {
                "rank": rank,
                "qualified_address": node_id,
                "node_type": node_type,
                "in_degree": out_deg,
                "referenced_by_formula_count": formula_refs,
                "referenced_by_vba_count": vba_refs,
                "value_type": cell.value_type if cell else "",
                "raw_value": cell.raw_value if cell else "",
            }
        )
    return rows
