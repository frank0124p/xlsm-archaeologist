"""Detect orphan formula cells: formula_cell nodes with no incoming edges."""

from __future__ import annotations

import networkx as nx

from xlsm_archaeologist.utils.logging import get_logger

logger = get_logger(__name__)


def detect_orphans(graph: nx.DiGraph) -> list[str]:
    """Return sorted list of orphan formula cell node ids.

    An orphan is a node with node_type='formula_cell' and in_degree == 0,
    meaning nothing else depends on its output.

    Args:
        graph: Directed dependency graph.

    Returns:
        Sorted list of qualified address strings for orphan formulas.
    """
    orphans = [
        node
        for node, attrs in graph.nodes(data=True)
        if attrs.get("node_type") == "formula_cell" and graph.in_degree(node) == 0
    ]
    return sorted(orphans)
