"""Detect cycles in a dependency DiGraph."""

from __future__ import annotations

import networkx as nx

from xlsm_archaeologist.models.dependency import CycleRecord
from xlsm_archaeologist.utils.logging import get_logger

logger = get_logger(__name__)


def detect_cycles(graph: nx.DiGraph) -> list[CycleRecord]:
    """Find all simple cycles in the dependency graph.

    Self-loops (A → A) are excluded.

    Args:
        graph: Directed NetworkX graph.

    Returns:
        Sorted list of CycleRecord (by cycle length, then first node).
    """
    records: list[CycleRecord] = []
    cycle_id = 0

    try:
        raw_cycles = list(nx.simple_cycles(graph))
    except Exception:  # noqa: BLE001
        logger.warning("Cycle detection failed")
        return []

    for cycle_nodes in raw_cycles:
        if len(cycle_nodes) < 2:
            continue  # skip self-loops

        edges_via: list[str] = []
        for i, src in enumerate(cycle_nodes):
            tgt = cycle_nodes[(i + 1) % len(cycle_nodes)]
            edge_data = graph.get_edge_data(src, tgt) or {}
            edges_via.append(str(edge_data.get("via", "unknown")))

        cycle_id += 1
        records.append(
            CycleRecord(
                cycle_id=cycle_id,
                length=len(cycle_nodes),
                nodes=list(cycle_nodes),
                edges_via=edges_via,
            )
        )

    return sorted(records, key=lambda c: (c.length, c.nodes[0] if c.nodes else ""))
