"""Orchestrate Phase 5: build graph, detect cycles/orphans, backfill is_referenced."""

from __future__ import annotations

from typing import TYPE_CHECKING

import networkx as nx

from xlsm_archaeologist.analyzers.cycle_detector import detect_cycles
from xlsm_archaeologist.analyzers.dependency_graph_builder import build_graph
from xlsm_archaeologist.analyzers.orphan_detector import detect_orphans
from xlsm_archaeologist.models.dependency import CycleRecord, DependencyEdge
from xlsm_archaeologist.utils.logging import get_logger

if TYPE_CHECKING:
    from xlsm_archaeologist.models.cell import CellRecord, ValidationRecord
    from xlsm_archaeologist.models.formula import FormulaRecord
    from xlsm_archaeologist.models.named_range import NamedRangeRecord
    from xlsm_archaeologist.models.vba import VbaProcedureRecord

logger = get_logger(__name__)


def run_dependency_analysis(
    formulas: list[FormulaRecord],
    vba_procedures: list[VbaProcedureRecord],
    named_ranges: list[NamedRangeRecord],
    cells: list[CellRecord],
    validations: list[ValidationRecord],
    warnings: list[str],
) -> tuple[
    nx.DiGraph,
    list[DependencyEdge],
    list[CycleRecord],
    list[str],
    list[CellRecord],
]:
    """Run the full Phase 5 dependency analysis.

    Returns:
        Tuple of (graph, edges, cycles, orphan_ids, updated_cells).
    """
    graph = build_graph(formulas, vba_procedures, named_ranges, cells, validations)

    # Extract sorted edge list
    edges: list[DependencyEdge] = []
    sorted_edges = sorted(graph.edges(data=True), key=lambda e: (e[0], e[1]))
    for edge_id, (src, tgt, attrs) in enumerate(sorted_edges, start=1):
        edges.append(
            DependencyEdge(
                dependency_id=edge_id,
                source_qualified_address=src,
                target_qualified_address=tgt,
                via=attrs.get("via", "formula"),
                via_detail=str(attrs.get("via_detail", "")),
                is_cross_sheet=bool(attrs.get("is_cross_sheet", False)),
            )
        )

    cycles = detect_cycles(graph)
    if cycles:
        warnings.append(f"Detected {len(cycles)} circular reference cycle(s)")

    orphan_ids = detect_orphans(graph)
    if orphan_ids:
        warnings.append(f"Detected {len(orphan_ids)} orphan formula(s)")

    # Backfill is_referenced: True when a cell has outgoing edges
    # (i.e., some formula/VBA uses it as input — it appears in the graph as a source)
    updated_cells = []
    for cell in cells:
        out_deg = (
            graph.out_degree(cell.qualified_address)
            if graph.has_node(cell.qualified_address)
            else 0
        )
        is_ref = out_deg > 0
        if is_ref != cell.is_referenced:
            updated_cells.append(cell.model_copy(update={"is_referenced": is_ref}))
        else:
            updated_cells.append(cell)

    return graph, edges, cycles, orphan_ids, updated_cells


def graph_to_json(graph: nx.DiGraph, cycles: list[CycleRecord]) -> dict[str, object]:
    """Serialize the graph to a JSON-friendly dict (node-link format)."""
    nodes = []
    for node, attrs in sorted(graph.nodes(data=True)):
        nodes.append(
            {
                "id": node,
                "node_type": attrs.get("node_type", "input_cell"),
                "value_type": attrs.get("value_type"),
                "in_degree": graph.in_degree(node),
                "out_degree": graph.out_degree(node),
            }
        )

    edges_list = []
    for src, tgt, attrs in sorted(graph.edges(data=True), key=lambda e: (e[0], e[1])):
        edges_list.append(
            {
                "source": src,
                "target": tgt,
                "via": attrs.get("via"),
                "via_detail": attrs.get("via_detail", ""),
            }
        )

    has_cycles = len(cycles) > 0
    wcc = nx.number_weakly_connected_components(graph)

    return {
        "directed": True,
        "multigraph": False,
        "graph": {
            "node_count": graph.number_of_nodes(),
            "edge_count": graph.number_of_edges(),
            "has_cycles": has_cycles,
            "cycle_count": len(cycles),
            "weakly_connected_component_count": wcc,
        },
        "nodes": nodes,
        "edges": edges_list,
    }
