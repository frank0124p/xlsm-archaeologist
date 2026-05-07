"""Build a NetworkX DiGraph representing cell/formula/VBA dependencies."""

from __future__ import annotations

from typing import TYPE_CHECKING

import networkx as nx

if TYPE_CHECKING:
    from xlsm_archaeologist.models.cell import CellRecord, ValidationRecord
    from xlsm_archaeologist.models.formula import FormulaRecord
    from xlsm_archaeologist.models.named_range import NamedRangeRecord
    from xlsm_archaeologist.models.vba import VbaProcedureRecord

from xlsm_archaeologist.utils.logging import get_logger

logger = get_logger(__name__)

_NAMED_PREFIX = "_named:"
_VBA_PREFIX = "_vba:"


def _node_type(node_id: str, formula_cells: set[str]) -> str:
    """Return the node_type string for a given node id."""
    if node_id.startswith(_NAMED_PREFIX):
        return "named_range"
    if node_id.startswith(_VBA_PREFIX):
        return "vba_procedure"
    if node_id in formula_cells:
        return "formula_cell"
    return "input_cell"


def _sheet_of(node_id: str) -> str | None:
    """Return sheet name from a qualified address, or None for special nodes."""
    if node_id.startswith((_NAMED_PREFIX, _VBA_PREFIX)):
        return None
    if "!" in node_id:
        return node_id.split("!", 1)[0]
    return None


def _is_cross_sheet(src: str, tgt: str) -> bool:
    s1 = _sheet_of(src)
    s2 = _sheet_of(tgt)
    return s1 is not None and s2 is not None and s1 != s2


def build_graph(
    formulas: list[FormulaRecord],
    vba_procedures: list[VbaProcedureRecord],
    named_ranges: list[NamedRangeRecord],
    cells: list[CellRecord],
    validations: list[ValidationRecord],
) -> nx.DiGraph:
    """Build a directed dependency graph from all extracted data.

    Nodes:
    - Each qualified cell address that appears as a cell, formula, or reference
    - _named:<name> for every named range
    - _vba:<module>.<proc> for every VBA procedure

    Edges (source → target means "target depends on source"):
    - Formula cell → referenced cells (via='formula')
    - Named range → formula cells that use it (via='named_range')
    - VBA procedure ↔ cells it reads/writes (via='vba_read_write')
    - Validation source cells → validated cells (via='validation')

    Args:
        formulas: All FormulaRecords from Phase 3.
        vba_procedures: All VbaProcedureRecords from Phase 4.
        named_ranges: All NamedRangeRecords from Phase 2.
        cells: All CellRecords from Phase 2.
        validations: All ValidationRecords from Phase 2.

    Returns:
        Directed NetworkX DiGraph with node/edge attributes.
    """
    graph: nx.DiGraph = nx.DiGraph()

    formula_cell_ids = {f.qualified_address for f in formulas}

    # --- Add cell nodes ---
    for cell in cells:
        ntype = "formula_cell" if cell.qualified_address in formula_cell_ids else "input_cell"
        graph.add_node(
            cell.qualified_address,
            node_type=ntype,
            value_type=cell.value_type,
        )

    # --- Add named range nodes ---
    for nr in named_ranges:
        node_id = f"{_NAMED_PREFIX}{nr.range_name}"
        graph.add_node(node_id, node_type="named_range", refers_to=nr.refers_to)

    # --- Add VBA procedure nodes ---
    for proc in vba_procedures:
        node_id = f"{_VBA_PREFIX}{proc.procedure_name}"
        graph.add_node(node_id, node_type="vba_procedure")

    # --- Formula → referenced-cell edges ---
    for formula in formulas:
        tgt = formula.qualified_address
        if not graph.has_node(tgt):
            graph.add_node(tgt, node_type="formula_cell", value_type="string")

        for ref in formula.referenced_cells:
            sheet = ref.sheet
            addr = ref.address
            if sheet:
                src = f"{sheet}!{addr}"
            else:
                # Same sheet as formula
                src_sheet = tgt.split("!", 1)[0] if "!" in tgt else ""
                src = f"{src_sheet}!{addr}" if src_sheet else addr
            if not graph.has_node(src):
                graph.add_node(src, node_type="input_cell", value_type="empty")
            graph.add_edge(
                src,
                tgt,
                via="formula",
                via_detail=str(formula.formula_id),
                is_cross_sheet=_is_cross_sheet(src, tgt),
            )

        # Named range → formula cell
        for nr_name in formula.referenced_named_ranges:
            nr_node = f"{_NAMED_PREFIX}{nr_name}"
            if not graph.has_node(nr_node):
                graph.add_node(nr_node, node_type="named_range")
            graph.add_edge(
                nr_node,
                tgt,
                via="named_range",
                via_detail=nr_name,
                is_cross_sheet=False,
            )

    # --- VBA range read/write edges ---
    for proc in vba_procedures:
        proc_node = f"{_VBA_PREFIX}{proc.procedure_name}"

        for ra in proc.reads:
            if ra.sheet and ra.range_ref:
                cell_node = f"{ra.sheet}!{ra.range_ref}"
            elif ra.range_ref and ra.range_ref != "(cells)":
                cell_node = ra.range_ref
            else:
                continue
            if not graph.has_node(cell_node):
                graph.add_node(cell_node, node_type="input_cell", value_type="empty")
            graph.add_edge(
                cell_node,
                proc_node,
                via="vba_read_write",
                via_detail=f"read:{proc.procedure_name}",
                is_cross_sheet=False,
            )

        for wa in proc.writes:
            if wa.sheet and wa.range_ref:
                cell_node = f"{wa.sheet}!{wa.range_ref}"
            elif wa.range_ref and wa.range_ref != "(cells)":
                cell_node = wa.range_ref
            else:
                continue
            if not graph.has_node(cell_node):
                graph.add_node(cell_node, node_type="input_cell", value_type="empty")
            graph.add_edge(
                proc_node,
                cell_node,
                via="vba_read_write",
                via_detail=f"write:{proc.procedure_name}",
                is_cross_sheet=False,
            )

    logger.debug(
        "Dependency graph: %d nodes, %d edges", graph.number_of_nodes(), graph.number_of_edges()
    )
    return graph
