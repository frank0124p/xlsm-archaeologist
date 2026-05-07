"""VBA behavior summary: one row per procedure."""

from __future__ import annotations

from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from xlsm_archaeologist.models.vba import VbaModuleRecord, VbaProcedureRecord


def build_vba_behavior(
    procedures: list[VbaProcedureRecord],
    modules: list[VbaModuleRecord],
) -> list[dict[str, object]]:
    """Build one row per VBA procedure.

    Returns:
        List of row dicts matching vba_behavior.csv schema.
    """
    module_map = {m.vba_module_id: m for m in modules}

    rows: list[dict[str, object]] = []
    for proc in sorted(procedures, key=lambda p: (p.vba_module_id, p.procedure_name)):
        module = module_map.get(proc.vba_module_id)
        module_name = module.module_name if module else "unknown"

        cross_sheet_reads = sum(1 for r in proc.reads if r.sheet is not None)
        cross_sheet_writes = sum(1 for w in proc.writes if w.sheet is not None)

        rows.append(
            {
                "module_name": module_name,
                "procedure_name": proc.procedure_name,
                "procedure_type": proc.procedure_type,
                "line_count": proc.line_count,
                "read_count": len(proc.reads),
                "write_count": len(proc.writes),
                "cross_sheet_read_count": cross_sheet_reads,
                "cross_sheet_write_count": cross_sheet_writes,
                "has_dynamic_range": proc.has_dynamic_range,
                "has_event_trigger": len(proc.triggers) > 0,
                "call_count": len(proc.calls),
                "complexity_score": proc.complexity_score,
            }
        )
    return rows
