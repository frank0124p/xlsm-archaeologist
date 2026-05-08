"""Data flow report — explains how the workbook operates for developers."""

from __future__ import annotations

from collections import defaultdict
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from xlsm_archaeologist.models.cell import CellRecord, ValidationRecord
    from xlsm_archaeologist.models.dependency import DependencyEdge
    from xlsm_archaeologist.models.formula import FormulaRecord
    from xlsm_archaeologist.models.vba import VBAProcedureRecord
    from xlsm_archaeologist.models.workbook import SheetRecord


def _sheet_from(addr: str) -> str:
    return addr.split("!", 1)[0] if "!" in addr else addr


def build_data_flow_md(
    sheets: list[SheetRecord],
    cells: list[CellRecord],
    formulas: list[FormulaRecord],
    validations: list[ValidationRecord],
    dep_edges: list[DependencyEdge],
    vba_procedures: list[VBAProcedureRecord],
    source_file: str,
) -> str:
    """Generate a developer-facing data flow document.

    Explains each sheet's role, key formulas, validations, and VBA interactions.
    Returns Markdown string for reports/data_flow.md.
    """
    # Group by sheet
    formulas_by_sheet: dict[str, list[FormulaRecord]] = defaultdict(list)
    for f in formulas:
        formulas_by_sheet[_sheet_from(f.qualified_address)].append(f)

    validations_by_sheet: dict[str, list[ValidationRecord]] = defaultdict(list)
    for v in validations:
        validations_by_sheet[_sheet_from(v.qualified_address)].append(v)

    # Cross-sheet dependency counts per sheet
    outgoing: dict[str, int] = defaultdict(int)
    incoming: dict[str, int] = defaultdict(int)
    for edge in dep_edges:
        if not edge.is_cross_sheet:
            continue
        outgoing[_sheet_from(edge.source_qualified_address)] += 1
        incoming[_sheet_from(edge.target_qualified_address)] += 1

    # VBA procedures that touch each sheet
    vba_by_sheet: dict[str, list[str]] = defaultdict(list)
    for proc in vba_procedures:
        touched: set[str] = set()
        for ref in proc.read_cells + proc.write_cells:
            s = _sheet_from(ref)
            if s:
                touched.add(s)
        for s in touched:
            vba_by_sheet[s].append(proc.procedure_name)

    # Formula category summary per sheet
    def cat_summary(flist: list[FormulaRecord]) -> str:
        counts: dict[str, int] = defaultdict(int)
        for f in flist:
            counts[f.formula_category] += 1
        return "、".join(f"{cat}×{n}" for cat, n in sorted(counts.items(), key=lambda x: -x[1]))

    # Build sheet sections
    sections: list[str] = []
    for s in sheets:
        name = s.sheet_name
        hidden_note = "（隱藏）" if s.is_hidden == "true" else ""
        flist = formulas_by_sheet.get(name, [])
        vlist = validations_by_sheet.get(name, [])
        vba = vba_by_sheet.get(name, [])
        out = outgoing.get(name, 0)
        inc = incoming.get(name, 0)

        # Sheet role
        if s.is_hidden == "true":
            role = "🔒 隱藏工作表（輔助資料或設定）"
        elif inc > 0 and out > 0:
            role = "⚙️ 計算層（接收並輸出跨工作表資料）"
        elif inc > 0:
            role = "📤 輸出層（彙整其他工作表的計算結果）"
        elif len(flist) == 0 and len(vlist) > 0:
            role = "📥 輸入層（使用者填寫資料）"
        elif len(flist) == 0:
            role = "📋 資料層（靜態參考資料）"
        else:
            role = "📥 輸入層（含部分計算）"

        lines = [f"### {name} {hidden_note}", "", f"**角色：** {role}", ""]
        lines.append(f"- 範圍：{s.row_count} 列 × {s.col_count} 欄，共 {s.cell_count} 個有效儲存格")
        lines.append(f"- 公式數：{len(flist)}，資料驗證數：{len(vlist)}")
        if out or inc:
            lines.append(f"- 跨工作表：輸出 {out} 個引用至他表，接收來自他表 {inc} 個引用")

        if flist:
            cat = cat_summary(flist)
            lines.append(f"- 公式分類：{cat}")
            top = sorted(flist, key=lambda f: f.complexity_score, reverse=True)[:3]
            lines.append("")
            lines.append("**最複雜的公式（前 3）：**")
            lines.append("")
            lines.append("| 位址 | 分類 | 複雜度 | 公式 |")
            lines.append("|---|---|---|---|")
            for f in top:
                formula_short = f.formula_text[:80] + "…" if len(f.formula_text) > 80 else f.formula_text  # noqa: E501
                lines.append(f"| `{f.qualified_address}` | {f.formula_category} | {f.complexity_score} | `{formula_short}` |")  # noqa: E501

        if vlist:
            lines.append("")
            lines.append("**資料驗證規則：**")
            lines.append("")
            lines.append("| 位址 | 類型 | 允許值 |")
            lines.append("|---|---|---|")
            for v in vlist[:10]:
                allowed = v.enum_values[:60] if v.enum_values else v.formula1 or "—"
                lines.append(f"| `{v.qualified_address}` | {v.validation_type} | {allowed} |")
            if len(vlist) > 10:
                lines.append(f"| … | 共 {len(vlist)} 條驗證規則 | |")

        if vba:
            lines.append("")
            lines.append(f"**相關 VBA 程序：** {', '.join(f'`{p}`' for p in vba[:5])}")

        sections.append("\n".join(lines))

    # Overall VBA summary
    vba_section = ""
    if vba_procedures:
        vba_lines = [
            "## VBA 程序操作摘要",
            "",
            "| 模組 | 程序 | 類型 | 讀取 cell 數 | 寫入 cell 數 | 動態 range |",
            "|---|---|---|---|---|---|",
        ]
        for proc in sorted(vba_procedures, key=lambda p: (p.module_name, p.procedure_name)):
            dynamic = "⚠️ 是" if proc.has_dynamic_range else "否"
            vba_lines.append(
                f"| {proc.module_name} | {proc.procedure_name} | {proc.procedure_type} "
                f"| {len(proc.read_cells)} | {len(proc.write_cells)} | {dynamic} |"
            )
        vba_section = "\n".join(vba_lines)

    sheets_md = "\n\n---\n\n".join(sections)

    return f"""# 操作說明 — {source_file}

> 由 xlsm-archaeologist 自動生成。描述每個工作表的角色、公式邏輯與 VBA 行為，供開發者理解此 Excel 的運作方式。

---

## 工作表角色說明

{sheets_md}

---

{vba_section}

---

*此文件由 `xlsm-archaeologist analyze` 自動產生，詳細原始資料見 `05_formulas.json`、`06_validations.csv`、`08_vba_procedures.json`。*
"""
