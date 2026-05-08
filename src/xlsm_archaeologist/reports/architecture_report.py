"""Architecture report — Mermaid flowchart of sheet-level dependencies."""

from __future__ import annotations

from collections import defaultdict
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from xlsm_archaeologist.models.dependency import DependencyEdge
    from xlsm_archaeologist.models.formula import FormulaRecord
    from xlsm_archaeologist.models.vba import VBAModuleRecord
    from xlsm_archaeologist.models.workbook import SheetRecord


def _sheet_from(addr: str) -> str:
    return addr.split("!", 1)[0] if "!" in addr else addr


def _safe_id(name: str) -> str:
    """Convert sheet name to a Mermaid-safe node ID."""
    return name.replace(" ", "_").replace("-", "_").replace("'", "")


def build_architecture_md(
    sheets: list[SheetRecord],
    dep_edges: list[DependencyEdge],
    formulas: list[FormulaRecord],
    vba_modules: list[VBAModuleRecord],
    source_file: str,
) -> str:
    """Generate a Markdown document with a Mermaid sheet-level architecture diagram.

    Returns the full Markdown string to be written as reports/architecture.md.
    """
    # Count formulas per sheet
    formula_count: dict[str, int] = defaultdict(int)
    for f in formulas:
        sheet = _sheet_from(f.qualified_address)
        formula_count[sheet] += 1

    # Build cross-sheet dependency pairs (source_sheet → target_sheet)
    cross: dict[tuple[str, str], int] = defaultdict(int)
    for edge in dep_edges:
        if not edge.is_cross_sheet:
            continue
        src = _sheet_from(edge.source_qualified_address)
        tgt = _sheet_from(edge.target_qualified_address)
        if src and tgt and src != tgt:
            cross[(src, tgt)] += 1

    # Classify sheets: input (no incoming cross-sheet), output (no outgoing cross-sheet), compute
    sheets_with_outgoing = {src for src, _ in cross}
    sheets_with_incoming = {tgt for _, tgt in cross}

    def classify(name: str) -> str:
        has_out = name in sheets_with_outgoing
        has_in = name in sheets_with_incoming
        if has_in and has_out:
            return "compute"
        if has_in and not has_out:
            return "output"
        return "input"

    # Build Mermaid diagram
    lines: list[str] = ["flowchart LR"]

    # Subgraph per sheet type
    type_groups: dict[str, list[str]] = defaultdict(list)
    for s in sheets:
        if s.is_hidden == "true":
            continue
        c = classify(s.sheet_name)
        type_groups[c].append(s.sheet_name)

    type_labels = {"input": "輸入 Input", "compute": "計算 Compute", "output": "輸出 Output"}
    for group in ["input", "compute", "output"]:
        members = type_groups.get(group, [])
        if not members:
            continue
        lines.append(f'    subgraph {group}["{type_labels[group]}"]')
        for name in members:
            sid = _safe_id(name)
            fc = formula_count.get(name, 0)
            label = f"{name}\\n({fc} formulas)" if fc else name
            shape = f'["{label}"]'
            lines.append(f"        {sid}{shape}")
        lines.append("    end")

    # Hidden sheets node (collapsed)
    hidden = [s.sheet_name for s in sheets if s.is_hidden == "true"]
    if hidden:
        lines.append('    subgraph hidden["隱藏工作表 Hidden"]')
        for name in hidden:
            lines.append(f'        {_safe_id(name)}["{name}"]')
        lines.append("    end")

    # VBA modules as separate nodes
    if vba_modules:
        lines.append('    subgraph vba["VBA 模組"]')
        for mod in vba_modules:
            mid = _safe_id(f"vba_{mod.module_name}")
            lines.append(f'        {mid}(["{mod.module_name}\\n{mod.module_type}"])')
        lines.append("    end")

    # Cross-sheet edges (show edge weight)
    for (src, tgt), count in sorted(cross.items(), key=lambda x: -x[1]):
        label = f"{count} refs"
        lines.append(f'    {_safe_id(src)} -->|"{label}"| {_safe_id(tgt)}')

    mermaid_block = "\n".join(lines)

    # Build sheet detail table
    sheet_rows: list[str] = []
    for s in sheets:
        c = "hidden" if s.is_hidden == "true" else classify(s.sheet_name)
        fc = formula_count.get(s.sheet_name, 0)
        out_count = sum(v for (src, _), v in cross.items() if src == s.sheet_name)
        in_count = sum(v for (_, tgt), v in cross.items() if tgt == s.sheet_name)
        sheet_rows.append(
            f"| {s.sheet_name} | {c} | {s.row_count} | {s.col_count} "
            f"| {fc} | {in_count} | {out_count} |"
        )

    sheet_table = "\n".join(sheet_rows)

    return f"""# 架構圖 — {source_file}

> 由 xlsm-archaeologist 自動生成。顯示工作表層級的資料流向與分類。

## 工作表資料流程圖

```mermaid
{mermaid_block}
```

**節點分類說明：**
- **Input（輸入）**：無其他工作表資料流入，為原始資料來源
- **Compute（計算）**：同時接收並輸出跨工作表資料，為中間計算層
- **Output（輸出）**：只接收資料、不向外輸出，為最終結果呈現

---

## 工作表統計

| 工作表 | 分類 | 列數 | 欄數 | 公式數 | 被引用次數 | 引用他人次數 |
|---|---|---|---|---|---|---|
{sheet_table}

---

## 跨工作表依賴清單

| 來源工作表 | 目標工作表 | 引用次數 |
|---|---|---|
{"".join(f"| {src} | {tgt} | {cnt} |\\n" for (src, tgt), cnt in sorted(cross.items(), key=lambda x: -x[1]))}
---

*此文件由 `xlsm-archaeologist analyze` 自動產生，詳細資料見 `09_dependencies.csv` 與 `10_dependency_graph.json`。*
"""
