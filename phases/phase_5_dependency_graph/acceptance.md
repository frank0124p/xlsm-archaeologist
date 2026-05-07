# Phase 5 — Acceptance Checklist

## 圖正確性

- [ ] `circular.xlsm` 偵測到至少 1 個 cycle,length=2
- [ ] `orphan_formula.xlsm` orphans.csv 至少 1 row
- [ ] `cross_sheet_chain.xlsm`:
    - `nx.descendants(G, "Sheet1!A1")` 包含 `Sheet2!B1` 與 `Sheet3!C1`
    - `nx.ancestors(G, "Sheet3!C1")` 包含 `Sheet1!A1` 與 `Sheet2!B1`

## VBA 整合

- [ ] `vba_basic.xlsm` 跑完後:
    - graph 含 `_vba:Module1.<proc>` 節點
    - 該節點的 in_edges 含 `Sheet!A1 → _vba:Module1.<proc>` (read)
    - 該節點的 out_edges 含 `_vba:Module1.<proc> → Sheet!B1` (write)

## is_referenced 回填

- [ ] 跑完 Phase 5 後 `04_cells.csv` 中:
    - 至少有些 cell 的 `is_referenced` 從 false 改成 true
    - 純 input 但沒被引用的 cell `is_referenced=false`

## Schema

- [ ] `09_dependencies.csv` header 完全照 DATA_MODEL.md
- [ ] `10_dependency_graph.json` 含 schema_version + 統計欄位
- [ ] `reports/cycles.json` 與 `reports/orphans.csv` 存在 (即使 0 cycles)

## Determinism

- [ ] 同一份 fixture 跑兩次,所有 graph 輸出檔案一致
- [ ] 邊清單按 (source, target, via) 排序穩定

## 程式品質

- [ ] `uv run pytest tests/test_phase_5_graph.py` 全 pass
- [ ] `uv run ruff check .` 零錯誤
- [ ] `uv run mypy src` 零錯誤

## 不能出現

- [ ] **沒有** 把 range (如 A1:A10) 展開成 10 個 cell-level edge
- [ ] **沒有** 漏掉 named range 的中介節點
- [ ] **沒有** 漏掉 VBA 動態 range 標記為粗粒度邊
- [ ] **沒有** 把 self-loop (A1 → A1 — 例如 `=A1+1` 寫到 A1 時) 算成 cycle
