# Phase 6 — Acceptance Checklist

## 輸出完整性

- [ ] `00_summary.json` 存在且 valid pydantic
- [ ] `reports/formula_categories.csv` 存在
- [ ] `reports/top_complex_formulas.csv` 存在,row 數 ≤ 50
- [ ] `reports/hotspot_cells.csv` 存在,row 數 ≤ 50
- [ ] `reports/vba_behavior.csv` 存在
- [ ] `reports/cross_sheet_refs.csv` 存在
- [ ] `reports/cycles.json` 存在
- [ ] `reports/orphans.csv` 存在

## Complexity Score

- [ ] 公式照 README 文件
- [ ] 跑 `simple.xlsm` (無公式無 VBA) → score 接近 0,difficulty=`low`
- [ ] 跑 `formulas_complex.xlsm` → score > 0,difficulty 至少 `medium`

## Reports 正確性

- [ ] formula_categories.csv:`pct_of_total` 加總 = 100% (容許四捨五入 ±0.1%)
- [ ] top_complex_formulas.csv:第 1 row 的 complexity_score ≥ 第 2 row
- [ ] hotspot_cells.csv:第 1 row 的 in_degree ≥ 第 2 row
- [ ] vba_behavior.csv:每個 procedure 都有對應 row

## Warnings

- [ ] warnings 排序:level (error→warning→info) 然後 category 然後 location
- [ ] 所有前面 phase 的 warnings 都被收集到

## Schema

- [ ] 所有 JSON 含 `schema_version: "1.0"`
- [ ] 所有 CSV header 照 DATA_MODEL.md

## End-to-End

- [ ] 整套 pipeline 對 `formulas_complex.xlsm` 跑通,無 fatal error
- [ ] 跑兩次 diff 為空 (deterministic)
- [ ] CLI 最終訊息含 complexity_score 與 difficulty

## 真實檔案

- [ ] 在使用者提供的真實 .xlsm 上跑過一次
- [ ] 跑出來的 stats 經使用者目視 review,認為合理
- [ ] `RUN_REPORT.md` 完成

## 程式品質

- [ ] `uv run pytest` 全 pass
- [ ] `uv run ruff check .` 零錯誤
- [ ] `uv run mypy src` 零錯誤
- [ ] 整體覆蓋率 ≥ 75%
- [ ] 公式分類器、VBA 讀寫識別 ≥ 90%

## 不能出現

- [ ] **沒有** 假資料 (所有數字都從 phase 1-5 計算來)
- [ ] **沒有** hard-coded 路徑
- [ ] **沒有** 漏掉任何一張報告
