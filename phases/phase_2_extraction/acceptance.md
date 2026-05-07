# Phase 2 — Acceptance Checklist

## 輸出正確性

- [ ] 跑 `analyze tests/fixtures/simple.xlsm`:
    - `01_workbook.json` schema_version, sha256, size 正確
    - `02_sheets.csv` 1 row, sheet_name=`Data`, formula_cell_count=0
    - `04_cells.csv` 0 row (無有意義 cell)
    - `06_validations.csv` 0 row
- [ ] 跑 `analyze tests/fixtures/with_validation.xlsm`:
    - `06_validations.csv` 至少 2 row
    - list literal 那條 `enum_values` 解析成功
    - list range 那條 `enum_values` 也解析成功
- [ ] 跑 `analyze tests/fixtures/with_named_range.xlsm`:
    - `03_named_ranges.csv` 至少 2 row
    - dynamic 那條 `has_dynamic_formula = true`
- [ ] 跑 `analyze tests/fixtures/hidden_sheets.xlsm`:
    - `02_sheets.csv` 中 hidden 與 very_hidden sheet 標記正確

## Schema 一致性

- [ ] 所有 JSON 含 `"schema_version": "1.0"`
- [ ] 所有 CSV header 字串照 `DATA_MODEL.md` 一字不差
- [ ] 所有布林欄位用 `is_*` / `has_*` (CSV 中為 `true` / `false`,JSON 為 boolean)

## Determinism

- [ ] 同一份 fixture 跑兩次 — `diff` 全部輸出檔案,無差異
- [ ] CSV row 順序穩定 (按 primary key 排序)
- [ ] JSON object key 順序穩定

## 程式品質

- [ ] `uv run ruff check .` 零錯誤
- [ ] `uv run mypy src` 零錯誤
- [ ] `uv run pytest tests/test_phase_2_extraction.py` 全 pass
- [ ] 覆蓋率 ≥ 80%

## 不能出現

- [ ] **沒有** 公式 AST 解析 (留 Phase 3)
- [ ] **沒有** VBA 邏輯 (留 Phase 4)
- [ ] **沒有** 用 `read_only=True` 開檔 (拿不到 named range)
- [ ] **沒有** 把 `is_referenced` 直接從 cell 引用算出來 (要等 Phase 5)
