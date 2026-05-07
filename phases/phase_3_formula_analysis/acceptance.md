# Phase 3 — Acceptance Checklist

## 功能正確性

- [ ] 6 種分類各一條公式測試,分類結果正確
- [ ] mixed 公式 (含兩種以上類別) 分類為 `mixed`
- [ ] 純引用 `=Sheet2!A1` 分類為 `reference`
- [ ] 純算術 `=A1+B1*2` 分類為 `compute`
- [ ] 跨 sheet 引用的 sheet 名稱正確保留
- [ ] Named range 引用正確抽出 (在 referenced_named_ranges)
- [ ] LAMBDA/LET 公式標記為 `is_parsable=false` 並寫入 warnings
- [ ] OFFSET / INDIRECT 等 volatile 函式正確標記 `is_volatile=true`
- [ ] 外部活頁簿引用正確標記 `has_external_reference=true`

## 複雜度

- [ ] 簡單公式 `=A1+B1` complexity_score < 5
- [ ] 巢狀深度 ≥ 5 的公式 complexity_score ≥ 15
- [ ] `nesting_depth` 計算正確 (測試 case:`=IF(IF(IF(A1>0,1,0)>0,1,0)>0,1,0)` 應 = 3)

## Schema

- [ ] `05_formulas.json` 含 `schema_version: "1.0"`
- [ ] 每筆 FormulaRecord 含所有欄位 (照 `DATA_MODEL.md`)
- [ ] AST 序列化為 JSON,結構符合 README 定義

## Determinism

- [ ] 同一份 fixture 跑兩次,`05_formulas.json` 完全一致
- [ ] formula_id 從 1 連續遞增
- [ ] referenced_cells / referenced_named_ranges 排序穩定

## 程式品質

- [ ] `uv run ruff check .` 零錯誤
- [ ] `uv run mypy src` 零錯誤
- [ ] `uv run pytest tests/test_phase_3_formula.py` 全 pass
- [ ] formula 模組覆蓋率 ≥ 90%

## 不能出現

- [ ] **沒有** 試圖計算公式結果 (例如不准用 `formulas` 套件的 evaluator)
- [ ] **沒有** 把 unparsable 公式直接 raise 中斷整個 run (要降級為 record + warning)
- [ ] **沒有** 漏掉 cross-sheet 的 sheet 名稱
