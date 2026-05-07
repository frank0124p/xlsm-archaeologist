# Phase 6 — Tasks

## Models

- [ ] 建立 `models/summary.py`:
    - `Stats`(sheet_count, named_range_count, formula_count, ...)
    - `RiskIndicators`(circular_reference_count, external_reference_count,
                      volatile_function_count, dynamic_vba_range_count,
                      deeply_nested_formula_count, orphan_formula_count,
                      cross_sheet_dependency_count)
    - `Warning`(level, category, location, message)
    - `MigrationDifficulty` enum
    - `SummaryRecord`(schema_version, tool_version, analyzed_at, input_file,
                     stats, risk_indicators, complexity_score, migration_difficulty,
                     warnings)
- [ ] commit: `feat(models): add summary and warning models`

## Summary Builder

- [ ] `analyzers/summary_analyzer.py`:
    - `compute_stats(...) -> Stats`
    - `compute_risk_indicators(...) -> RiskIndicators`
    - `compute_complexity_score(risk_indicators, stats) -> int`
    - `derive_migration_difficulty(score) -> MigrationDifficulty`
- [ ] `reports/summary_builder.py`:
    - `build_summary(...) -> SummaryRecord`
    - 整合所有 phase 結果 + warnings
- [ ] commit: `feat(reports): add summary builder`

## 五張報告

- [ ] `reports/formula_categories_report.py`:
    - `build_categories_report(formulas) -> list[CategoryRow]`
    - 寫成 CSV
- [ ] `reports/top_complex_formulas_report.py`:
    - 取 top 50,輸出 CSV
- [ ] `reports/hotspot_cells_report.py`:
    - 從 graph 取 in_degree top 50
    - 拆解 in_edges 統計 formula vs vba 來源
- [ ] `reports/vba_behavior_report.py`:
    - 每個 procedure 一 row
    - 計算 cross_sheet_read/write
- [ ] `reports/cross_sheet_refs_report.py`:
    - 從 graph 過濾 is_cross_sheet=true 的邊
- [ ] commit (每個一個): `feat(reports): <report name>`

## Warning Aggregation

- [ ] `utils/warnings.py`:
    - `WarningCollector` 類別 — 全 run 共用
    - `add(level, category, location, message)`
    - `flush() -> list[Warning]` (排序)
- [ ] 各 phase 改用此 collector (回頭改 phase 2-5)
- [ ] commit: `refactor: centralize warning collection`

## CLI 串接

- [ ] 修改 `cli.py`:
    - Phase 6 在最後執行
    - 印出最終 summary 摘要 (complexity_score、difficulty、warning 數)
    - 處理 `--no-reports` flag
- [ ] commit: `feat(cli): wire reports phase and final summary print`

## Tests

- [ ] `tests/test_phase_6_reports.py`:
    - test_complexity_score_calculation
    - test_migration_difficulty_thresholds
    - test_formula_categories_aggregation
    - test_top_complex_top50_limit
    - test_hotspot_in_degree_correctness
    - test_vba_behavior_cross_sheet_count
    - test_warnings_sorted_by_level
    - test_summary_schema_version
- [ ] commit: `test(reports): comprehensive report tests`

## End-to-End

- [ ] 建立 `tests/test_e2e.py`:
    - 跑 quickstart 對 `formulas_complex.xlsm` 做完整 analyze
    - 驗證所有 11 個輸出檔案都存在
    - 驗證 `00_summary.json` 結構合法
    - 驗證跑兩次 diff 為空
- [ ] commit: `test(e2e): add end-to-end test`

## Quality

- [ ] `uv run pytest` 全 pass
- [ ] `uv run ruff check .` 零錯誤
- [ ] `uv run mypy src` 零錯誤
- [ ] 整體覆蓋率 ≥ 75%

## 真實檔案驗證

- [ ] 跟使用者要一份真實 .xlsm 跑過一次
- [ ] 沒有 fatal error
- [ ] 寫 `RUN_REPORT.md`,內容:
    - 該 .xlsm 跑出來的 stats
    - complexity_score 與 difficulty
    - warning 摘要 (top 10 categories)
    - 跑 wallclock 時間
- [ ] commit: `docs: add real-world run report`

## 收尾

- [ ] 寫 `phase_6_summary.md` + 更新 `README.md` 加 sample output 連結
- [ ] **任務完成,等待最終 review**
