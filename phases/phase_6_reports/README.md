# Phase 6: Reports & Scoring

## 目標

整合前面所有 phase 的結果,產出:
1. `00_summary.json` — 總覽 + complexity_score + warnings
2. `reports/formula_categories.csv`
3. `reports/top_complex_formulas.csv` (Top 50)
4. `reports/hotspot_cells.csv` (Top 50)
5. `reports/vba_behavior.csv`
6. `reports/cross_sheet_refs.csv`
7. (cycles.json 與 orphans.csv 已在 Phase 5 寫,這裡確保檔案在)

## 模組

```
src/xlsm_archaeologist/
├── reports/
│   ├── summary_builder.py            # 00_summary.json
│   ├── formula_categories_report.py
│   ├── top_complex_formulas_report.py
│   ├── hotspot_cells_report.py
│   ├── vba_behavior_report.py
│   └── cross_sheet_refs_report.py
└── analyzers/
    └── summary_analyzer.py           # complexity_score 與 risk_indicators 計算
```

## Complexity Score 公式

```python
complexity_score = (
    formula_count * 1
  + deeply_nested_formula_count * 5     # nesting_depth >= 5
  + dynamic_vba_range_count * 10
  + circular_reference_count * 20
  + cross_sheet_dependency_count * 0.5
  + orphan_formula_count * 0.3
)

migration_difficulty = (
    "low"       if score < 200
    else "medium"     if score < 500
    else "high"       if score < 1000
    else "very_high"
)
```

## Reports 細節

### formula_categories.csv

| Column | Source |
|---|---|
| `category` | enum 字串 |
| `formula_count` | 該類公式數 |
| `total_complexity` | 該類複雜度加總 |
| `avg_complexity` | 該類平均複雜度 |
| `max_complexity` | 該類最大複雜度 |
| `pct_of_total` | 占總公式數百分比 |

排序:`formula_count` 由高到低。

### top_complex_formulas.csv

| Column | Source |
|---|---|
| `rank` | 1-50 |
| `qualified_address` | |
| `formula_text` | 截斷至 200 字 |
| `formula_category` | |
| `nesting_depth` | |
| `function_count` | |
| `referenced_cell_count` | |
| `complexity_score` | |

排序:`complexity_score` 由高到低,並列時依 qualified_address 字典序。

### hotspot_cells.csv

| Column | Source |
|---|---|
| `rank` | 1-50 |
| `qualified_address` | |
| `node_type` | |
| `in_degree` | 多少 cell 引用它 |
| `referenced_by_formula_count` | |
| `referenced_by_vba_count` | |
| `value_type` | |
| `raw_value` | 截斷 100 字 |

排序:`in_degree` 由高到低。

### vba_behavior.csv

每個 procedure 一 row:

| Column | Source |
|---|---|
| `module_name` | |
| `procedure_name` | |
| `procedure_type` | |
| `line_count` | |
| `read_count` | |
| `write_count` | |
| `cross_sheet_read_count` | reads 中 sheet ≠ procedure 所在 sheet |
| `cross_sheet_write_count` | |
| `has_dynamic_range` | |
| `has_event_trigger` | |
| `call_count` | calls 長度 |
| `complexity_score` | |

排序:`complexity_score` 由高到低。

### cross_sheet_refs.csv

從 Phase 5 graph 過濾:

| Column |
|---|
| `source_qualified_address` |
| `source_sheet` |
| `target_qualified_address` |
| `target_sheet` |
| `via` |
| `via_detail` |

排序:source 字典序。

## Warnings 收集

`00_summary.json#warnings` 收集前 5 個 phase 累計的所有 warnings:

```json
{
  "level": "warning" | "error" | "info",
  "category": "extraction" | "formula" | "vba" | "graph",
  "location": "Module1.UpdateRows#L23",
  "message": "..."
}
```

排序:level (error > warning > info),然後 category,然後 location。

## 驗收

見 `acceptance.md`。
