# Data Model

> 所有輸出檔案的 schema 定義。這份是契約,下游工具會依賴這個格式。

## 輸出資料夾結構

```
archaeology_output/
├── 00_summary.json
├── 01_workbook.json
├── 02_sheets.csv
├── 03_named_ranges.csv
├── 04_cells.csv
├── 05_formulas.json
├── 06_validations.csv
├── 07_vba_modules.json
├── 08_vba_procedures.json
├── 09_dependencies.csv
├── 10_dependency_graph.json
└── reports/
    ├── formula_categories.csv
    ├── top_complex_formulas.csv
    ├── hotspot_cells.csv
    ├── vba_behavior.csv
    ├── cycles.json
    ├── orphans.csv
    └── cross_sheet_refs.csv
```

## Schema 版本

當前版本:`schema_version: "1.0"`

每個 JSON 檔案頂層必須包含 `"schema_version": "1.0"` 欄位。CSV 檔案的版本記錄在 `00_summary.json`。

---

## 00_summary.json

```json
{
  "schema_version": "1.0",
  "tool_version": "0.1.0",
  "analyzed_at": "2026-05-07T14:23:11+08:00",
  "input_file": {
    "path": "complex_macro.xlsm",
    "sha256": "...",
    "size_bytes": 1234567
  },
  "stats": {
    "sheet_count": 32,
    "named_range_count": 47,
    "formula_count": 1834,
    "validation_count": 89,
    "vba_module_count": 6,
    "vba_procedure_count": 41,
    "dependency_edge_count": 3210
  },
  "risk_indicators": {
    "circular_reference_count": 2,
    "external_reference_count": 5,
    "volatile_function_count": 18,
    "dynamic_vba_range_count": 7,
    "deeply_nested_formula_count": 23,
    "orphan_formula_count": 41,
    "cross_sheet_dependency_count": 156
  },
  "complexity_score": 847,
  "migration_difficulty": "high",
  "warnings": [
    {
      "level": "warning",
      "category": "vba",
      "location": "Module1.UpdateRows",
      "message": "Dynamic range detected: Range(\"A\" & lastRow). Marked as has_dynamic_range."
    }
  ]
}
```

`complexity_score` 計算規則見 `reference/output_schema.md`。
`migration_difficulty` enum: `low` / `medium` / `high` / `very_high`。

---

## 01_workbook.json

```json
{
  "schema_version": "1.0",
  "workbook": {
    "file_path": "complex_macro.xlsm",
    "file_sha256": "...",
    "size_bytes": 1234567,
    "has_vba": true,
    "has_external_links": false,
    "default_sheet": "Input",
    "created": "2024-03-15T...",
    "modified": "2026-04-30T...",
    "author": "...",
    "last_modified_by": "..."
  }
}
```

---

## 02_sheets.csv

| Column | Type | Description |
|---|---|---|
| `sheet_id` | int | 序號 (1-based) |
| `sheet_name` | str | sheet 名稱 |
| `sheet_index` | int | 在 workbook 中的位置 (0-based) |
| `is_hidden` | bool | 是否隱藏 |
| `is_very_hidden` | bool | 是否「very hidden」(只能透過 VBA 顯示) |
| `used_range` | str | 例如 `A1:Z100` |
| `row_count` | int | used range 的 row 數 |
| `col_count` | int | used range 的 col 數 |
| `cell_count` | int | 非空 cell 數 |
| `formula_cell_count` | int | 含公式的 cell 數 |

---

## 03_named_ranges.csv

| Column | Type | Description |
|---|---|---|
| `named_range_id` | int | 序號 |
| `range_name` | str | 名稱 (如 `TaxRate`) |
| `scope` | str | `workbook` 或 sheet 名稱 |
| `refers_to` | str | 例如 `Params!$B$2` |
| `has_dynamic_formula` | bool | 是否含 OFFSET/INDIRECT 等動態函式 |
| `is_valid` | bool | refers_to 是否解析成功 (False = #REF!) |

---

## 04_cells.csv

> 只記錄「有意義」的 cell — 含公式、含 validation、被 named range 指到、或被其他 cell 引用。
> 純值且無人引用的 cell 不寫入,避免檔案膨脹。

| Column | Type | Description |
|---|---|---|
| `cell_id` | int | 序號 |
| `sheet_name` | str | |
| `cell_address` | str | 例如 `B7` (不含 sheet 前綴) |
| `qualified_address` | str | 例如 `Calc!B7` (含 sheet 前綴,作為跨表唯一 key) |
| `cell_row` | int | 1-based |
| `cell_col` | int | 1-based |
| `has_formula` | bool | |
| `has_validation` | bool | |
| `is_named` | bool | 是否被 named range 指到 |
| `is_referenced` | bool | 是否被其他 cell 或 VBA 引用 |
| `value_type` | str | `number` / `string` / `boolean` / `date` / `error` / `empty` |
| `raw_value` | str | 原始值 (number 也轉字串,空字串代表 empty) |

---

## 05_formulas.json

```json
{
  "schema_version": "1.0",
  "formulas": [
    {
      "formula_id": 1,
      "qualified_address": "Calc!B7",
      "formula_text": "=IF(VLOOKUP(A1,Params!A:B,2,0)>100, \"高\", \"低\")",
      "formula_category": "mixed",
      "function_list": ["IF", "VLOOKUP"],
      "referenced_cells": [
        {"sheet": "Calc", "address": "A1"},
        {"sheet": "Params", "address": "A:B"}
      ],
      "referenced_named_ranges": [],
      "has_external_reference": false,
      "is_volatile": false,
      "is_array_formula": false,
      "nesting_depth": 2,
      "function_count": 2,
      "complexity_score": 6,
      "ast": {
        "type": "function",
        "name": "IF",
        "args": [...]
      }
    }
  ]
}
```

`formula_category` enum: `lookup` / `branch` / `compute` / `aggregate` / `text` / `reference` / `mixed`
分類規則見 `reference/formula_categories.md`。

`complexity_score = nesting_depth * 2 + function_count + len(referenced_cells)`

---

## 06_validations.csv

| Column | Type | Description |
|---|---|---|
| `validation_id` | int | |
| `qualified_address` | str | 套用 validation 的 cell (或 range 起點) |
| `range_text` | str | 完整套用範圍 (如 `A2:A100`) |
| `validation_type` | str | `list` / `whole` / `decimal` / `date` / `time` / `length` / `custom` |
| `formula1` | str | 第一個條件 (list 來源 / 上限) |
| `formula2` | str | 第二個條件 (下限,可空) |
| `enum_values` | str | 解析後的下拉選項,管道分隔 (`A|B|C`),非 list 類型為空 |
| `allow_blank` | bool | |
| `error_title` | str | |
| `error_message` | str | |

---

## 07_vba_modules.json

```json
{
  "schema_version": "1.0",
  "modules": [
    {
      "vba_module_id": 1,
      "module_name": "Module1",
      "module_type": "standard",
      "line_count": 234,
      "procedure_count": 8,
      "source_code": "..."
    }
  ]
}
```

`module_type` enum: `standard` / `class` / `form` / `sheet` / `workbook` / `unknown`

---

## 08_vba_procedures.json

```json
{
  "schema_version": "1.0",
  "procedures": [
    {
      "vba_procedure_id": 1,
      "vba_module_id": 1,
      "procedure_name": "UpdateFormDetails",
      "procedure_type": "sub",
      "is_public": true,
      "parameters": [
        {"name": "rowIndex", "type_hint": "Long", "is_optional": false}
      ],
      "line_count": 47,
      "reads": [
        {"sheet": "Input", "range": "A2:A100", "via": "explicit_range"},
        {"sheet": "Params", "range": "TaxRate", "via": "named_range"}
      ],
      "writes": [
        {"sheet": "Output", "range": "B2:Z100", "via": "explicit_range"}
      ],
      "calls": ["CalculateRow", "FormatOutput"],
      "triggers": [
        {"event": "Worksheet_Change", "target": "Input!A:A"}
      ],
      "has_dynamic_range": true,
      "dynamic_range_notes": [
        "Line 23: Range(\"A\" & lastRow) — runtime computed"
      ],
      "complexity_score": 23,
      "source_code": "..."
    }
  ]
}
```

`procedure_type` enum: `sub` / `function` / `property_get` / `property_let` / `property_set`
`via` enum: `explicit_range` / `cells_method` / `named_range` / `dynamic` / `unknown`

---

## 09_dependencies.csv

> Cell-to-cell 依賴邊清單。每一條代表「source 變化會影響 target」。

| Column | Type | Description |
|---|---|---|
| `dependency_id` | int | |
| `source_qualified_address` | str | 被依賴方 (如 `Params!A1`) |
| `target_qualified_address` | str | 依賴方 (如 `Calc!B7`) |
| `via` | str | `formula` / `vba_read_write` / `validation` / `named_range` |
| `via_detail` | str | 對應公式 ID 或 VBA procedure ID |
| `is_cross_sheet` | bool | source 與 target 是否跨 sheet |

---

## 10_dependency_graph.json

> 完整 DAG,可用 NetworkX 重建 (node-link format)。

```json
{
  "schema_version": "1.0",
  "directed": true,
  "graph": {
    "node_count": 1234,
    "edge_count": 3210,
    "has_cycles": true,
    "cycle_count": 2,
    "weakly_connected_component_count": 47
  },
  "nodes": [
    {
      "id": "Calc!B7",
      "node_type": "formula_cell",
      "value_type": "number",
      "in_degree": 3,
      "out_degree": 2
    }
  ],
  "edges": [
    {
      "source": "Params!A1",
      "target": "Calc!B7",
      "via": "formula"
    }
  ]
}
```

`node_type` enum: `input_cell` / `formula_cell` / `output_cell` / `named_range` / `vba_procedure`

---

## reports/ 目錄

每張報告都是 CSV 或 JSON,設計給人 + 程式雙用。詳細欄位見 `reference/output_schema.md`。

| 檔名 | 內容 | 排序 |
|---|---|---|
| `formula_categories.csv` | 公式分類統計 | category 字典序 |
| `top_complex_formulas.csv` | 複雜度 Top 50 | complexity_score 由高到低 |
| `hotspot_cells.csv` | 被引用最多次的 cell | in_degree 由高到低 |
| `vba_behavior.csv` | VBA 讀寫概況 | procedure_name 字典序 |
| `cycles.json` | 循環引用清單 | cycle_id |
| `orphans.csv` | 孤島公式 (沒人引用) | qualified_address 字典序 |
| `cross_sheet_refs.csv` | 跨 sheet 依賴邊 | source 字典序 |
