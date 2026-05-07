# Output Schema Reference

> 補充 `DATA_MODEL.md`,給 agent 實作時的精確 schema 對照。

## Schema 版本管理

- 當前版本:`1.0`
- 改 schema 必須:bump version + 在本檔案附 changelog + commit message 標 BREAKING CHANGE

## JSON 共通規則

- 所有頂層 object 都有 `"schema_version": "1.0"`
- `indent=2`
- `ensure_ascii=False` (允許中文)
- `sort_keys=True` 在頂層;陣列內物件按 schema 定義順序
- timestamp 一律 ISO 8601 含 timezone (`+08:00` for Taiwan)

## 完整 JSON 範例

### 00_summary.json

```json
{
  "schema_version": "1.0",
  "tool_version": "0.1.0",
  "analyzed_at": "2026-05-07T14:23:11+08:00",
  "input_file": {
    "path": "complex_macro.xlsm",
    "sha256": "e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855",
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
      "location": "Module1.UpdateRows#L23",
      "message": "Dynamic range detected: Range(\"A\" & lastRow). Marked as has_dynamic_range."
    }
  ]
}
```

### 05_formulas.json

```json
{
  "schema_version": "1.0",
  "formulas": [
    {
      "formula_id": 1,
      "qualified_address": "Calc!B7",
      "formula_text": "=IF(VLOOKUP(A1,Params!A:B,2,0)>100,\"高\",\"低\")",
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
      "is_parsable": true,
      "parse_error": null,
      "nesting_depth": 2,
      "function_count": 2,
      "complexity_score": 6,
      "ast": {
        "type": "function",
        "name": "IF",
        "args": [
          {
            "type": "operator",
            "op": ">",
            "left": {
              "type": "function",
              "name": "VLOOKUP",
              "args": [
                {"type": "range", "sheet": null, "address": "A1"},
                {"type": "range", "sheet": "Params", "address": "A:B"},
                {"type": "operand", "operand_type": "number", "value": "2"},
                {"type": "operand", "operand_type": "number", "value": "0"}
              ]
            },
            "right": {"type": "operand", "operand_type": "number", "value": "100"}
          },
          {"type": "operand", "operand_type": "text", "value": "高"},
          {"type": "operand", "operand_type": "text", "value": "低"}
        ]
      }
    }
  ]
}
```

### 10_dependency_graph.json (NetworkX node-link 格式)

```json
{
  "schema_version": "1.0",
  "directed": true,
  "multigraph": false,
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
    },
    {
      "id": "_named:TaxRate",
      "node_type": "named_range",
      "refers_to": "Params!$B$2",
      "in_degree": 0,
      "out_degree": 12
    },
    {
      "id": "_vba:Module1.UpdateRows",
      "node_type": "vba_procedure",
      "in_degree": 5,
      "out_degree": 8
    }
  ],
  "edges": [
    {
      "source": "Params!A1",
      "target": "Calc!B7",
      "via": "formula",
      "formula_id": 1
    },
    {
      "source": "_vba:Module1.UpdateRows",
      "target": "Output!B2:Z100",
      "via": "vba_read_write"
    }
  ]
}
```

### reports/cycles.json

```json
{
  "schema_version": "1.0",
  "cycles": [
    {
      "cycle_id": 1,
      "length": 2,
      "nodes": ["Sheet1!A1", "Sheet1!B1"],
      "edges_via": ["formula", "formula"]
    }
  ]
}
```

## CSV 共通規則

- UTF-8 BOM (供 Excel 開啟時不亂碼)
- 引號:當值含逗號、換行、引號時加雙引號;否則不加
- 換行符:`\n` (不用 CRLF)
- header 在第一行
- 布林:小寫 `true` / `false`
- 空值:空字串 (不用 `null`)

## 完整 CSV header 對照

### 02_sheets.csv

```
sheet_id,sheet_name,sheet_index,is_hidden,is_very_hidden,used_range,row_count,col_count,cell_count,formula_cell_count
```

### 03_named_ranges.csv

```
named_range_id,range_name,scope,refers_to,has_dynamic_formula,is_valid
```

### 04_cells.csv

```
cell_id,sheet_name,cell_address,qualified_address,cell_row,cell_col,has_formula,has_validation,is_named,is_referenced,value_type,raw_value
```

### 06_validations.csv

```
validation_id,qualified_address,range_text,validation_type,formula1,formula2,enum_values,allow_blank,error_title,error_message
```

### 09_dependencies.csv

```
dependency_id,source_qualified_address,target_qualified_address,via,via_detail,is_cross_sheet
```

### reports/formula_categories.csv

```
category,formula_count,total_complexity,avg_complexity,max_complexity,pct_of_total
```

### reports/top_complex_formulas.csv

```
rank,qualified_address,formula_text,formula_category,nesting_depth,function_count,referenced_cell_count,complexity_score
```

### reports/hotspot_cells.csv

```
rank,qualified_address,node_type,in_degree,referenced_by_formula_count,referenced_by_vba_count,value_type,raw_value
```

### reports/vba_behavior.csv

```
module_name,procedure_name,procedure_type,line_count,read_count,write_count,cross_sheet_read_count,cross_sheet_write_count,has_dynamic_range,has_event_trigger,call_count,complexity_score
```

### reports/orphans.csv

```
qualified_address,formula_id,formula_category,complexity_score,reason
```

### reports/cross_sheet_refs.csv

```
source_qualified_address,source_sheet,target_qualified_address,target_sheet,via,via_detail
```

## Enum 值清單

統一在這列出,所有 phase 都要照這個範圍:

| Enum | 允許值 |
|---|---|
| `value_type` | `number` / `string` / `boolean` / `date` / `error` / `empty` |
| `validation_type` | `list` / `whole` / `decimal` / `date` / `time` / `length` / `custom` |
| `formula_category` | `lookup` / `branch` / `compute` / `aggregate` / `text` / `reference` / `mixed` |
| `module_type` | `standard` / `class` / `form` / `sheet` / `workbook` / `unknown` |
| `procedure_type` | `sub` / `function` / `property_get` / `property_let` / `property_set` |
| `range_access_via` | `explicit_range` / `cells_method` / `named_range` / `dynamic` / `unknown` |
| `node_type` | `input_cell` / `formula_cell` / `output_cell` / `named_range` / `vba_procedure` |
| `dependency_via` | `formula` / `vba_read_write` / `validation` / `named_range` |
| `migration_difficulty` | `low` / `medium` / `high` / `very_high` |
| `warning_level` | `info` / `warning` / `error` |
