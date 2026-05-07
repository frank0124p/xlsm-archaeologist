# CSV Schemas Reference

> 每個 CSV 欄位的精細定義、資料型別、限制、範例。
> Header 順序就是 column 順序,**不可調整**。

## 通則

| 項目 | 規則 |
|---|---|
| 編碼 | UTF-8 with BOM |
| 換行 | `\n` (LF) |
| 分隔符 | `,` (逗號) |
| Quote 規則 | 含逗號/換行/雙引號才加雙引號;雙引號內 escape 為 `""` |
| 布林表示 | 全小寫 `true` / `false` |
| 空值 | 空字串 |
| 整數 | 純數字,無千分位 |
| 浮點 | 點號小數,最多 4 位 |
| 字串截斷 | `formula_text` / `raw_value` 等大字串截斷規則見對應欄位 |

---

## 02_sheets.csv

| Column | Type | Nullable | Description | 範例 |
|---|---|---|---|---|
| `sheet_id` | int | no | 1-based 唯一 ID | `1` |
| `sheet_name` | str | no | sheet 名稱 (原始,含特殊字元保留) | `"Calc"` |
| `sheet_index` | int | no | 0-based 在 workbook 中位置 | `0` |
| `is_hidden` | bool | no | 是否隱藏 (一般隱藏) | `false` |
| `is_very_hidden` | bool | no | 是否「very hidden」(只能 VBA 顯示) | `false` |
| `used_range` | str | no | A1 標記範圍 | `"A1:Z100"` |
| `row_count` | int | no | used range 的 row 數 | `100` |
| `col_count` | int | no | used range 的 col 數 | `26` |
| `cell_count` | int | no | 非空 cell 總數 (不限 meaningful) | `2400` |
| `formula_cell_count` | int | no | 含公式的 cell 數 | `156` |

排序:`sheet_index` 升冪。

---

## 03_named_ranges.csv

| Column | Type | Nullable | Description | 範例 |
|---|---|---|---|---|
| `named_range_id` | int | no | | `1` |
| `range_name` | str | no | named range 名稱 (原始大小寫) | `"TaxRate"` |
| `scope` | str | no | `workbook` 或具體 sheet 名稱 | `"workbook"` |
| `refers_to` | str | no | 完整 refers_to 字串 (含 `$`) | `"=Params!$B$2"` |
| `has_dynamic_formula` | bool | no | 是否含 OFFSET/INDIRECT/INDEX | `false` |
| `is_valid` | bool | no | refers_to 是否解析成功 | `true` |

排序:`range_name` 字典序。

---

## 04_cells.csv

> 只記「有意義」的 cell。判定條件見 `phases/phase_2_extraction/README.md`。

| Column | Type | Nullable | Description | 範例 |
|---|---|---|---|---|
| `cell_id` | int | no | | `1` |
| `sheet_name` | str | no | | `"Calc"` |
| `cell_address` | str | no | A1 標記,不含 sheet | `"B7"` |
| `qualified_address` | str | no | 含 sheet 前綴的唯一 ID | `"Calc!B7"` |
| `cell_row` | int | no | 1-based | `7` |
| `cell_col` | int | no | 1-based (A=1) | `2` |
| `has_formula` | bool | no | | `true` |
| `has_validation` | bool | no | | `false` |
| `is_named` | bool | no | 是否被 named range 指到 | `false` |
| `is_referenced` | bool | no | 是否被其他 cell 或 VBA 引用 (Phase 5 回填) | `true` |
| `value_type` | enum | no | 見 enum 清單 | `"number"` |
| `raw_value` | str | yes | cell 原始值 (空字串=empty) | `"42"` |

排序:`qualified_address` 字典序 (sheet name → row → col)。

---

## 06_validations.csv

| Column | Type | Nullable | Description | 範例 |
|---|---|---|---|---|
| `validation_id` | int | no | | `1` |
| `qualified_address` | str | no | 套用 cell (或 range 起點) | `"Input!A2"` |
| `range_text` | str | no | 完整套用範圍 | `"A2:A100"` |
| `validation_type` | enum | no | | `"list"` |
| `formula1` | str | yes | 條件 1 | `"=Params!$A$2:$A$10"` |
| `formula2` | str | yes | 條件 2 (可空) | `""` |
| `enum_values` | str | yes | list 類型解析結果,管道分隔 | `"A\|B\|C"` |
| `allow_blank` | bool | no | | `true` |
| `error_title` | str | yes | | `"無效輸入"` |
| `error_message` | str | yes | | `"請從清單選擇"` |

排序:`qualified_address` 字典序。

備註:`enum_values` 中的 `\|` 在 CSV 中為實際 pipe 字元,值含 pipe 時需 escape (用 `\\|`)。

---

## 09_dependencies.csv

| Column | Type | Nullable | Description | 範例 |
|---|---|---|---|---|
| `dependency_id` | int | no | | `1` |
| `source_qualified_address` | str | no | 被依賴方 | `"Params!A1"` |
| `target_qualified_address` | str | no | 依賴方 | `"Calc!B7"` |
| `via` | enum | no | 依賴成因 | `"formula"` |
| `via_detail` | str | yes | 對應 formula_id 或 procedure ID | `"formula:42"` |
| `is_cross_sheet` | bool | no | source 與 target 是否跨 sheet | `true` |

排序:`(source_qualified_address, target_qualified_address, via)` tuple 字典序。

via_detail 格式約定:
- 若 via=formula → `"formula:{formula_id}"`
- 若 via=vba_read_write → `"vba_procedure:{vba_procedure_id}"`
- 若 via=validation → `"validation:{validation_id}"`
- 若 via=named_range → `"named_range:{named_range_id}"`

---

## reports/formula_categories.csv

| Column | Type | Nullable | Description |
|---|---|---|---|
| `category` | enum | no | formula_category enum 值 |
| `formula_count` | int | no | |
| `total_complexity` | int | no | 該類所有公式 complexity_score 加總 |
| `avg_complexity` | float | no | 平均 (4 位小數) |
| `max_complexity` | int | no | |
| `pct_of_total` | float | no | 百分比 (帶 1 位小數,如 `23.4`) |

排序:`formula_count` 由高到低。

---

## reports/top_complex_formulas.csv

| Column | Type | Nullable | Description |
|---|---|---|---|
| `rank` | int | no | 1-50 |
| `qualified_address` | str | no | |
| `formula_text` | str | no | 截斷至 200 字 (超過附 `...`) |
| `formula_category` | enum | no | |
| `nesting_depth` | int | no | |
| `function_count` | int | no | |
| `referenced_cell_count` | int | no | |
| `complexity_score` | int | no | |

排序:`complexity_score` 由高到低,並列依 `qualified_address` 字典序。
最多 50 row。

---

## reports/hotspot_cells.csv

| Column | Type | Nullable | Description |
|---|---|---|---|
| `rank` | int | no | 1-50 |
| `qualified_address` | str | no | |
| `node_type` | enum | no | |
| `in_degree` | int | no | |
| `referenced_by_formula_count` | int | no | |
| `referenced_by_vba_count` | int | no | |
| `value_type` | enum | yes | 對 named_range / vba_procedure node 為空 |
| `raw_value` | str | yes | 截斷至 100 字 |

排序:`in_degree` 由高到低,並列依 `qualified_address` 字典序。
最多 50 row。

---

## reports/vba_behavior.csv

| Column | Type | Nullable | Description |
|---|---|---|---|
| `module_name` | str | no | |
| `procedure_name` | str | no | |
| `procedure_type` | enum | no | |
| `line_count` | int | no | |
| `read_count` | int | no | reads 陣列長度 |
| `write_count` | int | no | writes 陣列長度 |
| `cross_sheet_read_count` | int | no | reads 中 sheet 不等於 procedure 所在 sheet 的數量 (對 standard module 全部視為 cross) |
| `cross_sheet_write_count` | int | no | 同上 |
| `has_dynamic_range` | bool | no | |
| `has_event_trigger` | bool | no | triggers 非空 |
| `call_count` | int | no | calls 陣列長度 |
| `complexity_score` | int | no | |

排序:`complexity_score` 由高到低,並列依 `(module_name, procedure_name)` 字典序。

---

## reports/orphans.csv

| Column | Type | Nullable | Description |
|---|---|---|---|
| `qualified_address` | str | no | |
| `formula_id` | int | no | |
| `formula_category` | enum | no | |
| `complexity_score` | int | no | |
| `reason` | str | no | 為什麼判定為 orphan,目前固定 `"in_degree=0"` |

排序:`qualified_address` 字典序。

---

## reports/cross_sheet_refs.csv

| Column | Type | Nullable | Description |
|---|---|---|---|
| `source_qualified_address` | str | no | |
| `source_sheet` | str | no | |
| `target_qualified_address` | str | no | |
| `target_sheet` | str | no | |
| `via` | enum | no | |
| `via_detail` | str | yes | |

排序:`(source_sheet, source_qualified_address, target_qualified_address)` tuple 字典序。
