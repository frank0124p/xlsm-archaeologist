# Test Plan

> 各 phase 必跑的測試案例清單。**這是 acceptance 的補充**:agent 寫測試時對照這份。

## Phase 1 — Skeleton

`tests/test_phase_1_skeleton.py`

| Test | 驗證 |
|---|---|
| `test_version_command_works` | `version` 印出版本資訊 |
| `test_analyze_command_callable` | `analyze` 可被呼叫,exit 0 |
| `test_inspect_command_callable` | `inspect` 可被呼叫,exit 0 |
| `test_help_lists_three_commands` | `--help` 列出 version/analyze/inspect |
| `test_logging_writes_to_stderr` | log 訊息走 stderr 不污染 stdout |

## Phase 2 — Extraction

`tests/test_phase_2_extraction.py`

### Workbook metadata

| Test | Fixture | 驗證 |
|---|---|---|
| `test_workbook_extracts_sha256` | simple | sha256 值非空且 64 字元 |
| `test_workbook_detects_has_vba_true` | vba_basic | has_vba=True |
| `test_workbook_detects_has_vba_false` | simple | has_vba=False |

### Sheets

| Test | Fixture | 驗證 |
|---|---|---|
| `test_sheets_extracted` | simple | row 數正確 |
| `test_hidden_sheet_marked` | hidden_sheets | is_hidden=True 的 sheet |
| `test_very_hidden_sheet_marked` | hidden_sheets | is_very_hidden=True |
| `test_used_range_calculated` | formulas_basic | used_range 字串非空 |

### Named ranges

| Test | Fixture | 驗證 |
|---|---|---|
| `test_named_range_basic` | with_named_range | 至少 1 條 |
| `test_named_range_dynamic_detection` | with_named_range | 含 OFFSET 的標 has_dynamic_formula=True |
| `test_named_range_invalid_marked` | (建一份 #REF!) | is_valid=False |

### Cells

| Test | Fixture | 驗證 |
|---|---|---|
| `test_cell_filter_meaningful_only` | simple | row 數 = 0 (無公式無 validation 無 named) |
| `test_formula_cell_recorded` | formulas_basic | 含公式 cell 都有 record |
| `test_validation_cell_recorded` | with_validation | 套 validation 的 cell 在 records 內 |
| `test_is_referenced_initially_false` | formulas_basic | Phase 2 一律 False |

### Validations

| Test | Fixture | 驗證 |
|---|---|---|
| `test_validation_list_literal_parsed` | with_validation | enum_values 解析 |
| `test_validation_list_range_parsed` | with_validation | range 引用解析 |
| `test_validation_whole_number_extracted` | (補) | formula1 / formula2 抓出 |

### Determinism

| Test | 驗證 |
|---|---|
| `test_extraction_deterministic` | 跑兩次 diff 為空 |
| `test_csv_row_order_stable` | row 排序穩定 |

## Phase 3 — Formula Analysis

`tests/test_phase_3_formula.py`

### Tokenizer

| Test | 驗證 |
|---|---|
| `test_tokenize_simple` | `=A1+B1` token 化結果合理 |
| `test_tokenize_with_or_without_equals` | `=A1+B1` 與 `A1+B1` 結果同 |

### Parser

| Test | 公式 | 驗證 |
|---|---|---|
| `test_parse_simple_compute` | `=A1+B1` | AST 是 OperatorNode(+) |
| `test_parse_function_call` | `=SUM(A1:A10)` | FunctionNode(name=SUM) |
| `test_parse_nested_if` | `=IF(IF(A1>0,1,0)>0,1,0)` | depth=2 |
| `test_parse_cross_sheet_range` | `=Sheet2!A1` | RangeNode(sheet=Sheet2) |
| `test_parse_named_range` | `=TaxRate*A1` | 含 NamedRangeNode |
| `test_parse_string_literal` | `=A1&"-"&B1` | 含 OperandNode(text) |
| `test_parse_unparsable_lambda` | `=LAMBDA(x, x*2)(A1)` | is_parsable=False |

### Classifier

| Test | 公式 | 預期分類 |
|---|---|---|
| `test_classify_lookup` | `=VLOOKUP(A1,B:C,2,0)` | lookup |
| `test_classify_branch` | `=IF(A1>0,1,0)` | branch |
| `test_classify_compute` | `=A1*B1+C1` | compute |
| `test_classify_aggregate` | `=SUM(A1:A10)` | aggregate |
| `test_classify_text` | `=CONCAT(A1,B1)` | text |
| `test_classify_reference` | `=Sheet2!A1` | reference |
| `test_classify_pure_named_range` | `=TaxRate` | reference |
| `test_classify_mixed_branch_lookup` | `=IF(VLOOKUP(...)>0,1,0)` | mixed |
| `test_classify_mixed_aggr_branch` | `=SUM(IF(A:A>0,A:A,0))` | mixed |
| `test_classify_branch_with_compute` | `=IF(A1>0,A1+1,0)` | branch (compute 不算進 mixed) |

### Complexity

| Test | 驗證 |
|---|---|
| `test_complexity_simple` | `=A1+B1` score < 5 |
| `test_complexity_deeply_nested` | depth ≥ 5 公式 score ≥ 15 |
| `test_complexity_many_refs` | `=SUM(A1,B1,C1,...,Z1)` score 升高 |

### Metadata

| Test | 驗證 |
|---|---|
| `test_volatile_detected` | `=NOW()` is_volatile=True |
| `test_volatile_offset` | `=OFFSET(A1,0,0)` is_volatile=True |
| `test_external_reference_detected` | `=[Other.xlsx]Sheet1!A1` has_external_reference=True |
| `test_array_formula_marked` | (CSE-array fixture) is_array_formula=True |

### Determinism

| Test | 驗證 |
|---|---|
| `test_formula_output_deterministic` | 跑兩次無差異 |
| `test_function_list_sorted` | function_list 是字典序 |
| `test_referenced_cells_sorted` | referenced_cells 排序穩定 |

## Phase 4 — VBA Analysis

`tests/test_phase_4_vba.py`

### Module extraction

| Test | Fixture | 驗證 |
|---|---|---|
| `test_module_extraction_basic` | vba_basic | 至少 1 個 module |
| `test_module_type_classification` | vba_event_trigger | sheet module 標 module_type=sheet |
| `test_no_vba_returns_empty` | simple | 0 modules |

### Procedure splitter

| Test | 驗證 |
|---|---|
| `test_split_single_sub` | 1 procedure |
| `test_split_multiple_subs_and_functions` | 多個 procedure 切分正確 |
| `test_split_property_get_let_set` | 三個獨立 procedure |
| `test_continuation_merged` | `_` 換行被合併處理 |

### Range detector

| Test | 驗證 |
|---|---|
| `test_range_static_read` | `Range("A1")` 在 expression 中 → read |
| `test_range_static_write` | `Range("A1") = 1` → write |
| `test_cells_static` | `Cells(1,1) = 1` → write Sheet!A1 |
| `test_square_bracket` | `[A1] = 1` → write A1 |
| `test_sheet_qualified` | `Sheets("X").Range("A1") = 1` → write X!A1 |
| `test_named_range_reference` | identifier 匹配 named range → via=named_range |
| `test_dynamic_concat_marked` | `Range("A" & i)` → has_dynamic_range=True |
| `test_dynamic_var_marked` | `Range(var)` → has_dynamic_range=True |
| `test_alias_tracking` | `Set rng = Range("A1"); rng = 1` → write A1 |

### Event triggers

| Test | 驗證 |
|---|---|
| `test_worksheet_change_detected` | triggers 含對應 event |
| `test_intersect_target_extracted` | target 從 Intersect 抓出 |
| `test_workbook_open_detected` | Workbook_Open 在 workbook module 偵測到 |

### Call graph

| Test | 驗證 |
|---|---|
| `test_call_graph_simple` | A 呼叫 B,A.calls=[B] |
| `test_call_graph_cross_module` | Module1 中 A 呼叫 Module2.B |
| `test_builtin_excluded` | MsgBox/Range/Cells 不在 calls |
| `test_string_call_excluded` | 字串 `"Foo"` 中含其他 procedure 名不算 call |

### 異常情境

| Test | 驗證 |
|---|---|
| `test_encrypted_vba_skipped` | 寫 warning,不 raise |
| `test_module_with_no_procedures` | 空 module,procedure_count=0 |

## Phase 5 — Dependency Graph

`tests/test_phase_5_graph.py`

| Test | Fixture | 驗證 |
|---|---|---|
| `test_graph_built_from_formulas` | formulas_basic | edge 數合理 |
| `test_named_range_intermediary_node` | with_named_range | 圖中有 `_named:Foo` 節點 |
| `test_vba_procedure_node` | vba_basic | 圖中有 `_vba:...` 節點 |
| `test_vba_read_edge` | vba_basic | source → vba_procedure 邊 |
| `test_vba_write_edge` | vba_basic | vba_procedure → target 邊 |
| `test_cycle_detection_simple` | circular | cycles 至少 1 條,length=2 |
| `test_cycle_self_loop_excluded` | (建 self-ref fixture) | self-loop 不算 cycle |
| `test_orphan_detection` | orphan_formula | orphans 至少 1 |
| `test_is_referenced_backfilled` | formulas_basic | 04_cells.csv 中 referenced cell 標 True |
| `test_cross_sheet_marked` | cross_sheet_chain | is_cross_sheet=True 的邊出現 |
| `test_descendants_traversal` | cross_sheet_chain | `nx.descendants` 算到三層下 |
| `test_serialization_roundtrip` | formulas_basic | 寫出 JSON 載入後圖結構一致 |
| `test_range_not_expanded` | aggregate fixture | A1:A10 是一個 node 不是十個 |

## Phase 6 — Reports

`tests/test_phase_6_reports.py`

| Test | 驗證 |
|---|---|
| `test_complexity_score_formula` | 給定 indicators 算出 score |
| `test_difficulty_low_threshold` | score=199 → low |
| `test_difficulty_medium_threshold` | score=500 → medium |
| `test_difficulty_high_threshold` | score=999 → high |
| `test_difficulty_very_high_threshold` | score=1000 → very_high |
| `test_categories_pct_sum_to_100` | pct_of_total 加總 100 ±0.1 |
| `test_top_complex_top50_limit` | 公式 100+ 條,輸出限 50 |
| `test_top_complex_sorted` | row 1 score ≥ row 2 score |
| `test_hotspot_in_degree_correct` | 引用 5 次的 cell in_degree=5 |
| `test_vba_behavior_cross_sheet_count` | cross_sheet_count 計算正確 |
| `test_warnings_sorted_by_level` | error 在前、info 在後 |

## E2E

`tests/test_e2e.py`

| Test | 驗證 |
|---|---|
| `test_e2e_simple_xlsm` | simple.xlsm 跑完,11 檔 + reports 都存在 |
| `test_e2e_formulas_complex` | formulas_complex 完整跑通 |
| `test_e2e_deterministic` | 同一 fixture 跑兩次,所有檔案 diff 為空 |
| `test_e2e_phases_partial` | `--phases 1,2` 只跑前兩 phase |
| `test_e2e_no_vba_flag` | `--no-vba` 跳過 Phase 4 |
| `test_e2e_force_overwrite` | `--force` 覆寫非空輸出資料夾 |

## CLI

`tests/test_cli.py`

| Test | 驗證 |
|---|---|
| `test_cli_invalid_path_exits_1` | 不存在檔案 exit 1 |
| `test_cli_bad_args_exits_2` | 缺必要參數 exit 2 |
| `test_cli_output_dir_not_empty_exits_3` | 已有輸出且未 --force exit 3 |
| `test_cli_quiet_no_progress` | `--quiet` 不顯示 progress |
| `test_cli_log_level_debug` | `--log-level debug` 多印 debug log |
