# Phase 4 — Tasks

## Models

- [ ] 建立 `models/vba.py`:
    - `VbaModuleType` enum (`standard`/`class`/`form`/`sheet`/`workbook`/`unknown`)
    - `ProcedureType` enum (`sub`/`function`/`property_get`/`property_let`/`property_set`)
    - `RangeAccessVia` enum (`explicit_range`/`cells_method`/`named_range`/`dynamic`/`unknown`)
    - `RangeAccess`(sheet, range, via)
    - `EventTrigger`(event, target)
    - `Parameter`(name, type_hint, is_optional)
    - `VbaModuleRecord`(vba_module_id, module_name, module_type, line_count,
                       procedure_count, source_code)
    - `VbaProcedureRecord`(vba_procedure_id, vba_module_id, procedure_name, procedure_type,
                          is_public, parameters, line_count, reads, writes, calls, triggers,
                          has_dynamic_range, dynamic_range_notes, complexity_score, source_code)
- [ ] commit: `feat(models): add vba analysis models`

## Extractor

- [ ] `extractors/vba_extractor.py`:
    - `extract_vba_modules(path) -> Iterator[VbaModuleRecord]`
    - 用 `oletools.olevba3.VBA_Parser`
    - 偵測 module_type:
        - 路徑含 `Sheet` / 物件名匹配 sheet 名 → `sheet`
        - 路徑含 `ThisWorkbook` → `workbook`
        - `cls` 副檔名 → `class`
        - `frm` → `form`
        - 其他 → `standard`
    - 加密 vbaProject → warning + 回傳空
- [ ] commit: `feat(vba): add module extractor`

## Procedure Splitter

- [ ] `analyzers/vba_procedure_splitter.py`:
    - `split_procedures(module: VbaModuleRecord) -> Iterator[ProcedureChunk]`
    - regex 匹配 procedure 開頭與 End
    - 預處理:合併 `_` continuation、移除註解 (但保留行數計算)
    - 處理巢狀 (其實 VBA 不允許巢狀 procedure,但要 detect 並 raise)
- [ ] commit: `feat(vba): add procedure splitter`

## Range Detector

- [ ] `analyzers/vba_range_detector.py`:
    - `detect_range_accesses(code: str) -> tuple[list[RangeAccess], list[RangeAccess], list[str]]`
    - 三個回傳值:reads, writes, dynamic_notes
    - 內部:
        - 一遍掃 alias (`Set var = ...Range...`)
        - 一遍掃 read/write (用 `=` 位置判定)
        - 一遍掃動態 pattern,加 dynamic_notes
- [ ] 實作 alias 追蹤 (procedure scope)
- [ ] 實作 named range 引用偵測 (透過全域 named range 清單)
- [ ] commit: `feat(vba): add range read-write detector`

## Event Trigger Detector

- [ ] `analyzers/vba_range_detector.py::detect_triggers`:
    - 對 sheet/workbook module 才掃
    - 匹配 `Worksheet_Change` / `Worksheet_SelectionChange` / `Workbook_Open` 等
    - 嘗試從 body 內 Intersect(Target, Range("...")) 抓 target range (heuristic)
- [ ] commit: `feat(vba): add event trigger detector`

## Call Graph

- [ ] `analyzers/vba_call_graph.py`:
    - `extract_calls(code: str, all_procedure_names: set[str], builtin_whitelist: set[str]) -> list[str]`
    - 兩階段:
        1. 收集所有 known procedure name (Phase 4 內已知)
        2. 在 source 中找對它們的引用
    - 排除 declaration、字串、註解內出現
- [ ] commit: `feat(vba): add procedure call graph builder`

## 主協調者

- [ ] `analyzers/vba_analyzer.py::analyze_vba`:
    - 兩 pass:
        1. extract modules + split procedures (取得 all_procedure_names)
        2. 對每個 procedure 跑 range_detect / trigger_detect / call_extract
    - 計算 complexity_score (line_count + ref_count + branch_count)
- [ ] commit: `feat(vba): add main analyzer pipeline`

## CLI 串接

- [ ] 修改 `cli.py`:
    - Phase 4 寫 `07_vba_modules.json` 與 `08_vba_procedures.json`
    - 處理 `--no-vba` flag
- [ ] commit: `feat(cli): wire vba analysis into pipeline`

## Tests

- [ ] 建立 fixture `tests/fixtures/vba_basic.xlsm`:
    - Module1 含一個 Sub:讀 A1、寫 B1
    - 用 `Range("A1")` 與 `Range("B1") = ...`
- [ ] 建立 fixture `tests/fixtures/vba_dynamic_range.xlsm`:
    - 一個 Sub 用 `Range("A" & lastRow)` 寫值
    - 一個 Sub 用 `Cells(i, j)` 在迴圈內讀寫
- [ ] 建立 fixture `tests/fixtures/vba_event_trigger.xlsm`:
    - Sheet1 module 含 `Worksheet_Change` event handler
    - Intersect 限定在 A:A
- [ ] 建立 fixture `tests/fixtures/vba_call_graph.xlsm`:
    - Module1 有 Sub Main 呼叫 SubA、SubB
    - Module2 有 SubA 呼叫 SubC
- [ ] `tests/test_phase_4_vba.py`:
    - test_module_extraction
    - test_procedure_splitter_basic
    - test_procedure_splitter_property
    - test_range_detect_static_read
    - test_range_detect_static_write
    - test_range_detect_dynamic_marked
    - test_range_detect_alias_tracking
    - test_event_trigger_detection
    - test_call_graph_simple
    - test_encrypted_vba_skipped_with_warning (mock 一個 case)
- [ ] commit: `test(vba): comprehensive analyzer tests`

## Quality

- [ ] `uv run pytest tests/test_phase_4_vba.py` 全 pass
- [ ] `uv run ruff check .` 零錯誤
- [ ] `uv run mypy src` 零錯誤
- [ ] vba 模組覆蓋率 ≥ 90%

## 收尾

- [ ] 寫 `phase_4_summary.md`
- [ ] **停下來等 review**
