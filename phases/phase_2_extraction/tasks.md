# Phase 2 — Tasks

## Models 先行

- [ ] 建立 `models/workbook.py`:
    - `WorkbookRecord`(file_path, file_sha256, size_bytes, has_vba, has_external_links, ...)
    - `SheetRecord`(sheet_id, sheet_name, sheet_index, is_hidden, is_very_hidden, ...)
- [ ] 建立 `models/named_range.py`:
    - `NamedRangeRecord`(named_range_id, range_name, scope, refers_to, has_dynamic_formula, is_valid)
- [ ] 建立 `models/cell.py`:
    - `CellRecord`(cell_id, sheet_name, cell_address, qualified_address, ...,
                  has_formula, has_validation, is_named, is_referenced, value_type, raw_value)
    - `ValidationRecord`(validation_id, qualified_address, range_text, validation_type,
                        formula1, formula2, enum_values, allow_blank, error_title, error_message)
- [ ] 所有 model 加 docstring + `model_config = ConfigDict(frozen=True)` (immutable)
- [ ] commit: `feat(models): add extraction-phase pydantic models`

## Extractors

- [ ] `extractors/workbook_extractor.py`:
    - `extract_workbook(path: Path) -> WorkbookRecord`
    - 計算 sha256、size、has_vba (檢查 keep_vba 後 wb.vba_archive)
- [ ] `extractors/sheet_extractor.py`:
    - `extract_sheets(wb) -> Iterator[SheetRecord]`
    - 處理 hidden / very_hidden 區別
    - used_range 用 `sheet.calculate_dimension()`
- [ ] `extractors/named_range_extractor.py`:
    - `extract_named_ranges(wb) -> Iterator[NamedRangeRecord]`
    - 解析 dynamic_formula 偵測
    - `is_valid` False 時 refers_to 含 `#REF!`
- [ ] `extractors/validation_extractor.py`:
    - `extract_validations(wb) -> Iterator[ValidationRecord]`
    - list 類型解析 enum_values (支援字面值與 range 引用兩種)
- [ ] `extractors/cell_extractor.py`:
    - `extract_cells(wb, named_addresses, validation_addresses) -> Iterator[CellRecord]`
    - 只 yield 「有意義」cell
    - `is_referenced` 先填 False
- [ ] commit (每個 extractor 一個): `feat(extraction): <extractor name>`

## Serializers

- [ ] `serializers/json_writer.py`:
    - `write_json(path, data, schema_version="1.0")`
    - 強制 indent=2、sort_keys=True、ensure_ascii=False
- [ ] `serializers/csv_writer.py`:
    - `write_csv(path, records, columns)`
    - 用內建 `csv.DictWriter`,quote 模式統一
    - 確保 column 順序固定 (照 `DATA_MODEL.md`)
- [ ] commit: `feat(serializers): add json and csv writers`

## CLI 串接

- [ ] 修改 `cli.py` 的 `analyze` command:
    - 真正呼叫 extractors,寫 01-04, 06 號檔案
    - 用 rich progress bar 顯示進度
    - 仍保留 phase 5/6 的 placeholder
- [ ] 處理 `--phases` 參數 (允許跳過 Phase 3+)
- [ ] commit: `feat(cli): wire extraction phase into analyze command`

## Tests

- [ ] 建立 fixture `tests/fixtures/simple.xlsm`:
    - 1 sheet (`Data`),A1:C3 純值
    - 0 公式、0 VBA、0 named range
- [ ] 建立 fixture `tests/fixtures/with_validation.xlsm`:
    - 2 sheets
    - 1 個 list validation (字面值)、1 個 list validation (range 引用)
- [ ] 建立 fixture `tests/fixtures/with_named_range.xlsm`:
    - 2 個 named range:1 個正常、1 個含 OFFSET (dynamic)
- [ ] 建立 fixture `tests/fixtures/hidden_sheets.xlsm`:
    - 含 hidden 與 very_hidden sheet 各一
- [ ] `tests/test_phase_2_extraction.py`:
    - test_workbook_metadata
    - test_sheet_extraction
    - test_named_range_dynamic_detection
    - test_validation_list_literal
    - test_validation_list_range_reference
    - test_cell_filter_meaningful_only
    - test_extraction_deterministic (跑兩次比對輸出)
- [ ] commit: `test(extraction): add fixture-based tests`

## Quality

- [ ] `uv run pytest tests/test_phase_2_extraction.py` 全 pass
- [ ] `uv run ruff check .` 零錯誤
- [ ] `uv run mypy src` 零錯誤
- [ ] 覆蓋率 ≥ 80% (`uv run pytest --cov=src/xlsm_archaeologist/extractors`)

## 收尾

- [ ] 寫 `phase_2_summary.md`:
    - 完成清單
    - 拿任一 fixture 跑 `analyze` 的 output 範例
    - 已知限制 (e.g. is_referenced 待 Phase 5 補)
- [ ] **停下來等 review**
