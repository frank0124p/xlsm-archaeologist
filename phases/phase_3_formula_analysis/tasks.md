# Phase 3 — Tasks

## Models

- [ ] 建立 `models/formula.py`:
    - `FormulaCategory` enum (Literal)
    - `AstNode` base + `FunctionNode` / `OperandNode` / `OperatorNode` / `RangeNode` / `NamedRangeNode`
    - `CellRef`(sheet, address)
    - `FormulaRecord`(formula_id, qualified_address, formula_text, formula_category,
                     function_list, referenced_cells, referenced_named_ranges,
                     has_external_reference, is_volatile, is_array_formula,
                     nesting_depth, function_count, complexity_score, ast,
                     is_parsable, parse_error)
- [ ] commit: `feat(models): add formula analysis models`

## Tokenizer 包裝

- [ ] `analyzers/formula_tokenizer.py`:
    - `tokenize(formula_text: str) -> list[Token]`
    - 包裝 `openpyxl.formula.tokenizer.Tokenizer`
    - 統一處理 `=` 開頭/不開頭
- [ ] commit: `feat(formula): add formula tokenizer wrapper`

## Parser

- [ ] `analyzers/formula_parser.py`:
    - `parse(tokens: list[Token]) -> AstNode`
    - 處理:
        - 函式呼叫 `FUNC(arg1, arg2)`
        - 巢狀 `IF(IF(...), ...)`
        - 二元運算 `A + B * C` (注意優先級)
        - 一元運算 `-A`
        - 範圍 `A1:B10`、`Sheet2!A1:B10`
        - 字面值 (number/text/bool/error)
        - named range
    - parse 失敗:回傳特殊 `UnparsableNode`,raise 收集到 warnings
- [ ] commit: `feat(formula): add token-to-ast parser`

## Classifier

- [ ] `analyzers/formula_classifier.py`:
    - `classify(ast: AstNode) -> FormulaCategory`
    - 內部用 set lookup (LOOKUP_FUNCS / BRANCH_FUNCS / AGGR_FUNCS / TEXT_FUNCS)
    - 詳細列表見 `reference/formula_categories.md`
- [ ] commit: `feat(formula): add ast-based classifier`

## Complexity

- [ ] `analyzers/formula_complexity.py`:
    - `compute_complexity(ast, references) -> tuple[depth, func_count, score]`
    - 巢狀深度:遞迴算 FunctionNode 最大層
    - 函式數:遞迴累加 FunctionNode
- [ ] commit: `feat(formula): add complexity calculator`

## Reference Extractor

- [ ] `analyzers/formula_analyzer.py`:
    - `extract_references(ast) -> tuple[list[CellRef], list[str]]`
    - 第一個 list:所有 RangeNode (去重、排序)
    - 第二個 list:所有 NamedRangeNode (去重、排序)
- [ ] `is_volatile` 偵測 (VOLATILE_FUNCS)
- [ ] `has_external_reference` 偵測 (檢查 RangeNode 是否含 `[...]` 模式)
- [ ] commit: `feat(formula): add reference extractor and metadata flags`

## 主協調者

- [ ] `analyzers/formula_analyzer.py::analyze_formulas`:
    - 輸入:Phase 2 抽出的 cell list (only has_formula=true)
    - 輸出:`Iterator[FormulaRecord]`
    - 每條公式跑 tokenize → parse → classify → complexity → references
    - parse 失敗 → record 設 is_parsable=false,加 warnings
- [ ] commit: `feat(formula): add main analyzer pipeline`

## Serializer 串接

- [ ] 修改 `cli.py`:
    - Phase 3 寫 `05_formulas.json`
    - 用 rich progress bar (per formula)
- [ ] commit: `feat(cli): wire formula analysis into pipeline`

## Tests

- [ ] 建立 fixture `tests/fixtures/formulas_basic.xlsm`:
    - 各類公式各一條:
        - lookup: `=VLOOKUP(A1, Params!A:B, 2, 0)`
        - branch: `=IF(A1>0, "正", "負")`
        - compute: `=A1*B1+C1`
        - aggregate: `=SUM(A1:A10)`
        - text: `=A1 & "-" & B1`
        - reference: `=Sheet2!A1`
- [ ] 建立 fixture `tests/fixtures/formulas_complex.xlsm`:
    - 巢狀 IF (深度 5+)
    - mixed: `=IF(VLOOKUP(...)>0, SUM(...), 0)`
    - cross-sheet: `=Sheet1!A1 + Sheet2!B1`
    - volatile: `=OFFSET(A1, 0, 0)`
    - named range: `=TaxRate * BasePrice`
- [ ] `tests/test_phase_3_formula.py`:
    - test_tokenize_simple
    - test_parse_nested_if (深度 ≥ 3)
    - test_classify_each_category (6 個 fixture 對應 6 個分類)
    - test_classify_mixed
    - test_complexity_score_simple
    - test_complexity_score_deeply_nested
    - test_extract_references_cross_sheet
    - test_extract_references_named_range
    - test_volatile_detection
    - test_external_reference_detection
    - test_unparsable_lambda (LAMBDA / LET 標記為 unparsable)
- [ ] commit: `test(formula): comprehensive analyzer tests`

## Quality

- [ ] `uv run pytest tests/test_phase_3_formula.py` 全 pass
- [ ] `uv run ruff check .` 零錯誤
- [ ] `uv run mypy src` 零錯誤
- [ ] 覆蓋率 ≥ 90% (`uv run pytest --cov=src/xlsm_archaeologist/analyzers`)

## 收尾

- [ ] 寫 `phase_3_summary.md`
- [ ] **停下來等 review**
