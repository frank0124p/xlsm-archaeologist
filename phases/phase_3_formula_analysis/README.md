# Phase 3: Formula Analysis

## 目標

對 Phase 2 抽出的「含公式 cell」做:
1. Tokenize 公式
2. 建立 simple AST
3. 公式分類 (lookup / branch / compute / aggregate / text / reference / mixed)
4. 計算複雜度
5. 抽取 referenced_cells / referenced_named_ranges
6. 寫進 `05_formulas.json`

## 為什麼自實作 parser 而不用 `formulas` 套件

- `formulas` 套件能力強但較重,且我們**不執行公式**,只需要結構
- `openpyxl.formula.tokenizer.Tokenizer` 內建,夠用做分類與依賴抽取
- 自寫 parser 約 200-300 行,可控、易測

## 模組

```
src/xlsm_archaeologist/
├── analyzers/
│   ├── formula_analyzer.py       # 主協調者
│   ├── formula_tokenizer.py      # 包裝 openpyxl tokenizer
│   ├── formula_parser.py         # token → AST
│   ├── formula_classifier.py     # AST → category
│   └── formula_complexity.py     # 複雜度計算
└── models/
    └── formula.py                # FormulaRecord, FormulaCategory enum
```

## AST 設計 (簡化版)

```python
class AstNode(BaseModel):
    type: Literal["function", "operand", "operator", "range", "named_range"]

class FunctionNode(AstNode):
    type: Literal["function"] = "function"
    name: str                      # IF / VLOOKUP / SUM ...
    args: list[AstNode]

class OperandNode(AstNode):
    type: Literal["operand"] = "operand"
    operand_type: Literal["number", "text", "logical", "error"]
    value: str

class OperatorNode(AstNode):
    type: Literal["operator"] = "operator"
    op: str                        # + - * / ^ & = > < <= >= <>
    left: AstNode
    right: AstNode

class RangeNode(AstNode):
    type: Literal["range"] = "range"
    sheet: str | None              # None 代表同 sheet
    address: str                   # A1 / A1:B10 / A:A

class NamedRangeNode(AstNode):
    type: Literal["named_range"] = "named_range"
    name: str
```

## 分類規則 (摘要)

詳細規則見 `reference/formula_categories.md`。摘要:

```python
def classify(ast: AstNode) -> FormulaCategory:
    funcs = collect_function_names(ast)

    if not funcs:
        # 純引用 / 純運算
        if has_only_range_or_named(ast):
            return "reference"
        return "compute"

    has_lookup = funcs & LOOKUP_FUNCS    # VLOOKUP, HLOOKUP, XLOOKUP, INDEX, MATCH, LOOKUP
    has_branch = funcs & BRANCH_FUNCS    # IF, IFS, SWITCH, CHOOSE, IFERROR
    has_aggr   = funcs & AGGR_FUNCS      # SUM, SUMIF, COUNT, AVG, MAX, MIN, ...
    has_text   = funcs & TEXT_FUNCS      # CONCAT, LEFT, RIGHT, TEXTJOIN, ...

    type_count = sum([has_lookup, has_branch, has_aggr, has_text])
    if type_count >= 2:
        return "mixed"

    if has_lookup:  return "lookup"
    if has_branch:  return "branch"
    if has_aggr:    return "aggregate"
    if has_text:    return "text"
    return "compute"
```

## 複雜度計算

```python
complexity_score = (
    nesting_depth * 2
    + function_count
    + len(referenced_cells)
)
```

`nesting_depth`:AST 中 FunctionNode 嵌套最深層數
`function_count`:AST 中 FunctionNode 總數
`referenced_cells`:RangeNode 唯一去重後的數量

## Volatile 偵測

```python
VOLATILE_FUNCS = {"NOW", "TODAY", "RAND", "RANDBETWEEN", "OFFSET", "INDIRECT", "INFO", "CELL"}
is_volatile = bool(funcs & VOLATILE_FUNCS)
```

## 已知限制

- ⚠ Dynamic array 公式 (含 `@`、`#`) 標記為 `unparsable: true`,保留原文,加 warning
- ⚠ LET / LAMBDA / 自訂 LAMBDA 標記為 `unparsable: true`
- ⚠ 含外部活頁簿引用 (`[OtherFile.xlsx]Sheet1!A1`) 拆出 `has_external_reference: true`,
  但不解析外部檔案
- ⚠ 跨表引用 (`Sheet2!A1`) 必須完整保留 sheet 名稱

## 驗收

見 `acceptance.md`。
