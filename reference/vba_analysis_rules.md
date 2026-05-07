# VBA Analysis Rules

> Phase 4 的權威規則文件。VBA 讀寫識別、動態 range 標記、event trigger 偵測。

## 設計哲學

**靜態分析做不到的事不要假裝做到。** 標記為 `has_dynamic_range=true` 比假裝解出來更有價值 —
下游可以根據旗標決定要 manual review 哪些 procedure。

## 規則 1:Procedure 切分

### Pattern

```python
PROCEDURE_PATTERN = re.compile(
    r"^[ \t]*"
    r"(?P<scope>Public\s+|Private\s+|Friend\s+)?"
    r"(?P<static>Static\s+)?"
    r"(?P<kind>Sub|Function|Property\s+(?:Get|Let|Set))\s+"
    r"(?P<name>[A-Za-z_]\w*)\s*"
    r"(?:\((?P<params>[^)]*)\))?"
    r"(?:\s+As\s+(?P<return_type>\w+))?"
    r"\s*$",
    re.IGNORECASE | re.MULTILINE,
)

END_PATTERN = re.compile(
    r"^[ \t]*End\s+(Sub|Function|Property)\s*$",
    re.IGNORECASE | re.MULTILINE,
)
```

### 預處理

VBA 的 `_` continuation 會把一行分成多行。掃描前必須先合併:

```python
def merge_continuations(code: str) -> str:
    # 把 " _\n   " 合併
    return re.sub(r"\s+_\s*\n\s*", " ", code)
```

### 處理順序

1. 合併 continuations
2. 找出所有 PROCEDURE_PATTERN 的位置
3. 找出所有 END_PATTERN 的位置
4. 配對 (procedure_start[i] 到下一個 end_pattern)
5. 切出 source_code

### 邊界

- VBA 不允許巢狀 procedure — 若偵測到應 raise warning
- Property Get/Let/Set 是三個獨立 procedure (即使同名)
- Declare 語句 (`Declare Function`) 不是 procedure,是 API binding,要 skip

## 規則 2:Range 讀寫識別 — 靜態可解 patterns

### RANGE_LITERAL — `Range("...")`

```python
RANGE_LITERAL = re.compile(
    r'Range\s*\(\s*"(?P<addr>[^"]+)"\s*\)',
    re.IGNORECASE,
)
# Match: Range("A1") / Range("A1:B10") / Range("A:A")
# 不 match: Range("A" & i) — 含 & 不在 quote 內
```

### CELLS_LITERAL — `Cells(int, int)`

```python
CELLS_LITERAL = re.compile(
    r'Cells\s*\(\s*(?P<row>\d+)\s*,\s*(?P<col>\d+|"[A-Z]+")\s*\)',
    re.IGNORECASE,
)
# Match: Cells(1, 1) / Cells(5, "A")
```

`Cells(row, col)` 中:
- col 為整數 → A1 標記轉換 (1 → A, 27 → AA)
- col 為字串 → 直接用

### SQUARE_BRACKET — `[A1]` 標記

```python
SQUARE_BRACKET = re.compile(
    r'\[\s*(?P<addr>[A-Za-z]+\d+(?::[A-Za-z]+\d+)?)\s*\]',
)
# Match: [A1] / [A1:B10]
# 注意:VBA 中 [...] 也可能是 collection access,要排除
# 排除規則:前面緊接著識別字 (如 .Item[1]) 就不是 range
```

### SHEET_RANGE — `Sheets("X").Range("Y")`

```python
SHEET_RANGE = re.compile(
    r'(?:Sheets|Worksheets)\s*\(\s*"(?P<sheet>[^"]+)"\s*\)\s*\.\s*Range\s*\(\s*"(?P<addr>[^"]+)"\s*\)',
    re.IGNORECASE,
)
SHEET_CELLS = re.compile(
    r'(?:Sheets|Worksheets)\s*\(\s*"(?P<sheet>[^"]+)"\s*\)\s*\.\s*Cells\s*\(\s*(?P<row>\d+)\s*,\s*(?P<col>\d+|"[A-Z]+")\s*\)',
    re.IGNORECASE,
)
```

### Named Range Reference

當 procedure 內出現 identifier `X`,且 X 是已知 named range:

```python
def detect_named_range_refs(
    code: str,
    known_named_ranges: set[str],
) -> list[str]:
    found = []
    for nr_name in known_named_ranges:
        # 用 \b 確保完整字
        pattern = rf'\b{re.escape(nr_name)}\b'
        if re.search(pattern, code):
            found.append(nr_name)
    return found
```

注意:這會 false-positive — 當變數名跟 named range 同名。可以接受,標記為 `via=named_range` 即可。

## 規則 3:Range 讀寫識別 — 動態 patterns (標記 has_dynamic_range)

```python
DYNAMIC_PATTERNS = [
    # Range("A" & i) — 字串串接
    re.compile(r'Range\s*\(\s*"[^"]*"\s*&', re.IGNORECASE),

    # Range(varName) — 變數
    re.compile(r'Range\s*\(\s*[A-Za-z_]\w*\s*\)', re.IGNORECASE),

    # Cells(i, j) — 變數
    re.compile(r'Cells\s*\(\s*[A-Za-z_]\w*', re.IGNORECASE),

    # Range(Cells(...), Cells(...))
    re.compile(r'Range\s*\(\s*Cells\s*\(', re.IGNORECASE),
]
```

當任一 pattern 命中,該 procedure 標記:
- `has_dynamic_range = true`
- `dynamic_range_notes` 加一條,包含 line number 與 code 片段 (前後各 30 字)

## 規則 4:讀 vs 寫 判定

VBA 不像 SQL 有明確 SELECT/UPDATE。判定基於:

### 寫 (write)

- `<range_expr> = ...` — `=` 在 range 表達式右側
- `<range_expr>.Value = ...`
- `<range_expr>.Formula = ...`
- `<range_expr>.FormulaR1C1 = ...`
- `<range_expr>.ClearContents`
- `<range_expr>.Clear`
- `<range_expr>.Delete`
- `<range_expr>.PasteSpecial ...`
- `Source.Copy <range_expr>` (Copy 的 destination 是 write)
- `Source.Cut <range_expr>`

### 讀 (read)

- 出現在 expression 中 (= 右側、傳入函式、條件判斷)
- `<range_expr>.Value` (讀屬性)
- `<range_expr>.Formula`
- `Source = <range_expr>` 或 `Source = <range_expr>.Value`
- 配合 `Set` 的 alias 賦值

### 不算讀寫 (僅形式存在)

- `Dim r As Range`
- `<range_expr>.Select`
- `<range_expr>.Activate`

## 規則 5:Variable Alias 追蹤

範圍:**procedure 級** — 不跨 procedure。

```vba
Sub Example()
    Dim rng As Range
    Set rng = Sheets("Output").Range("B2:Z100")  ← 建立 alias

    rng.Value = 1                                 ← 視為 write Sheets("Output")!B2:Z100
    Debug.Print rng.Cells(1,1).Value              ← 視為 read 同一個 range
End Sub
```

### 演算法

```python
def track_aliases(proc_code: str) -> dict[str, RangeRef]:
    """掃描 procedure source,建立 alias 表。"""
    aliases = {}

    # 找所有 Set var = ...range_expr...
    set_pattern = re.compile(
        r'^[ \t]*Set\s+(?P<var>[A-Za-z_]\w*)\s*=\s*(?P<expr>.+?)$',
        re.IGNORECASE | re.MULTILINE,
    )

    for m in set_pattern.finditer(proc_code):
        var = m.group("var")
        expr = m.group("expr")
        # 嘗試從 expr 抽出靜態可解的 range
        ref = try_extract_static_range(expr)
        if ref is not None:
            aliases[var] = ref
        # 解不出來 → 不建 alias,後續用該變數的操作會落入 dynamic

    return aliases
```

### 已知限制

- alias 跨 procedure (透過參數) — 不追,標記 dynamic
- alias 重新賦值 (`Set rng = ...; rng = ...; Set rng = OtherRange`) — 簡化為「最後一次 Set 為準」
- alias 經過運算 (`Set rng2 = rng.Offset(1,0)`) — 標記 dynamic

## 規則 6:Event Trigger 偵測

### 標準事件名

只在 `module_type == "sheet"` 或 `"workbook"` 的模組才掃。

```python
SHEET_EVENTS = {
    "Worksheet_Change",
    "Worksheet_SelectionChange",
    "Worksheet_BeforeDoubleClick",
    "Worksheet_BeforeRightClick",
    "Worksheet_Activate",
    "Worksheet_Deactivate",
    "Worksheet_Calculate",
    "Worksheet_FollowHyperlink",
    "Worksheet_PivotTableUpdate",
}

WORKBOOK_EVENTS = {
    "Workbook_Open",
    "Workbook_BeforeClose",
    "Workbook_BeforeSave",
    "Workbook_AfterSave",
    "Workbook_BeforePrint",
    "Workbook_NewSheet",
    "Workbook_SheetActivate",
    "Workbook_SheetChange",
    "Workbook_SheetSelectionChange",
}
```

### Target 範圍偵測 (heuristic)

`Worksheet_Change(ByVal Target As Range)` 的 procedure body 通常開頭會有 Intersect 過濾:

```vba
Private Sub Worksheet_Change(ByVal Target As Range)
    If Intersect(Target, Range("A:A")) Is Nothing Then Exit Sub
    ...
End Sub
```

抓 pattern:

```python
INTERSECT_PATTERN = re.compile(
    r'Intersect\s*\(\s*Target\s*,\s*Range\s*\(\s*"(?P<target>[^"]+)"\s*\)',
    re.IGNORECASE,
)
```

抓不到就 target 留空 (代表「整個 sheet 任一 cell 變動都觸發」)。

## 規則 7:Procedure Call Graph

### 兩階段

1. **Pass 1**:收集所有 procedure name (跨模組)
2. **Pass 2**:對每個 procedure body,找出對其他 procedure name 的引用

### Pass 2 Pattern

```python
def find_calls(body: str, known_procs: set[str], builtins: set[str]) -> list[str]:
    calls = []
    for proc_name in known_procs:
        # 用 \b 確保完整字、不要在字串/註解內
        pattern = rf'\b{re.escape(proc_name)}\b'
        # 簡單清理:移除註解 (' ... 到行尾) 和字串
        cleaned_body = remove_comments_and_strings(body)
        if re.search(pattern, cleaned_body):
            calls.append(proc_name)
    return [c for c in calls if c not in builtins]
```

### Builtin Whitelist

不該被當成 call 的:

```python
VBA_BUILTINS = {
    # IO
    "MsgBox", "InputBox", "Debug", "Print",
    # Range/Sheet object methods
    "Range", "Cells", "Rows", "Columns", "Worksheets", "Sheets", "Workbooks",
    "Application", "ActiveSheet", "ActiveCell", "ActiveWorkbook", "Selection",
    # 字串/型別轉換
    "CStr", "CInt", "CLng", "CDbl", "CBool", "CDate", "CDec",
    "Format", "Trim", "LTrim", "RTrim",
    "Left", "Right", "Mid", "Len", "InStr", "InStrRev",
    "UCase", "LCase", "Replace", "Split", "Join",
    # 數學
    "Abs", "Int", "Fix", "Round", "Sgn", "Sqr",
    # 集合/陣列
    "Array", "UBound", "LBound", "IsArray",
    # 型別判定
    "IsEmpty", "IsNull", "IsNumeric", "IsDate", "IsObject", "IsError", "IsMissing",
    # 流程
    "Exit", "Stop", "End", "GoTo", "On",
    # Sub-objects
    "Set", "Let", "Get", "Property",
    # 日期
    "Now", "Date", "Time", "Year", "Month", "Day",
    "Hour", "Minute", "Second", "DateAdd", "DateDiff", "DateSerial",
}
```

當有疑問時加進 whitelist 就好,寧願少抓也不要 false-positive。

## 規則 8:Complexity Score

```python
def compute_procedure_complexity(proc) -> int:
    return (
        proc.line_count // 10                        # 每 10 行 +1
        + (proc.read_count + proc.write_count)       # 每筆 r/w +1
        + (10 if proc.has_dynamic_range else 0)      # 動態 range 重罰
        + (5 if proc.has_event_trigger else 0)       # event trigger 加分
        + len(proc.calls) * 2                        # 每 call +2
        + count_branches(proc.source_code)           # 每 If/Select Case +1
        + count_loops(proc.source_code) * 2          # 每 For/Do/While +2
    )
```

## 規則 9:已知失敗模式

寫進 `dynamic_range_notes` 或 `00_summary.warnings` 的情境:

| 情境 | 處理 |
|---|---|
| Encrypted vbaProject | warning + skip 整個 module |
| Procedure 巢狀 (理論不該存在) | warning,outer procedure 視為未閉合 |
| 字串內含 procedure name (false-positive call) | 透過 remove strings 預處理避免 |
| `Application.Run "macro_name"` | 列入 calls,但 via 標 `dynamic_invocation` |
| Form events (Click 等) | 暫不分類,標 `procedure_type=sub`,不加進 triggers |
| `On Error Resume Next` 改變 control flow | 不影響靜態分析,但加 warning 提示重構需注意 |
