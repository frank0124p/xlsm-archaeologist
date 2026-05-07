# Phase 4: VBA Analysis

## 目標

抽出 .xlsm 中所有 VBA 模組的 source code,並對每個 Sub/Function 做:
1. 切分 procedure
2. 識別讀寫的 cell/range (含動態 range 標記)
3. 識別 procedure 之間的呼叫關係
4. 識別 event triggers (`Worksheet_Change` 等)
5. 計算 procedure-level 複雜度
6. 寫進 `07_vba_modules.json` 與 `08_vba_procedures.json`

## 為什麼 VBA 分析最複雜

- VBA 沒有官方 Python AST library,要自己用 token + regex 拼
- 動態 range 例如 `Range("A" & lastRow)` 無法靜態解析,只能標記
- 變數 alias (例如 `Set rng = Range("A1"); rng.Value = 1`) 要做簡單追蹤
- Event trigger 散落在 sheet/workbook module 裡,要特別處理

## 模組

```
src/xlsm_archaeologist/
├── extractors/
│   └── vba_extractor.py         # 用 oletools 抽 source code
├── analyzers/
│   ├── vba_analyzer.py          # 主協調者
│   ├── vba_procedure_splitter.py # 切分 Sub/Function
│   ├── vba_range_detector.py    # 識別讀寫 cell/range
│   └── vba_call_graph.py        # procedure call graph
└── models/
    └── vba.py                   # VbaModuleRecord, VbaProcedureRecord
```

## VBA 抽取

```python
from oletools.olevba3 import VBA_Parser

vba = VBA_Parser(path)
if not vba.detect_vba_macros():
    return []  # no VBA

for (filename, stream_path, vba_filename, vba_code) in vba.extract_macros():
    # filename:  "complex_macro.xlsm"
    # vba_filename: "Module1" / "Sheet1" / "ThisWorkbook" / "Class1"
    # vba_code: 純文字原始碼
    yield VbaModuleRecord(...)
```

注意:
- 加密 VBA (`encrypted vbaProject.bin`) 抽不出來,設 warning + skip
- VBA forms 抽出來的是 `.frx` binary 加 form code,只取 form code 部分

## Procedure 切分

```python
PROCEDURE_PATTERN = re.compile(
    r"^(?P<scope>Public\s+|Private\s+)?"
    r"(?P<static>Static\s+)?"
    r"(?P<kind>Sub|Function|Property\s+(?:Get|Let|Set))\s+"
    r"(?P<name>\w+)\s*"
    r"(?:\((?P<params>[^)]*)\))?\s*"
    r"(?:\s+As\s+(?P<return_type>\w+))?",
    re.IGNORECASE | re.MULTILINE,
)

END_PATTERN = re.compile(r"^End\s+(Sub|Function|Property)", re.IGNORECASE | re.MULTILINE)
```

注意:
- VBA 對大小寫不敏感,regex 用 IGNORECASE
- 處理 `Property Get/Let/Set`
- 多行 `_` continuation 要先合併再切

## Range Detection 規則

詳細規則見 `reference/vba_analysis_rules.md`。摘要:

```python
# 靜態可解
RANGE_LITERAL = r'Range\("([^"]+)"\)'           # Range("A1")  / Range("A1:B10")
CELLS_LITERAL = r'Cells\((\d+),\s*(\d+)\)'      # Cells(1, 1)
SQUARE_BRACKET = r'\[([A-Z]+\d+(?::[A-Z]+\d+)?)\]'  # [A1] / [A1:B10]
SHEET_RANGE = r'(?:Sheets|Worksheets)\("([^"]+)"\)\.Range\("([^"]+)"\)'

# 動態 (標記 has_dynamic_range=true)
RANGE_DYNAMIC = r'Range\(\s*"[^"]*"\s*&\s*\w+'   # Range("A" & i)
CELLS_DYNAMIC = r'Cells\([^,)]*[a-zA-Z_]\w*'     # Cells(i, j) — 變數
```

讀 vs 寫的判定:
- 出現在 `=` 左邊 → write (Range("A1") = ...)
- 出現在 `=` 右邊或 expression 中 → read
- `.Value = ...` 也算 write
- `.Copy`, `.PasteSpecial`, `.ClearContents` 算 write
- `.Cells.Value`, `.Range.Formula` 算 read

## 變數 alias 追蹤 (簡單版)

```vba
Set rng = Sheets("Output").Range("B2:Z100")
rng.Value = ...
```

策略:
- scope 限定在 procedure 內
- 偵測 `Set <var> = ...Range...` 建 alias 表
- 後續 `<var>.<method>` 視為對 aliased range 的操作

無法追蹤的情況 (例如 alias 跨函式傳遞) → 標記 `has_dynamic_range=true`。

## Event Trigger 偵測

- Sheet module 內的 `Worksheet_Change(ByVal Target As Range)` → trigger event=`Worksheet_Change`,
  target 從 procedure 第一行 Intersect 邏輯抓 (heuristic)
- Workbook module 內的 `Workbook_Open` / `Workbook_BeforeSave` 等
- `Application.OnTime` 動態 trigger → 標記 `has_dynamic_trigger=true`

## Call Graph

對每個 procedure:
- 掃描 source code,找其他 procedure name 的識別字
- 純識別字 + 後接 `(` 或 SoL → 視為呼叫
- 內建函式 (`MsgBox`, `Range`, `Format`, ...) 維護 whitelist 排除
- 結果寫進 `calls: list[str]`

## 已知限制

- ⚠ 動態 range 不嘗試 evaluate,只標記
- ⚠ 不解析 ActiveX form 內的事件 (太罕見)
- ⚠ 不展開 `For Each` / `Do While` 的迴圈體 — 只記錄 procedure body 整體讀寫
- ⚠ 不解析跨檔案 `Application.Run("OtherFile.xlsm!Macro")`
- ⚠ 加密 VBA 直接 skip + warning

## 驗收

見 `acceptance.md`。
