# Test Fixtures

> 手寫的 mini .xlsm 測試檔案說明。每份檔案對應一個明確場景。

## Fixture 清單

| 檔名 | 用於哪個 phase | 主要場景 |
|---|---|---|
| `simple.xlsm` | 2, e2e | 純值,無公式無 VBA |
| `formulas_basic.xlsm` | 2, 3 | 各類公式各一條 |
| `formulas_complex.xlsm` | 3, 5, 6, e2e | 巢狀 IF + 跨 sheet + named range |
| `vba_basic.xlsm` | 4, 5 | 基本讀寫 cell |
| `vba_dynamic_range.xlsm` | 4 | 動態 range 標記測試 |
| `vba_event_trigger.xlsm` | 4 | Worksheet_Change 偵測 |
| `vba_call_graph.xlsm` | 4 | procedure 互呼叫 |
| `circular.xlsm` | 5 | 循環引用偵測 |
| `orphan_formula.xlsm` | 5 | 孤島偵測 |
| `cross_sheet_chain.xlsm` | 5 | 跨 sheet 鏈式依賴 |
| `with_validation.xlsm` | 2 | list validation 兩種來源 |
| `with_named_range.xlsm` | 2, 3 | 一般 + 動態 named range |
| `hidden_sheets.xlsm` | 2 | hidden 與 very_hidden |

## 怎麼產出這些 fixture

**Fixture 不能用 git 直接 commit binary** — 雖然 .xlsm 是 binary 但相對小,
我們 commit 它們進 repo 沒問題,但**也要附產生腳本**讓檔案可重新生成。

放在 `scripts/build_fixtures.py`,跑法:

```bash
uv run python scripts/build_fixtures.py
```

腳本會把所有 fixture 重新產出到 `tests/fixtures/`。

## 不含 VBA 的 fixture

可用純 openpyxl 產生:

```python
# scripts/build_fixtures.py 範例片段

from openpyxl import Workbook
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from pathlib import Path

OUT = Path(__file__).parent.parent / "tests" / "fixtures"


def build_simple():
    wb = Workbook()
    ws = wb.active
    ws.title = "Data"
    for r in range(1, 4):
        for c in range(1, 4):
            ws.cell(row=r, column=c, value=r * c)
    wb.save(OUT / "simple.xlsm")  # 即使無 VBA 也用 .xlsm 副檔名


def build_formulas_basic():
    wb = Workbook()
    ws = wb.active
    ws.title = "Calc"

    # 各類公式
    ws["A1"] = 10
    ws["B1"] = 20
    ws["A2"] = "=VLOOKUP(A1,Params!A:B,2,0)"  # lookup
    ws["B2"] = "=IF(A1>0, \"正\", \"負\")"      # branch
    ws["C2"] = "=A1*B1+1"                       # compute
    ws["D2"] = "=SUM(A1:B1)"                    # aggregate
    ws["E2"] = "=A1 & \"-\" & B1"               # text
    ws["F2"] = "=Calc!A1"                       # reference (self-sheet)

    # 第二個 sheet
    ws2 = wb.create_sheet("Params")
    ws2["A1"] = 10
    ws2["B1"] = "十"

    wb.save(OUT / "formulas_basic.xlsm")


def build_with_named_range():
    wb = Workbook()
    ws = wb.active
    ws.title = "Data"
    ws["A1"] = "TaxRate"
    ws["B1"] = 0.05
    ws["A2"] = "DynamicRange"
    ws["B2"] = 100

    # 一般 named range
    wb.defined_names["TaxRate"] = DefinedName(
        name="TaxRate", attr_text="Data!$B$1"
    )

    # 動態 named range (含 OFFSET)
    wb.defined_names["DynamicArea"] = DefinedName(
        name="DynamicArea", attr_text="OFFSET(Data!$B$2,0,0,COUNT(Data!$B:$B),1)"
    )

    wb.save(OUT / "with_named_range.xlsm")


def build_hidden_sheets():
    wb = Workbook()
    visible = wb.active
    visible.title = "Visible"
    visible["A1"] = "I am visible"

    hidden = wb.create_sheet("Hidden")
    hidden.sheet_state = "hidden"
    hidden["A1"] = "I am hidden"

    very_hidden = wb.create_sheet("VeryHidden")
    very_hidden.sheet_state = "veryHidden"
    very_hidden["A1"] = "Very secret"

    wb.save(OUT / "hidden_sheets.xlsm")


def build_circular():
    wb = Workbook()
    ws = wb.active
    ws.title = "Loop"
    ws["A1"] = "=B1+1"
    ws["B1"] = "=A1+1"
    wb.save(OUT / "circular.xlsm")


def build_cross_sheet_chain():
    wb = Workbook()
    s1 = wb.active
    s1.title = "S1"
    s1["A1"] = 100

    s2 = wb.create_sheet("S2")
    s2["B1"] = "=S1!A1*2"

    s3 = wb.create_sheet("S3")
    s3["C1"] = "=S2!B1+1"

    wb.save(OUT / "cross_sheet_chain.xlsm")


def build_orphan_formula():
    wb = Workbook()
    ws = wb.active
    ws.title = "Orphan"
    ws["A1"] = 10
    ws["B1"] = 20
    ws["C1"] = "=A1+B1"  # 沒人引用的公式 = orphan
    wb.save(OUT / "orphan_formula.xlsm")


# ... 等等
```

完整 build_fixtures.py 由 Phase 2 跟 Phase 3 對應的 task 階段陸續補充。

## 含 VBA 的 fixture

openpyxl 不能寫 VBA。三種做法:

### 做法 A:base 檔 + 修改 (推薦)

1. 用 Excel 手動建立一份「最小 .xlsm」(空 module),命名 `_base_with_vba.xlsm`
2. base 檔裡用 olevba 工具直接 inject 你要的 VBA code
3. 用 openpyxl 載入 base 檔 + 補資料 + 用 `keep_vba=True` 存回 fixture

### 做法 B:預先手寫,直接 commit

最直接 — 這些 fixture 不變動,人工建好後 commit 到 repo。
缺點:無法靠腳本重建。

### 做法 C:用 `python-pptx` 風格的 mini library

實際上不存在現成的「Python 寫 VBA」工具。略過。

## 推薦做法

對於 v0.1.0:**做法 B**。手動建好 6 份 VBA fixture,commit。
因為 fixture 內容簡單 (10-30 行 VBA),用 Excel/LibreOffice 5 分鐘搞定。

之後若需要更多 VBA fixture,再考慮做法 A。

## VBA Fixture 的 VBA 內容

每份 fixture 的 VBA 應該這樣寫 (以下是參考範本,實際 build 時要逐字照抄到 Excel):

### vba_basic.xlsm — Module1

```vba
Public Sub UpdateB1()
    Range("B1").Value = Range("A1").Value * 2
End Sub
```

預期偵測:
- 1 module、1 procedure
- reads: Sheet1!A1
- writes: Sheet1!B1
- has_dynamic_range: false

### vba_dynamic_range.xlsm — Module1

```vba
Public Sub FillToLastRow()
    Dim lastRow As Long
    lastRow = Range("A1").End(xlDown).Row
    Range("B" & lastRow).Value = 999
    
    Dim i As Long
    For i = 1 To lastRow
        Cells(i, 3).Value = i
    Next i
End Sub
```

預期偵測:
- has_dynamic_range: true
- dynamic_range_notes 至少 2 條 (Range("B" & ...) 與 Cells(i, 3))

### vba_event_trigger.xlsm — Sheet1 module

```vba
Private Sub Worksheet_Change(ByVal Target As Range)
    If Intersect(Target, Range("A:A")) Is Nothing Then Exit Sub
    Application.EnableEvents = False
    Target.Offset(0, 1).Value = "Modified"
    Application.EnableEvents = True
End Sub
```

預期偵測:
- triggers: [{event: Worksheet_Change, target: Sheet1!A:A}]

### vba_call_graph.xlsm — Module1

```vba
Public Sub Main()
    Call SubA
    Call SubB
End Sub

Private Sub SubA()
    Call SubC
    MsgBox "A done"
End Sub

Private Sub SubB()
    MsgBox "B done"
End Sub

Private Sub SubC()
    Range("A1").Value = "from C"
End Sub
```

預期偵測:
- Main.calls = ["SubA", "SubB"]
- SubA.calls = ["SubC"]  (MsgBox 不該被收進來)
- SubB.calls = []
- SubC.calls = []

## Fixture 維護

- 改了 fixture 後 git commit 帶說明
- 任何測試壞了不要改 fixture 適配 — 應該檢查邏輯是否錯
- fixture 不是「合約」,但有改動要在 PR description 說明
