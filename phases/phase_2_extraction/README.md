# Phase 2: Extraction

## 目標

用 openpyxl 把 .xlsm 的「結構性資料」全部抽出來,寫進 `01-04, 06` 號檔案。
**不做公式 AST 解析**(留 Phase 3),**不做 VBA**(留 Phase 4),**不做依賴圖**(留 Phase 5)。

## 範圍

### 抽什麼

- ✅ workbook metadata:檔案 sha256、大小、has_vba、是否含外部連結、author 等
- ✅ sheet 清單:名稱、index、隱藏狀態、used_range、row/col count、公式 cell 數
- ✅ named range 清單:含 `has_dynamic_formula` 偵測 (refers_to 含 OFFSET/INDIRECT)
- ✅ cell 清單:**只記「有意義」的 cell**,定義為下列任一:
    - 含公式
    - 有 validation
    - 被 named range 指到
    - (Phase 5 之後才能補:被其他 cell 或 VBA 引用 — 本 phase 暫填 false)
- ✅ validation:含下拉選單,解析出 enum_values

### 對應產出

| 檔案 | Phase 2 完成度 |
|---|---|
| `01_workbook.json` | 100% |
| `02_sheets.csv` | 100% |
| `03_named_ranges.csv` | 100% |
| `04_cells.csv` | 95% (`is_referenced` 欄位先給 false,Phase 5 回填) |
| `06_validations.csv` | 100% |

## 模組

```
src/xlsm_archaeologist/
├── extractors/
│   ├── workbook_extractor.py    # 主協調者
│   ├── sheet_extractor.py
│   ├── named_range_extractor.py
│   ├── cell_extractor.py
│   └── validation_extractor.py
├── models/
│   ├── workbook.py              # WorkbookRecord, SheetRecord
│   ├── cell.py                  # CellRecord, ValidationRecord
│   └── named_range.py           # NamedRangeRecord
└── serializers/
    ├── json_writer.py           # 寫 01_workbook.json
    └── csv_writer.py            # 寫 02-04, 06.csv
```

## 重要實作細節

### 開檔模式

```python
from openpyxl import load_workbook

wb = load_workbook(
    filename=path,
    read_only=False,       # named range 用 read-only 模式抓不到
    data_only=False,       # 我們要原始公式,不要 cached value
    keep_vba=True,         # 保留 VBA (Phase 4 用)
)
```

> ⚠ `read_only=True` 雖然快但拿不到 named range / validation,**不能用**。
> 大檔案的速度問題用 progress bar + 邊抽邊寫紓解。

### "有意義 cell" 篩選

不要把整個 used_range 都 dump,會撐爆 04_cells.csv。
**只記滿足下列任一的 cell**:

```python
def is_meaningful(cell) -> bool:
    return (
        cell.data_type == "f"            # 含公式
        or cell.coordinate in validation_addresses  # 有 validation
        or cell.coordinate in named_addresses        # 被 named range 指到
    )
```

`is_referenced` 欄位先填 `False`,Phase 5 建完依賴圖後回填正確值。

### Validation 解析

```python
for sheet in wb.worksheets:
    for dv in sheet.data_validations.dataValidation:
        # dv.sqref = "A2:A100 C5"  ← 多個 range 空格分隔
        # dv.type, dv.formula1, dv.formula2
        # 若 type == "list",formula1 可能是:
        #   - 字面值: '"A,B,C"' → split by comma
        #   - range 引用: '=Params!$A$2:$A$10' → 解析後讀內容
        ...
```

### Named Range 動態偵測

```python
DYNAMIC_FUNCS = {"OFFSET", "INDIRECT", "INDEX"}  # INDEX 嚴格說不一定 volatile
                                                    # 但 referenced from a name 通常是動態
def has_dynamic_formula(refers_to: str) -> bool:
    upper = refers_to.upper()
    return any(f"{f}(" in upper for f in DYNAMIC_FUNCS)
```

## 限制與已知 trade-off

- ⚠ openpyxl 對某些罕見 cell 屬性 (如 rich text、conditional format formula) 抽取不完整
   — Phase 2 不處理,在 warnings 列出
- ⚠ 巨大 .xlsx (> 50MB) 載入慢 — 接受,加 progress bar
- ⚠ `is_referenced` 在 Phase 2 全填 false,要在 Phase 5 補完

## 驗收

見 `acceptance.md`。
