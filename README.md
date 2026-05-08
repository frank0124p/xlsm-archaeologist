# xlsm-archaeologist

把複雜的 `.xlsm`（帶 VBA 巨集的 Excel）完整考古成結構化的 JSON/CSV，包含工作表、公式 AST、VBA 模組/程序、儲存格依賴圖與移植難度評分。輸出可直接餵給後續規則引擎設計、LLM 分析，或搭配 [Form Assembly Service](https://github.com/frank0124p/form-assembly) 自動生成互動式表單。

[![Python](https://img.shields.io/badge/python-3.12%20%7C%203.13%20%7C%203.14-blue)](https://www.python.org)
[![License: MIT](https://img.shields.io/badge/license-MIT-green)](LICENSE)

---

## 安裝

### 前置需求

- **Python 3.12 / 3.13 / 3.14**（不支援 3.11 以下）

確認版本：

```bash
python3 --version
# 應顯示 Python 3.12.x 以上
```

若版本不符，至 https://www.python.org/downloads/ 下載安裝。

---

### 方法一：pip 直接從 GitHub 安裝（推薦）

```bash
pip install git+https://github.com/frank0124p/xlsm-archaeologist.git
```

若系統有多個 Python 版本，請明確指定：

```bash
python3.12 -m pip install git+https://github.com/frank0124p/xlsm-archaeologist.git
```

確認安裝成功：

```bash
xlsm-archaeologist version
```

預期輸出：
```
xlsm-archaeologist 0.1.1
schema_version: 1.0
python: 3.12.x
openpyxl: 3.1.x
oletools: 0.60.x
```

> **Windows 用戶注意**：若出現 `not recognized as the name of a cmdlet` 錯誤，請見下方 [Windows PATH 設定](#windows-path-設定)。

---

### 方法二：下載 wheel 手動安裝

至 [Releases 頁面](https://github.com/frank0124p/xlsm-archaeologist/releases/latest) 下載 `xlsm_archaeologist-*.whl`：

```bash
pip install xlsm_archaeologist-0.1.1-py3-none-any.whl
```

---

### 方法三：原始碼安裝（開發 / 修改用）

```bash
git clone https://github.com/frank0124p/xlsm-archaeologist.git
cd xlsm-archaeologist
pip install -e ".[dev]"
```

---

## Windows PATH 設定

安裝後在 Windows 出現 `not recognized as the name of a cmdlet` 時，代表 pip 的 Scripts 資料夾尚未加入 PATH。

**方法一：加入系統 PATH（一勞永逸）**

在 PowerShell 執行，找出 Scripts 路徑：

```powershell
python -c "import sys, os; print(os.path.join(os.path.dirname(sys.executable), 'Scripts'))"
```

輸出類似：
```
C:\Users\你的名字\AppData\Local\Programs\Python\Python312\Scripts
```

把這個路徑加入系統 PATH：
1. 搜尋「編輯系統環境變數」
2. 點「環境變數」→ 選取「Path」→「編輯」→「新增」
3. 貼上上面的路徑 → 確定
4. 重新開啟 PowerShell

之後即可直接使用 `xlsm-archaeologist` 指令。

**方法二：用 `python -m` 執行（不需改 PATH）**

```powershell
python -m xlsm_archaeologist analyze 你的檔案.xlsm
python -m xlsm_archaeologist version
python -m xlsm_archaeologist inspect 你的檔案.xlsm
```

---

## 快速開始

安裝完成後，直接用 `xlsm-archaeologist` 指令：

```bash
# 分析一個 .xlsm 檔案（輸出到 ./archaeology_output/）
xlsm-archaeologist analyze MyFile.xlsm

# 指定輸出目錄
xlsm-archaeologist analyze MyFile.xlsm --output ./my_output

# 快速預覽基本資訊（不寫檔）
xlsm-archaeologist inspect MyFile.xlsm
```

---

## 指令參考

### `analyze` — 完整分析

```
xlsm-archaeologist analyze [OPTIONS] INPUT_PATH
```

| 選項 | 說明 | 預設 |
|---|---|---|
| `-o, --output PATH` | 輸出目錄 | `archaeology_output` |
| `--no-vba` | 跳過 VBA 分析 | — |
| `--no-graph` | 跳過依賴圖 | — |
| `--no-reports` | 跳過報告產生 | — |
| `--phases TEXT` | 只跑指定 phase（逗號分隔），或 `all` | `all` |
| `--max-formula-depth INT` | 公式 AST 最大巢狀深度 | `20` |
| `--log-level TEXT` | `debug` / `info` / `warning` / `error` | `info` |
| `-q, --quiet` | 隱藏 progress bar | — |
| `--force` | 覆蓋已有輸出目錄 | — |

**常用範例：**

```bash
# 完整分析
xlsm-archaeologist analyze Invoice.xlsm -o ./out

# 只分析結構（最快，跳過 VBA 與依賴圖）
xlsm-archaeologist analyze Invoice.xlsm --no-vba --no-graph --no-reports

# 強制覆蓋已有輸出目錄
xlsm-archaeologist analyze Invoice.xlsm --force

# 詳細 debug 模式
xlsm-archaeologist analyze Invoice.xlsm --log-level debug
```

### `inspect` — 快速預覽

顯示 workbook 基本資訊，**不寫入任何檔案**：

```bash
xlsm-archaeologist inspect MyFile.xlsm
```

### `version` — 版本資訊

```bash
xlsm-archaeologist version
```

---

## 輸出檔案說明

分析完成後，輸出目錄會產生以下檔案：

```
archaeology_output/
├── 00_summary.json          ← 總覽：stats、風險指標、複雜度、移植難度
├── 01_workbook.json         ← Workbook 元資料（SHA256、VBA 有無、外部連結）
├── 02_sheets.csv            ← 每個工作表的統計（row/col/formula 數量）
├── 03_named_ranges.csv      ← 所有命名範圍（動態公式偵測、作用域）
├── 04_cells.csv             ← 所有「有意義」儲存格（含公式/驗證/命名的格）
├── 05_formulas.json         ← 每條公式的 AST、分類、複雜度、引用清單
├── 06_validations.csv       ← Data Validation 規則（含下拉選單列舉值）
├── 07_vba_modules.json      ← VBA 模組（類型、行數、原始碼）
├── 08_vba_procedures.json   ← VBA 程序（讀寫 cell 行為、事件觸發、call graph）
├── 09_dependencies.csv      ← 依賴邊（source → target，標記跨工作表）
├── 10_dependency_graph.json ← 完整有向圖（node-link 格式，可用 networkx 載入）
└── reports/
    ├── cycles.json               ← 循環引用偵測結果
    ├── formula_categories.csv    ← 公式分類統計（計數、平均複雜度）
    ├── top_complex_formulas.csv  ← 複雜度排行榜 Top 50
    ├── hotspot_cells.csv         ← 被最多其他儲存格引用的格（hotspot 排行）
    ├── vba_behavior.csv          ← 每個 VBA 程序的讀寫行為摘要
    └── cross_sheet_refs.csv      ← 所有跨工作表依賴邊
```

### `00_summary.json` 欄位說明

```jsonc
{
  "schema_version": "1.0",
  "tool_version": "0.1.0",
  "analyzed_at": "2026-05-08T00:00:00+00:00",
  "input_file": { "path": "/path/to/MyFile.xlsm", "sha256": "...", "size_bytes": 102400 },

  "stats": {
    "sheet_count": 5,
    "named_range_count": 12,
    "formula_count": 834,
    "validation_count": 23,
    "vba_module_count": 3,
    "vba_procedure_count": 17,
    "dependency_edge_count": 1204
  },

  "risk_indicators": {
    "circular_reference_count": 0,
    "external_reference_count": 2,
    "volatile_function_count": 8,
    "dynamic_vba_range_count": 3,
    "deeply_nested_formula_count": 5,
    "orphan_formula_count": 12,
    "cross_sheet_dependency_count": 67
  },

  "complexity_score": 312,
  "migration_difficulty": "high",   // low / medium / high / very_high

  "warnings": [
    {
      "level": "warning",
      "category": "formula",
      "location": "Sheet1!B7",
      "message": "Parse failed: '=LAMBDA(x,x+1)'"
    }
  ]
}
```

**移植難度分級：**

| 複雜度分數 | 難度 |
|---|---|
| 0 – 49 | `low` |
| 50 – 199 | `medium` |
| 200 – 499 | `high` |
| ≥ 500 | `very_high` |

---

## 公式分類

每條公式自動分類到七種類別之一（`formula_category` 欄位）：

| 類別 | 代表函式 |
|---|---|
| `lookup` | VLOOKUP、HLOOKUP、XLOOKUP、INDEX/MATCH |
| `branch` | IF、IFS、IFERROR、AND/OR/NOT |
| `aggregate` | SUM、SUMIF、COUNT、AVERAGE、MAX/MIN |
| `text` | CONCAT、LEFT/RIGHT/MID、TRIM、TEXT |
| `compute` | 四則運算、數學/日期函式 |
| `reference` | 純儲存格引用（`=A1` 或 `=Sheet2!A1`） |
| `mixed` | 同時含兩種以上類別 |

---

## 分析管線

```
.xlsm 檔案
    │
    ├─ Phase 2: 結構抽取
    │   ├─ workbook_extractor    → 01_workbook.json
    │   ├─ sheet_extractor       → 02_sheets.csv
    │   ├─ named_range_extractor → 03_named_ranges.csv
    │   ├─ cell_extractor        → 04_cells.csv（初版）
    │   └─ validation_extractor  → 06_validations.csv
    │
    ├─ Phase 3: 公式分析
    │   └─ formula_analyzer      → 05_formulas.json
    │       ├─ tokenizer
    │       ├─ recursive-descent parser（AST）
    │       ├─ classifier（7 類）
    │       └─ complexity scorer
    │
    ├─ Phase 4: VBA 分析（oletools）
    │   ├─ vba_extractor         → 07_vba_modules.json
    │   ├─ procedure_splitter    → 08_vba_procedures.json
    │   ├─ range_detector（標記動態 range）
    │   └─ call_graph_extractor
    │
    ├─ Phase 5: 依賴圖（networkx）
    │   ├─ graph_builder         → 09_dependencies.csv
    │   ├─ cycle_detector        → reports/cycles.json
    │   ├─ orphan_detector
    │   └─ 回填 is_referenced    → 04_cells.csv（最終版）
    │                              10_dependency_graph.json
    │
    └─ Phase 6: 報告與評分
        ├─ complexity scorer
        ├─ migration difficulty → 00_summary.json
        └─ 5 張報告 CSV/JSON   → reports/
```

---

## 專案結構

```
src/xlsm_archaeologist/
├── cli.py                     # CLI 入口（typer）
├── runner.py                  # 主流程協調者
├── config.py                  # pydantic-settings 全域設定
├── errors.py                  # 自訂例外
├── models/                    # Pydantic v2 資料模型
├── extractors/                # Layer 1：從 openpyxl 原始抽取
├── analyzers/                 # Layer 2：語意分析
├── reports/                   # 報告產生器
└── serializers/               # JSON / CSV 輸出
```

---

## 設計原則

| 原則 | 說明 |
|---|---|
| **純靜態分析** | 不執行任何 VBA 或公式，只讀取原始 XML |
| **唯讀** | 永遠不修改原始 `.xlsm` |
| **決定性輸出** | 同一檔案跑兩次輸出完全相同（排序穩定） |
| **容錯** | 單一 cell/公式/程序失敗不中斷整個 run，錯誤收進 warnings |
| **動態 VBA range 標記** | 無法靜態解析的 range 標記 `has_dynamic_range: true` |
| **UTF-8-with-BOM** | CSV 供 Excel 直接開啟不亂碼 |

---

## 已知限制

- **動態 VBA range**：`Range("A" & lastRow)` 無法解析確切位址，標記 `has_dynamic_range: true` 並列進 warnings
- **LAMBDA / LET**：新式自訂函式標記 `is_parsable: false`
- **外部 workbook 連結**：偵測 `[Book.xlsx]Sheet!A1` 模式，但不追入外部檔案
- **加密 VBA**：oletools 無法解密，回傳空模組並記錄 warning

---

## 開發

```bash
git clone https://github.com/frank0124p/xlsm-archaeologist.git
cd xlsm-archaeologist

# 安裝含開發工具
pip install -e ".[dev]"

# 執行測試
pytest
pytest --cov                  # 含覆蓋率報告

# Lint & 型別檢查
ruff check src/
mypy src/
```

---

## 搭配 Form Assembly Service

分析完成後，把輸出目錄路徑貼入 [Form Assembly](https://github.com/frank0124p/form-assembly)，即可自動從 `06_validations.csv` 生成互動式表單，並搭配 `05_formulas.json` 計算欄位即時計算。

---

## License

MIT
