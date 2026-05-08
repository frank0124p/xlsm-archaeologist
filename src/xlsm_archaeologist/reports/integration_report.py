"""Integration guide — downstream system connection for developers."""

from __future__ import annotations

from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from xlsm_archaeologist.models.cell import ValidationRecord
    from xlsm_archaeologist.models.formula import FormulaRecord
    from xlsm_archaeologist.models.summary import SummaryRecord
    from xlsm_archaeologist.models.workbook import SheetRecord


def build_integration_md(
    source_file: str,
    summary: SummaryRecord,
    sheets: list[SheetRecord],
    validations: list[ValidationRecord],
    formulas: list[FormulaRecord],
    output_dir: str,
) -> str:
    """Generate a downstream integration guide for developers.

    Returns Markdown string for reports/integration.md.
    """
    # Top formula functions across all formulas
    func_counts: dict[str, int] = {}
    for f in formulas:
        for fn in f.function_list:
            func_counts[fn] = func_counts.get(fn, 0) + 1
    top_funcs = sorted(func_counts.items(), key=lambda x: -x[1])[:10]
    top_funcs_md = "\n".join(f"  - `{fn}` — {n} 次" for fn, n in top_funcs) if top_funcs else "  （無）"  # noqa: E501

    # Difficulty note
    difficulty = summary.migration_difficulty
    difficulty_note = {
        "low": "✅ 低複雜度，可直接對應規則或公式遷移",
        "medium": "⚠️ 中等複雜度，建議逐一驗證核心計算邏輯",
        "high": "🔴 高複雜度，建議分階段遷移並保留 Excel 作為對照",
        "very_high": "🔴🔴 極高複雜度，需人工深度審查才能安全遷移",
    }.get(difficulty, "")

    return f"""# 下游系統串接指引 — {source_file}

> 由 xlsm-archaeologist 自動生成。說明如何將本工具的輸出接入下游系統。

---

## 總覽

| 項目 | 值 |
|---|---|
| 移植難度 | **{difficulty}** {difficulty_note} |
| 複雜度分數 | {summary.complexity_score} |
| 工作表數 | {summary.stats.sheet_count} |
| 公式數 | {summary.stats.formula_count} |
| 資料驗證數 | {summary.stats.validation_count} |
| VBA 模組數 | {summary.stats.vba_module_count} |
| 依賴邊數 | {summary.stats.dependency_edge_count} |

---

## 輸出檔案對應表

| 檔案 | 用途 | 建議接入方式 |
|---|---|---|
| `00_summary.json` | 整體品質與風險評估 | CI/CD 品質門檻判斷、Dashboard 顯示 |
| `01_workbook.json` | Workbook 元資料 | 存入 DB 的 workbook 主表 |
| `02_sheets.csv` | 工作表清單 | DB sheets 表，或 API schema 生成 |
| `03_named_ranges.csv` | 命名範圍 | 對應規則引擎的常數/變數定義 |
| `04_cells.csv` | 有意義的儲存格 | 欄位清單、ETL 欄位對應 |
| `05_formulas.json` | 公式 AST + 分類 | 規則引擎轉換、LLM 語意分析輸入 |
| `06_validations.csv` | 資料驗證規則 | 表單系統欄位定義（見下方） |
| `07_vba_modules.json` | VBA 原始碼 | 人工審查、LLM 翻譯成目標語言 |
| `08_vba_procedures.json` | VBA 程序讀寫行為 | 業務邏輯對應、API 端點規劃 |
| `09_dependencies.csv` | 依賴邊清單 | 計算順序排序、DAG 執行引擎輸入 |
| `10_dependency_graph.json` | 完整有向圖 | networkx / 圖資料庫（Neo4j 等）匯入 |
| `reports/architecture.md` | 工作表架構圖 | 技術文件、系統設計討論 |
| `reports/data_flow.md` | 操作說明文件 | 開發者交接、需求分析 |

---

## 串接場景指引

### 1. 表單系統（Form Assembly）

`06_validations.csv` 直接對應表單欄位定義：

```python
import csv

with open("archaeology_output/06_validations.csv", encoding="utf-8-sig") as f:
    reader = csv.DictReader(f)
    for row in reader:
        field = {{
            "address": row["qualified_address"],
            "type":    row["validation_type"],   # list / whole / decimal / date
            "options": row["enum_values"].split("|") if row["enum_values"] else [],
            "min":     row["formula1"] or None,
            "max":     row["formula2"] or None,
            "required": row["allow_blank"] == "false",
        }}
```

搭配 [Form Assembly Service](https://github.com/frank0124p/form-assembly) 可自動從此 CSV 生成互動式表單。

---

### 2. 資料庫匯入（DB）

建議 DB schema：

```sql
CREATE TABLE workbook_sheets (
    sheet_name    VARCHAR(128) PRIMARY KEY,
    row_count     INT,
    col_count     INT,
    formula_count INT,
    is_hidden     BOOLEAN
);

CREATE TABLE workbook_formulas (
    qualified_address VARCHAR(256) PRIMARY KEY,
    formula_text      TEXT,
    category          VARCHAR(32),
    complexity_score  INT,
    nesting_depth     INT
);

CREATE TABLE workbook_validations (
    qualified_address VARCHAR(256) PRIMARY KEY,
    validation_type   VARCHAR(32),
    enum_values       TEXT,
    formula1          VARCHAR(256),
    formula2          VARCHAR(256),
    allow_blank       BOOLEAN
);
```

讀取範例（Python）：

```python
import json, csv

# 公式
with open("archaeology_output/05_formulas.json", encoding="utf-8") as f:
    formulas = json.load(f)["formulas"]

# 驗證規則
with open("archaeology_output/06_validations.csv", encoding="utf-8-sig") as f:
    validations = list(csv.DictReader(f))
```

---

### 3. 規則引擎 / 業務邏輯遷移

`05_formulas.json` 包含每條公式的 AST，可用於自動轉換成目標語言：

```python
with open("archaeology_output/05_formulas.json", encoding="utf-8") as f:
    data = json.load(f)

# 篩選高複雜度公式（需人工審查）
complex_formulas = [
    f for f in data["formulas"]
    if f["complexity_score"] > 10 or not f["is_parsable"]
]

# 依分類分組
from collections import defaultdict
by_category = defaultdict(list)
for f in data["formulas"]:
    by_category[f["formula_category"]].append(f)
```

---

### 4. 依賴圖分析（計算順序 / 衝擊分析）

`10_dependency_graph.json` 為 networkx node-link 格式，可直接載入：

```python
import json
import networkx as nx

with open("archaeology_output/10_dependency_graph.json", encoding="utf-8") as f:
    data = json.load(f)

G = nx.node_link_graph(data)

# 拓樸排序（計算執行順序）
order = list(nx.topological_sort(G))

# 衝擊分析：如果 Sheet1!A1 改變，哪些 cell 會受影響？
affected = nx.descendants(G, "Sheet1!A1")
```

---

### 5. LLM 輔助分析

將關鍵輸出餵給 LLM 做語意解讀：

```python
import json

with open("archaeology_output/00_summary.json", encoding="utf-8") as f:
    summary = json.load(f)

with open("archaeology_output/05_formulas.json", encoding="utf-8") as f:
    formulas = json.load(f)["formulas"][:50]  # 取前 50 條

# 建議 prompt 結構：
# 1. summary 作為 system context
# 2. formulas 逐一說明用途
# 3. vba_procedures 翻譯成 pseudocode
```

---

## 已使用的函式清單

{top_funcs_md}

---

## 已知風險

{chr(10).join(f"- ⚠️ {w.category} @ `{w.location}`: {w.message}" for w in (summary.warnings or [])[:10]) or "無"}

---

*此文件由 `xlsm-archaeologist analyze` 自動產生，輸出目錄：`{output_dir}`*
"""
