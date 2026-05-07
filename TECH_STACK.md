# Tech Stack

> 套件選型、版本鎖定範圍、選型理由。新增/換套件前先 review 本檔案。

## 為什麼選 Python (不選 Node.js)

關鍵 blocker:

1. **VBA 抽取**:Python `oletools` 是業界標準,十幾年成熟。Node.js 沒有等價套件,要自己刻
   binary parser (CFBF + PerformanceCache 解壓),工作量翻 5 倍以上。
2. **公式 AST**:Python `formulas` 套件 + `openpyxl.formula.tokenizer` 內建。
   Node.js `xlsx (SheetJS)` 只給字串,要自己 tokenize。
3. **依賴圖**:Python `networkx` 是十年磨一劍。Node.js `graphology` 還在追趕。
4. **LLM 生態系**:後續若要做語意化分析,Python 套件 (LangChain、Anthropic SDK) 也更完整。

> 想跟 Node.js 環境整合的話,不會耦合在語言層 — 透過 CLI + JSON 輸出橋接即可。

## 環境

| 項目 | 版本 | 為什麼 |
|---|---|---|
| Python | **3.12.x** | 最新穩定,type hint 表達力強;不用 3.13 因為部分套件還在追兼容 |
| OS | macOS / Linux | Mac mini M4 跑沒問題;Windows 沒測 |
| 套件管理 | **uv** (≥ 0.4) | 比 Poetry 快 10x,Astral 主推 |
| Lint + Format | **ruff** (≥ 0.6) | 一個工具搞定,比 black + flake8 + isort 快 |
| Test | **pytest** (≥ 8.0) | 標配 |

## Runtime 套件 (whitelist)

```toml
# pyproject.toml [project] dependencies
[project]
dependencies = [
  "openpyxl >= 3.1, < 4.0",
  "oletools >= 0.60, < 1.0",
  "networkx >= 3.2, < 4.0",
  "typer >= 0.12, < 1.0",
  "pydantic >= 2.5, < 3.0",
  "pydantic-settings >= 2.1, < 3.0",
  "rich >= 13.7, < 14.0",
  "jinja2 >= 3.1, < 4.0",
]
```

### 用途說明

| 套件 | 用途 | 用在哪一層 |
|---|---|---|
| **openpyxl** | 讀 .xlsm 結構、cell、公式字串、named range、validation | Layer 1 Extractor |
| **oletools** | `olevba3` 抽 VBA source code | Layer 1 VBA Extractor |
| **networkx** | 依賴圖建構、拓樸排序、循環偵測、連通分量 | Layer 2 Dependency Analyzer |
| **typer** | CLI 介面 (型別友善、自動產 help) | CLI 進入點 |
| **pydantic** | 所有資料 model + JSON schema | Layer 1.5 Models |
| **pydantic-settings** | 全域設定 (從 env / config file 讀) | config.py |
| **rich** | progress bar、彩色 log、表格輸出 | utils/progress、CLI |
| **jinja2** | reports markdown 模板 (若有需要) | Layer 4 Reports |

## Dev 套件

```toml
[dependency-groups]
dev = [
  "pytest >= 8.0, < 9.0",
  "pytest-cov >= 5.0, < 6.0",
  "ruff >= 0.6, < 1.0",
  "mypy >= 1.8, < 2.0",
]
```

## 不允許的套件

- ❌ **pandas** — 過殺。Cell-level 分析直接用 pydantic 跟 csv 模組更輕。
- ❌ **xlsxwriter** — 我們不寫 .xlsx,只讀。
- ❌ **xlrd** — 已棄用,不支援 .xlsx 公式。
- ❌ **pywin32 / xlwings** — 需要 Excel 安裝,違反 cross-platform 原則。
- ❌ **任何需要 Excel/LibreOffice 執行公式的套件** — 違反「不執行公式」原則。
- ❌ **sqlite3 / sqlalchemy / pymysql** — 本工具不寫 DB。
- ❌ **flask / fastapi** — 本工具不做 server。

## 為什麼這樣鎖版本

- 主版本鎖死 (`< 4.0`) — 避免 breaking change 偷襲
- 最低版本鎖到 release date 後幾個月 — 確保套件穩定後才用
- `uv.lock` 記錄精確版本 — CI 跟本地完全一致

## 升級流程

升級任何套件:

1. 新增/修改 `pyproject.toml`
2. 跑 `uv lock --upgrade-package <name>`
3. 跑全部測試
4. 在 commit message 寫升級理由
5. PR 標題加上 `chore(deps):` 前綴

## 安裝指令範例

```bash
# 開發者初始化
git clone <repo>
cd xlsm-archaeologist
uv sync                     # 裝所有 deps (含 dev)
uv run xlsm-archaeologist version

# 跑測試
uv run pytest
uv run pytest --cov

# Lint + format
uv run ruff check .
uv run ruff format .
uv run mypy src
```
