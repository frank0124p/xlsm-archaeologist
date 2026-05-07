# CLAUDE.md — Working Rules for the Agent

> 這份檔案是給執行此專案的 Claude Code agent 看的。讀完並內化這些規則後再動手。

## 你是誰、要做什麼

你是被指派來實作 `xlsm-archaeologist` 的 agent。先讀:

1. `README.md` — 專案總覽
2. `PROJECT.md` — 願景、scope、non-goals
3. `ARCHITECTURE.md` — 三層架構
4. `PHASE_PLAN.md` — Phase 1-6 怎麼走
5. **本檔案** — 工作守則

## 工作模式

### 每個 phase 的標準流程

```
1. 讀 phases/phase_N_xxx/README.md (該 phase 目標)
2. 讀 phases/phase_N_xxx/tasks.md   (細項任務)
3. 列出執行計畫,呈報給人類確認
4. 開始實作 — 一個 task 完成 commit 一次
5. 跑 phases/phase_N_xxx/acceptance.md 的驗收 checklist
6. 寫一份 phase summary 給人類 review
7. 等人類說 "go phase N+1" 才進下一階段
```

**不要連續做多個 phase 不停下來**。每個 phase 結束都要等人類 review。

### Commit 規範

每個有意義的步驟都 commit。message 格式:

```
<type>(<scope>): <subject>

<body — 為什麼這樣做、有什麼權衡>
```

types: `feat`、`fix`、`refactor`、`test`、`docs`、`chore`

範例:
```
feat(extraction): add named_range extractor

Uses openpyxl.workbook.defined_names. Skips ranges with
#REF! errors and logs them to warnings list. See
reference/output_schema.md for output shape.
```

## 技術規範 (硬性)

### 環境

- **Python 3.12** (不是 3.11、不是 3.13)
- **uv** 做套件管理 (不是 pip、不是 poetry)
- **ruff** 做 lint + format (不是 black + flake8)
- **pytest** 做測試
- **typer** 做 CLI (不是 click 不是 argparse)
- **pydantic v2** 做資料模型 (不是 dataclass、不是 attrs)

### 套件白名單

只能用 `TECH_STACK.md` 列出的套件。要加新套件必須先在 commit message 解釋為什麼,
並更新 `TECH_STACK.md`。

### 不能做的事

- ❌ **不執行任何 VBA 或公式** — 只做靜態分析。任何嘗試「計算結果」的程式碼都禁止。
- ❌ **不修改原始 .xlsm** — 永遠以 read-only 模式打開。輸出永遠在另一個資料夾。
- ❌ **不寫 DB 連線** — 本工具只輸出檔案。LOAD DATA INFILE 是下游的事。
- ❌ **不做 web UI、不做 API server** — 純 CLI。
- ❌ **不裝 Node.js 套件** — 純 Python。
- ❌ **不在輸出檔案中混入 binary** — JSON/CSV/MD only。
- ❌ **不假裝解出無法靜態解析的 VBA 動態 range** — 標記 `has_dynamic_range: true`,
     在 warnings 列出,讓人類接手。

### 必須做的事

- ✅ 所有輸出都符合 `reference/output_schema.md` 與 `reference/csv_schemas.md` 定義
- ✅ 所有 pydantic model 都有 docstring,說明每個欄位代表什麼
- ✅ 所有複雜邏輯 (公式分類、VBA 讀寫識別) 都有對應的 unit test
- ✅ 輸出永遠 deterministic — 同樣輸入跑兩次結果一樣 (排序穩定、不用 set 直接序列化)
- ✅ 大檔案處理用 streaming 或 generator,不要把整份載入記憶體後才處理

## Code Style

### 命名

沿用使用者既有規範 (見 `CONVENTIONS.md`):

- 布林欄位:`is_*`、`has_*`、`can_*` (不能用 `_flag` 後綴)
- 多狀態列舉欄位:`*_status`,值為 enum
- snake_case 一律用 (Python 變數、JSON key、CSV column)
- 模組/檔案命名:小寫底線

### 結構

```python
# 一個典型的 extractor module 長這樣
from pathlib import Path
from typing import Iterator

from pydantic import BaseModel
from openpyxl import load_workbook

from xlsm_archaeologist.models import FormulaRecord
from xlsm_archaeologist.config import Settings


def extract_formulas(
    workbook_path: Path,
    settings: Settings,
) -> Iterator[FormulaRecord]:
    """逐一抽取所有 cell 中的公式。

    Yields FormulaRecord per cell containing a formula.
    Skips empty cells and pure-value cells.

    Args:
        workbook_path: 讀取目標 .xlsm 路徑 (read-only opened)
        settings: 全域設定 (含分類器、複雜度計算參數)

    Yields:
        FormulaRecord with cell_address, formula_text, ast, category, ...
    """
    ...
```

### Type hints

- 所有 public function 必須有完整 type hint
- 用 Python 3.12 syntax (`list[str]` 不用 `List[str]`、`X | Y` 不用 `Union[X, Y]`)
- pydantic v2 的 `Field` 用來加描述

### 錯誤處理

- 抽取失敗的 cell/sheet/procedure 不能讓整個 run 死掉
- 用 `warnings` 機制收集問題,寫進 `00_summary.json` 的 `warnings` 區塊
- 真正的 fatal error 才 raise,並讓 CLI 用 exit code 1 退出

## 測試規範

### Fixture

`tests/fixtures/` 有手寫的 mini .xlsm 檔案。每個 fixture 對應一個明確場景:

- `simple.xlsm` — 純資料、無公式無 VBA (最小 case)
- `formulas_basic.xlsm` — 各類公式 (lookup/branch/compute) 各一條
- `formulas_complex.xlsm` — 巢狀 IF + VLOOKUP 跨 sheet
- `vba_basic.xlsm` — 一個 Sub 讀 A1 寫 B1
- `vba_dynamic_range.xlsm` — 動態 range 場景 (測試 `has_dynamic_range` 標記)
- `circular.xlsm` — 含循環引用 (測試循環偵測)

### 測試覆蓋目標

- 每個 phase 的核心模組:**80%+**
- 公式分類器、VBA 讀寫識別:**90%+** (這兩塊最容易出錯)
- CLI 主要 flow:smoke test 即可

### 跑測試

```bash
uv run pytest                    # 全部
uv run pytest tests/test_phase_2 # 單一 phase
uv run pytest -k formula         # 關鍵字篩選
uv run pytest --cov              # 覆蓋率
```

## 輸出契約

下游系統 (Schema Studio、Rule Catalog、未來的 Form Assembly Service) 會依賴你的輸出格式。
**輸出 schema 一旦定下來就不能隨便改**。如果真的需要改:

1. 在 `reference/output_schema.md` 標明 schema version bump
2. 在 commit message 寫 BREAKING CHANGE
3. 提供 migration note

## 何時停下來問人

- 任務描述模糊到無法執行
- 兩種設計方案各有優劣,需要人類拍板
- 發現 fixture 缺少某類場景,需要人類補充
- 進度落後超過 50%
- 有任何安全/隱私疑慮 (例如 fixture 不小心包含真實業務資料)

## 何時不要問,直接做

- 命名選擇 (照 CONVENTIONS.md)
- 套件次要選擇 (照 TECH_STACK.md)
- code 結構細節 (照本檔案 § Code Style)
- commit message 寫法 (照本檔案 § Commit 規範)

## 完成標準

每個 phase 的完成 = `phases/phase_N_xxx/acceptance.md` 全部打勾 + 人類 review 通過。

整個專案完成 = Phase 6 acceptance 通過 + 拿真實 .xlsm 跑過一次 + summary report 產出。
