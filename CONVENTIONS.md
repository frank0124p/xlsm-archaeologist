# Conventions

> 命名、code style、commit、目錄結構規範。Agent 必須遵守。

## 命名規範

### Python

- **模組/檔案**:`snake_case.py`
- **類別**:`PascalCase`
- **函式/變數**:`snake_case`
- **常數**:`UPPER_SNAKE_CASE`
- **私有**:單底線 `_private`,雙底線只用在 name mangling

### 布林欄位 (重要,沿用使用者既有規範)

- ✅ `is_hidden`、`is_referenced`、`is_volatile`、`is_array_formula`
- ✅ `has_formula`、`has_validation`、`has_vba`、`has_dynamic_range`
- ✅ `can_be_parsed`、`can_resolve`
- ❌ **不准** `*_flag` 後綴 (如 `hidden_flag`)
- ❌ **不准** 反義 (如 `is_not_visible`)

### 多狀態欄位

- 用 `*_status` 結尾,值為 enum (string literal)
- 範例:`migration_difficulty: "low" | "medium" | "high" | "very_high"`
- 範例:`module_type: "standard" | "class" | "form" | "sheet" | "workbook" | "unknown"`

### JSON / CSV key

- 一律 `snake_case`
- 與 Python pydantic field 同名 (避免 alias 帶來的混淆)
- 不用縮寫,除非業界標準 (如 `vba`、`url`、`sha256`)

### Cell address

- 內部一律使用 `qualified_address` 格式:`SheetName!A1`
- 不含 sheet 前綴的版本叫 `cell_address`,只在 CSV 的 sheet 欄位明確時用
- Range 用 `Sheet!A1:B10` 格式

## Code Style

### 強制

```bash
# CI / pre-commit 跑這三個
uv run ruff check .       # lint
uv run ruff format .      # format
uv run mypy src           # type check
```

### 函式長度

- 一個函式不超過 **50 行** (含 docstring)
- 超過就拆 helper

### 巢狀深度

- 最多 3 層 if/for 巢狀
- 更深要重構成 early return 或拆 helper

### Import 順序 (ruff 自動處理)

1. 標準庫
2. 第三方套件
3. 本專案

### Docstring

所有 public 模組、類別、函式都要有 docstring。格式:

```python
def extract_formulas(workbook_path: Path) -> Iterator[FormulaRecord]:
    """逐一抽取所有 cell 中的公式。

    Yields FormulaRecord per cell containing a formula.
    Skips empty cells and pure-value cells.

    Args:
        workbook_path: 讀取目標 .xlsm 路徑 (read-only opened)

    Yields:
        FormulaRecord with cell_address, formula_text, ast, category, ...

    Raises:
        FileNotFoundError: 檔案不存在
        InvalidFileError: 不是有效的 xlsx/xlsm
    """
```

### Type hints

- 所有 public function 參數與回傳都要 type hint
- 用 Python 3.12 syntax:`list[str]`、`dict[str, int]`、`X | Y`、`X | None`
- 不用 `typing.List` / `typing.Optional` (舊 style)

### 例外處理

- 只 catch 具體例外,不要 bare `except`
- 預期得到的錯誤用自訂 exception 類別 (在 `xlsm_archaeologist/errors.py`)
- 例外鏈:`raise NewError("...") from original_error`

## Commit 規範

### 格式

```
<type>(<scope>): <subject>

<body>
```

- subject 不超過 50 字,不結尾句號
- body 用來說明「為什麼」,不是「做了什麼」(diff 已經說了)

### Types

| Type | 用在 |
|---|---|
| `feat` | 新功能 |
| `fix` | bug 修正 |
| `refactor` | 重構,不改外部行為 |
| `test` | 加 / 改測試 |
| `docs` | 文件 (含 docstring) |
| `chore` | 雜項 (deps、CI、lint config) |
| `perf` | 效能優化 |

### Scopes

| Scope | 對應 |
|---|---|
| `cli` | CLI 介面 |
| `extraction` | Phase 2 |
| `formula` | Phase 3 |
| `vba` | Phase 4 |
| `graph` | Phase 5 |
| `reports` | Phase 6 |
| `models` | pydantic model |
| `config` | 設定 |
| `tests` | 測試 |
| `deps` | 套件相依 |

### 範例

```
feat(extraction): add named_range extractor

Uses openpyxl.workbook.defined_names. Skips ranges with
#REF! errors and logs them to warnings list.
```

```
fix(formula): handle empty IF branches

Previously raised IndexError when IF had only one arg
(e.g. =IF(A1>0)). Now treats missing branches as empty
string. Test added in test_formula_analyzer.py.
```

## 目錄結構規範

### Source

```
src/xlsm_archaeologist/
├── __init__.py
├── __main__.py
├── cli.py
├── config.py
├── errors.py
├── models/
├── extractors/
├── analyzers/
├── serializers/
├── reports/
└── utils/
```

### Tests

```
tests/
├── conftest.py              # pytest fixtures (mini .xlsm 路徑)
├── fixtures/                # 手寫的 .xlsm 測試檔
│   ├── simple.xlsm
│   ├── formulas_basic.xlsm
│   ├── formulas_complex.xlsm
│   ├── vba_basic.xlsm
│   ├── vba_dynamic_range.xlsm
│   └── circular.xlsm
├── test_phase_1_skeleton.py
├── test_phase_2_extraction.py
├── test_phase_3_formula.py
├── test_phase_4_vba.py
├── test_phase_5_graph.py
├── test_phase_6_reports.py
└── test_cli.py
```

### Pytest 慣例

- 測試函式名:`test_<unit>_<scenario>` (如 `test_formula_classifier_handles_nested_if`)
- fixture 寫在 `conftest.py`
- 用 `tmp_path` 而不是 `/tmp/...`
- 一個 test 一個 assertion 概念 (除非真的需要多個 assert)

## 檔案頭

每個 .py 檔案開頭:

```python
"""模組目的的一句話描述。

詳細描述 (可選)。
"""
from __future__ import annotations  # 不是必要,但 3.12 仍建議

# imports...
```

不需要 license header (已在 LICENSE 檔)。
不需要 `# -*- coding: utf-8 -*-` (Python 3 預設)。

## 字串 quote

- 一律用 **double quote** `"..."` (ruff 預設)
- 例外:字串內含 `"` 時用 single quote 避免 escape

## 錯誤訊息

- 中文 OK (使用者環境是繁中)
- 但變數名、技術術語維持英文
- 範例:`f"無法解析公式 {formula_text}: {error}"`

## Logging

- 用標準 `logging` 模組 + rich handler
- 全域 logger 從 `xlsm_archaeologist.utils.logging.get_logger(__name__)` 取
- log level 透過 CLI `--log-level` 控制
- 不准用 `print()` 做 log (CLI 真正的輸出例外)
