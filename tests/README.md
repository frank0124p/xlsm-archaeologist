# Tests

## 測試策略

三層金字塔:

```
        ┌─────────────────┐
        │   E2E (1-2 個)   │  跑完整 pipeline 對 fixture
        └─────────────────┘
       ┌───────────────────┐
       │ Phase 整合測 (6)   │  每 phase 一個整合測,確認 phase 內的協調
       └───────────────────┘
     ┌─────────────────────────┐
     │ Unit Tests (大多數)      │  個別 analyzer / extractor 的單元測
     └─────────────────────────┘
```

## 不依賴 mock

我們的測試**不用 unittest.mock 或 pytest-mock**。原因:

- openpyxl 與 oletools 的 mock 太脆弱,版本一升就壞
- 用真實 fixture .xlsm 才能涵蓋真實 edge case
- 整套工具都是 read-only,沒有副作用,不需要 mock 隔離

例外:
- 加密 vbaProject 場景沒辦法用 fixture (建檔複雜),這個 case 用 monkeypatch + minimal mock

## Fixture 命名

`tests/fixtures/<scenario>.xlsm`,scenario 名稱見 `fixtures/README.md`。

## 跑測試

```bash
# 全部
uv run pytest

# 個別 phase
uv run pytest tests/test_phase_2_extraction.py

# 關鍵字
uv run pytest -k "formula and complex"

# 帶覆蓋率
uv run pytest --cov=src/xlsm_archaeologist --cov-report=term-missing

# 詳細輸出
uv run pytest -v --tb=short

# 只跑失敗的
uv run pytest --lf
```

## 覆蓋率目標

| 模組 | 目標 |
|---|---|
| `extractors/` | 80% |
| `analyzers/formula_*` | 90% |
| `analyzers/vba_*` | 90% |
| `analyzers/dependency_*` | 85% |
| `serializers/` | 80% |
| `reports/` | 80% |
| `cli.py` | 70% (smoke tests 即可) |
| 整體 | ≥ 75% |

公式分類器跟 VBA 讀寫識別覆蓋率特別高的原因:這兩塊最容易出錯,
而且邏輯密集 — 同樣的 LOC 蓋住更多分支。

## 測試命名

`test_<unit>_<scenario>` 格式:

- ✅ `test_formula_classifier_handles_nested_if`
- ✅ `test_vba_range_detector_marks_dynamic_when_concat`
- ❌ `test_works` (沒 scenario)
- ❌ `test1` (沒含意)

一個 test 一個 assertion 概念 (允許多個 assert,但都要圍繞同一個 scenario)。

## 慣例

- fixture 路徑用 `tests.fixtures` 模組相對路徑找,不寫絕對路徑
- 檔案 IO 用 `tmp_path` (pytest 內建),不寫 `/tmp/...`
- `conftest.py` 提供共用 fixture (例如已載入的 workbook)

範例 conftest:

```python
import pytest
from pathlib import Path

FIXTURES_DIR = Path(__file__).parent / "fixtures"


@pytest.fixture
def simple_xlsm() -> Path:
    return FIXTURES_DIR / "simple.xlsm"


@pytest.fixture
def formulas_basic_xlsm() -> Path:
    return FIXTURES_DIR / "formulas_basic.xlsm"


@pytest.fixture
def vba_basic_xlsm() -> Path:
    return FIXTURES_DIR / "vba_basic.xlsm"


# ... 每個 fixture 一個 pytest fixture
```

## CI

雖然這個 repo 不一定要 CI,但建議至少在 pre-commit hook 跑:

```yaml
# .pre-commit-config.yaml (建議)
repos:
  - repo: local
    hooks:
      - id: ruff-check
        name: ruff check
        entry: uv run ruff check
        language: system
        pass_filenames: false
      - id: ruff-format
        name: ruff format
        entry: uv run ruff format --check
        language: system
        pass_filenames: false
      - id: mypy
        name: mypy
        entry: uv run mypy src
        language: system
        pass_filenames: false
      - id: pytest
        name: pytest
        entry: uv run pytest
        language: system
        pass_filenames: false
```

## 手寫 fixture .xlsm 的方式

詳見 `fixtures/README.md`。簡單版:

```python
# scripts/build_fixtures.py (一次性 build 腳本,不在 src/ 裡)
from openpyxl import Workbook

def build_simple():
    wb = Workbook()
    ws = wb.active
    ws.title = "Data"
    for i in range(1, 4):
        for j in range(1, 4):
            ws.cell(i, j, value=i * j)
    wb.save("tests/fixtures/simple.xlsm")
```

帶 VBA 的 fixture 不能用 openpyxl 純粹寫 — VBA 需要事先有一份 .xlsm 包含目標 VBA,
然後用 openpyxl 讀進來改其他部分,save 出去保留 VBA。

詳細步驟見 `fixtures/README.md`。
