# Phase 1 — Tasks

逐項打勾。完成一項 commit 一次。

## Setup

- [ ] `git init` + 第一個 commit (handoff package)
- [ ] 建立 `pyproject.toml`,內容包含:
    - [ ] `[project]` metadata (name=`xlsm-archaeologist`, version=`0.1.0`, requires-python=`>=3.12,<3.13`)
    - [ ] dependencies (照 `TECH_STACK.md`)
    - [ ] dev dependencies group
    - [ ] `[project.scripts]` 註冊 `xlsm-archaeologist = "xlsm_archaeologist.cli:app"`
    - [ ] `[tool.ruff]` 設定 (line-length=100, target-version=py312)
    - [ ] `[tool.pytest.ini_options]` 設定 (testpaths=tests)
    - [ ] `[tool.mypy]` 設定 (strict=true)
- [ ] 建立 `.python-version` 內容 `3.12`
- [ ] 建立 `.gitignore` (Python 標配 + `archaeology_output/`)
- [ ] 跑 `uv sync` 產出 `uv.lock`
- [ ] commit: `chore: bootstrap project with uv and pyproject.toml`

## Source skeleton

- [ ] 建立 `src/xlsm_archaeologist/` 整個目錄樹 (見 README.md)
- [ ] `__init__.py` 內 `__version__ = "0.1.0"`
- [ ] `errors.py`:
    - `XlsmArchaeologistError` (base)
    - `InvalidFileError`, `ExtractionError`, `AnalysisError` (placeholders)
- [ ] `config.py`:
    - `Settings(BaseSettings)`,欄位:`max_formula_depth: int = 20`、`log_level: str = "info"`
- [ ] `utils/logging.py`:
    - `get_logger(name: str) -> Logger`
    - 用 `rich.logging.RichHandler`
- [ ] `utils/progress.py`:
    - `ProgressBar` 類別 / context manager 骨架 (內部用 `rich.progress`)
- [ ] commit: `feat(skeleton): add source package structure`

## CLI

- [ ] `cli.py`:
    - 用 `typer.Typer()` 建 app
    - 註冊 `version`、`analyze`、`inspect` 三個 command
    - `version` 真的實作:印出 `"xlsm-archaeologist {version}"` 等四行
    - `analyze` / `inspect` 印 `"not implemented in phase 1"` 並 exit 0
- [ ] `__main__.py`: `from xlsm_archaeologist.cli import app; app()`
- [ ] commit: `feat(cli): implement CLI skeleton with version command`

## Tests

- [ ] `tests/conftest.py`:基本 pytest 設定 (沒有 fixture 也 OK,先有檔案)
- [ ] `tests/test_phase_1_skeleton.py`:
    - test_version_command_works
    - test_analyze_command_callable
    - test_inspect_command_callable
- [ ] 跑 `uv run pytest`,全部 pass
- [ ] commit: `test(skeleton): add smoke tests for CLI`

## Quality

- [ ] `uv run ruff check .` 零錯誤
- [ ] `uv run ruff format .` 沒有 diff
- [ ] `uv run mypy src` 零錯誤
- [ ] commit (若需): `chore: fix lint/type warnings`

## 收尾

- [ ] 寫一份 `phase_1_summary.md` 給人類 review,內容:
    - 完成的 task 清單
    - 跑 `uv run xlsm-archaeologist version` 的輸出
    - 跑 `uv run pytest` 的輸出
    - 任何 deviation from the plan
- [ ] **停下來等人類 review**,不要進入 Phase 2
