# Phase 1 — Acceptance Checklist

## 結構

- [ ] `pyproject.toml` 存在且 valid (`uv sync` 不報錯)
- [ ] `uv.lock` 已產生並 commit
- [ ] `.python-version` 內容是 `3.12`
- [ ] 目錄結構與 `phases/phase_1_skeleton/README.md` 完全一致
- [ ] 所有 `__init__.py` 存在 (即使是空檔)

## CLI 行為

- [ ] `uv run xlsm-archaeologist --help` 顯示三個 command:`version`, `analyze`, `inspect`
- [ ] `uv run xlsm-archaeologist version` 輸出至少四行 (tool/schema/python/openpyxl 版本)
- [ ] `uv run xlsm-archaeologist analyze ./anything.xlsm` 印 `"not implemented in phase 1"` 並 exit 0
- [ ] `uv run xlsm-archaeologist inspect ./anything.xlsm` 印 `"not implemented in phase 1"` 並 exit 0
- [ ] 所有 progress / log 走 stderr,不污染 stdout

## 程式品質

- [ ] `uv run ruff check .` 零錯誤
- [ ] `uv run ruff format --check .` 零差異
- [ ] `uv run mypy src` 零錯誤
- [ ] `uv run pytest` 全部 pass
- [ ] 至少 3 個 smoke test (對應三個 command)

## 文件

- [ ] 簡版 `README.md` 在專案根目錄,引導到 handoff package
- [ ] `phase_1_summary.md` 已寫,描述完成狀態與下一步建議

## 不能出現

- [ ] **沒有** 業務邏輯 (沒有解析 .xlsm 的 code)
- [ ] **沒有** import openpyxl / oletools / networkx
- [ ] **沒有** 違反 `CONVENTIONS.md` 的命名 (布林欄位若有,必須 `is_/has_/can_`)
