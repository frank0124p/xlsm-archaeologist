# Phase 1: Skeleton

## 目標

建立可執行的 CLI 骨架 — 跑 `uv run xlsm-archaeologist version` / `--help` 都可以,
所有目錄結構就位、設定 / logging / 測試框架都能動。**不寫業務邏輯**。

## Deliverables

```
xlsm-archaeologist/
├── pyproject.toml
├── uv.lock
├── README.md (簡版,引導到 handoff README)
├── .gitignore
├── .python-version
├── ruff.toml (或內嵌 pyproject.toml)
├── src/xlsm_archaeologist/
│   ├── __init__.py
│   ├── __main__.py
│   ├── cli.py                     # typer app, version / analyze / inspect 三個指令的骨架
│   ├── config.py                  # pydantic-settings Settings class
│   ├── errors.py                  # 自訂 exception 類別骨架
│   ├── models/__init__.py         # 空檔
│   ├── extractors/__init__.py     # 空檔
│   ├── analyzers/__init__.py      # 空檔
│   ├── serializers/__init__.py    # 空檔
│   ├── reports/__init__.py        # 空檔
│   └── utils/
│       ├── __init__.py
│       ├── logging.py             # rich handler 設定
│       └── progress.py            # rich progress wrapper 骨架
└── tests/
    ├── conftest.py
    ├── fixtures/                  # 空目錄 + .gitkeep
    └── test_phase_1_skeleton.py   # smoke test
```

## 重點

- CLI 三個指令必須都「能呼叫」,但 `analyze` 跟 `inspect` 印 `"not implemented in phase 1"` 即可
- `version` 必須真的可用,輸出格式照 `CLI_CONTRACT.md` § version
- logging 走 rich handler,`--log-level` 控制

## 驗收

見 `acceptance.md`。

## 限制與注意

- 不要寫任何 .xlsm 解析邏輯
- 不要 import openpyxl / oletools / networkx (留到對應 phase)
- 但 `pyproject.toml` 已經把它們列進 dependencies,確保 `uv sync` 一次裝齊
