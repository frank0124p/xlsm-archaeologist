# Architecture

## 三層架構

```
┌──────────────────────────────────────────────────────────────────┐
│ Layer 4: 報告層 (Reports)                                         │
│   reports/ 內五張分析報告 + 健康度評分                             │
│   消費者:工程師、PM、未來的 LLM 分析                              │
├──────────────────────────────────────────────────────────────────┤
│ Layer 3: 結構化儲存層 (Structured Output)                         │
│   01-10 號 JSON/CSV 檔案,所有資料正規化                          │
│   消費者:下游工具 (Schema Studio、MariaDB import、LLM)           │
├──────────────────────────────────────────────────────────────────┤
│ Layer 2: 分析層 (Analyzers)                                       │
│   - FormulaAnalyzer:公式 AST、分類、複雜度                       │
│   - VbaAnalyzer:VBA token 解析、讀寫 cell 識別                   │
│   - DependencyAnalyzer:cell-to-cell DAG、循環、孤島              │
│   - SummaryAnalyzer:統計、健康度評分                             │
├──────────────────────────────────────────────────────────────────┤
│ Layer 1: 抽取層 (Extractors)                                      │
│   - WorkbookExtractor:openpyxl 讀 sheet/cell/named range         │
│   - ValidationExtractor:資料驗證 / 下拉選單                       │
│   - VbaExtractor:oletools 抽 VBA source code                     │
└──────────────────────────────────────────────────────────────────┘
                              ↑
                         Input: .xlsm
```

## 資料流

```
.xlsm 檔案
   │
   ▼
┌─────────────┐    ┌──────────────────┐    ┌────────────────────┐
│ Extractors  │ →  │ Pydantic Models  │ →  │ Analyzers          │
│ (Layer 1)   │    │ (in-memory)      │    │ (Layer 2)          │
└─────────────┘    └──────────────────┘    └────────────────────┘
                                                     │
                                                     ▼
                                           ┌────────────────────┐
                                           │ Serializers        │
                                           │ (JSON/CSV writers) │
                                           └────────────────────┘
                                                     │
                                                     ▼
                                           archaeology_output/
                                           ├── 00_summary.json
                                           ├── 01-10 ...
                                           └── reports/
```

關鍵設計:
- Layer 1 只負責「讀出來變 pydantic model」,不做分析
- Layer 2 只在 in-memory model 上跑,不再碰 .xlsm 檔案
- Serializers 跟 Analyzers 分離,讓相同 model 可以輸出多種格式

## 模組劃分

```
src/xlsm_archaeologist/
├── __init__.py
├── __main__.py                  # python -m xlsm_archaeologist
├── cli.py                       # typer CLI 進入點
├── config.py                    # pydantic-settings 全域設定
├── models/                      # Pydantic models (Layer 1.5)
│   ├── __init__.py
│   ├── workbook.py              # WorkbookRecord, SheetRecord
│   ├── cell.py                  # CellRecord, ValidationRecord
│   ├── formula.py               # FormulaRecord, FormulaCategory
│   ├── vba.py                   # VbaModuleRecord, VbaProcedureRecord
│   ├── dependency.py            # DependencyEdge, DependencyGraph
│   └── summary.py               # SummaryRecord, HealthScore
├── extractors/                  # Layer 1
│   ├── __init__.py
│   ├── workbook_extractor.py
│   ├── validation_extractor.py
│   └── vba_extractor.py
├── analyzers/                   # Layer 2
│   ├── __init__.py
│   ├── formula_analyzer.py
│   ├── vba_analyzer.py
│   ├── dependency_analyzer.py
│   └── summary_analyzer.py
├── serializers/                 # Layer 3 writer
│   ├── __init__.py
│   ├── json_writer.py
│   ├── csv_writer.py
│   └── markdown_writer.py
├── reports/                     # Layer 4
│   ├── __init__.py
│   ├── formula_categories.py
│   ├── top_complex.py
│   ├── hotspots.py
│   ├── vba_behavior.py
│   ├── cycles.py
│   ├── orphans.py
│   └── cross_sheet_refs.py
└── utils/
    ├── __init__.py
    ├── address.py               # cell address 解析 (A1 ↔ (col, row))
    ├── logging.py               # 統一 logger 設定
    └── progress.py              # rich progress bar wrapper
```

## 跨層約定

### Pydantic Model 是契約

所有跨層傳遞的資料都是 pydantic model。**不允許**用 dict 在 layer 之間傳。
這讓:
- IDE 補全跟 type check 都有效
- JSON 輸出 schema 自動跟 model 同步
- 改欄位時 mypy 會抓出所有受影響的地方

### Cell Address 統一格式

所有 cell address 在內部一律用 `"SheetName!A1"` 字串格式 (含 sheet 名稱)。
解析跟組裝都透過 `utils/address.py` 提供的 helper。

### Logging 與 Warnings

- 一般進度:`logger.info` + rich progress bar
- 抽取/分析中遇到的可恢復錯誤:`warnings_collector.add(...)`,最後寫進 `00_summary.json`
- 致命錯誤:raise + CLI exit 1

### Determinism

所有輸出必須 deterministic:
- 字典/集合輸出前先 `sorted()`
- pydantic model 序列化用 `model_dump_json(indent=2)`
- CSV 輸出前先按 primary key 排序
- 圖節點/邊輸出前按 ID 排序
