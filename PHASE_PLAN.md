# Phase Plan

> 6 個 phase 的執行順序、依賴關係、每個 phase 的範圍邊界。
> 每個 phase 細節見 `phases/phase_N_xxx/`。

## 為什麼分 6 個 phase

每個 phase 滿足三個條件才能成為一個 phase:

1. **獨立可驗收** — 不依賴後續 phase 也能 review 這階段的產出
2. **聚焦單一面向** — 一個 phase 只做一類事 (抽取 / 分析 / 報告)
3. **容易 rollback** — 出問題不會牽連太多

## Phase 依賴圖

```
        Phase 1: Skeleton
             │
             ▼
        Phase 2: Extraction (sheet/cell/named_range/validation)
             │
       ┌─────┴─────┐
       ▼           ▼
   Phase 3:    Phase 4:
   Formula     VBA
   Analysis    Analysis
       │           │
       └─────┬─────┘
             ▼
        Phase 5: Dependency Graph
             │
             ▼
        Phase 6: Reports & Scoring
```

Phase 3 跟 Phase 4 可以並行做 (都依賴 Phase 2),但 agent 一次做一個比較不會亂。

## Phase 一覽

| # | 名稱 | 主要產出 | 依賴 | 預估工作量 |
|---|---|---|---|---|
| 1 | Skeleton | 可跑的 CLI 骨架、設定、logging | — | S (0.5d) |
| 2 | Extraction | `01-04, 06.csv/json` | 1 | M (1.5d) |
| 3 | Formula Analysis | `05_formulas.json` | 2 | M (1.5d) |
| 4 | VBA Analysis | `07-08.json` | 2 | L (2d) |
| 5 | Dependency Graph | `09-10.csv/json` | 3, 4 | M (1d) |
| 6 | Reports | `00_summary.json` + `reports/*` | 5 | S (0.5d) |

預估總工作量:7 天 (給 agent 跑大概是 1-2 天 wallclock)。

## Phase 1: Skeleton

**目標**:可以跑 `uv run xlsm-archaeologist version` 跟 `--help`,所有目錄結構就位。

**重點**:
- 不寫業務邏輯
- 設定 typer CLI、pydantic-settings、rich logging、ruff/mypy 規則
- `pyproject.toml` 完整、`uv.lock` 產生、`pytest` 能跑空測試

**驗收**:`phases/phase_1_skeleton/acceptance.md`

## Phase 2: Extraction

**目標**:用 openpyxl 把 .xlsm 的「結構性資料」全部抽出來,寫進 `01-04, 06` 號檔案。

**範圍**:
- ✅ workbook metadata (sha256、size、has_vba)
- ✅ sheet 清單 (名稱、隱藏狀態、used range、size)
- ✅ named range 清單 (含 dynamic 偵測)
- ✅ cell 清單 (只記「有意義的」 — 含公式 / validation / 被引用)
- ✅ validation 清單 (含 enum_values 解析)

**不在這個 phase 做**:
- ❌ 公式 AST 解析 (Phase 3)
- ❌ VBA (Phase 4)
- ❌ 依賴圖 (Phase 5)

**驗收**:`phases/phase_2_extraction/acceptance.md`

## Phase 3: Formula Analysis

**目標**:對 Phase 2 抽出的公式做 AST 解析、分類、複雜度計算,寫進 `05_formulas.json`。

**範圍**:
- ✅ 用 `openpyxl.formula.tokenizer.Tokenizer` token 化
- ✅ 自實作 simple parser 建 AST (用 token list)
- ✅ 公式分類 (見 `reference/formula_categories.md`)
- ✅ 複雜度計算 (`nesting_depth * 2 + function_count + ref_count`)
- ✅ referenced_cells 與 referenced_named_ranges 抽取

**已知限制**:
- ⚠ `formulas` 套件能力更強但較重,本 phase 先用內建 tokenizer。需要時再 upgrade。
- ⚠ 部分罕見公式 (如 dynamic array 帶 `@`、LET/LAMBDA) 標記為 `unparsable: true`,
  保留原文,在 `00_summary.warnings` 列出。

**驗收**:`phases/phase_3_formula_analysis/acceptance.md`

## Phase 4: VBA Analysis

**目標**:用 oletools 抽 VBA source code,做模組與 procedure 層級的讀寫 cell 識別。

**範圍**:
- ✅ 用 `oletools.olevba3.VBA_Parser` 抽所有模組原始碼
- ✅ 切分 Sub/Function/Property,建 procedure 列表
- ✅ 對每個 procedure 跑「讀寫 cell 識別」(見 `reference/vba_analysis_rules.md`):
  - 抓 `Range("...")`、`Cells(r, c)`、`[A1]`、`Sheets("...").Range`、named range 引用
  - 動態 range (含變數 / 串接) 標記 `has_dynamic_range: true`
- ✅ 偵測 event triggers (`Worksheet_Change` 等)
- ✅ 提取 procedure-to-procedure call graph (writes `calls` 欄位)

**已知限制**:
- ⚠ 無法解析 runtime computed range (如 `Range("A" & lastRow)`),只標記旗標
- ⚠ 無法跨檔案追 VBA 呼叫
- ⚠ 加密 VBA project 無法解析,寫 warning 並跳過

**驗收**:`phases/phase_4_vba_analysis/acceptance.md`

## Phase 5: Dependency Graph

**目標**:整合 Phase 3 (公式) 與 Phase 4 (VBA) 的引用資訊,建 cell-level DAG。

**範圍**:
- ✅ 用 networkx 建 DiGraph
- ✅ 加入 nodes (formula_cell / input_cell / output_cell / named_range / vba_procedure)
- ✅ 加入 edges (via: formula / vba_read_write / validation / named_range)
- ✅ 偵測循環 (`nx.simple_cycles`)
- ✅ 偵測孤島 (in_degree == 0 且不是 input)
- ✅ 計算 in_degree、out_degree、weakly connected components
- ✅ 輸出 `09_dependencies.csv` (邊清單) 與 `10_dependency_graph.json` (NetworkX node-link format)

**已知限制**:
- ⚠ VBA 動態 range 無法精確建邊,以「procedure → sheet (粗粒度)」表示
- ⚠ Range (如 `A1:A10`) 引用會展開成單一 node `Sheet!A1:A10`,不展開到每個 cell
  (避免 graph 爆炸)

**驗收**:`phases/phase_5_dependency_graph/acceptance.md`

## Phase 6: Reports & Scoring

**目標**:產出 `00_summary.json` 與 `reports/*`,給人看的版本。

**範圍**:
- ✅ `00_summary.json`:統計 + risk_indicators + complexity_score + warnings
- ✅ `formula_categories.csv`:每類公式數量
- ✅ `top_complex_formulas.csv`:複雜度 Top 50
- ✅ `hotspot_cells.csv`:被引用最多次的 Top 50 cell
- ✅ `vba_behavior.csv`:每個 procedure 一 row,讀寫 cell 概況
- ✅ `cycles.json`:所有循環引用
- ✅ `orphans.csv`:孤島公式
- ✅ `cross_sheet_refs.csv`:跨 sheet 依賴邊

**Complexity Score 公式**:

```
complexity_score =
    formula_count * 1
  + deeply_nested_formula_count * 5
  + dynamic_vba_range_count * 10
  + circular_reference_count * 20
  + cross_sheet_dependency_count * 0.5
  + orphan_formula_count * 0.3
```

`migration_difficulty`:
- < 200 → `low`
- 200-500 → `medium`
- 500-1000 → `high`
- ≥ 1000 → `very_high`

**驗收**:`phases/phase_6_reports/acceptance.md`

## 全專案完成標準

1. 所有 6 個 phase 的 acceptance 都打勾
2. `tests/fixtures/` 的所有 fixture 都跑得過
3. 拿一份**真實 .xlsm** (使用者提供) 跑過一次,輸出檢查無 fatal error
4. README 的 quickstart 步驟一字不漏照做能成功
5. 產出 `RUN_REPORT.md`:在真實 .xlsm 上跑出的關鍵指標
