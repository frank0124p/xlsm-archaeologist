# xlsm-archaeologist

> 一個 Python CLI 工具,把複雜的 .xlsm (帶 VBA 巨集的 Excel) 完整考古成結構化的 JSON/CSV 資料,
> 包含 sheet/cell/公式/VBA/依賴圖。輸出可直接餵給後續的規則引擎設計、LLM 分析、或進 DB。

## 這個 handoff package 給誰看

- **主要對象**:Claude Code agent (自主執行)
- **次要對象**:工程師 (review、debug、接手)

整包設計成 agent 從 `README.md` → `PROJECT.md` → `CLAUDE.md` → `phases/phase_1_skeleton/` 一路讀下去就知道該做什麼。

## 怎麼開始

```bash
# 1. 解壓並 init 一個新 repo
unzip xlsm_archaeologist_handoff.zip
cd xlsm_archaeologist_handoff
git init && git add . && git commit -m "chore: import handoff package"

# 2. 啟動 Claude Code
claude

# 3. 給 agent 的開場指令 (複製貼上)
請完整閱讀 README.md → PROJECT.md → CLAUDE.md → phases/phase_1_skeleton/README.md,
理解整個專案目標與第一階段任務後,先列出你的執行計畫給我確認,再開始動手。
```

## Phase 一覽 (依依賴順序)

| Phase | 主題 | 輸出 | 依賴 |
|---|---|---|---|
| 1 | Skeleton & CLI | 可跑 `xlsm-archaeologist --version` | — |
| 2 | Structure Extraction | `01-04, 06.csv/json` (sheet/cell/named_range/validation) | 1 |
| 3 | Formula Analysis | `05_formulas.json` (AST、分類、複雜度) | 2 |
| 4 | VBA Analysis | `07-08.json` (modules、procedures、讀寫 cell) | 2 |
| 5 | Dependency Graph | `09-10.csv/json` (DAG、循環、孤島) | 3, 4 |
| 6 | Reports & Scoring | `00_summary.json` + `reports/*` | 5 |

每個 phase 都是**獨立可驗收**的。Agent 跑完一個 phase 就停下來,讓人 review,再進下一個。

## 專案結構導覽

```
xlsm_archaeologist_handoff/
├── README.md              ← 你在這裡
├── PROJECT.md             ← 專案目標、為什麼存在、不做什麼
├── CLAUDE.md              ← 給 agent 的工作守則 (規範、技術選型、禁止事項)
├── ARCHITECTURE.md        ← 三層架構、模組劃分、資料流
├── DATA_MODEL.md          ← 所有 JSON/CSV 的欄位定義
├── CLI_CONTRACT.md        ← CLI 介面與輸出契約
├── TECH_STACK.md          ← 套件選型、版本鎖定、選型理由
├── PHASE_PLAN.md          ← Phase 1-6 的詳細計畫 + 依賴關係
├── ACCEPTANCE.md          ← 每個 phase 的驗收標準
├── CONVENTIONS.md         ← 命名、code style、commit、目錄結構
│
├── phases/                ← 每個 phase 一個資料夾
│   ├── phase_1_skeleton/
│   ├── phase_2_extraction/
│   ├── phase_3_formula_analysis/
│   ├── phase_4_vba_analysis/
│   ├── phase_5_dependency_graph/
│   └── phase_6_reports/
│       └── 每個 phase 內部:
│           ├── README.md       ← 該 phase 的目標、deliverables、limitations
│           ├── tasks.md        ← 細部 task 清單 (給 agent 逐項打勾)
│           └── acceptance.md   ← 該 phase 的驗收 checklist
│
├── reference/             ← 規範與規則 (agent 要參照)
│   ├── output_schema.md        ← 所有 JSON 的 schema 定義
│   ├── csv_schemas.md          ← 所有 CSV 的欄位定義
│   ├── formula_categories.md   ← 公式分類規則 + edge case
│   ├── vba_analysis_rules.md   ← VBA 讀寫 cell 識別規則 + 限制
│   └── example_outputs/        ← 預期輸出範例 (給 agent 比對)
│
├── tests/                 ← 測試策略與 fixture 設計
│   ├── README.md          ← 測試策略
│   ├── test_plan.md       ← 各 phase 的測試案例清單
│   └── fixtures/          ← mini .xlsm fixture 設計說明
│
└── prompts/               ← 給 agent 的關鍵 prompt 範本
    └── kickoff.md         ← Phase 1 開工的 prompt
```

## FAQ

**Q: 為什麼是 Python 不是 Node.js?**
→ `TECH_STACK.md` § 為什麼選 Python

**Q: 為什麼分 6 個 phase?**
→ `PHASE_PLAN.md` § Phase 邊界與依賴

**Q: VBA 動態 range 怎麼處理?**
→ `reference/vba_analysis_rules.md` § 已知限制

**Q: 公式分類規則怎麼定?**
→ `reference/formula_categories.md`

**Q: 我想看實際輸出長怎樣?**
→ `reference/example_outputs/`

## Status

Handoff package complete. Ready for kickoff.
