# Acceptance Criteria

> 整個專案的驗收標準總覽。每個 phase 細部驗收見 `phases/phase_N_xxx/acceptance.md`。

## 全專案驗收 Checklist

### Phase 完整性

- [ ] Phase 1 acceptance 全打勾
- [ ] Phase 2 acceptance 全打勾
- [ ] Phase 3 acceptance 全打勾
- [ ] Phase 4 acceptance 全打勾
- [ ] Phase 5 acceptance 全打勾
- [ ] Phase 6 acceptance 全打勾

### 程式碼品質

- [ ] `uv run ruff check .` 零錯誤
- [ ] `uv run ruff format --check .` 零差異
- [ ] `uv run mypy src` 零錯誤
- [ ] `uv run pytest` 全部 pass
- [ ] 整體覆蓋率 ≥ 75%
- [ ] 公式分類器、VBA 讀寫識別 ≥ 90%

### 輸出契約

- [ ] 所有 JSON 輸出含 `schema_version: "1.0"`
- [ ] 所有 CSV header 與 `DATA_MODEL.md` 完全一致
- [ ] 所有 enum 欄位的值都在 `DATA_MODEL.md` 規定的範圍內
- [ ] 同一份 .xlsm 跑兩次,輸出完全一致 (deterministic)

### 文件

- [ ] README quickstart 步驟可一字不漏執行成功
- [ ] 所有 public 模組/函式有 docstring
- [ ] `RUN_REPORT.md` 提供:在真實 .xlsm 上跑出的指標摘要

### 真實場景驗證

- [ ] 在使用者提供的真實 .xlsm 上跑過一次
- [ ] 沒有 fatal error
- [ ] warnings 都有合理解釋 (在 `00_summary.json#warnings`)
- [ ] 使用者目視 review 過 `top_complex_formulas.csv` 跟 `hotspot_cells.csv`,認為結果合理

### Fixture 覆蓋

- [ ] `simple.xlsm` 跑得過,輸出符合預期
- [ ] `formulas_basic.xlsm` 各類公式分類正確
- [ ] `formulas_complex.xlsm` 巢狀 IF + VLOOKUP 跨 sheet 解析正確
- [ ] `vba_basic.xlsm` 讀寫 cell 識別正確
- [ ] `vba_dynamic_range.xlsm` 動態 range 正確標記 `has_dynamic_range: true`
- [ ] `circular.xlsm` 循環引用正確偵測

## 三類 bug 嚴重度

### Critical (必修才能交付)

- 跑某類 fixture 直接 crash
- 公式分類錯誤導致整體統計失真
- VBA 動態 range **沒**標記為 `has_dynamic_range`,導致下游誤信
- JSON schema 違反 `DATA_MODEL.md` 定義
- Output 不 deterministic

### Major (建議修,可帶 known issue 交付)

- 罕見公式 (如 LET/LAMBDA) 分類為 `unparsable`
- VBA 加密 project 無法解析,但有寫 warning
- 部分跨 sheet 引用解析不完整,但有寫 warning

### Minor (記錄到 backlog,後續版本)

- Progress bar 顯示偶有閃爍
- Markdown 報告排版不夠美

## 驗收會議的 Demo 流程

人類 review 時建議照這個順序:

1. **跑 quickstart**:照 README 步驟,從 clone repo 到產出第一份 archaeology_output
2. **看 summary**:打開 `00_summary.json`,確認指標合理
3. **抽看 5 條公式**:從 `05_formulas.json` 隨機挑 5 條,確認分類與依賴解析正確
4. **抽看 1 個 VBA procedure**:從 `08_vba_procedures.json` 挑一個,確認讀寫識別
5. **驗依賴圖**:用 Python REPL 載入 `10_dependency_graph.json` 重建 NetworkX 圖,
   驗證某個已知關係 (如「改 Params!A1 會影響 Output!Z1」)
6. **看 reports**:檢視五張報告是否能用人眼快速消化

如果這 6 步都過,就算交付成功。
