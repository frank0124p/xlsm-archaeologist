# Kickoff Prompt

> 把這份 handoff package 餵給 Claude Code 時,複製下面這段作為第一條訊息。

---

## 開場 prompt (複製貼上給 agent)

```
你是被指派執行 xlsm-archaeologist 專案的 agent。

請先完整閱讀以下檔案,理解後再開始動手:

1. README.md          ← 專案總覽與目錄結構
2. PROJECT.md         ← 為什麼存在、scope、non-goals
3. CLAUDE.md          ← 給你的工作守則 (重要!裡面講禁止事項與必做事項)
4. ARCHITECTURE.md    ← 三層架構
5. PHASE_PLAN.md      ← 6 個 phase 的依賴關係
6. CONVENTIONS.md     ← 命名與 code 規範
7. TECH_STACK.md      ← 套件選型
8. DATA_MODEL.md      ← 輸出契約
9. CLI_CONTRACT.md    ← CLI 介面契約
10. phases/phase_1_skeleton/README.md
11. phases/phase_1_skeleton/tasks.md
12. phases/phase_1_skeleton/acceptance.md

讀完後請做兩件事,**不要動手寫 code**:

(A) 列出你的 Phase 1 執行計畫,內容包含:
    - 你打算建立哪些檔案
    - commit 的順序
    - 任何你看不懂或想釐清的點 (對應到具體檔案的具體段落)

(B) 列出你對 reference/ 與 tests/ 的理解:
    - 你打算什麼時候參照 reference/output_schema.md 與 reference/csv_schemas.md
    - 你計畫如何看待 tests/test_plan.md (是要全做還是先做子集)

我看完後會給你 go signal,你才開始動手 Phase 1。

幾個重要原則 (從 CLAUDE.md 摘出來,確認你看到了):

✗ 不允許執行任何 VBA 或公式
✗ 不允許修改原始 .xlsm
✗ 不允許寫 DB 連線
✗ 不允許用 pandas / xlrd / xlwings / pywin32
✗ 不允許在每個 phase 之間跳關 — 一個 phase 完成 stop 等 review

✓ 只能用 TECH_STACK.md 列出的套件
✓ 所有輸出必須 deterministic
✓ VBA 動態 range 必須誠實標 has_dynamic_range,不准假裝解出來
✓ 所有布林欄位必須 is_/has_/can_ 前綴

開始吧。
```

---

## 後續 phase 的 prompt 範本

每個 phase review 完後,進下一 phase 的 prompt:

```
Phase N 完成,review 通過。進入 Phase N+1。

請先讀:

1. phases/phase_{N+1}_xxx/README.md
2. phases/phase_{N+1}_xxx/tasks.md
3. phases/phase_{N+1}_xxx/acceptance.md
4. (相關的 reference/*)

列出你的 Phase N+1 執行計畫,等我確認後再動手。
```

## Debug / 卡住時的 prompt

```
請暫停目前的 task。看 CLAUDE.md § "何時停下來問人",判斷你目前的狀況屬於哪一類。

- 如果是「任務描述模糊」、「兩種設計各有優劣」、「fixture 缺場景」、「進度落後 50%」、
  「安全/隱私疑慮」 — 請給我一個簡短的問題,告訴我你卡在哪、需要什麼決定。

- 如果不是上述任何一類 — 請直接做下去,不要問。
```

## 最終驗收的 prompt

```
所有 6 個 phase 完成。請執行最終驗收流程:

1. 跑 ACCEPTANCE.md 的全專案 checklist,逐項打勾
2. 拿 (我提供的) 真實 .xlsm 跑 `xlsm-archaeologist analyze ...`
3. 在 RUN_REPORT.md 寫:
    - 該檔案的 stats / complexity_score / migration_difficulty
    - top 5 warnings 摘要
    - 跑 wallclock 時間
    - 任何在真實檔案上發現但 fixture 沒涵蓋的場景
4. 如果發現任何 critical / major issue,在 RUN_REPORT.md 標明,並建議下一個 backlog item

完成後給我最終報告。
```
