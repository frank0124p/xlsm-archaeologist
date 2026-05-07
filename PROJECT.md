# xlsm-archaeologist — Project Vision

> 把一份「公式 + VBA 巨集 + 跨 sheet 引用」糾纏在一起的 .xlsm 檔案,
> 完整考古成結構化資料,讓你在動手重構前,先**看清楚到底有什麼**。

## 為什麼存在

很多企業有這樣一份 Excel:

- 幾十個 sheet
- 上千條公式,巢狀 IF + VLOOKUP + 跨 sheet 引用
- 一堆 VBA 巨集,Worksheet_Change 事件、動態 Range、自動展開 row
- 每次業務需求變動,工程師花一整天在追「這個公式為什麼會算錯」
- 想重構,但沒人敢動,因為**沒人知道全貌**

xlsm-archaeologist 的核心價值是:在你決定怎麼重構之前,**先把這份 Excel 完整考古一遍**,
讓所有公式、VBA、依賴關係都變成可查詢、可分析、可量化的結構化資料。

## 它能回答什麼問題

跑完之後,你可以立刻回答這些問題:

1. 這份檔案裡到底有**多少條公式**?分別屬於哪幾類 (lookup / branch / compute / aggregate)?
2. **最複雜的 50 條公式**長什麼樣?它們是規則引擎重構的優先目標。
3. **被引用最多次的 cell** 是哪些?它們通常是「核心參數」(稅率、匯率、單價基準),
   在新系統會變成 `editable_parameters`。
4. 改了 `Params!A1` 會**影響哪些 cell**?要透過哪些公式/VBA 影響?
5. `Output!Z7` 的值是**從哪些輸入一層層算出來的**?
6. 哪些 VBA Sub 讀寫了哪些 cell?哪些是 event-triggered (`Worksheet_Change`)?
7. 有沒有**循環引用**?有沒有**孤島公式** (沒人引用、可以砍)?
8. 整體**遷移難度評分**多少?

## 它不做什麼 (Non-goals)

明確列出非目標,避免 scope creep:

- ❌ **不執行公式、不算結果** — 只做靜態分析。動態追蹤 (餵測試資料跑新舊比對) 是後續工具的事。
- ❌ **不重寫 .xlsm** — 純讀取,輸出永遠在另一個資料夾。原檔案不動。
- ❌ **不做 UI** — 純 CLI 工具,輸出 JSON/CSV。視覺化交給下游。
- ❌ **不直接寫 DB** — 輸出檔案後,要進 MariaDB/ClickHouse 是下游工具的事。
- ❌ **不處理跨 workbook 引用** — 單檔分析。跨檔留給未來版本。
- ❌ **不做規則引擎** — 那是「Form Assembly Service」的事。本工具只是它的前置考古。

## 它的下游 (這個工具產出後可以接什麼)

```
xlsm-archaeologist (本工具)
    ↓ 產出 JSON/CSV
    ↓
    ├──→ Schema Studio:把 sheet 結構轉成 form schema
    ├──→ Rule Catalog:把公式分類匯入成 rule definitions
    ├──→ MariaDB:LOAD DATA INFILE 進關聯資料庫做 SQL 查詢
    ├──→ LLM 分析:把複雜公式餵給 Claude 做語意化說明
    └──→ Form Assembly Service:作為新規則引擎的設計依據
```

## 目標使用者

- **主要**:工程師 (你),要設計新系統取代複雜 Excel 巨集
- **次要**:接手或新進工程師,要快速理解現有 Excel 的結構
- **非目標**:業務人員 (這是內部分析工具,不需要 UI)

## 關鍵設計原則

1. **靜態分析優先** — 不執行 VBA、不算公式,只看「程式碼長什麼樣」與「引用了誰」。
2. **誠實標註限制** — VBA 動態 range 解不出來就標 `has_dynamic_range: true`,不假裝解出來。
3. **輸出結構穩定** — JSON/CSV schema 是契約,版本要鎖定,讓下游可以信任。
4. **單檔可重入** — 同一份 .xlsm 跑兩次結果一樣 (deterministic)。
5. **每個 phase 獨立可驗收** — 不必整包做完才能 review。

## Status

Vision locked. See `PHASE_PLAN.md` for execution.
