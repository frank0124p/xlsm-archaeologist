# Example Outputs

> 給 agent 比對用的「典型輸出」範例。**這些不是真實 .xlsm 跑出來的結果**,
> 是手寫的範例,展示輸出檔案應該長什麼樣。

## 用途

- Agent 在實作 serializer 時,可以拿這些檔案做 fixture 比對
- 人類 review 時,看真實輸出是否與這些範例「結構一致」
- 寫測試時,這些檔案可以作為 expected output

## 檔案

- `00_summary.example.json` — 一份小型 fixture 跑出的 summary
- `05_formulas.example.json` — 兩條公式的範例 (一條 compute、一條 mixed)
- `08_vba_procedures.example.json` — 兩個 procedure 範例 (一個有動態 range、一個 event handler)

## 注意

**這些檔案的具體 sha256 / line_count 等數值是虛構的**,只用來示範**結構**。
真實跑出來的數值會不同 — 那是正確的,不是 bug。

agent 寫測試時,要驗證的是:
- ✅ schema 結構符合
- ✅ 欄位名稱與型別正確
- ✅ enum 值在合法範圍內
- ❌ 不要用 hash 比對整檔 (數值會差)
