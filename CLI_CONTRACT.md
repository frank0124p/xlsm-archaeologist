# CLI Contract

> CLI 的所有指令、參數、輸出與 exit code。這是契約。

## 指令總覽

```
xlsm-archaeologist [OPTIONS] COMMAND [ARGS]...

Commands:
  analyze       對一份 .xlsm 做完整考古分析
  inspect       快速檢視 .xlsm 概況 (不寫檔)
  version       顯示版本
```

## `analyze` — 主指令

```bash
xlsm-archaeologist analyze INPUT_PATH [OPTIONS]
```

### 參數

| Argument / Option | Type | Required | Default | Description |
|---|---|---|---|---|
| `INPUT_PATH` | path | yes | — | 要分析的 .xlsm 檔案路徑 |
| `--output`, `-o` | path | no | `./archaeology_output` | 輸出資料夾 (不存在會自動建立) |
| `--phases` | str | no | `all` | 跑哪些 phase,逗號分隔 (如 `2,3,5`),`all` 表示全跑 |
| `--no-vba` | flag | no | false | 跳過 VBA 分析 (適用無 VBA 的 .xlsx) |
| `--no-graph` | flag | no | false | 跳過依賴圖 (大檔案省時用) |
| `--no-reports` | flag | no | false | 跳過報告生成 |
| `--max-formula-depth` | int | no | 20 | 公式 AST 解析最大巢狀深度,超過截斷 |
| `--log-level` | str | no | `info` | `debug` / `info` / `warning` / `error` |
| `--quiet`, `-q` | flag | no | false | 不顯示 progress bar |
| `--force` | flag | no | false | 輸出資料夾非空時強制覆寫 |

### 範例

```bash
# 最常用:完整分析
xlsm-archaeologist analyze ./complex_macro.xlsm -o ./out

# 只跑公式分析 (Phase 2 + 3)
xlsm-archaeologist analyze ./complex_macro.xlsm --phases 2,3

# 大檔案先不跑依賴圖
xlsm-archaeologist analyze ./huge.xlsm --no-graph

# CI 用 (不要 progress bar)
xlsm-archaeologist analyze ./complex_macro.xlsm -q
```

### 輸出

成功時:

```
✔ Phase 1: Skeleton check (0.1s)
✔ Phase 2: Extracted 32 sheets, 47 named ranges (2.3s)
✔ Phase 3: Analyzed 1834 formulas (4.1s)
✔ Phase 4: Parsed 6 VBA modules, 41 procedures (1.2s)
✔ Phase 5: Built dependency graph (3210 edges, 2 cycles) (1.8s)
✔ Phase 6: Generated 5 reports (0.4s)

📦 Output: ./out
   00_summary.json (complexity_score: 847, difficulty: high)
   ...
   reports/ (5 files)

⚠ 7 warnings — see 00_summary.json#warnings
```

### Exit Codes

| Code | Meaning |
|---|---|
| 0 | 成功 (即使有 warnings) |
| 1 | 致命錯誤 (檔案不存在、解析失敗、IO 錯誤) |
| 2 | 參數錯誤 (CLI 用法錯) |
| 3 | 輸出資料夾非空且未加 `--force` |

---

## `inspect` — 快速概況

```bash
xlsm-archaeologist inspect INPUT_PATH
```

不寫檔,只在 stdout 印一份 summary。給「先看這份檔大概什麼狀況」用。

### 範例輸出

```
📄 complex_macro.xlsm (1.2 MB, sha256: abc123...)

  Sheets:           32 (3 hidden, 1 very_hidden)
  Named ranges:     47
  Formulas:         1834
    └─ branches:    412 (深巢狀: 23)
    └─ lookups:     287
    └─ computes:    856
    └─ aggregates:  189
    └─ mixed:       90
  Validations:      89
  VBA modules:      6 (234 + 87 + ... lines)
  VBA procedures:   41 (8 with dynamic range)

  Estimated complexity: high
  Run `xlsm-archaeologist analyze` for full output.
```

---

## `version`

```bash
xlsm-archaeologist version
```

輸出:

```
xlsm-archaeologist 0.1.0
schema_version: 1.0
python: 3.12.x
openpyxl: 3.1.x
oletools: 0.60.x
```

---

## stderr 行為

- 所有 progress、informational message 走 stderr
- 真正的「資料輸出」(如 `inspect` 的 summary) 才走 stdout
- 這樣 `inspect | jq` 之類的 pipeline 不會被 progress 污染

## Logging

`--log-level debug` 時額外印出:

- 每個 sheet 抽取耗時
- 公式解析失敗的詳細資訊
- VBA 動態 range 的 line number

debug log 都走 stderr。
