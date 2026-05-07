# Formula Categories

> 公式分類規則的權威文件。Phase 3 的 classifier 必須完全照這個實作。

## 七種分類

| Category | 含義 | 在新系統中可能對應 |
|---|---|---|
| `lookup` | 查表類:從另一個 range/table 找對應值 | `LookupRule` |
| `branch` | 條件分支:依條件選擇值 | `BranchRule` |
| `compute` | 純計算:四則運算、數學函式 | `ComputeRule` |
| `aggregate` | 聚合:SUM/COUNT/AVG/MAX/MIN 等 | `AggregateRule` |
| `text` | 文字處理:concat、left/right、format | `TextRule` |
| `reference` | 純引用:`=Sheet2!A1` 或單純複製其他 cell | `ReferenceRule` |
| `mixed` | 兩種以上類別混合 | 拆解後各自映射 |

## 函式分類清單

> Phase 3 的 classifier 用這幾個 set 做判定。

### LOOKUP_FUNCS

```python
LOOKUP_FUNCS = {
    "VLOOKUP", "HLOOKUP", "XLOOKUP", "LOOKUP",
    "INDEX", "MATCH", "XMATCH",
    "CHOOSEROWS", "CHOOSECOLS",
    "FILTER", "UNIQUE", "SORT", "SORTBY",
}
```

注意:
- `INDEX` 嚴格說可以做 lookup 也可以做 reference,但統計上 INDEX/MATCH 組合是 lookup 主力,
  歸 lookup
- `CHOOSE` 不是 lookup,它是 branch 性質

### BRANCH_FUNCS

```python
BRANCH_FUNCS = {
    "IF", "IFS", "SWITCH", "CHOOSE",
    "IFERROR", "IFNA",
    "AND", "OR", "NOT", "XOR",  # 邏輯運算 — 雖然不直接分支,但常與 IF 配對,單獨出現也算 branch
}
```

### AGGR_FUNCS

```python
AGGR_FUNCS = {
    "SUM", "SUMIF", "SUMIFS", "SUMPRODUCT",
    "COUNT", "COUNTA", "COUNTBLANK", "COUNTIF", "COUNTIFS",
    "AVERAGE", "AVERAGEIF", "AVERAGEIFS",
    "MAX", "MAXIFS", "MIN", "MINIFS",
    "MEDIAN", "MODE", "MODE.SNGL", "MODE.MULT",
    "STDEV", "STDEV.S", "STDEV.P", "VAR", "VAR.S", "VAR.P",
    "AGGREGATE", "SUBTOTAL",
    "LARGE", "SMALL", "RANK", "RANK.EQ", "RANK.AVG",
}
```

### TEXT_FUNCS

```python
TEXT_FUNCS = {
    "CONCAT", "CONCATENATE", "TEXTJOIN",
    "LEFT", "RIGHT", "MID", "LEN",
    "UPPER", "LOWER", "PROPER",
    "TRIM", "CLEAN", "SUBSTITUTE", "REPLACE",
    "TEXT", "VALUE", "NUMBERVALUE",
    "FIND", "SEARCH",
    "TEXTBEFORE", "TEXTAFTER", "TEXTSPLIT",
    "REPT", "T", "DOLLAR", "FIXED",
    "EXACT",
}
```

### COMPUTE_FUNCS (預設,不需明列;沒被前面四類涵蓋的數學/日期函式都算這類)

```python
COMPUTE_FUNCS = {
    # 數學
    "ABS", "SIGN", "ROUND", "ROUNDUP", "ROUNDDOWN", "TRUNC",
    "INT", "MOD", "QUOTIENT", "POWER", "SQRT", "EXP", "LN", "LOG", "LOG10",
    "FLOOR", "CEILING", "FLOOR.MATH", "CEILING.MATH",
    "PI", "DEGREES", "RADIANS",
    "SIN", "COS", "TAN", "ASIN", "ACOS", "ATAN", "ATAN2",
    # 日期
    "DATE", "DATEVALUE", "TIME", "TIMEVALUE",
    "YEAR", "MONTH", "DAY", "HOUR", "MINUTE", "SECOND",
    "WEEKDAY", "WEEKNUM",
    "EDATE", "EOMONTH",
    "DATEDIF", "DAYS", "DAYS360", "NETWORKDAYS", "WORKDAY",
    # 財務 (常見少數)
    "NPV", "IRR", "PMT", "FV", "PV", "RATE",
}
```

### VOLATILE_FUNCS

```python
VOLATILE_FUNCS = {
    "NOW", "TODAY",                  # 時間
    "RAND", "RANDBETWEEN", "RANDARRAY",  # 隨機
    "OFFSET", "INDIRECT",            # 動態引用
    "INFO", "CELL",                  # 系統資訊
}
```

### UNPARSABLE_FUNCS (放棄解析,標記 unparsable=true)

```python
UNPARSABLE_FUNCS = {
    "LAMBDA", "LET",                 # 自訂函式定義 — 語意需要更深的 binding 分析
    "BYROW", "BYCOL",                # 高階函式
    "REDUCE", "MAP", "SCAN",
    "MAKEARRAY",
}
```

對這些函式仍要記錄 `function_list`,但設 `is_parsable=false` 並加 warning。

## 分類演算法

```python
from xlsm_archaeologist.models.formula import FormulaCategory

def classify(funcs: set[str], has_only_refs: bool) -> FormulaCategory:
    """根據出現的函式集合決定分類。

    funcs: 公式中出現的所有函式名稱 (大寫)
    has_only_refs: 公式內容是否「純粹只有引用 + 字面值」(無函式無運算)
    """
    if not funcs:
        # 沒有任何函式 — 看是純引用還是純運算
        return "reference" if has_only_refs else "compute"

    has_lookup = bool(funcs & LOOKUP_FUNCS)
    has_branch = bool(funcs & BRANCH_FUNCS)
    has_aggr   = bool(funcs & AGGR_FUNCS)
    has_text   = bool(funcs & TEXT_FUNCS)

    type_count = sum([has_lookup, has_branch, has_aggr, has_text])

    if type_count >= 2:
        return "mixed"
    if has_lookup:
        return "lookup"
    if has_branch:
        return "branch"
    if has_aggr:
        return "aggregate"
    if has_text:
        return "text"

    # 全是 compute / volatile 類函式
    return "compute"
```

## 邊界 Case

### 純引用 vs compute

```
=A1                    → reference  (純複製)
=Sheet2!A1             → reference  (純跨表複製)
=A1+0                  → compute    (有運算,即使結果一樣)
=A1&""                 → text       (有 text 函式 / & 運算)
```

注意 `&` 運算子算 text 性質,但若公式只有 `&` 沒有其他 TEXT_FUNCS,
*目前簡化處理為 compute*。如果統計顯示需要,未來可加偵測 `&` operator 路徑。

### 邊界 Mixed

```
=IF(VLOOKUP(...)>0, ...)        → mixed (branch + lookup)
=IFERROR(VLOOKUP(...), 0)       → mixed (branch + lookup)
=SUM(IF(...))                   → mixed (aggregate + branch) [array formula]
=IF(A1>0, B1+C1, 0)             → branch (compute 不算進 type_count)
=IF(A1>0, "X", "Y")             → branch (text literals 不算 text 類)
```

關鍵原則:**type_count 只看 lookup/branch/aggregate/text 四大類**。
compute / reference 不參與 mixed 判定。

### 純邏輯函式

```
=AND(A1>0, B1>0)                → branch (含 AND)
=NOT(A1=B1)                     → branch
```

### Volatile + 普通函式

```
=NOW() + 1                      → compute (NOW 不影響分類,只標 is_volatile=true)
=OFFSET(A1, 0, 0)               → reference (OFFSET 是動態引用,不是 lookup)
                                  也設 is_volatile=true
=INDIRECT("A1")                 → reference (同上)
                                  也設 is_volatile=true
```

注意 `OFFSET` / `INDIRECT` 雖然產生引用,但不是「查表」,所以歸 `reference` 而不是 `lookup`。

### 公式裡含 named range 但無函式

```
=TaxRate                        → reference
=TaxRate * BasePrice            → compute
```

## function_list 格式

- 大寫
- 去重
- 字典序排序

範例:
- `=IF(VLOOKUP(A1,B:C,2,0)>0,IFERROR(SUM(D:D),0),"")` → `["IF", "IFERROR", "SUM", "VLOOKUP"]`
