# Phase 4 — Acceptance Checklist

## 功能正確性

- [ ] 跑 `vba_basic.xlsm`:
    - 1 個 module、1 個 procedure
    - reads 含 `{sheet, range="A1", via="explicit_range"}`
    - writes 含 `{sheet, range="B1", via="explicit_range"}`
- [ ] 跑 `vba_dynamic_range.xlsm`:
    - `has_dynamic_range = true`
    - dynamic_range_notes 至少一條,包含 line number 與原始 code 片段
- [ ] 跑 `vba_event_trigger.xlsm`:
    - triggers 含 `{event="Worksheet_Change", target="Sheet1!A:A"}`
- [ ] 跑 `vba_call_graph.xlsm`:
    - Main procedure 的 calls 含 `["SubA", "SubB"]`
    - SubA 的 calls 含 `["SubC"]`
- [ ] alias 追蹤:`Set rng = Range("A1"): rng.Value = 1` 正確識別為 write A1

## Schema

- [ ] `07_vba_modules.json` 與 `08_vba_procedures.json` 都含 `schema_version`
- [ ] 所有 enum 值在 DATA_MODEL.md 規定範圍內
- [ ] 加密 VBA 不會 crash,寫 warning + skip

## Determinism

- [ ] 同一份 fixture 跑兩次,VBA 輸出檔案無差異
- [ ] procedure 排序穩定 (依 module 然後 procedure_name)

## 程式品質

- [ ] `uv run pytest tests/test_phase_4_vba.py` 全 pass
- [ ] `uv run ruff check .` 零錯誤
- [ ] `uv run mypy src` 零錯誤
- [ ] vba 模組覆蓋率 ≥ 90%

## 不能出現

- [ ] **沒有** 嘗試執行 VBA 程式碼
- [ ] **沒有** 把動態 range 假裝解出來具體位置
- [ ] **沒有** 漏掉 has_dynamic_range 旗標 (這會嚴重誤導下游)
- [ ] **沒有** 把內建函式 (MsgBox / Range / Cells) 當成 procedure call 收進 calls
