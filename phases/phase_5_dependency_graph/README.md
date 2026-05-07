# Phase 5: Dependency Graph

## 目標

整合 Phase 3 (公式) 與 Phase 4 (VBA) 的引用資訊,用 networkx 建立 cell-level DAG。
產出 `09_dependencies.csv` (邊清單) 與 `10_dependency_graph.json` (完整圖)。

## 為什麼這個 phase 是核心

這是使用者最在意的能力 — 「**連動態也要**」。
跑完 Phase 5 之後可以回答:
- 改 `Params!A1` 會影響哪些 cell?(`nx.descendants(G, "Params!A1")`)
- `Output!Z1` 是從哪些 input 算出來的?(`nx.ancestors(G, "Output!Z1")`)
- 有哪些循環引用?(`list(nx.simple_cycles(G))`)
- 有哪些孤島公式? (`[n for n in G if G.in_degree(n)==0 and G.nodes[n]["type"]=="formula_cell"]`)

## 模組

```
src/xlsm_archaeologist/
├── analyzers/
│   ├── dependency_analyzer.py        # 主協調者
│   ├── dependency_graph_builder.py   # 建 NetworkX DiGraph
│   ├── cycle_detector.py             # 循環偵測
│   └── orphan_detector.py            # 孤島偵測
└── models/
    └── dependency.py                 # DependencyEdge, DependencyGraph
```

## 圖節點類型

```python
class NodeType(str, Enum):
    INPUT_CELL = "input_cell"          # 沒有公式、被引用過的 cell
    FORMULA_CELL = "formula_cell"      # 含公式
    OUTPUT_CELL = "output_cell"        # 暫時不分,先全部叫 formula_cell
    NAMED_RANGE = "named_range"        # named range (作為中介節點)
    VBA_PROCEDURE = "vba_procedure"    # VBA Sub/Function
```

## 邊類型

```python
class DependencyVia(str, Enum):
    FORMULA = "formula"               # cell A 的公式引用 cell B
    VBA_READ_WRITE = "vba_read_write" # VBA procedure 讀 X 寫 Y
    VALIDATION = "validation"         # cell 的 validation list 來源
    NAMED_RANGE = "named_range"       # 透過 named range 中介
```

## 建圖步驟

```python
def build_graph(formulas, vba_procedures, named_ranges, validations) -> DiGraph:
    G = nx.DiGraph()

    # Step 1: 加 named_range nodes
    for nr in named_ranges:
        G.add_node(f"_named:{nr.range_name}", node_type="named_range",
                   refers_to=nr.refers_to)

    # Step 2: 加 vba_procedure nodes
    for proc in vba_procedures:
        G.add_node(f"_vba:{proc.module_name}.{proc.procedure_name}",
                   node_type="vba_procedure")

    # Step 3: 處理公式 — 為每個 referenced cell 加 source→target 邊
    for formula in formulas:
        target = formula.qualified_address
        G.add_node(target, node_type="formula_cell")

        # 公式內每個 referenced cell
        for ref in formula.referenced_cells:
            source = f"{ref.sheet}!{ref.address}"
            if source not in G:
                G.add_node(source, node_type="input_cell")
            G.add_edge(source, target, via="formula", formula_id=formula.formula_id)

        # 公式內每個 named range
        for nr_name in formula.referenced_named_ranges:
            G.add_edge(f"_named:{nr_name}", target, via="named_range",
                       formula_id=formula.formula_id)

    # Step 4: 處理 VBA reads/writes
    for proc in vba_procedures:
        proc_node = f"_vba:{proc.module_name}.{proc.procedure_name}"

        # VBA 讀的 cell → 該 procedure
        for r in proc.reads:
            source = f"{r.sheet}!{r.range}"
            G.add_node(source, node_type="input_cell", maybe=True)
            G.add_edge(source, proc_node, via="vba_read_write")

        # VBA procedure → 寫的 cell
        for w in proc.writes:
            target = f"{w.sheet}!{w.range}"
            G.add_node(target, node_type="formula_cell", maybe=True)
            G.add_edge(proc_node, target, via="vba_read_write")

    # Step 5: 處理 validation list (列表來源也是依賴)
    for v in validations:
        if v.validation_type == "list" and v.formula1.startswith("="):
            ...

    return G
```

## Range 引用展開原則

這個重要,**不展開**到單一 cell:

- 公式 `=SUM(A1:A10)` → 加一條邊 `Sheet!A1:A10 → target`
- **不**加 10 條邊到每個 cell
- 這避免 graph 爆炸 (一個 `SUM(A:A)` 就會炸出百萬條邊)

下游若需要展開,可以用 NetworkX API 自己做:遇到 range 形態的 node 才展開。

## 循環偵測

```python
cycles = list(nx.simple_cycles(G))
# 過濾長度 ≥ 2 的 cycle (排除自指)
real_cycles = [c for c in cycles if len(c) >= 2]
```

每個 cycle 在輸出中:
```json
{
  "cycle_id": 1,
  "length": 3,
  "nodes": ["A!B1", "A!B2", "A!B3"],
  "edges_via": ["formula", "formula", "formula"]
}
```

## 孤島偵測

```python
orphans = [
    n for n in G.nodes()
    if G.in_degree(n) == 0
    and G.nodes[n]["node_type"] == "formula_cell"
]
```

含義:這個 cell 有公式,但沒有任何其他公式或 VBA 引用它。
通常是廢棄的計算 — 重構時可以砍。

## 回填 `is_referenced`

Phase 2 留下的 `is_referenced=false` 在這裡更新:

```python
for cell in cells:
    cell.is_referenced = G.in_degree(cell.qualified_address) > 0
```

更新後重寫 `04_cells.csv`。

## Graph 序列化

用 NetworkX 的 node-link format:

```python
import networkx as nx
data = nx.node_link_data(G, edges="edges")
# data 是 {"directed": True, "graph": {...}, "nodes": [...], "edges": [...]}
```

加上自己的統計欄位 (node_count, edge_count, has_cycles, ...) 寫進 JSON。

## 已知限制

- ⚠ Range 引用以「整段 range」為節點,不展開個別 cell
- ⚠ VBA 動態 range 以 procedure 與 sheet 之間的粗粒度邊表示
- ⚠ Graph 大時 JSON 檔可能很大 (10MB+) — 可接受,因為下游會 stream 解析

## 驗收

見 `acceptance.md`。
