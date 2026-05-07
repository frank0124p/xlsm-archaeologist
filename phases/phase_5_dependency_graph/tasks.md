# Phase 5 — Tasks

## Models

- [ ] 建立 `models/dependency.py`:
    - `NodeType` enum
    - `DependencyVia` enum
    - `GraphNode`(id, node_type, value_type?, in_degree, out_degree, is_referenced=True if in_degree>0)
    - `DependencyEdge`(dependency_id, source_qualified_address, target_qualified_address,
                      via, via_detail, is_cross_sheet)
    - `CycleRecord`(cycle_id, length, nodes, edges_via)
- [ ] commit: `feat(models): add dependency graph models`

## Builder

- [ ] `analyzers/dependency_graph_builder.py`:
    - `build_graph(formulas, vba_procedures, named_ranges, validations) -> DiGraph`
    - 五個步驟見 README
    - 用 NetworkX DiGraph,node attributes 與 edge attributes 都明確 typed
- [ ] commit: `feat(graph): build cell-level DAG`

## Cycle Detection

- [ ] `analyzers/cycle_detector.py`:
    - `detect_cycles(G: DiGraph) -> list[CycleRecord]`
    - `nx.simple_cycles` + 過濾自指
    - 每個 cycle 收集 nodes 與經過的 via
- [ ] commit: `feat(graph): detect cycles`

## Orphan Detection

- [ ] `analyzers/orphan_detector.py`:
    - `detect_orphans(G: DiGraph) -> list[str]`
    - 條件:formula_cell 且 in_degree==0
- [ ] commit: `feat(graph): detect orphan formulas`

## is_referenced 回填

- [ ] `analyzers/dependency_analyzer.py::backfill_is_referenced`:
    - 對每個 cell record,根據 G.in_degree 更新 is_referenced
    - 觸發重寫 `04_cells.csv`
- [ ] commit: `feat(graph): backfill is_referenced flag`

## Serialization

- [ ] 寫 `09_dependencies.csv`:
    - 從 G.edges() 展開
    - 每邊一 row,標記 is_cross_sheet
- [ ] 寫 `10_dependency_graph.json`:
    - `nx.node_link_data` + 統計欄位
- [ ] 寫 `cycles.json` (放 reports/) 與 `orphans.csv` (放 reports/)
- [ ] commit: `feat(graph): serialize graph and cycle/orphan reports`

## CLI 串接

- [ ] 修改 `cli.py`:
    - Phase 5 在 Phase 3+4 後執行
    - 處理 `--no-graph` flag
- [ ] commit: `feat(cli): wire dependency graph phase`

## Tests

- [ ] 建立 fixture `tests/fixtures/circular.xlsm`:
    - A1 = B1 + 1, B1 = A1 + 1 (簡單循環)
- [ ] 建立 fixture `tests/fixtures/orphan_formula.xlsm`:
    - 含一個沒人引用的公式 cell
- [ ] 建立 fixture `tests/fixtures/cross_sheet_chain.xlsm`:
    - Sheet1!A1 → Sheet2!B1 → Sheet3!C1 形成跨 sheet 鏈
- [ ] `tests/test_phase_5_graph.py`:
    - test_graph_basic_formula_dependency
    - test_graph_named_range_node
    - test_graph_vba_read_write_edges
    - test_cycle_detection_simple
    - test_cycle_detection_long
    - test_orphan_detection
    - test_is_referenced_backfilled
    - test_cross_sheet_marked
    - test_graph_serialization_roundtrip (寫出再用 nx.node_link_graph 載入,結構一致)
- [ ] commit: `test(graph): comprehensive graph tests`

## Quality

- [ ] `uv run pytest tests/test_phase_5_graph.py` 全 pass
- [ ] `uv run ruff check .` 零錯誤
- [ ] `uv run mypy src` 零錯誤

## 收尾

- [ ] 寫 `phase_5_summary.md`:
    - 用 cross_sheet_chain fixture 示範:`nx.descendants(G, "Sheet1!A1")` 應包含 Sheet3!C1
- [ ] **停下來等 review**
