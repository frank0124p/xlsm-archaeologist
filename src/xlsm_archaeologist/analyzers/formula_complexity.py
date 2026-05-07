"""Compute formula complexity metrics from an AST."""

from __future__ import annotations

from xlsm_archaeologist.models.formula import AstNode, CellRef, FunctionNode, OperatorNode


def _max_depth(node: AstNode, current: int = 0) -> int:
    """Return maximum FunctionNode nesting depth in the AST."""
    if isinstance(node, FunctionNode):
        child_depths = [_max_depth(arg, current + 1) for arg in node.args] or [current + 1]
        return max(child_depths)
    if isinstance(node, OperatorNode):
        depths = []
        if node.left is not None:
            depths.append(_max_depth(node.left, current))
        depths.append(_max_depth(node.right, current))
        return max(depths) if depths else current
    return current


def _count_functions(node: AstNode) -> int:
    """Return total number of FunctionNode instances (with duplicates)."""
    if isinstance(node, FunctionNode):
        return 1 + sum(_count_functions(arg) for arg in node.args)
    if isinstance(node, OperatorNode):
        total = 0
        if node.left is not None:
            total += _count_functions(node.left)
        total += _count_functions(node.right)
        return total
    return 0


def compute_complexity(
    ast: AstNode,
    referenced_cells: list[CellRef],
) -> tuple[int, int, int]:
    """Compute complexity metrics for a formula.

    Formula: score = nesting_depth * 2 + function_count + len(referenced_cells)

    Args:
        ast: Root AST node.
        referenced_cells: List of cell references extracted from the formula.

    Returns:
        Tuple of (nesting_depth, function_count, complexity_score).
    """
    depth = _max_depth(ast)
    func_count = _count_functions(ast)
    score = depth * 2 + func_count + len(referenced_cells)
    return depth, func_count, score
