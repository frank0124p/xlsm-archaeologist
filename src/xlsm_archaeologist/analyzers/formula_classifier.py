"""Classify a formula AST into one of the seven FormulaCategory values."""

from __future__ import annotations

from xlsm_archaeologist.models.formula import AstNode, FormulaCategory, FunctionNode, OperatorNode

LOOKUP_FUNCS: frozenset[str] = frozenset(
    {
        "VLOOKUP", "HLOOKUP", "XLOOKUP", "LOOKUP",
        "INDEX", "MATCH", "XMATCH",
        "CHOOSEROWS", "CHOOSECOLS",
        "FILTER", "UNIQUE", "SORT", "SORTBY",
    }
)

BRANCH_FUNCS: frozenset[str] = frozenset(
    {
        "IF", "IFS", "SWITCH", "CHOOSE",
        "IFERROR", "IFNA",
        "AND", "OR", "NOT", "XOR",
    }
)

AGGR_FUNCS: frozenset[str] = frozenset(
    {
        "SUM", "SUMIF", "SUMIFS", "SUMPRODUCT",
        "COUNT", "COUNTA", "COUNTBLANK", "COUNTIF", "COUNTIFS",
        "AVERAGE", "AVERAGEIF", "AVERAGEIFS",
        "MAX", "MAXIFS", "MIN", "MINIFS",
        "MEDIAN", "MODE", "MODE.SNGL", "MODE.MULT",
        "STDEV", "STDEV.S", "STDEV.P", "VAR", "VAR.S", "VAR.P",
        "AGGREGATE", "SUBTOTAL",
        "LARGE", "SMALL", "RANK", "RANK.EQ", "RANK.AVG",
    }
)

TEXT_FUNCS: frozenset[str] = frozenset(
    {
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
)


def _collect_functions(node: AstNode) -> list[str]:
    """Recursively collect all function names from the AST."""
    names: list[str] = []
    if isinstance(node, FunctionNode):
        names.append(node.name.upper())
        for arg in node.args:
            names.extend(_collect_functions(arg))
    elif isinstance(node, OperatorNode):
        if node.left is not None:
            names.extend(_collect_functions(node.left))
        names.extend(_collect_functions(node.right))
    return names


def _is_pure_reference(node: AstNode) -> bool:
    """Return True if the formula is a bare cell/range ref or named range with no operators."""
    from xlsm_archaeologist.models.formula import NamedRangeNode, RangeNode

    return isinstance(node, (RangeNode, NamedRangeNode))


def classify(ast: AstNode) -> FormulaCategory:
    """Classify a formula AST into one of the seven FormulaCategory values.

    Classification rules:
    - Pure cell/named-range reference → 'reference'
    - Two or more of {lookup, branch, aggregate, text} → 'mixed'
    - Else pick the single matching category, or 'compute' as default

    Args:
        ast: Root AST node from the parser.

    Returns:
        FormulaCategory string.
    """
    if _is_pure_reference(ast):
        return "reference"

    funcs = set(_collect_functions(ast))

    has_lookup = bool(funcs & LOOKUP_FUNCS)
    has_branch = bool(funcs & BRANCH_FUNCS)
    has_aggr = bool(funcs & AGGR_FUNCS)
    has_text = bool(funcs & TEXT_FUNCS)

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
    return "compute"
