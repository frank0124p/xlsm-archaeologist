"""Build an AST from openpyxl formula tokens."""

from __future__ import annotations

from openpyxl.formula.tokenizer import Token

from xlsm_archaeologist.models.formula import (
    AstNode,
    FunctionNode,
    NamedRangeNode,
    OperandNode,
    OperatorNode,
    RangeNode,
    UnparsableNode,
)
from xlsm_archaeologist.utils.logging import get_logger

logger = get_logger(__name__)

# openpyxl Token.type values (verified against openpyxl source)
_T_OPERAND = "OPERAND"
_T_FUNC = "FUNC"
_T_SEP = "SEP"
_T_OP_INFIX = "OPERATOR-INFIX"
_T_OP_PREFIX = "OPERATOR-PREFIX"
_T_OP_POSTFIX = "OPERATOR-POSTFIX"

# Token.subtype values
_SUB_RANGE = "RANGE"
_SUB_TEXT = "TEXT"
_SUB_NUMBER = "NUMBER"
_SUB_LOGICAL = "LOGICAL"
_SUB_ERROR = "ERROR"
_SUB_OPEN = "OPEN"
_SUB_CLOSE = "CLOSE"
_SUB_ARG = "ARG"  # argument separator ","


def _is_range_address(value: str) -> bool:
    """Return True if the operand value looks like a cell/range reference."""
    v = value
    if "!" in v or ":" in v:
        return True
    # strip quotes around sheet names
    stripped = v.strip("'").lstrip("$")
    if not stripped:
        return False
    i = 0
    while i < len(stripped) and stripped[i].isalpha():
        i += 1
    if i == 0:
        return False
    j = i
    while j < len(stripped) and (stripped[j].isdigit() or stripped[j] == "$"):
        j += 1
    return j == len(stripped) and j > i


def _parse_range_token(value: str) -> RangeNode:
    """Parse a cell/range reference token into a RangeNode."""
    if "!" in value:
        sheet_part, addr_part = value.split("!", 1)
        sheet = sheet_part.strip("'$")
        return RangeNode(sheet=sheet, address=addr_part)
    return RangeNode(sheet=None, address=value)


class _Parser:
    """Recursive-descent parser over a flat token list.

    openpyxl tokens for =IF(A1>0,IF(B1>0,1,0),0) look like:
        FUNC OPEN 'IF('
        OPERAND RANGE 'A1'
        OPERATOR-INFIX '' '>'
        OPERAND NUMBER '0'
        SEP ARG ','
        FUNC OPEN 'IF('
        ...
        FUNC CLOSE ')'
        SEP ARG ','
        OPERAND NUMBER '0'
        FUNC CLOSE ')'
    """

    def __init__(self, tokens: list[Token]) -> None:
        self._tokens = tokens
        self._pos = 0

    def _peek(self) -> Token | None:
        if self._pos < len(self._tokens):
            return self._tokens[self._pos]
        return None

    def _consume(self) -> Token:
        tok = self._tokens[self._pos]
        self._pos += 1
        return tok

    def parse(self) -> AstNode:
        if not self._tokens:
            return UnparsableNode(raw="")
        node = self._parse_expr()
        return node

    def _parse_expr(self) -> AstNode:
        """Parse expression, handling binary infix operators."""
        left = self._parse_unary()
        while True:
            tok = self._peek()
            if tok is None or tok.type != _T_OP_INFIX:
                break
            op_tok = self._consume()
            right = self._parse_unary()
            left = OperatorNode(op=op_tok.value, left=left, right=right)
        return left

    def _parse_unary(self) -> AstNode:
        tok = self._peek()
        if tok is not None and tok.type in (_T_OP_PREFIX, _T_OP_POSTFIX):
            self._consume()
            operand = self._parse_primary()
            return OperatorNode(op=tok.value, left=None, right=operand)
        return self._parse_primary()

    def _parse_primary(self) -> AstNode:
        tok = self._peek()
        if tok is None:
            return UnparsableNode(raw="<eof>")

        # Function call: FUNC OPEN token starts it
        if tok.type == _T_FUNC and tok.subtype == _SUB_OPEN:
            return self._parse_function()

        # Operand: cell/range reference or literal
        if tok.type == _T_OPERAND:
            self._consume()
            if tok.subtype == _SUB_RANGE or _is_range_address(tok.value):
                return _parse_range_token(tok.value)
            if tok.subtype == _SUB_TEXT:
                return OperandNode(operand_type="text", value=tok.value)
            if tok.subtype == _SUB_NUMBER:
                return OperandNode(operand_type="number", value=tok.value)
            if tok.subtype == _SUB_LOGICAL:
                return OperandNode(operand_type="boolean", value=tok.value.upper())
            if tok.subtype == _SUB_ERROR:
                return OperandNode(operand_type="error", value=tok.value)
            # Unknown operand subtype — could be a named range
            return NamedRangeNode(name=tok.value)

        # Skip unexpected tokens gracefully
        self._consume()
        return UnparsableNode(raw=tok.value)

    def _parse_function(self) -> FunctionNode:
        """Parse FUNC OPEN ... args ... FUNC CLOSE."""
        open_tok = self._consume()  # FUNC OPEN e.g. 'IF('
        name = open_tok.value.rstrip("(").upper()

        # Collect argument token sub-lists, separated by SEP ARG tokens at depth 1
        depth = 1
        arg_tokens: list[Token] = []
        arg_groups: list[list[Token]] = []

        while self._pos < len(self._tokens):
            tok = self._tokens[self._pos]

            if tok.type == _T_FUNC and tok.subtype == _SUB_OPEN:
                depth += 1
                arg_tokens.append(tok)
                self._pos += 1

            elif tok.type == _T_FUNC and tok.subtype == _SUB_CLOSE:
                depth -= 1
                self._pos += 1
                if depth == 0:
                    # End of this function's argument list
                    arg_groups.append(arg_tokens)
                    arg_tokens = []
                    break
                arg_tokens.append(tok)

            elif tok.type == _T_SEP and tok.subtype == _SUB_ARG and depth == 1:
                # Argument separator at this function's level
                arg_groups.append(arg_tokens)
                arg_tokens = []
                self._pos += 1

            else:
                arg_tokens.append(tok)
                self._pos += 1

        # Parse each arg sub-list into an AstNode
        args: list[AstNode] = []
        for group in arg_groups:
            if group:
                args.append(_Parser(group).parse())
            # skip empty groups (trailing commas etc.)

        return FunctionNode(name=name, args=args)


def parse(tokens: list[Token]) -> AstNode:
    """Build an AST from a token list produced by the tokenizer.

    Returns an UnparsableNode if the token list is empty or parsing fails.

    Args:
        tokens: Flat list of openpyxl Token objects.

    Returns:
        Root AstNode of the parsed formula.
    """
    if not tokens:
        return UnparsableNode(raw="")
    try:
        return _Parser(tokens).parse()
    except Exception as exc:  # noqa: BLE001
        logger.debug("Parse error: %s", exc)
        raw = "".join(t.value for t in tokens)
        return UnparsableNode(raw=raw)
