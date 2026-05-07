"""Thin wrapper around openpyxl's formula tokenizer."""

from __future__ import annotations

from openpyxl.formula.tokenizer import Token, Tokenizer

from xlsm_archaeologist.utils.logging import get_logger

logger = get_logger(__name__)


def tokenize(formula_text: str) -> list[Token]:
    """Tokenize an Excel formula string into openpyxl Token objects.

    Handles formulas with or without the leading '='.

    Args:
        formula_text: Raw formula string, e.g. '=IF(A1>0,1,0)' or 'A1+B1'.

    Returns:
        List of openpyxl Token objects. Empty list if tokenization fails.
    """
    text = formula_text if formula_text.startswith("=") else f"={formula_text}"
    try:
        tok = Tokenizer(text)
        return list(tok.items)
    except Exception as exc:  # noqa: BLE001
        logger.debug("Tokenization failed for %r: %s", formula_text, exc)
        return []
