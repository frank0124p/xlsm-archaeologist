"""Extract inter-procedure call references from VBA source code."""

from __future__ import annotations

import re

_COMMENT_RE = re.compile(r"'.*$", re.MULTILINE)
_STRING_RE = re.compile(r'"[^"]*"')


def _strip_code(source: str) -> str:
    """Remove string literals and comments so we don't match keywords inside them."""
    # Strip string literals first, then comments
    source = _STRING_RE.sub('""', source)
    source = _COMMENT_RE.sub("", source)
    return source


def extract_calls(
    source_code: str,
    all_procedure_names: set[str],
) -> list[str]:
    """Return sorted list of known procedure names called in the given source.

    Looks for bare name references preceded by Call keyword or standalone
    identifier tokens that match known procedure names.

    Args:
        source_code: Procedure source code to scan.
        all_procedure_names: Set of all known procedure names in the VBA project.

    Returns:
        Sorted, deduplicated list of called procedure names.
    """
    clean = _strip_code(source_code)
    found: set[str] = set()

    for name in all_procedure_names:
        # Match as a word boundary call: Call FuncName(...) or standalone line
        pattern = re.compile(
            r"(?:Call\s+)?" + re.escape(name) + r"\s*(?:\(|$|\s*,|\s+[^=])",
            re.IGNORECASE,
        )
        if pattern.search(clean):
            found.add(name)

    return sorted(found)
