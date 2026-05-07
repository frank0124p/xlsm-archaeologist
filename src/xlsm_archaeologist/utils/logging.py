"""Unified logger factory using rich's RichHandler."""

from __future__ import annotations

import logging

from rich.logging import RichHandler

_configured = False


def _configure_root(level: str = "info") -> None:
    """Configure the root logger with RichHandler once."""
    global _configured
    if _configured:
        return
    logging.basicConfig(
        level=level.upper(),
        format="%(message)s",
        datefmt="[%X]",
        handlers=[RichHandler(rich_tracebacks=True, markup=True)],
    )
    _configured = True


def get_logger(name: str, level: str = "info") -> logging.Logger:
    """Return a named logger backed by RichHandler.

    Args:
        name: Module name, typically ``__name__``.
        level: Log level string (default: ``"info"``).

    Returns:
        Configured Logger instance.
    """
    _configure_root(level)
    return logging.getLogger(name)
