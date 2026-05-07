"""Custom exception hierarchy for xlsm-archaeologist."""

from __future__ import annotations


class XlsmArchaeologistError(Exception):
    """Base exception for all xlsm-archaeologist errors."""


class InvalidFileError(XlsmArchaeologistError):
    """Raised when the input file is missing, unreadable, or not a valid .xlsm/.xlsx."""


class ExtractionError(XlsmArchaeologistError):
    """Raised when a fatal error occurs during Layer 1 extraction."""


class AnalysisError(XlsmArchaeologistError):
    """Raised when a fatal error occurs during Layer 2 analysis."""
