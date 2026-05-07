"""Shared pytest fixtures for all test phases."""

from __future__ import annotations

from pathlib import Path

import pytest


@pytest.fixture()
def fixtures_dir() -> Path:
    """Return the path to tests/fixtures/."""
    return Path(__file__).parent / "fixtures"
