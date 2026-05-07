"""Global settings loaded from environment variables or .env file."""

from __future__ import annotations

from pydantic_settings import BaseSettings


class Settings(BaseSettings):
    """Global configuration for xlsm-archaeologist.

    Fields are read from environment variables (prefix: XLSM_) or a .env file.
    CLI flags override these at runtime.
    """

    max_formula_depth: int = 20
    """Maximum AST nesting depth before truncation."""

    log_level: str = "info"
    """Logging verbosity: debug / info / warning / error."""

    model_config = {"env_prefix": "XLSM_"}
