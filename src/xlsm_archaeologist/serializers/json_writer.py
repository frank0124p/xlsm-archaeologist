"""Write JSON output files with enforced schema conventions."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any


def write_json(path: Path, data: dict[str, Any], schema_version: str = "1.0") -> None:
    """Serialize *data* to a JSON file with project-standard formatting.

    The ``schema_version`` key is injected at the top level (before all other
    keys) so downstream readers can version-check without scanning the file.

    Args:
        path: Destination file path; parent directory must exist.
        data: Payload dict to serialize.  Must not already contain
            ``schema_version`` (it is added here).
        schema_version: Schema version string to inject.
    """
    payload: dict[str, Any] = {"schema_version": schema_version}
    payload.update(data)

    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8") as fh:
        json.dump(payload, fh, indent=2, sort_keys=True, ensure_ascii=False)
        fh.write("\n")
