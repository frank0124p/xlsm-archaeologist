"""Write CSV output files with UTF-8 BOM and fixed column ordering."""

from __future__ import annotations

import csv
import io
from pathlib import Path
from typing import Any, Sequence


def _to_csv_value(v: Any) -> str:
    """Convert a Python value to its CSV string representation.

    - bool → lowercase 'true' / 'false'
    - None → empty string
    - everything else → str()
    """
    if isinstance(v, bool):
        return "true" if v else "false"
    if v is None:
        return ""
    return str(v)


def write_csv(
    path: Path,
    records: Sequence[dict[str, Any]],
    columns: list[str],
) -> None:
    """Write *records* to a UTF-8-with-BOM CSV file.

    Args:
        path: Destination file path; parent directory must exist.
        records: Sequence of dicts, one per row.  Extra keys beyond
            *columns* are silently ignored; missing keys produce empty strings.
        columns: Ordered list of column names — this defines both the header
            row and the column order in every data row.
    """
    path.parent.mkdir(parents=True, exist_ok=True)

    # Build in-memory first so we can write BOM + content atomically
    buf = io.StringIO()
    writer = csv.DictWriter(
        buf,
        fieldnames=columns,
        extrasaction="ignore",
        lineterminator="\n",
        quoting=csv.QUOTE_MINIMAL,
    )
    writer.writeheader()
    for record in records:
        writer.writerow({col: _to_csv_value(record.get(col)) for col in columns})

    with path.open("w", encoding="utf-8-sig", newline="") as fh:
        fh.write(buf.getvalue())
