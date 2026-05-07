"""Extract top-level workbook metadata into a WorkbookRecord."""

from __future__ import annotations

import hashlib
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.workbook.workbook import Workbook

from xlsm_archaeologist.models.workbook import WorkbookRecord
from xlsm_archaeologist.utils.logging import get_logger

logger = get_logger(__name__)

_EXTERNAL_REF_MARKERS = ("[", "]")


def _sha256(path: Path) -> str:
    """Return hex SHA-256 digest of the file at *path*."""
    h = hashlib.sha256()
    with path.open("rb") as fh:
        for chunk in iter(lambda: fh.read(65536), b""):
            h.update(chunk)
    return h.hexdigest()


def _detect_external_links(wb: Workbook) -> bool:
    """Return True if any defined name or cell formula references an external file."""
    for name in wb.defined_names:
        dn = wb.defined_names[name]
        attr = getattr(dn, "attr_text", None) or getattr(dn, "value", "") or ""
        if "[" in str(attr):
            return True
    return False


def _iso_timestamp(dt: object) -> str | None:
    """Convert a datetime-like object to ISO-8601 string, or None."""
    if dt is None:
        return None
    if hasattr(dt, "isoformat"):
        result = dt.isoformat()
        return str(result)
    return str(dt)


def extract_workbook(path: Path) -> tuple[WorkbookRecord, Workbook]:
    """Load the workbook and build a WorkbookRecord.

    Args:
        path: Absolute path to the .xlsm / .xlsx file.

    Returns:
        Tuple of (WorkbookRecord, open Workbook) so callers can continue
        extracting sheets/cells without re-opening the file.

    Raises:
        FileNotFoundError: *path* does not exist.
        InvalidFileError: File cannot be opened by openpyxl.
    """
    logger.info("Opening workbook: %s", path)
    sha = _sha256(path)
    size = path.stat().st_size

    wb = load_workbook(
        filename=path,
        read_only=False,
        data_only=False,
        keep_vba=True,
    )

    # vba_archive is non-None for any file opened with keep_vba=True.
    # Real VBA presence requires xl/vbaProject.bin inside the archive.
    has_vba = wb.vba_archive is not None and "xl/vbaProject.bin" in wb.vba_archive.namelist()
    has_ext = _detect_external_links(wb)

    props = wb.properties
    record = WorkbookRecord(
        file_path=str(path),
        file_sha256=sha,
        size_bytes=size,
        has_vba=has_vba,
        has_external_links=has_ext,
        default_sheet=wb.active.title if wb.active else None,
        created=_iso_timestamp(props.created) if props else None,
        modified=_iso_timestamp(props.modified) if props else None,
        author=props.creator if props else None,
        last_modified_by=props.lastModifiedBy if props else None,
    )
    logger.info("Workbook metadata extracted (sha256=%s…)", sha[:8])
    return record, wb
