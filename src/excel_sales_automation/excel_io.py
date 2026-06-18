"""Excel file input/output boundaries."""

from __future__ import annotations

from pathlib import Path


def ensure_excel_path(path: str | Path) -> Path:
    """Validate that a path looks like a supported Excel workbook path."""
    workbook_path = Path(path)
    if workbook_path.suffix.lower() not in {".xlsx", ".xlsm", ".xls"}:
        raise ValueError(f"Unsupported Excel file extension: {workbook_path.suffix}")
    return workbook_path
