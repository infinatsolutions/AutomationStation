"""Excel report generation placeholders."""

from __future__ import annotations

from pathlib import Path


def validate_report_output_path(path: str | Path) -> Path:
    """Validate the destination path for a generated Excel report."""
    output_path = Path(path)
    if output_path.suffix.lower() != ".xlsx":
        raise ValueError("Report output path must use the .xlsx extension.")
    return output_path
