"""Excel file input/output boundaries."""

from __future__ import annotations

import importlib
import importlib.util
import logging
from collections.abc import Sequence
from pathlib import Path

from excel_sales_automation.models import SheetName

pd = importlib.import_module("pandas") if importlib.util.find_spec("pandas") is not None else None

logger = logging.getLogger(__name__)

SUPPORTED_READ_EXTENSIONS = {".xlsx", ".xlsm"}
LEGACY_EXTENSION = ".xls"


class ExcelInputError(ValueError):
    """Raised when an Excel workbook cannot be loaded."""


class MissingColumnError(ExcelInputError):
    """Raised when a workbook is missing required columns."""


def ensure_excel_path(path: str | Path) -> Path:
    """Validate that a path looks like a supported Excel workbook path."""
    workbook_path = Path(path)
    if workbook_path.suffix.lower() not in {*SUPPORTED_READ_EXTENSIONS, LEGACY_EXTENSION}:
        raise ExcelInputError(f"Unsupported Excel file extension: {workbook_path.suffix}")
    return workbook_path


def read_excel_workbook(
    path: str | Path,
    *,
    sheet_name: SheetName = None,
    required_columns: Sequence[str] = (),
    text_columns: Sequence[str] = (),
) -> pd.DataFrame:
    """Read an Excel workbook into a DataFrame and validate required columns."""
    workbook_path = ensure_excel_path(path)
    if not workbook_path.exists():
        raise ExcelInputError(f"Excel file does not exist: {workbook_path}")
    if not workbook_path.is_file():
        raise ExcelInputError(f"Excel path is not a file: {workbook_path}")
    if workbook_path.suffix.lower() == LEGACY_EXTENSION:
        raise ExcelInputError(
            "Legacy .xls files are not supported in this phase. "
            f"Please convert to .xlsx: {workbook_path}"
        )

    pandas_sheet_name = 0 if sheet_name is None else sheet_name
    if pd is None:
        raise ExcelInputError(
            "pandas is required to read Excel files. Install the project dependencies first."
        )

    dtype = {column: "string" for column in text_columns}
    logger.info("Reading Excel workbook %s sheet=%s", workbook_path, pandas_sheet_name)

    try:
        frame = pd.read_excel(
            workbook_path,
            sheet_name=pandas_sheet_name,
            engine="openpyxl",
            dtype=dtype or None,
        )
    except ValueError as exc:
        raise ExcelInputError(
            f"Excel sheet could not be loaded from {workbook_path}: {pandas_sheet_name}: {exc}"
        ) from exc
    except OSError as exc:
        raise ExcelInputError(f"Excel file could not be read: {workbook_path}: {exc}") from exc

    if not isinstance(frame, pd.DataFrame):
        raise ExcelInputError(f"Expected one worksheet from {workbook_path}, got multiple sheets.")

    validate_required_columns(frame, required_columns, workbook_path)
    return frame


def validate_required_columns(
    frame: pd.DataFrame,
    required_columns: Sequence[str],
    workbook_path: Path | None = None,
) -> None:
    """Validate that all required columns are present in a DataFrame."""
    missing_columns = [column for column in required_columns if column not in frame.columns]
    if not missing_columns:
        return

    location = f" in {workbook_path}" if workbook_path is not None else ""
    available_columns = [str(column) for column in frame.columns]
    raise MissingColumnError(
        f"Missing required columns{location}: {missing_columns}. "
        f"Available columns: {available_columns}."
    )
