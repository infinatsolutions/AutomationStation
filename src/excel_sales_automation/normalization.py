"""Data normalization helpers for messy sales spreadsheets."""

from __future__ import annotations

import logging
import math
import re
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from typing import Any

from excel_sales_automation.models import InvalidRow

logger = logging.getLogger(__name__)

_MISSING_STRINGS = {"", "nan", "none", "null", "<na>"}
_CURRENCY_PATTERN = re.compile(r"[$€£¥]\s*")


@dataclass(frozen=True, slots=True)
class NormalizationResult:
    """A normalized DataFrame plus row-level data-quality issues."""

    data: Any
    invalid_rows: list[InvalidRow]


class NormalizationError(ValueError):
    """Raised when input data cannot be normalized safely."""


def normalize_product_code(value: Any) -> str | None:
    """Normalize a product code while preserving leading zeros in text values."""
    if _is_missing(value):
        return None
    if isinstance(value, float) and value.is_integer():
        normalized = str(int(value))
    else:
        normalized = str(value).strip()
    if normalized.lower() in _MISSING_STRINGS:
        return None
    return normalized


def clean_numeric_value(value: Any) -> float | None:
    """Convert common messy Excel numeric values into floats."""
    if _is_missing(value):
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, int | float):
        numeric_value = float(value)
        return None if math.isnan(numeric_value) else numeric_value
    if isinstance(value, Decimal):
        return float(value)

    text = str(value).strip()
    if text.lower() in _MISSING_STRINGS:
        return None
    cleaned = _CURRENCY_PATTERN.sub("", text).replace(",", "").strip()
    if cleaned.startswith("(") and cleaned.endswith(")"):
        cleaned = f"-{cleaned[1:-1]}"

    try:
        return float(Decimal(cleaned))
    except InvalidOperation:
        return None


def normalize_sales_dataframe(
    frame: Any,
    *,
    product_code_column: str,
    price_column: str,
    cost_column: str | None = None,
) -> NormalizationResult:
    """Normalize sales rows without mutating the input DataFrame."""
    _validate_normalization_columns(frame, product_code_column, price_column, cost_column)
    normalized = frame.copy(deep=True)
    invalid_rows: list[InvalidRow] = []

    for index, row in frame.iterrows():
        excel_row_number = _excel_row_number(index)

        raw_product_code = row[product_code_column]
        product_code = normalize_product_code(raw_product_code)
        normalized.at[index, product_code_column] = product_code
        if product_code is None:
            invalid_rows.append(
                _invalid_row(
                    excel_row_number,
                    product_code_column,
                    "Missing product code.",
                    raw_product_code,
                    row,
                )
            )

        raw_price = row[price_column]
        price = clean_numeric_value(raw_price)
        normalized.at[index, price_column] = price if price is not None else math.nan
        if price is None:
            invalid_rows.append(
                _invalid_row(
                    excel_row_number, price_column, "Invalid required price.", raw_price, row
                )
            )

        if cost_column is not None and cost_column in frame.columns:
            raw_cost = row[cost_column]
            cost = clean_numeric_value(raw_cost)
            normalized.at[index, cost_column] = cost if cost is not None else math.nan
            if cost is None and not _is_blank_value(raw_cost):
                invalid_rows.append(
                    _invalid_row(
                        excel_row_number, cost_column, "Invalid optional cost.", raw_cost, row
                    )
                )

    logger.info("Normalized %s rows with %s invalid row issues", len(frame), len(invalid_rows))
    return NormalizationResult(data=normalized, invalid_rows=invalid_rows)


def _validate_normalization_columns(
    frame: Any,
    product_code_column: str,
    price_column: str,
    cost_column: str | None,
) -> None:
    columns = set(frame.columns)
    required_columns = [product_code_column, price_column]
    missing_columns = [column for column in required_columns if column not in columns]
    if missing_columns:
        raise NormalizationError(
            f"Cannot normalize data because required columns are missing: {missing_columns}. "
            f"Available columns: {sorted(str(column) for column in columns)}."
        )
    if cost_column is not None and cost_column not in columns:
        logger.info(
            "Optional cost column %s is not present; cost normalization skipped.", cost_column
        )


def _is_missing(value: Any) -> bool:
    if value is None:
        return True
    if isinstance(value, float) and math.isnan(value):
        return True
    if str(type(value)).endswith("NAType'>"):
        return True
    return False


def _is_blank_value(value: Any) -> bool:
    if _is_missing(value):
        return True
    return isinstance(value, str) and value.strip().lower() in _MISSING_STRINGS


def _excel_row_number(index: Any) -> int:
    return int(index) + 2 if isinstance(index, int) else 0


def _invalid_row(
    row_number: int,
    column_name: str,
    reason: str,
    raw_value: Any,
    row: Any,
) -> InvalidRow:
    return InvalidRow(
        row_number=row_number,
        column_name=column_name,
        reason=reason,
        raw_value=raw_value,
        raw_values=row.to_dict(),
    )
