"""Financial metric calculations for matched sales rows."""

from __future__ import annotations

import importlib
import importlib.util
import logging
import math
from dataclasses import dataclass
from typing import Any

from excel_sales_automation.models import MarginFormula

pd = importlib.import_module("pandas") if importlib.util.find_spec("pandas") is not None else None

logger = logging.getLogger(__name__)

STATUS_OK = "ok"
STATUS_WARNING = "warning"
STATUS_INVALID = "invalid"


class CalculationError(ValueError):
    """Raised when financial metrics cannot be calculated safely."""


@dataclass(frozen=True, slots=True)
class SimpleDataFrame:
    """Small records-backed DataFrame fallback for dependency-constrained tests."""

    records: list[dict[str, Any]]

    @property
    def columns(self) -> list[str]:
        columns: list[str] = []
        for record in self.records:
            for column in record:
                if column not in columns:
                    columns.append(column)
        return columns

    def __len__(self) -> int:
        return len(self.records)

    def to_dict(self, orient: str = "records") -> list[dict[str, Any]]:
        if orient != "records":
            raise ValueError("SimpleDataFrame only supports orient='records'.")
        return [dict(record) for record in self.records]


def calculate_financial_metrics(
    matched_rows: Any,
    *,
    baseline_price_column: str,
    comparison_price_column: str,
    cost_column: str | None = None,
    margin_formula: MarginFormula = "auto",
    rounding_decimals: int = 2,
) -> Any:
    """Add price and margin metrics to matched rows without mutating the input."""
    if margin_formula not in {"auto", "price_difference_over_comparison", "gross_margin"}:
        raise CalculationError(f"Unsupported margin formula: {margin_formula}")
    if rounding_decimals < 0:
        raise CalculationError("rounding_decimals must be non-negative.")

    if _is_pandas_frame(matched_rows):
        return _calculate_with_pandas(
            matched_rows,
            baseline_price_column=baseline_price_column,
            comparison_price_column=comparison_price_column,
            cost_column=cost_column,
            margin_formula=margin_formula,
            rounding_decimals=rounding_decimals,
        )

    return _calculate_records(
        _records_from_table(matched_rows),
        baseline_price_column=baseline_price_column,
        comparison_price_column=comparison_price_column,
        cost_column=cost_column,
        margin_formula=margin_formula,
        rounding_decimals=rounding_decimals,
    )


def _calculate_records(
    records: list[dict[str, Any]],
    *,
    baseline_price_column: str,
    comparison_price_column: str,
    cost_column: str | None,
    margin_formula: MarginFormula,
    rounding_decimals: int,
) -> SimpleDataFrame:
    baseline_column = f"{baseline_price_column}_baseline"
    comparison_column = f"{comparison_price_column}_comparison"
    if not records:
        return SimpleDataFrame([])
    _validate_record_columns(records, [baseline_column, comparison_column])
    calculated: list[dict[str, Any]] = []

    for record in records:
        output = dict(record)
        baseline_price = _to_number(record.get(baseline_column))
        comparison_price = _to_number(record.get(comparison_column))
        warnings: list[str] = []
        status = STATUS_OK

        if baseline_price is None:
            warnings.append("missing baseline price")
        if comparison_price is None:
            warnings.append("missing comparison price")
        if warnings:
            output.update(_empty_metrics(STATUS_INVALID, warnings))
            calculated.append(output)
            continue

        price_difference = comparison_price - baseline_price
        price_difference_rate = _safe_divide(price_difference, baseline_price)
        if price_difference_rate is None:
            warnings.append("baseline price is zero; price difference rate unavailable")

        margin_rate, formula_used, margin_warnings = _calculate_margin_rate(
            price_difference=price_difference,
            comparison_price=comparison_price,
            record=record,
            cost_column=cost_column,
            margin_formula=margin_formula,
        )
        warnings.extend(margin_warnings)
        if warnings:
            status = STATUS_WARNING

        output.update(
            {
                "price_difference": _round_or_nan(price_difference, rounding_decimals),
                "price_difference_rate": _round_or_nan(price_difference_rate, rounding_decimals),
                "margin_rate": _round_or_nan(margin_rate, rounding_decimals),
                "margin_formula_used": formula_used,
                "calculation_status": status,
                "calculation_warning": "; ".join(warnings),
            }
        )
        calculated.append(output)

    logger.info("Calculated financial metrics for %s matched rows", len(calculated))
    return SimpleDataFrame(calculated)


def _calculate_with_pandas(
    matched_rows: Any,
    *,
    baseline_price_column: str,
    comparison_price_column: str,
    cost_column: str | None,
    margin_formula: MarginFormula,
    rounding_decimals: int,
) -> Any:
    baseline_column = f"{baseline_price_column}_baseline"
    comparison_column = f"{comparison_price_column}_comparison"
    _validate_columns(matched_rows.columns, [baseline_column, comparison_column])

    if len(matched_rows) == 0:
        output = matched_rows.copy()
        for column in [
            "price_difference",
            "price_difference_rate",
            "margin_rate",
            "margin_formula_used",
            "calculation_status",
            "calculation_warning",
        ]:
            output[column] = []
        return output

    output = matched_rows.copy()
    baseline_price = pd.to_numeric(output[baseline_column], errors="coerce")
    comparison_price = pd.to_numeric(output[comparison_column], errors="coerce")
    output["price_difference"] = comparison_price - baseline_price
    output["price_difference_rate"] = output["price_difference"] / baseline_price

    records = output.to_dict("records")
    calculated = _calculate_records(
        records,
        baseline_price_column=baseline_price_column,
        comparison_price_column=comparison_price_column,
        cost_column=cost_column,
        margin_formula=margin_formula,
        rounding_decimals=rounding_decimals,
    )
    return pd.DataFrame(calculated.to_dict("records"))


def _calculate_margin_rate(
    *,
    price_difference: float,
    comparison_price: float,
    record: dict[str, Any],
    cost_column: str | None,
    margin_formula: MarginFormula,
) -> tuple[float | None, str, list[str]]:
    warnings: list[str] = []
    if comparison_price == 0:
        return (
            None,
            _formula_used_for_zero_price(margin_formula, cost_column, record),
            ["comparison price is zero; margin rate unavailable"],
        )

    if margin_formula == "price_difference_over_comparison":
        return price_difference / comparison_price, "price_difference_over_comparison", warnings

    cost = _find_cost(record, cost_column)
    if margin_formula == "gross_margin":
        if cost is None:
            return None, "gross_margin", ["cost is required for gross_margin"]
        return (comparison_price - cost) / comparison_price, "gross_margin", warnings

    if cost is not None:
        return (comparison_price - cost) / comparison_price, "gross_margin", warnings
    return price_difference / comparison_price, "price_difference_over_comparison", warnings


def _formula_used_for_zero_price(
    margin_formula: MarginFormula, cost_column: str | None, record: dict[str, Any]
) -> str:
    if margin_formula != "auto":
        return margin_formula
    return (
        "gross_margin"
        if _find_cost(record, cost_column) is not None
        else "price_difference_over_comparison"
    )


def _find_cost(record: dict[str, Any], cost_column: str | None) -> float | None:
    if cost_column is None:
        return None
    comparison_cost = _to_number(record.get(f"{cost_column}_comparison"))
    if comparison_cost is not None:
        return comparison_cost
    return _to_number(record.get(f"{cost_column}_baseline"))


def _validate_columns(columns: Any, required_columns: list[str]) -> None:
    missing = [column for column in required_columns if column not in columns]
    if missing:
        raise CalculationError(f"Missing required calculation columns: {missing}")


def _validate_record_columns(records: list[dict[str, Any]], required_columns: list[str]) -> None:
    available_columns = {column for record in records for column in record}
    _validate_columns(available_columns, required_columns)


def _records_from_table(table: Any) -> list[dict[str, Any]]:
    if isinstance(table, list):
        return [dict(record) for record in table]
    if hasattr(table, "to_dict"):
        return [dict(record) for record in table.to_dict("records")]
    if hasattr(table, "records"):
        return [dict(record) for record in table.records]
    raise CalculationError("Unsupported table type for financial calculations.")


def _is_pandas_frame(table: Any) -> bool:
    return pd is not None and isinstance(table, pd.DataFrame)


def _to_number(value: Any) -> float | None:
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    try:
        numeric = float(value)
    except (TypeError, ValueError):
        return None
    return None if math.isnan(numeric) else numeric


def _safe_divide(numerator: float, denominator: float) -> float | None:
    if denominator == 0:
        return None
    return numerator / denominator


def _round_or_nan(value: float | None, rounding_decimals: int) -> float:
    return math.nan if value is None else round(value, rounding_decimals)


def _empty_metrics(status: str, warnings: list[str]) -> dict[str, Any]:
    return {
        "price_difference": math.nan,
        "price_difference_rate": math.nan,
        "margin_rate": math.nan,
        "margin_formula_used": "",
        "calculation_status": status,
        "calculation_warning": "; ".join(warnings),
    }
