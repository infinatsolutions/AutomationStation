"""Tests for financial metric calculations."""

from __future__ import annotations

import math

import pytest

from excel_sales_automation.calculations import CalculationError, calculate_financial_metrics


def records(table: object) -> list[dict[str, object]]:
    return table.to_dict("records")


def calculate(rows: list[dict[str, object]], **kwargs: object) -> list[dict[str, object]]:
    result = calculate_financial_metrics(
        rows,
        baseline_price_column="price",
        comparison_price_column="price",
        **kwargs,
    )
    return records(result)


def test_positive_price_difference() -> None:
    row = calculate([{"product_code": "001", "price_baseline": 10, "price_comparison": 12}])[0]

    assert row["price_difference"] == 2
    assert row["calculation_status"] == "ok"


def test_negative_price_difference() -> None:
    row = calculate([{"product_code": "001", "price_baseline": 12, "price_comparison": 10}])[0]

    assert row["price_difference"] == -2


def test_zero_price_difference() -> None:
    row = calculate([{"product_code": "001", "price_baseline": 10, "price_comparison": 10}])[0]

    assert row["price_difference"] == 0
    assert row["price_difference_rate"] == 0


def test_price_difference_rate_uses_baseline_price() -> None:
    row = calculate([{"product_code": "001", "price_baseline": 10, "price_comparison": 12}])[0]

    assert row["price_difference_rate"] == 0.2


def test_gross_margin_formula_with_comparison_cost() -> None:
    row = calculate(
        [
            {
                "product_code": "001",
                "price_baseline": 10,
                "price_comparison": 20,
                "cost_comparison": 8,
            }
        ],
        cost_column="cost",
        margin_formula="gross_margin",
    )[0]

    assert row["margin_rate"] == 0.6
    assert row["margin_formula_used"] == "gross_margin"


def test_gross_margin_formula_falls_back_to_baseline_cost() -> None:
    row = calculate(
        [{"product_code": "001", "price_baseline": 10, "price_comparison": 20, "cost_baseline": 9}],
        cost_column="cost",
        margin_formula="gross_margin",
    )[0]

    assert row["margin_rate"] == 0.55


def test_auto_formula_chooses_gross_margin_when_cost_exists() -> None:
    row = calculate(
        [
            {
                "product_code": "001",
                "price_baseline": 10,
                "price_comparison": 20,
                "cost_comparison": 8,
            }
        ],
        cost_column="cost",
        margin_formula="auto",
    )[0]

    assert row["margin_formula_used"] == "gross_margin"
    assert row["margin_rate"] == 0.6


def test_auto_formula_falls_back_when_cost_missing() -> None:
    row = calculate(
        [{"product_code": "001", "price_baseline": 10, "price_comparison": 20}],
        cost_column="cost",
        margin_formula="auto",
    )[0]

    assert row["margin_formula_used"] == "price_difference_over_comparison"
    assert row["margin_rate"] == 0.5


def test_price_difference_over_comparison_formula() -> None:
    row = calculate(
        [{"product_code": "001", "price_baseline": 10, "price_comparison": 20}],
        margin_formula="price_difference_over_comparison",
    )[0]

    assert row["margin_rate"] == 0.5


def test_division_by_zero_produces_nan_and_warning() -> None:
    row = calculate([{"product_code": "001", "price_baseline": 10, "price_comparison": 0}])[0]

    assert row["price_difference"] == -10
    assert math.isnan(row["margin_rate"])
    assert row["calculation_status"] == "warning"
    assert "comparison price is zero" in row["calculation_warning"]


def test_zero_baseline_price_warns_for_price_difference_rate() -> None:
    row = calculate([{"product_code": "001", "price_baseline": 0, "price_comparison": 10}])[0]

    assert math.isnan(row["price_difference_rate"])
    assert row["calculation_status"] == "warning"
    assert "baseline price is zero" in row["calculation_warning"]


def test_missing_baseline_price_marks_invalid() -> None:
    row = calculate([{"product_code": "001", "price_baseline": math.nan, "price_comparison": 10}])[
        0
    ]

    assert math.isnan(row["price_difference"])
    assert row["calculation_status"] == "invalid"
    assert "missing baseline price" in row["calculation_warning"]


def test_missing_comparison_price_marks_invalid() -> None:
    row = calculate([{"product_code": "001", "price_baseline": 10, "price_comparison": None}])[0]

    assert math.isnan(row["price_difference"])
    assert row["calculation_status"] == "invalid"
    assert "missing comparison price" in row["calculation_warning"]


def test_rounding_behavior() -> None:
    row = calculate(
        [{"product_code": "001", "price_baseline": 3, "price_comparison": 10}],
        rounding_decimals=3,
    )[0]

    assert row["price_difference_rate"] == 2.333
    assert row["margin_rate"] == 0.7


def test_gross_margin_missing_cost_warns() -> None:
    row = calculate(
        [{"product_code": "001", "price_baseline": 10, "price_comparison": 20}],
        cost_column="cost",
        margin_formula="gross_margin",
    )[0]

    assert math.isnan(row["margin_rate"])
    assert row["calculation_status"] == "warning"
    assert "cost is required" in row["calculation_warning"]


def test_input_rows_are_not_mutated() -> None:
    rows = [{"product_code": "001", "price_baseline": 10, "price_comparison": 12}]

    calculate(rows)

    assert rows == [{"product_code": "001", "price_baseline": 10, "price_comparison": 12}]


def test_missing_required_price_column_raises_error() -> None:
    with pytest.raises(CalculationError, match="price_baseline"):
        calculate_financial_metrics(
            [{"product_code": "001", "price_comparison": 12}],
            baseline_price_column="price",
            comparison_price_column="price",
        )


def test_empty_matched_rows_returns_empty_result() -> None:
    result = calculate_financial_metrics(
        [],
        baseline_price_column="price",
        comparison_price_column="price",
    )

    assert records(result) == []


def test_unsupported_margin_formula_raises_error() -> None:
    with pytest.raises(CalculationError, match="Unsupported margin formula"):
        calculate_financial_metrics(
            [{"product_code": "001", "price_baseline": 10, "price_comparison": 12}],
            baseline_price_column="price",
            comparison_price_column="price",
            margin_formula="unsupported",  # type: ignore[arg-type]
        )


def test_negative_rounding_decimals_raises_error() -> None:
    with pytest.raises(CalculationError, match="rounding_decimals"):
        calculate_financial_metrics(
            [{"product_code": "001", "price_baseline": 10, "price_comparison": 12}],
            baseline_price_column="price",
            comparison_price_column="price",
            rounding_decimals=-1,
        )
