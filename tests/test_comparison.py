"""Tests for product-code matching and comparison logic."""

from __future__ import annotations

import pytest

from excel_sales_automation.comparison import ComparisonError, compare_sales_dataframes
from excel_sales_automation.models import InvalidRow


def records(frame: object) -> list[dict[str, object]]:
    return frame.to_dict("records")


def test_compare_sales_dataframes_matches_exact_product_codes() -> None:
    result = compare_sales_dataframes(
        [{"product_code": "001", "price": 10.0}, {"product_code": "002", "price": 20.0}],
        [{"product_code": "001", "price": 11.0}, {"product_code": "002", "price": 21.0}],
        baseline_product_code_column="product_code",
        comparison_product_code_column="product_code",
    )

    matched = records(result.matched_rows)

    assert [row["product_code"] for row in matched] == ["001", "002"]
    assert matched[0]["price_baseline"] == 10.0
    assert matched[0]["price_comparison"] == 11.0
    assert result.summary["matched_count"] == 2


def test_compare_sales_dataframes_tracks_baseline_only_as_missing_in_comparison() -> None:
    result = compare_sales_dataframes(
        [{"product_code": "001", "price": 10.0}, {"product_code": "002", "price": 20.0}],
        [{"product_code": "001", "price": 11.0}],
        baseline_product_code_column="product_code",
        comparison_product_code_column="product_code",
    )

    missing = records(result.missing_in_comparison)

    assert missing == [
        {"product_code": "002", "product_code_baseline": "002", "price_baseline": 20.0}
    ]
    assert result.summary["missing_in_comparison_count"] == 1


def test_compare_sales_dataframes_tracks_comparison_only_as_missing_in_baseline() -> None:
    result = compare_sales_dataframes(
        [{"product_code": "001", "price": 10.0}],
        [{"product_code": "001", "price": 11.0}, {"product_code": "003", "price": 30.0}],
        baseline_product_code_column="product_code",
        comparison_product_code_column="product_code",
    )

    missing = records(result.missing_in_baseline)

    assert missing == [
        {"product_code": "003", "product_code_comparison": "003", "price_comparison": 30.0}
    ]
    assert result.summary["missing_in_baseline_count"] == 1


def test_duplicate_policy_report_continues_and_excludes_ambiguous_matches() -> None:
    result = compare_sales_dataframes(
        [{"product_code": "001", "price": 10.0}, {"product_code": "001", "price": 12.0}],
        [{"product_code": "001", "price": 11.0}, {"product_code": "002", "price": 20.0}],
        baseline_product_code_column="product_code",
        comparison_product_code_column="product_code",
        duplicate_policy="report",
    )

    assert records(result.matched_rows) == []
    assert [row["product_code"] for row in records(result.duplicates)] == ["001", "001"]
    assert records(result.missing_in_baseline) == [
        {"product_code": "002", "product_code_comparison": "002", "price_comparison": 20.0}
    ]
    assert result.summary["duplicate_count"] == 2


def test_duplicate_policy_fail_raises_clear_error() -> None:
    with pytest.raises(ComparisonError, match="baseline:001"):
        compare_sales_dataframes(
            [{"product_code": "001", "price": 10.0}, {"product_code": "001", "price": 12.0}],
            [{"product_code": "001", "price": 11.0}],
            baseline_product_code_column="product_code",
            comparison_product_code_column="product_code",
            duplicate_policy="fail",
        )


def test_missing_product_codes_are_invalid_and_not_matched() -> None:
    result = compare_sales_dataframes(
        [{"product_code": None, "price": 10.0}, {"product_code": "001", "price": 20.0}],
        [{"product_code": "", "price": 11.0}, {"product_code": "001", "price": 21.0}],
        baseline_product_code_column="product_code",
        comparison_product_code_column="product_code",
    )

    assert [row["product_code"] for row in records(result.matched_rows)] == ["001"]
    assert len(records(result.invalid_rows)) == 2
    assert result.summary["invalid_row_count"] == 2


def test_summary_counts_are_correct() -> None:
    result = compare_sales_dataframes(
        [
            {"product_code": "001", "price": 10.0},
            {"product_code": "002", "price": 20.0},
            {"product_code": "004", "price": 40.0},
            {"product_code": "004", "price": 41.0},
            {"product_code": None, "price": 50.0},
        ],
        [
            {"product_code": "001", "price": 11.0},
            {"product_code": "003", "price": 30.0},
        ],
        baseline_product_code_column="product_code",
        comparison_product_code_column="product_code",
        duplicate_policy="report",
    )

    assert result.summary == {
        "baseline_rows": 5,
        "comparison_rows": 2,
        "matched_count": 1,
        "missing_in_baseline_count": 1,
        "missing_in_comparison_count": 1,
        "duplicate_count": 2,
        "invalid_row_count": 1,
    }


def test_passed_invalid_rows_are_included_in_invalid_output() -> None:
    invalid = InvalidRow(
        row_number=2,
        column_name="price",
        reason="Invalid required price.",
        raw_value="bad",
        raw_values={"product_code": "999", "price": "bad"},
    )

    result = compare_sales_dataframes(
        [{"product_code": "001", "price": 10.0}],
        [{"product_code": "001", "price": 11.0}],
        baseline_product_code_column="product_code",
        comparison_product_code_column="product_code",
        baseline_invalid_rows=[invalid],
    )

    invalid_rows = records(result.invalid_rows)

    assert invalid_rows[0]["source"] == "baseline"
    assert invalid_rows[0]["reason"] == "Invalid required price."
    assert result.summary["invalid_row_count"] == 1


def test_duplicate_row_numbers_preserve_original_positions_after_missing_codes() -> None:
    result = compare_sales_dataframes(
        [
            {"product_code": None, "price": 5.0},
            {"product_code": "001", "price": 10.0},
            {"product_code": "001", "price": 12.0},
        ],
        [{"product_code": "002", "price": 20.0}],
        baseline_product_code_column="product_code",
        comparison_product_code_column="product_code",
        duplicate_policy="report",
    )

    assert [row["row_number"] for row in records(result.duplicates)] == [3, 4]
    assert records(result.invalid_rows)[0]["row_number"] == 2


def test_invalid_numeric_rows_can_still_match_by_product_code() -> None:
    invalid = InvalidRow(
        row_number=2,
        column_name="price",
        reason="Invalid required price.",
        raw_value="bad",
        raw_values={"product_code": "001", "price": "bad"},
    )

    result = compare_sales_dataframes(
        [{"product_code": "001", "price": float("nan")}],
        [{"product_code": "001", "price": 11.0}],
        baseline_product_code_column="product_code",
        comparison_product_code_column="product_code",
        baseline_invalid_rows=[invalid],
    )

    assert [row["product_code"] for row in records(result.matched_rows)] == ["001"]
    assert result.summary["invalid_row_count"] == 1
