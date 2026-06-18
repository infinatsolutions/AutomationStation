"""Tests for messy spreadsheet normalization."""

from __future__ import annotations

import math
from collections.abc import Iterator
from copy import deepcopy

from excel_sales_automation.normalization import (
    clean_numeric_value,
    normalize_product_code,
    normalize_sales_dataframe,
)


class FakeRow(dict):
    def to_dict(self) -> dict[str, object]:
        return dict(self)


class AtIndexer:
    def __init__(self, frame: FakeFrame) -> None:
        self.frame = frame

    def __setitem__(self, key: tuple[int, str], value: object) -> None:
        index, column = key
        self.frame.rows[index][column] = value


class FakeFrame:
    def __init__(self, rows: list[dict[str, object]]) -> None:
        self.rows = rows
        self.at = AtIndexer(self)

    def copy(self, deep: bool = True) -> FakeFrame:
        return FakeFrame(deepcopy(self.rows) if deep else list(self.rows))

    def iterrows(self) -> Iterator[tuple[int, FakeRow]]:
        for index, row in enumerate(self.rows):
            yield index, FakeRow(row)

    def __len__(self) -> int:
        return len(self.rows)

    @property
    def columns(self) -> list[str]:
        return list(self.rows[0]) if self.rows else []


def test_normalize_product_code_strips_whitespace() -> None:
    assert normalize_product_code("  ABC-123  ") == "ABC-123"


def test_normalize_product_code_preserves_leading_zero_text() -> None:
    assert normalize_product_code("00123") == "00123"


def test_normalize_product_code_converts_safe_integer_float() -> None:
    assert normalize_product_code(123.0) == "123"


def test_normalize_product_code_returns_none_for_missing_values() -> None:
    assert normalize_product_code("") is None
    assert normalize_product_code(" nan ") is None
    assert normalize_product_code("None") is None
    assert normalize_product_code(float("nan")) is None


def test_clean_numeric_value_accepts_common_numeric_values() -> None:
    assert clean_numeric_value(10) == 10.0
    assert clean_numeric_value(10.5) == 10.5
    assert clean_numeric_value("1,234.50") == 1234.5
    assert clean_numeric_value("$1,234.50") == 1234.5
    assert clean_numeric_value("($12.30)") == -12.3


def test_clean_numeric_value_returns_none_for_blank_or_invalid_values() -> None:
    assert clean_numeric_value("") is None
    assert clean_numeric_value("not a price") is None
    assert clean_numeric_value(True) is None


def test_normalize_sales_dataframe_returns_clean_copy_and_invalid_rows() -> None:
    original = FakeFrame(
        [
            {"product_code": " 001 ", "price": "$1,200.50", "cost": "800"},
            {"product_code": "", "price": "10", "cost": ""},
            {"product_code": "003", "price": "bad", "cost": "bad cost"},
            {"product_code": "004", "price": "15", "cost": "$5.00"},
        ]
    )

    result = normalize_sales_dataframe(
        original,
        product_code_column="product_code",
        price_column="price",
        cost_column="cost",
    )

    assert original.rows[0]["product_code"] == " 001 "
    assert result.data.rows[0]["product_code"] == "001"
    assert result.data.rows[0]["price"] == 1200.5
    assert result.data.rows[0]["cost"] == 800.0
    assert math.isnan(result.data.rows[2]["price"])
    assert math.isnan(result.data.rows[2]["cost"])
    assert len(result.invalid_rows) == 3
    assert {issue.column_name for issue in result.invalid_rows} == {"product_code", "price", "cost"}
    assert [issue.row_number for issue in result.invalid_rows] == [3, 4, 4]
