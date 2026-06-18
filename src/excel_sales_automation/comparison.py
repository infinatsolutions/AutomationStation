"""Product-code matching and comparison logic."""

from __future__ import annotations

import importlib
import importlib.util
import logging
from collections import defaultdict
from collections.abc import Iterable, Mapping, Sequence
from dataclasses import dataclass
from typing import Any

from excel_sales_automation.models import ComparisonResult, DuplicatePolicy, InvalidRow

pd = importlib.import_module("pandas") if importlib.util.find_spec("pandas") is not None else None

logger = logging.getLogger(__name__)

BASELINE_SOURCE = "baseline"
COMPARISON_SOURCE = "comparison"


class ComparisonError(ValueError):
    """Raised when product comparison cannot be completed safely."""


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


def compare_sales_dataframes(
    baseline_frame: Any,
    comparison_frame: Any,
    *,
    baseline_product_code_column: str,
    comparison_product_code_column: str,
    duplicate_policy: DuplicatePolicy = "report",
    baseline_invalid_rows: Sequence[InvalidRow] = (),
    comparison_invalid_rows: Sequence[InvalidRow] = (),
) -> ComparisonResult:
    """Compare normalized baseline and comparison rows by product code."""
    if duplicate_policy not in {"report", "fail"}:
        raise ComparisonError(f"Unsupported duplicate policy: {duplicate_policy}")

    if _is_pandas_frame(baseline_frame) and _is_pandas_frame(comparison_frame):
        return _compare_with_pandas(
            baseline_frame,
            comparison_frame,
            baseline_product_code_column=baseline_product_code_column,
            comparison_product_code_column=comparison_product_code_column,
            duplicate_policy=duplicate_policy,
            baseline_invalid_rows=baseline_invalid_rows,
            comparison_invalid_rows=comparison_invalid_rows,
        )

    return _compare_records(
        _with_source_positions(_records_from_frame(baseline_frame)),
        _with_source_positions(_records_from_frame(comparison_frame)),
        baseline_product_code_column=baseline_product_code_column,
        comparison_product_code_column=comparison_product_code_column,
        duplicate_policy=duplicate_policy,
        baseline_invalid_rows=baseline_invalid_rows,
        comparison_invalid_rows=comparison_invalid_rows,
    )


def _compare_with_pandas(
    baseline_frame: Any,
    comparison_frame: Any,
    *,
    baseline_product_code_column: str,
    comparison_product_code_column: str,
    duplicate_policy: DuplicatePolicy,
    baseline_invalid_rows: Sequence[InvalidRow],
    comparison_invalid_rows: Sequence[InvalidRow],
) -> ComparisonResult:
    baseline_frame = _pandas_with_source_positions(baseline_frame)
    comparison_frame = _pandas_with_source_positions(comparison_frame)
    baseline_valid, baseline_missing = _split_pandas_missing_codes(
        baseline_frame, baseline_product_code_column, BASELINE_SOURCE
    )
    comparison_valid, comparison_missing = _split_pandas_missing_codes(
        comparison_frame, comparison_product_code_column, COMPARISON_SOURCE
    )

    baseline_duplicates = _pandas_duplicate_records(
        baseline_valid, baseline_product_code_column, BASELINE_SOURCE
    )
    comparison_duplicates = _pandas_duplicate_records(
        comparison_valid, comparison_product_code_column, COMPARISON_SOURCE
    )
    duplicate_records = baseline_duplicates + comparison_duplicates
    _raise_for_duplicates_if_needed(duplicate_records, duplicate_policy)

    duplicate_codes = {record["product_code"] for record in duplicate_records}
    baseline_clean = baseline_valid[
        ~baseline_valid[baseline_product_code_column].isin(duplicate_codes)
    ]
    comparison_clean = comparison_valid[
        ~comparison_valid[comparison_product_code_column].isin(duplicate_codes)
    ]

    merged = baseline_clean.merge(
        comparison_clean,
        left_on=baseline_product_code_column,
        right_on=comparison_product_code_column,
        how="outer",
        suffixes=("_baseline", "_comparison"),
        indicator=True,
    )
    merged["product_code"] = merged[baseline_product_code_column].combine_first(
        merged[comparison_product_code_column]
    )

    matched = merged[merged["_merge"] == "both"].copy()
    missing_in_baseline = merged[merged["_merge"] == "right_only"].copy()
    missing_in_comparison = merged[merged["_merge"] == "left_only"].copy()
    duplicates = _frame_from_records(duplicate_records)
    invalid_rows = _frame_from_records(
        _invalid_row_records(baseline_invalid_rows, comparison_invalid_rows)
        + _missing_code_records(baseline_missing, comparison_missing)
    )

    return _build_result(
        baseline_rows=len(baseline_frame),
        comparison_rows=len(comparison_frame),
        matched=matched,
        missing_in_baseline=missing_in_baseline,
        missing_in_comparison=missing_in_comparison,
        duplicates=duplicates,
        invalid_rows=invalid_rows,
    )


def _compare_records(
    baseline_records: list[dict[str, Any]],
    comparison_records: list[dict[str, Any]],
    *,
    baseline_product_code_column: str,
    comparison_product_code_column: str,
    duplicate_policy: DuplicatePolicy,
    baseline_invalid_rows: Sequence[InvalidRow],
    comparison_invalid_rows: Sequence[InvalidRow],
) -> ComparisonResult:
    baseline_valid, baseline_missing = _split_missing_codes(
        baseline_records, baseline_product_code_column, BASELINE_SOURCE
    )
    comparison_valid, comparison_missing = _split_missing_codes(
        comparison_records, comparison_product_code_column, COMPARISON_SOURCE
    )

    baseline_duplicates = _duplicate_records(
        baseline_valid, baseline_product_code_column, BASELINE_SOURCE
    )
    comparison_duplicates = _duplicate_records(
        comparison_valid, comparison_product_code_column, COMPARISON_SOURCE
    )
    duplicate_records = baseline_duplicates + comparison_duplicates
    _raise_for_duplicates_if_needed(duplicate_records, duplicate_policy)

    duplicate_codes = {record["product_code"] for record in duplicate_records}
    baseline_by_code = _unique_records_by_code(
        baseline_valid, baseline_product_code_column, duplicate_codes
    )
    comparison_by_code = _unique_records_by_code(
        comparison_valid, comparison_product_code_column, duplicate_codes
    )

    baseline_codes = set(baseline_by_code)
    comparison_codes = set(comparison_by_code)

    matched = [
        _merged_record(code, baseline_by_code[code], comparison_by_code[code])
        for code in sorted(baseline_codes & comparison_codes)
    ]
    missing_in_baseline = [
        _comparison_only_record(code, comparison_by_code[code])
        for code in sorted(comparison_codes - baseline_codes)
    ]
    missing_in_comparison = [
        _baseline_only_record(code, baseline_by_code[code])
        for code in sorted(baseline_codes - comparison_codes)
    ]
    invalid_rows = (
        _invalid_row_records(baseline_invalid_rows, comparison_invalid_rows)
        + baseline_missing
        + comparison_missing
    )

    return _build_result(
        baseline_rows=len(baseline_records),
        comparison_rows=len(comparison_records),
        matched=SimpleDataFrame(matched),
        missing_in_baseline=SimpleDataFrame(missing_in_baseline),
        missing_in_comparison=SimpleDataFrame(missing_in_comparison),
        duplicates=SimpleDataFrame(duplicate_records),
        invalid_rows=SimpleDataFrame(invalid_rows),
    )


def _build_result(
    *,
    baseline_rows: int,
    comparison_rows: int,
    matched: Any,
    missing_in_baseline: Any,
    missing_in_comparison: Any,
    duplicates: Any,
    invalid_rows: Any,
) -> ComparisonResult:
    summary = {
        "baseline_rows": baseline_rows,
        "comparison_rows": comparison_rows,
        "matched_count": len(matched),
        "missing_in_baseline_count": len(missing_in_baseline),
        "missing_in_comparison_count": len(missing_in_comparison),
        "duplicate_count": len(duplicates),
        "invalid_row_count": len(invalid_rows),
    }
    logger.info("Comparison summary: %s", summary)
    return ComparisonResult(
        matched_rows=matched,
        missing_in_baseline=missing_in_baseline,
        missing_in_comparison=missing_in_comparison,
        duplicates=duplicates,
        invalid_rows=invalid_rows,
        summary=summary,
    )


def _is_pandas_frame(frame: Any) -> bool:
    return pd is not None and isinstance(frame, pd.DataFrame)


def _records_from_frame(frame: Any) -> list[dict[str, Any]]:
    if isinstance(frame, list):
        return [dict(record) for record in frame]
    if hasattr(frame, "to_dict"):
        return [dict(record) for record in frame.to_dict("records")]
    if hasattr(frame, "rows"):
        return [dict(record) for record in frame.rows]
    raise ComparisonError("Unsupported frame type for comparison.")


def _with_source_positions(records: list[dict[str, Any]]) -> list[dict[str, Any]]:
    positioned: list[dict[str, Any]] = []
    for index, record in enumerate(records):
        positioned.append({"__row_index": index, "__row_number": index + 2, **record})
    return positioned


def _record_row_index(record: Mapping[str, Any], fallback: int) -> int:
    value = record.get("__row_index", fallback)
    return int(value) if isinstance(value, int) else fallback


def _record_row_number(record: Mapping[str, Any], fallback: int) -> int:
    value = record.get("__row_number", fallback)
    return int(value) if isinstance(value, int) else fallback


def _public_record(record: Mapping[str, Any]) -> dict[str, Any]:
    return {key: value for key, value in record.items() if not str(key).startswith("__")}


def _split_missing_codes(
    records: Iterable[Mapping[str, Any]], product_code_column: str, source: str
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    valid: list[dict[str, Any]] = []
    invalid: list[dict[str, Any]] = []
    for index, record in enumerate(records):
        product_code = record.get(product_code_column)
        row_index = _record_row_index(record, index)
        row_number = _record_row_number(record, row_index + 2)
        if product_code is None or str(product_code).strip() == "":
            invalid.append(
                {
                    "source": source,
                    "product_code": product_code,
                    "row_index": row_index,
                    "row_number": row_number,
                    "reason": "Missing product code.",
                    **{f"{key}_{source}": value for key, value in _public_record(record).items()},
                }
            )
        else:
            valid.append(dict(record))
    return valid, invalid


def _duplicate_records(
    records: Sequence[Mapping[str, Any]], product_code_column: str, source: str
) -> list[dict[str, Any]]:
    grouped: dict[str, list[tuple[int, Mapping[str, Any]]]] = defaultdict(list)
    for index, record in enumerate(records):
        grouped[str(record[product_code_column])].append((index, record))

    duplicates: list[dict[str, Any]] = []
    for product_code, group in sorted(grouped.items()):
        if len(group) <= 1:
            continue
        for fallback_index, record in group:
            row_index = _record_row_index(record, fallback_index)
            duplicates.append(
                {
                    "source": source,
                    "product_code": product_code,
                    "row_index": row_index,
                    "row_number": _record_row_number(record, row_index + 2),
                    **{f"{key}_{source}": value for key, value in _public_record(record).items()},
                }
            )
    return duplicates


def _raise_for_duplicates_if_needed(
    duplicate_records: Sequence[Mapping[str, Any]], duplicate_policy: DuplicatePolicy
) -> None:
    if duplicate_policy != "fail" or not duplicate_records:
        return
    duplicate_descriptions = [
        f"{record['source']}:{record['product_code']}@row{record['row_number']}"
        for record in duplicate_records
    ]
    raise ComparisonError(
        "Duplicate product codes found while duplicate_policy=fail: "
        + ", ".join(duplicate_descriptions)
    )


def _unique_records_by_code(
    records: Sequence[Mapping[str, Any]], product_code_column: str, excluded_codes: set[str]
) -> dict[str, dict[str, Any]]:
    return {
        str(record[product_code_column]): dict(record)
        for record in records
        if str(record[product_code_column]) not in excluded_codes
    }


def _merged_record(
    product_code: str, baseline_record: Mapping[str, Any], comparison_record: Mapping[str, Any]
) -> dict[str, Any]:
    return {
        "product_code": product_code,
        **{f"{key}_baseline": value for key, value in _public_record(baseline_record).items()},
        **{f"{key}_comparison": value for key, value in _public_record(comparison_record).items()},
    }


def _baseline_only_record(product_code: str, baseline_record: Mapping[str, Any]) -> dict[str, Any]:
    return {
        "product_code": product_code,
        **{f"{key}_baseline": value for key, value in _public_record(baseline_record).items()},
    }


def _comparison_only_record(
    product_code: str, comparison_record: Mapping[str, Any]
) -> dict[str, Any]:
    return {
        "product_code": product_code,
        **{f"{key}_comparison": value for key, value in _public_record(comparison_record).items()},
    }


def _invalid_row_records(
    baseline_invalid_rows: Sequence[InvalidRow], comparison_invalid_rows: Sequence[InvalidRow]
) -> list[dict[str, Any]]:
    return [
        _invalid_row_record(BASELINE_SOURCE, invalid_row) for invalid_row in baseline_invalid_rows
    ] + [
        _invalid_row_record(COMPARISON_SOURCE, invalid_row)
        for invalid_row in comparison_invalid_rows
    ]


def _invalid_row_record(source: str, invalid_row: InvalidRow) -> dict[str, Any]:
    return {
        "source": source,
        "product_code": invalid_row.raw_values.get("product_code"),
        "row_number": invalid_row.row_number,
        "column_name": invalid_row.column_name,
        "reason": invalid_row.reason,
        "raw_value": invalid_row.raw_value,
        **{f"{key}_{source}": value for key, value in invalid_row.raw_values.items()},
    }


def _frame_from_records(records: list[dict[str, Any]]) -> Any:
    return pd.DataFrame(records) if pd is not None else SimpleDataFrame(records)


def _pandas_with_source_positions(frame: Any) -> Any:
    positioned = frame.copy()
    positioned["__row_index"] = list(range(len(positioned)))
    positioned["__row_number"] = positioned["__row_index"] + 2
    return positioned


def _split_pandas_missing_codes(
    frame: Any, product_code_column: str, source: str
) -> tuple[Any, Any]:
    missing_mask = frame[product_code_column].isna() | (
        frame[product_code_column].astype(str).str.strip() == ""
    )
    missing = frame[missing_mask].copy()
    valid = frame[~missing_mask].copy()
    missing_records = _split_missing_codes(missing.to_dict("records"), product_code_column, source)[
        1
    ]
    return valid, missing_records


def _pandas_duplicate_records(
    frame: Any, product_code_column: str, source: str
) -> list[dict[str, Any]]:
    duplicate_frame = frame[frame[product_code_column].duplicated(keep=False)].copy()
    return _duplicate_records(duplicate_frame.to_dict("records"), product_code_column, source)


def _missing_code_records(
    baseline_missing: list[dict[str, Any]], comparison_missing: list[dict[str, Any]]
) -> list[dict[str, Any]]:
    return baseline_missing + comparison_missing
