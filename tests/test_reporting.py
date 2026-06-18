"""Tests for formatted Excel report export."""

from __future__ import annotations

import zipfile
from datetime import UTC, datetime
from pathlib import Path

import pytest

from excel_sales_automation.comparison import SimpleDataFrame
from excel_sales_automation.models import ComparisonResult
from excel_sales_automation.reporting import (
    REQUIRED_SHEETS,
    ReportExportError,
    export_excel_report,
    table_to_records,
    validate_report_output_path,
)


def make_result() -> ComparisonResult:
    return ComparisonResult(
        matched_rows=SimpleDataFrame(
            [
                {
                    "product_code": "001",
                    "price_baseline": 10.0,
                    "price_comparison": 12.0,
                    "price_difference": 2.0,
                    "price_difference_rate": 0.2,
                    "margin_rate": 0.5,
                    "calculation_status": "ok",
                    "calculation_warning": "",
                }
            ]
        ),
        missing_in_baseline=SimpleDataFrame([{"product_code": "002"}]),
        missing_in_comparison=SimpleDataFrame([{"product_code": "003"}]),
        duplicates=SimpleDataFrame([{"source": "baseline", "product_code": "004"}]),
        invalid_rows=SimpleDataFrame([{"source": "comparison", "reason": "Missing product code."}]),
        summary={
            "baseline_rows": 3,
            "comparison_rows": 3,
            "matched_count": 1,
            "missing_in_baseline_count": 1,
            "missing_in_comparison_count": 1,
            "duplicate_count": 1,
            "invalid_row_count": 1,
        },
    )


def workbook_xml(path: Path, member: str) -> str:
    with zipfile.ZipFile(path) as archive:
        return archive.read(member).decode("utf-8")


def test_validate_report_output_path_accepts_xlsx(tmp_path: Path) -> None:
    path = tmp_path / "report.xlsx"

    assert validate_report_output_path(path) == path


def test_validate_report_output_path_rejects_non_xlsx(tmp_path: Path) -> None:
    with pytest.raises(ReportExportError, match=".xlsx"):
        validate_report_output_path(tmp_path / "report.csv")


def test_table_to_records_supports_list_and_to_dict() -> None:
    assert table_to_records([{"a": 1}]) == [{"a": 1}]
    assert table_to_records(SimpleDataFrame([{"b": 2}])) == [{"b": 2}]


def test_export_excel_report_creates_workbook_and_required_sheets(tmp_path: Path) -> None:
    output_path = tmp_path / "nested" / "report.xlsx"

    result_path = export_excel_report(
        make_result(),
        output_path,
        input_file_names=["baseline.xlsx", "comparison.xlsx"],
        config_path="config/default_config.yaml",
        generated_at=datetime(2026, 6, 18, 12, 0, tzinfo=UTC),
    )

    assert result_path == output_path
    assert output_path.exists()
    xml = workbook_xml(output_path, "xl/workbook.xml")
    for sheet_name in REQUIRED_SHEETS:
        assert f'name="{sheet_name}"' in xml


def test_summary_sheet_contains_counts_formulas_and_metadata(tmp_path: Path) -> None:
    output_path = export_excel_report(
        make_result(),
        tmp_path / "report.xlsx",
        input_file_names=["baseline.xlsx", "comparison.xlsx"],
        config_path="config/default_config.yaml",
        generated_at=datetime(2026, 6, 18, 12, 0, tzinfo=UTC),
    )

    summary_xml = workbook_xml(output_path, "xl/worksheets/sheet1.xml")

    assert "baseline_rows" in summary_xml
    assert "matched_count" in summary_xml
    assert "price_difference = comparison_price - baseline_price" in summary_xml
    assert "Product rows are matched by normalized product_code" in summary_xml
    assert "baseline.xlsx" in summary_xml
    assert "config/default_config.yaml" in summary_xml


def test_matched_sheet_contains_calculated_columns(tmp_path: Path) -> None:
    output_path = export_excel_report(make_result(), tmp_path / "report.xlsx")

    matched_xml = workbook_xml(output_path, "xl/worksheets/sheet2.xml")

    assert "product_code" in matched_xml
    assert "price_difference" in matched_xml
    assert "price_difference_rate" in matched_xml
    assert "margin_rate" in matched_xml
    assert "calculation_status" in matched_xml


def test_export_excel_report_does_not_write_nan_numeric_values(tmp_path: Path) -> None:
    result = make_result()
    result.matched_rows.records[0]["margin_rate"] = float("nan")

    output_path = export_excel_report(result, tmp_path / "report.xlsx")

    matched_xml = workbook_xml(output_path, "xl/worksheets/sheet2.xml")
    assert "<v>nan</v>" not in matched_xml


def test_table_to_records_rejects_unsupported_table_type() -> None:
    with pytest.raises(ReportExportError, match="Unsupported report table type"):
        table_to_records(object())
