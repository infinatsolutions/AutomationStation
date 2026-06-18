"""End-to-end local Excel comparison workflow orchestration."""

from __future__ import annotations

import logging
from dataclasses import replace
from pathlib import Path
from typing import Any

from excel_sales_automation.calculations import calculate_financial_metrics
from excel_sales_automation.comparison import compare_sales_dataframes
from excel_sales_automation.excel_io import read_excel_workbook
from excel_sales_automation.models import AppConfig, ComparisonResult, SheetName
from excel_sales_automation.normalization import normalize_sales_dataframe
from excel_sales_automation.reporting import export_excel_report

logger = logging.getLogger(__name__)


def run_excel_comparison_workflow(
    *,
    baseline_file: str | Path,
    comparison_file: str | Path,
    output_file: str | Path,
    config: AppConfig,
    baseline_sheet: SheetName = None,
    comparison_sheet: SheetName = None,
) -> tuple[Path, ComparisonResult]:
    """Run the complete local Excel comparison workflow and export a report."""
    columns = config.columns
    baseline_sheet_name = (
        config.runtime.baseline_sheet_name if baseline_sheet is None else baseline_sheet
    )
    comparison_sheet_name = (
        config.runtime.comparison_sheet_name if comparison_sheet is None else comparison_sheet
    )

    logger.info("Reading baseline workbook: %s", baseline_file)
    baseline_raw = read_excel_workbook(
        baseline_file,
        sheet_name=baseline_sheet_name,
        required_columns=[columns.baseline_product_code_column, columns.baseline_price_column],
        text_columns=[columns.baseline_product_code_column],
    )

    logger.info("Reading comparison workbook: %s", comparison_file)
    comparison_raw = read_excel_workbook(
        comparison_file,
        sheet_name=comparison_sheet_name,
        required_columns=[columns.comparison_product_code_column, columns.comparison_price_column],
        text_columns=[columns.comparison_product_code_column],
    )

    logger.info("Normalizing baseline rows")
    baseline_normalized = normalize_sales_dataframe(
        baseline_raw,
        product_code_column=columns.baseline_product_code_column,
        price_column=columns.baseline_price_column,
        cost_column=columns.cost_column,
    )

    logger.info("Normalizing comparison rows")
    comparison_normalized = normalize_sales_dataframe(
        comparison_raw,
        product_code_column=columns.comparison_product_code_column,
        price_column=columns.comparison_price_column,
        cost_column=columns.cost_column,
    )

    logger.info("Comparing product rows")
    comparison_result = compare_sales_dataframes(
        baseline_normalized.data,
        comparison_normalized.data,
        baseline_product_code_column=columns.baseline_product_code_column,
        comparison_product_code_column=columns.comparison_product_code_column,
        duplicate_policy=config.runtime.duplicate_policy,
        baseline_invalid_rows=baseline_normalized.invalid_rows,
        comparison_invalid_rows=comparison_normalized.invalid_rows,
    )

    logger.info("Calculating financial metrics")
    calculated_matched_rows = calculate_financial_metrics(
        comparison_result.matched_rows,
        baseline_price_column=columns.baseline_price_column,
        comparison_price_column=columns.comparison_price_column,
        cost_column=columns.cost_column,
        margin_formula=config.formulas.margin_formula,
        rounding_decimals=config.formulas.rounding_decimals,
    )
    report_result = replace(comparison_result, matched_rows=calculated_matched_rows)

    logger.info("Exporting report: %s", output_file)
    report_path = export_excel_report(
        report_result,
        output_file,
        input_file_names=[Path(baseline_file).name, Path(comparison_file).name],
        config_path=config.source_path,
    )
    logger.info("Report written to %s", report_path)
    return report_path, report_result


def summary_lines(summary: dict[str, Any]) -> list[str]:
    """Format workflow summary counts for console output."""
    ordered_keys = [
        "baseline_rows",
        "comparison_rows",
        "matched_count",
        "missing_in_baseline_count",
        "missing_in_comparison_count",
        "duplicate_count",
        "invalid_row_count",
    ]
    return [f"{key}: {summary.get(key, 0)}" for key in ordered_keys]
