"""Typed data models for comparison inputs, configuration, and outputs."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Literal

MarginFormula = Literal["auto", "price_difference_over_comparison", "gross_margin"]
DuplicatePolicy = Literal["report", "fail"]
SheetName = str | int | None


@dataclass(frozen=True, slots=True)
class WorkbookInput:
    """Local Excel workbook input descriptor."""

    path: Path
    sheet_name: SheetName = None


@dataclass(frozen=True, slots=True)
class ColumnConfig:
    """Column names used to match and compare workbook rows."""

    baseline_product_code_column: str
    comparison_product_code_column: str
    baseline_price_column: str
    comparison_price_column: str
    cost_column: str | None = None


@dataclass(frozen=True, slots=True)
class FormulaConfig:
    """Formula and rounding settings for calculated report metrics."""

    margin_formula: MarginFormula = "auto"
    rounding_decimals: int = 2


@dataclass(frozen=True, slots=True)
class RuntimeOptions:
    """Runtime behavior options for workbook loading and row matching."""

    baseline_sheet_name: SheetName = None
    comparison_sheet_name: SheetName = None
    duplicate_policy: DuplicatePolicy = "report"


@dataclass(frozen=True, slots=True)
class AppConfig:
    """Validated application configuration loaded from YAML."""

    columns: ColumnConfig
    formulas: FormulaConfig
    runtime: RuntimeOptions
    source_path: Path


@dataclass(frozen=True, slots=True)
class InvalidRow:
    """Row-level data-quality issue found during input normalization."""

    row_number: int
    column_name: str
    reason: str
    raw_value: object
    raw_values: dict[str, object]


@dataclass(frozen=True, slots=True)
class ComparisonResult:
    """Report-ready comparison tables and summary counts."""

    matched_rows: object
    missing_in_baseline: object
    missing_in_comparison: object
    duplicates: object
    invalid_rows: object
    summary: dict[str, int]
