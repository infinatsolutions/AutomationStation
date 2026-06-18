"""Tests for YAML configuration loading and validation."""

from __future__ import annotations

from pathlib import Path

import pytest

from excel_sales_automation.config import ConfigError, load_config
from excel_sales_automation.models import AppConfig


def write_config(path: Path, content: str) -> Path:
    path.write_text(content, encoding="utf-8")
    return path


def test_load_default_config() -> None:
    config = load_config()

    assert isinstance(config, AppConfig)
    assert config.columns.baseline_product_code_column == "product_code"
    assert config.columns.comparison_product_code_column == "product_code"
    assert config.columns.baseline_price_column == "baseline_price"
    assert config.columns.comparison_price_column == "comparison_price"
    assert config.columns.cost_column == "cost"
    assert config.formulas.margin_formula == "auto"
    assert config.formulas.rounding_decimals == 2
    assert config.runtime.baseline_sheet_name is None
    assert config.runtime.comparison_sheet_name is None
    assert config.runtime.duplicate_policy == "report"
    assert config.source_path == Path("config/default_config.yaml")


def test_load_custom_config(tmp_path: Path) -> None:
    config_path = write_config(
        tmp_path / "custom.yaml",
        """
columns:
  baseline_product_code_column: SKU
  comparison_product_code_column: sku_code
  baseline_price_column: Old Price
  comparison_price_column: New Price
  cost_column: Unit Cost
formulas:
  margin_formula: gross_margin
  rounding_decimals: 4
runtime:
  baseline_sheet_name: Baseline
  comparison_sheet_name: 1
  duplicate_policy: fail
""",
    )

    config = load_config(config_path)

    assert config.columns.baseline_product_code_column == "SKU"
    assert config.columns.comparison_product_code_column == "sku_code"
    assert config.columns.cost_column == "Unit Cost"
    assert config.formulas.margin_formula == "gross_margin"
    assert config.formulas.rounding_decimals == 4
    assert config.runtime.baseline_sheet_name == "Baseline"
    assert config.runtime.comparison_sheet_name == 1
    assert config.runtime.duplicate_policy == "fail"
    assert config.source_path == config_path


def test_optional_settings_use_safe_defaults(tmp_path: Path) -> None:
    config_path = write_config(
        tmp_path / "minimal.yaml",
        """
columns:
  baseline_product_code_column: product_code
  comparison_product_code_column: product_code
  baseline_price_column: baseline_price
  comparison_price_column: comparison_price
""",
    )

    config = load_config(config_path)

    assert config.columns.cost_column is None
    assert config.formulas.margin_formula == "auto"
    assert config.formulas.rounding_decimals == 2
    assert config.runtime.baseline_sheet_name is None
    assert config.runtime.comparison_sheet_name is None
    assert config.runtime.duplicate_policy == "report"


def test_invalid_margin_formula_raises_config_error(tmp_path: Path) -> None:
    config_path = write_config(
        tmp_path / "invalid-margin.yaml",
        """
columns:
  baseline_product_code_column: product_code
  comparison_product_code_column: product_code
  baseline_price_column: baseline_price
  comparison_price_column: comparison_price
formulas:
  margin_formula: unsupported
""",
    )

    with pytest.raises(ConfigError, match="formulas.margin_formula"):
        load_config(config_path)


def test_invalid_duplicate_policy_raises_config_error(tmp_path: Path) -> None:
    config_path = write_config(
        tmp_path / "invalid-duplicate-policy.yaml",
        """
columns:
  baseline_product_code_column: product_code
  comparison_product_code_column: product_code
  baseline_price_column: baseline_price
  comparison_price_column: comparison_price
runtime:
  duplicate_policy: ignore
""",
    )

    with pytest.raises(ConfigError, match="duplicate_policy"):
        load_config(config_path)


def test_missing_required_column_mapping_raises_config_error(tmp_path: Path) -> None:
    config_path = write_config(
        tmp_path / "missing-column.yaml",
        """
columns:
  baseline_product_code_column: product_code
  comparison_product_code_column: product_code
  comparison_price_column: comparison_price
""",
    )

    with pytest.raises(ConfigError, match="baseline_price_column"):
        load_config(config_path)


def test_missing_config_file_raises_config_error_with_path(tmp_path: Path) -> None:
    missing_path = tmp_path / "missing.yaml"

    with pytest.raises(ConfigError, match=str(missing_path)):
        load_config(missing_path)


def test_negative_rounding_decimals_raises_config_error(tmp_path: Path) -> None:
    config_path = write_config(
        tmp_path / "negative-rounding.yaml",
        """
columns:
  baseline_product_code_column: product_code
  comparison_product_code_column: product_code
  baseline_price_column: baseline_price
  comparison_price_column: comparison_price
formulas:
  rounding_decimals: -1
""",
    )

    with pytest.raises(ConfigError, match="rounding_decimals"):
        load_config(config_path)


def test_boolean_rounding_decimals_raises_config_error(tmp_path: Path) -> None:
    config_path = write_config(
        tmp_path / "boolean-rounding.yaml",
        """
columns:
  baseline_product_code_column: product_code
  comparison_product_code_column: product_code
  baseline_price_column: baseline_price
  comparison_price_column: comparison_price
formulas:
  rounding_decimals: true
""",
    )

    with pytest.raises(ConfigError, match="rounding_decimals"):
        load_config(config_path)


def test_boolean_sheet_name_raises_config_error(tmp_path: Path) -> None:
    config_path = write_config(
        tmp_path / "boolean-sheet.yaml",
        """
columns:
  baseline_product_code_column: product_code
  comparison_product_code_column: product_code
  baseline_price_column: baseline_price
  comparison_price_column: comparison_price
runtime:
  baseline_sheet_name: true
""",
    )

    with pytest.raises(ConfigError, match="baseline_sheet_name"):
        load_config(config_path)


def test_empty_config_file_raises_config_error(tmp_path: Path) -> None:
    config_path = write_config(tmp_path / "empty.yaml", "\n")

    with pytest.raises(ConfigError, match="empty"):
        load_config(config_path)


def test_top_level_list_raises_config_error(tmp_path: Path) -> None:
    config_path = write_config(tmp_path / "list.yaml", "- columns\n- formulas\n")

    with pytest.raises(ConfigError, match="top-level mapping|Invalid YAML syntax"):
        load_config(config_path)


def test_non_mapping_columns_section_raises_config_error(tmp_path: Path) -> None:
    config_path = write_config(tmp_path / "bad-columns.yaml", "columns: product_code\n")

    with pytest.raises(ConfigError, match="columns"):
        load_config(config_path)


def test_blank_required_column_name_raises_config_error(tmp_path: Path) -> None:
    config_path = write_config(
        tmp_path / "blank-column.yaml",
        """
columns:
  baseline_product_code_column: "   "
  comparison_product_code_column: product_code
  baseline_price_column: baseline_price
  comparison_price_column: comparison_price
""",
    )

    with pytest.raises(ConfigError, match="baseline_product_code_column"):
        load_config(config_path)


def test_legacy_matching_duplicate_policy_is_supported(tmp_path: Path) -> None:
    config_path = write_config(
        tmp_path / "legacy-matching.yaml",
        """
columns:
  baseline_product_code_column: product_code
  comparison_product_code_column: product_code
  baseline_price_column: baseline_price
  comparison_price_column: comparison_price
matching:
  duplicate_policy: fail
""",
    )

    config = load_config(config_path)

    assert config.runtime.duplicate_policy == "fail"


def test_load_config_accepts_string_path(tmp_path: Path) -> None:
    config_path = write_config(
        tmp_path / "string-path.yaml",
        """
columns:
  baseline_product_code_column: product_code
  comparison_product_code_column: product_code
  baseline_price_column: baseline_price
  comparison_price_column: comparison_price
""",
    )

    config = load_config(str(config_path))

    assert config.source_path == config_path
