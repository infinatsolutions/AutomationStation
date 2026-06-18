"""Tests for CLI configuration validation behavior."""

from __future__ import annotations

from pathlib import Path

import pytest

from excel_sales_automation.cli import main


def test_cli_accepts_valid_config_path(tmp_path: Path) -> None:
    config_path = tmp_path / "config.yaml"
    config_path.write_text(
        """
columns:
  baseline_product_code_column: product_code
  comparison_product_code_column: product_code
  baseline_price_column: baseline_price
  comparison_price_column: comparison_price
""",
        encoding="utf-8",
    )

    assert main(["--config", str(config_path)]) == 0


def test_cli_rejects_missing_config_path(tmp_path: Path) -> None:
    missing_path = tmp_path / "missing.yaml"

    with pytest.raises(SystemExit) as exc_info:
        main(["--config", str(missing_path)])

    assert exc_info.value.code == 2


def test_cli_version_exits_successfully() -> None:
    with pytest.raises(SystemExit) as exc_info:
        main(["--version"])

    assert exc_info.value.code == 0
