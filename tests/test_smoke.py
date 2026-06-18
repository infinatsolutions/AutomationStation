"""Smoke tests for the package scaffold."""

from __future__ import annotations

import excel_sales_automation


def test_version_string_exists() -> None:
    """The package exposes a non-empty version string."""
    assert isinstance(excel_sales_automation.__version__, str)
    assert excel_sales_automation.__version__
