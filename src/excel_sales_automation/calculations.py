"""Metric calculation helpers."""

from __future__ import annotations


def calculate_price_difference(baseline_price: float, comparison_price: float) -> float:
    """Calculate the absolute price difference."""
    return comparison_price - baseline_price
