"""Data normalization placeholders for product-code and numeric fields."""

from __future__ import annotations


def normalize_product_code(value: object) -> str:
    """Normalize a product code for matching."""
    return str(value).strip()
