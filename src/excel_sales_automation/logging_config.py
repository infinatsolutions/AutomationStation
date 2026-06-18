"""Structured logging setup for the CLI."""

from __future__ import annotations

import logging
import os


def configure_logging() -> None:
    """Configure application logging from environment variables."""
    level_name = os.getenv("EXCEL_SALES_AUTOMATION_LOG_LEVEL", "INFO").upper()
    level = getattr(logging, level_name, logging.INFO)
    logging.basicConfig(
        level=level,
        format="%(asctime)s %(levelname)s %(name)s - %(message)s",
    )
