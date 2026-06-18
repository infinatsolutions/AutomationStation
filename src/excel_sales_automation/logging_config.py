"""Structured logging setup for the CLI."""

from __future__ import annotations

import logging
import os
from pathlib import Path


def configure_logging(level_name: str | None = None, *, log_file: str | Path | None = None) -> None:
    """Configure application logging from CLI arguments or environment variables."""
    selected_level = (level_name or os.getenv("EXCEL_SALES_AUTOMATION_LOG_LEVEL", "INFO")).upper()
    level = getattr(logging, selected_level, logging.INFO)
    handlers: list[logging.Handler] = [logging.StreamHandler()]
    if log_file is not None:
        log_path = Path(log_file)
        log_path.parent.mkdir(parents=True, exist_ok=True)
        handlers.append(logging.FileHandler(log_path, encoding="utf-8"))
    logging.basicConfig(
        level=level,
        format="%(asctime)s %(levelname)s %(name)s - %(message)s",
        handlers=handlers,
        force=True,
    )
