"""Command-line entrypoint for the Excel sales automation tool."""

from __future__ import annotations

import argparse
import logging
from pathlib import Path

from excel_sales_automation import __version__
from excel_sales_automation.config import ConfigError, load_config
from excel_sales_automation.logging_config import configure_logging

logger = logging.getLogger(__name__)


def build_parser() -> argparse.ArgumentParser:
    """Build the CLI argument parser."""
    parser = argparse.ArgumentParser(
        prog="excel-sales-compare",
        description="Compare local Excel sales/product files and export a report.",
    )
    parser.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    parser.add_argument(
        "--config",
        default="config/default_config.yaml",
        help="Path to the YAML configuration file. Workflow execution is not implemented yet.",
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    """Run the placeholder CLI."""
    configure_logging()
    parser = build_parser()
    args = parser.parse_args(argv)

    try:
        config = load_config(Path(args.config))
    except ConfigError as exc:
        parser.error(str(exc))

    logger.info("Loaded configuration from %s", config.source_path)
    logger.info("Excel comparison workflow is not implemented in the scaffold phase.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
