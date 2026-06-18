"""Command-line entrypoint for the Excel sales automation tool."""

from __future__ import annotations

import argparse
import logging
from pathlib import Path

from excel_sales_automation import __version__
from excel_sales_automation.calculations import CalculationError
from excel_sales_automation.comparison import ComparisonError
from excel_sales_automation.config import ConfigError, load_config
from excel_sales_automation.excel_io import ExcelInputError
from excel_sales_automation.logging_config import configure_logging
from excel_sales_automation.normalization import NormalizationError
from excel_sales_automation.reporting import ReportExportError, validate_report_output_path
from excel_sales_automation.workflow import run_excel_comparison_workflow, summary_lines

logger = logging.getLogger(__name__)

_USER_FACING_ERRORS = (
    ConfigError,
    ExcelInputError,
    NormalizationError,
    ComparisonError,
    CalculationError,
    ReportExportError,
)


def build_parser() -> argparse.ArgumentParser:
    """Build the CLI argument parser."""
    parser = argparse.ArgumentParser(
        prog="excel-sales-compare",
        description="Compare local Excel sales/product files and export a formatted report.",
    )
    parser.add_argument("--version", action="version", version=f"%(prog)s {__version__}")

    subparsers = parser.add_subparsers(dest="command")
    compare_parser = subparsers.add_parser(
        "compare", help="Run the full Excel comparison workflow."
    )
    compare_parser.add_argument(
        "--baseline-file", required=True, help="Baseline Excel workbook path."
    )
    compare_parser.add_argument(
        "--comparison-file", required=True, help="Comparison Excel workbook path."
    )
    compare_parser.add_argument("--output-file", required=True, help="Output .xlsx report path.")
    compare_parser.add_argument(
        "--config",
        default="config/default_config.yaml",
        help="Path to the YAML configuration file.",
    )
    compare_parser.add_argument("--baseline-sheet", help="Override baseline sheet name or index.")
    compare_parser.add_argument(
        "--comparison-sheet", help="Override comparison sheet name or index."
    )
    compare_parser.add_argument(
        "--log-level", default=None, help="Console log level, e.g. INFO or DEBUG."
    )
    compare_parser.add_argument(
        "--debug", action="store_true", help="Show tracebacks for unexpected errors."
    )

    check_parser = subparsers.add_parser("check-config", help="Validate configuration and exit.")
    check_parser.add_argument(
        "--config",
        default="config/default_config.yaml",
        help="Path to the YAML configuration file.",
    )
    check_parser.add_argument(
        "--log-level", default=None, help="Console log level, e.g. INFO or DEBUG."
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    """Run the CLI."""
    parser = build_parser()
    args = parser.parse_args(argv)
    configure_logging(args.log_level if hasattr(args, "log_level") else None)

    if args.command == "check-config":
        config = _load_config_or_error(parser, args.config)
        print(f"Configuration OK: {config.source_path}")
        return 0

    if args.command == "compare":
        return _run_compare_command(parser, args)

    parser.print_help()
    return 0


def _run_compare_command(parser: argparse.ArgumentParser, args: argparse.Namespace) -> int:
    config = _load_config_or_error(parser, args.config)
    baseline_file = _existing_file_or_error(parser, args.baseline_file, "baseline file")
    comparison_file = _existing_file_or_error(parser, args.comparison_file, "comparison file")
    output_file = validate_report_output_path(args.output_file)

    try:
        report_path, result = run_excel_comparison_workflow(
            baseline_file=baseline_file,
            comparison_file=comparison_file,
            output_file=output_file,
            config=config,
            baseline_sheet=_parse_sheet_override(args.baseline_sheet),
            comparison_sheet=_parse_sheet_override(args.comparison_sheet),
        )
    except _USER_FACING_ERRORS as exc:
        if args.debug:
            raise
        parser.error(str(exc))
    except Exception as exc:
        if args.debug:
            raise
        logger.exception("Unexpected workflow failure")
        parser.error(f"Unexpected error. Re-run with --debug for details: {exc}")

    print(f"Report written to: {report_path}")
    for line in summary_lines(result.summary):
        print(line)
    return 0


def _load_config_or_error(parser: argparse.ArgumentParser, config_path: str) -> object:
    try:
        return load_config(config_path)
    except ConfigError as exc:
        parser.error(str(exc))


def _existing_file_or_error(parser: argparse.ArgumentParser, path: str, label: str) -> Path:
    candidate = Path(path)
    if not candidate.exists():
        parser.error(f"{label} does not exist: {candidate}")
    if not candidate.is_file():
        parser.error(f"{label} is not a file: {candidate}")
    return candidate


def _parse_sheet_override(value: str | None) -> str | int | None:
    if value is None:
        return None
    return int(value) if value.isdigit() else value


if __name__ == "__main__":
    raise SystemExit(main())
