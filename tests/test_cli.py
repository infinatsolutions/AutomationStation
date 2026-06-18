"""Tests for CLI behavior."""

from __future__ import annotations

from pathlib import Path

import pytest

from excel_sales_automation import workflow
from excel_sales_automation.cli import main
from excel_sales_automation.models import ComparisonResult
from excel_sales_automation.reporting import ReportExportError


class FakeRow(dict):
    def to_dict(self) -> dict[str, object]:
        return dict(self)


class AtIndexer:
    def __init__(self, frame: FakeFrame) -> None:
        self.frame = frame

    def __setitem__(self, key: tuple[int, str], value: object) -> None:
        index, column = key
        self.frame.rows[index][column] = value


class FakeFrame:
    def __init__(self, rows: list[dict[str, object]]) -> None:
        self.rows = rows
        self.at = AtIndexer(self)

    @property
    def columns(self) -> list[str]:
        return list(self.rows[0]) if self.rows else []

    def copy(self, deep: bool = True) -> FakeFrame:
        return FakeFrame([dict(row) for row in self.rows])

    def iterrows(self):
        for index, row in enumerate(self.rows):
            yield index, FakeRow(row)

    def __len__(self) -> int:
        return len(self.rows)

    def to_dict(self, orient: str = "records") -> list[dict[str, object]]:
        return [dict(row) for row in self.rows]


def write_config(path: Path, product_column: str = "product_code") -> Path:
    path.write_text(
        f"""
columns:
  baseline_product_code_column: {product_column}
  comparison_product_code_column: {product_column}
  baseline_price_column: price
  comparison_price_column: price
  cost_column: cost
formulas:
  margin_formula: auto
  rounding_decimals: 2
runtime:
  duplicate_policy: report
""",
        encoding="utf-8",
    )
    return path


def test_cli_help_command_succeeds(capsys: pytest.CaptureFixture[str]) -> None:
    with pytest.raises(SystemExit) as exc_info:
        main(["--help"])

    assert exc_info.value.code == 0
    assert "compare" in capsys.readouterr().out


def test_cli_check_config_succeeds(tmp_path: Path, capsys: pytest.CaptureFixture[str]) -> None:
    config_path = write_config(tmp_path / "config.yaml")

    assert main(["check-config", "--config", str(config_path)]) == 0
    assert "Configuration OK" in capsys.readouterr().out


def test_cli_missing_input_file_returns_nonzero(tmp_path: Path) -> None:
    config_path = write_config(tmp_path / "config.yaml")
    comparison_path = tmp_path / "comparison.xlsx"
    comparison_path.write_bytes(b"placeholder")

    with pytest.raises(SystemExit) as exc_info:
        main(
            [
                "compare",
                "--baseline-file",
                str(tmp_path / "missing.xlsx"),
                "--comparison-file",
                str(comparison_path),
                "--output-file",
                str(tmp_path / "report.xlsx"),
                "--config",
                str(config_path),
            ]
        )

    assert exc_info.value.code == 2


def test_cli_end_to_end_creates_report_with_sheet_overrides(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path, capsys: pytest.CaptureFixture[str]
) -> None:
    config_path = write_config(tmp_path / "config.yaml")
    baseline_path = tmp_path / "baseline.xlsx"
    comparison_path = tmp_path / "comparison.xlsx"
    output_path = tmp_path / "nested" / "report.xlsx"
    baseline_path.write_bytes(b"placeholder")
    comparison_path.write_bytes(b"placeholder")
    calls: list[dict[str, object]] = []

    def fake_read_excel_workbook(path: object, **kwargs: object) -> FakeFrame:
        calls.append({"path": path, **kwargs})
        if Path(path) == baseline_path:
            return FakeFrame([{"product_code": "001", "price": "$10", "cost": "5"}])
        return FakeFrame([{"product_code": "001", "price": "$12", "cost": "6"}])

    monkeypatch.setattr(workflow, "read_excel_workbook", fake_read_excel_workbook)

    exit_code = main(
        [
            "compare",
            "--baseline-file",
            str(baseline_path),
            "--comparison-file",
            str(comparison_path),
            "--output-file",
            str(output_path),
            "--config",
            str(config_path),
            "--baseline-sheet",
            "Products",
            "--comparison-sheet",
            "2",
        ]
    )

    output = capsys.readouterr().out
    assert exit_code == 0
    assert output_path.exists()
    assert "Report written to" in output
    assert "matched_count: 1" in output
    assert calls[0]["sheet_name"] == "Products"
    assert calls[1]["sheet_name"] == 2


def test_cli_config_override_is_used(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    config_path = write_config(tmp_path / "config.yaml", product_column="SKU")
    baseline_path = tmp_path / "baseline.xlsx"
    comparison_path = tmp_path / "comparison.xlsx"
    baseline_path.write_bytes(b"placeholder")
    comparison_path.write_bytes(b"placeholder")
    required_columns_seen: list[list[str]] = []

    def fake_read_excel_workbook(path: object, **kwargs: object) -> FakeFrame:
        required_columns_seen.append(list(kwargs["required_columns"]))
        return FakeFrame([{"SKU": "001", "price": "10", "cost": "5"}])

    monkeypatch.setattr(workflow, "read_excel_workbook", fake_read_excel_workbook)

    assert (
        main(
            [
                "compare",
                "--baseline-file",
                str(baseline_path),
                "--comparison-file",
                str(comparison_path),
                "--output-file",
                str(tmp_path / "report.xlsx"),
                "--config",
                str(config_path),
            ]
        )
        == 0
    )
    assert required_columns_seen == [["SKU", "price"], ["SKU", "price"]]


def test_cli_converts_workflow_error_to_friendly_parser_error(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    config_path = write_config(tmp_path / "config.yaml")
    baseline_path = tmp_path / "baseline.xlsx"
    comparison_path = tmp_path / "comparison.xlsx"
    baseline_path.write_bytes(b"placeholder")
    comparison_path.write_bytes(b"placeholder")

    def fail_workflow(**kwargs: object) -> tuple[Path, ComparisonResult]:
        raise ReportExportError("could not export report")

    monkeypatch.setattr("excel_sales_automation.cli.run_excel_comparison_workflow", fail_workflow)

    with pytest.raises(SystemExit) as exc_info:
        main(
            [
                "compare",
                "--baseline-file",
                str(baseline_path),
                "--comparison-file",
                str(comparison_path),
                "--output-file",
                str(tmp_path / "report.xlsx"),
                "--config",
                str(config_path),
            ]
        )

    assert exc_info.value.code == 2


def test_cli_version_exits_successfully() -> None:
    with pytest.raises(SystemExit) as exc_info:
        main(["--version"])

    assert exc_info.value.code == 0
