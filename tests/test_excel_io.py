"""Tests for Excel workbook loading."""

from __future__ import annotations

from pathlib import Path

import pytest

from excel_sales_automation import excel_io
from excel_sales_automation.excel_io import ExcelInputError, MissingColumnError, read_excel_workbook


class FakeFrame:
    def __init__(self, columns: list[str]) -> None:
        self.columns = columns


class FakePandas:
    DataFrame = FakeFrame

    def __init__(self, frame: FakeFrame) -> None:
        self.frame = frame
        self.calls: list[dict[str, object]] = []

    def read_excel(self, *args: object, **kwargs: object) -> FakeFrame:
        self.calls.append({"args": args, "kwargs": kwargs})
        return self.frame


def test_read_excel_workbook_loads_xlsx_with_openpyxl(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    workbook_path = tmp_path / "sales.xlsx"
    workbook_path.write_bytes(b"placeholder")
    fake_pandas = FakePandas(FakeFrame(["product_code", "price"]))
    monkeypatch.setattr(excel_io, "pd", fake_pandas)

    result = read_excel_workbook(
        workbook_path,
        required_columns=["product_code", "price"],
        text_columns=["product_code"],
    )

    assert result is fake_pandas.frame
    assert fake_pandas.calls[0]["kwargs"]["engine"] == "openpyxl"
    assert fake_pandas.calls[0]["kwargs"]["sheet_name"] == 0
    assert fake_pandas.calls[0]["kwargs"]["dtype"] == {"product_code": "string"}


def test_read_excel_workbook_supports_sheet_name(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    workbook_path = tmp_path / "sales.xlsx"
    workbook_path.write_bytes(b"placeholder")
    fake_pandas = FakePandas(FakeFrame(["product_code", "price"]))
    monkeypatch.setattr(excel_io, "pd", fake_pandas)

    read_excel_workbook(workbook_path, sheet_name="Products")

    assert fake_pandas.calls[0]["kwargs"]["sheet_name"] == "Products"


def test_read_excel_workbook_rejects_missing_file(tmp_path: Path) -> None:
    missing_path = tmp_path / "missing.xlsx"

    with pytest.raises(ExcelInputError, match=str(missing_path)):
        read_excel_workbook(missing_path)


def test_read_excel_workbook_rejects_unsupported_extension(tmp_path: Path) -> None:
    text_path = tmp_path / "sales.txt"
    text_path.write_text("not excel", encoding="utf-8")

    with pytest.raises(ExcelInputError, match="Unsupported Excel file extension"):
        read_excel_workbook(text_path)


def test_read_excel_workbook_rejects_missing_required_column(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    workbook_path = tmp_path / "sales.xlsx"
    workbook_path.write_bytes(b"placeholder")
    fake_pandas = FakePandas(FakeFrame(["product_code"]))
    monkeypatch.setattr(excel_io, "pd", fake_pandas)

    with pytest.raises(MissingColumnError, match="price"):
        read_excel_workbook(workbook_path, required_columns=["product_code", "price"])


def test_read_excel_workbook_rejects_legacy_xls_with_clear_error(tmp_path: Path) -> None:
    workbook_path = tmp_path / "legacy.xls"
    workbook_path.write_bytes(b"not a real workbook")

    with pytest.raises(ExcelInputError, match="Legacy .xls files are not supported"):
        read_excel_workbook(workbook_path)


def test_read_excel_workbook_rejects_missing_pandas(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    workbook_path = tmp_path / "sales.xlsx"
    workbook_path.write_bytes(b"placeholder")
    monkeypatch.setattr(excel_io, "pd", None)

    with pytest.raises(ExcelInputError, match="pandas is required"):
        read_excel_workbook(workbook_path)
