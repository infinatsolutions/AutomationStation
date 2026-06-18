"""Typed data models for comparison inputs and outputs."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path


@dataclass(frozen=True, slots=True)
class WorkbookInput:
    """Local Excel workbook input descriptor."""

    path: Path
    sheet_name: str | int | None = 0
