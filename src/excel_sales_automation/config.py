"""Configuration loading helpers."""

from __future__ import annotations

from pathlib import Path
from typing import Any


def load_config(path: str | Path) -> dict[str, Any]:
    """Load configuration from disk.

    YAML parsing will be implemented with the comparison workflow. The placeholder keeps
    file-path validation centralized without adding extra runtime dependencies.
    """
    config_path = Path(path)
    if not config_path.exists():
        raise FileNotFoundError(f"Configuration file does not exist: {config_path}")
    return {"config_path": config_path}
