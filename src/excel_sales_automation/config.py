"""Configuration loading and validation helpers."""

from __future__ import annotations

import importlib.util
import logging
from collections.abc import Mapping
from pathlib import Path
from typing import Any

if importlib.util.find_spec("yaml") is not None:
    import yaml
else:
    yaml = None

# The fallback parser keeps local tests usable in constrained environments where the
# declared PyYAML dependency cannot be installed. Production installs should use PyYAML.

from excel_sales_automation.models import AppConfig, ColumnConfig, FormulaConfig, RuntimeOptions

logger = logging.getLogger(__name__)

DEFAULT_CONFIG_PATH = Path("config/default_config.yaml")
SUPPORTED_MARGIN_FORMULAS = {"auto", "price_difference_over_comparison", "gross_margin"}
SUPPORTED_DUPLICATE_POLICIES = {"report", "fail"}


class ConfigError(ValueError):
    """Raised when configuration cannot be loaded or validated."""


def load_config(config_path: str | Path | None = None) -> AppConfig:
    """Load and validate application configuration from YAML."""
    resolved_path = Path(config_path) if config_path is not None else DEFAULT_CONFIG_PATH
    logger.debug("Loading configuration from %s", resolved_path)

    raw_config = _load_yaml_mapping(resolved_path)
    columns = _parse_columns(_require_mapping(raw_config, "columns"))
    formulas = _parse_formulas(_optional_mapping(raw_config, "formulas"))
    runtime = _parse_runtime(raw_config)

    return AppConfig(
        columns=columns,
        formulas=formulas,
        runtime=runtime,
        source_path=resolved_path,
    )


def _load_yaml_mapping(path: Path) -> Mapping[str, Any]:
    if not path.exists():
        raise ConfigError(f"Configuration file does not exist: {path}")
    if not path.is_file():
        raise ConfigError(f"Configuration path is not a file: {path}")

    try:
        text = path.read_text(encoding="utf-8")
    except OSError as exc:
        raise ConfigError(f"Configuration file could not be read: {path}: {exc}") from exc

    if yaml is not None:
        try:
            loaded = yaml.safe_load(text)
        except yaml.YAMLError as exc:
            raise ConfigError(f"Configuration file contains invalid YAML: {path}: {exc}") from exc
    else:
        loaded = _parse_simple_yaml(text, path)

    if loaded is None:
        raise ConfigError(f"Configuration file is empty: {path}")
    if not isinstance(loaded, Mapping):
        raise ConfigError(f"Configuration file must contain a top-level mapping: {path}")
    return loaded


def _parse_columns(section: Mapping[str, Any]) -> ColumnConfig:
    return ColumnConfig(
        baseline_product_code_column=_required_string(section, "baseline_product_code_column"),
        comparison_product_code_column=_required_string(section, "comparison_product_code_column"),
        baseline_price_column=_required_string(section, "baseline_price_column"),
        comparison_price_column=_required_string(section, "comparison_price_column"),
        cost_column=_optional_string(section, "cost_column"),
    )


def _parse_formulas(section: Mapping[str, Any]) -> FormulaConfig:
    margin_formula = _optional_string(section, "margin_formula", default="auto")
    if margin_formula not in SUPPORTED_MARGIN_FORMULAS:
        supported = ", ".join(sorted(SUPPORTED_MARGIN_FORMULAS))
        raise ConfigError(
            f"Invalid formulas.margin_formula '{margin_formula}'. Supported values: {supported}."
        )

    rounding_decimals = _optional_int(section, "rounding_decimals", default=2)
    if rounding_decimals < 0:
        raise ConfigError(
            "Invalid formulas.rounding_decimals: value must be a non-negative integer."
        )

    return FormulaConfig(margin_formula=margin_formula, rounding_decimals=rounding_decimals)


def _parse_runtime(raw_config: Mapping[str, Any]) -> RuntimeOptions:
    runtime_section = _optional_mapping(raw_config, "runtime")
    matching_section = _optional_mapping(raw_config, "matching")

    duplicate_policy = _optional_string(
        runtime_section,
        "duplicate_policy",
        default=_optional_string(matching_section, "duplicate_policy", default="report"),
    )
    if duplicate_policy not in SUPPORTED_DUPLICATE_POLICIES:
        supported = ", ".join(sorted(SUPPORTED_DUPLICATE_POLICIES))
        raise ConfigError(
            f"Invalid duplicate_policy '{duplicate_policy}'. Supported values: {supported}."
        )

    return RuntimeOptions(
        baseline_sheet_name=_optional_sheet_name(runtime_section, "baseline_sheet_name"),
        comparison_sheet_name=_optional_sheet_name(runtime_section, "comparison_sheet_name"),
        duplicate_policy=duplicate_policy,
    )


def _require_mapping(config: Mapping[str, Any], key: str) -> Mapping[str, Any]:
    if key not in config:
        raise ConfigError(f"Missing required configuration section: {key}.")
    value = config[key]
    if not isinstance(value, Mapping):
        raise ConfigError(f"Configuration section '{key}' must be a mapping.")
    return value


def _optional_mapping(config: Mapping[str, Any], key: str) -> Mapping[str, Any]:
    value = config.get(key, {})
    if value is None:
        return {}
    if not isinstance(value, Mapping):
        raise ConfigError(f"Configuration section '{key}' must be a mapping when provided.")
    return value


def _required_string(section: Mapping[str, Any], key: str) -> str:
    if key not in section:
        raise ConfigError(f"Missing required configuration field: {key}.")
    value = section[key]
    if not isinstance(value, str) or not value.strip():
        raise ConfigError(f"Configuration field '{key}' must be a non-empty string.")
    return value.strip()


def _optional_string(
    section: Mapping[str, Any], key: str, default: str | None = None
) -> str | None:
    value = section.get(key, default)
    if value is None:
        return None
    if not isinstance(value, str) or not value.strip():
        raise ConfigError(f"Configuration field '{key}' must be a non-empty string when provided.")
    return value.strip()


def _optional_int(section: Mapping[str, Any], key: str, default: int) -> int:
    value = section.get(key, default)
    if isinstance(value, bool) or not isinstance(value, int):
        raise ConfigError(f"Configuration field '{key}' must be an integer.")
    return value


def _optional_sheet_name(section: Mapping[str, Any], key: str) -> str | int | None:
    value = section.get(key)
    if value is None:
        return None
    if isinstance(value, bool):
        raise ConfigError(
            f"Configuration field '{key}' must be a sheet name string, integer, or null."
        )
    if isinstance(value, int):
        return value
    if isinstance(value, str) and value.strip():
        return value.strip()
    raise ConfigError(f"Configuration field '{key}' must be a sheet name string, integer, or null.")


def _parse_simple_yaml(text: str, path: Path) -> Mapping[str, Any] | None:
    """Parse the project's simple nested YAML shape when PyYAML is unavailable."""
    if not text.strip():
        return None

    parsed: dict[str, Any] = {}
    current_section: dict[str, Any] | None = None

    for line_number, raw_line in enumerate(text.splitlines(), start=1):
        if not raw_line.strip() or raw_line.lstrip().startswith("#"):
            continue
        indent = len(raw_line) - len(raw_line.lstrip(" "))
        stripped = raw_line.strip()
        if ":" not in stripped:
            raise ConfigError(f"Invalid YAML syntax in {path} at line {line_number}: {stripped}")
        key, raw_value = stripped.split(":", 1)
        key = key.strip()
        raw_value = raw_value.strip()

        if indent == 0:
            if raw_value:
                parsed[key] = _parse_scalar(raw_value)
                current_section = None
            else:
                current_section = {}
                parsed[key] = current_section
            continue

        if indent != 2 or current_section is None:
            raise ConfigError(f"Invalid YAML nesting in {path} at line {line_number}: {stripped}")
        current_section[key] = _parse_scalar(raw_value)

    return parsed


def _parse_scalar(value: str) -> str | int | bool | None:
    if value in {"null", "~"}:
        return None
    if value in {"true", "True"}:
        return True
    if value in {"false", "False"}:
        return False
    if value.startswith(('"', "'")) and value.endswith(('"', "'")):
        return value[1:-1]
    if value.lstrip("-").isdigit():
        return int(value)
    return value
