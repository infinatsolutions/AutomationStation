# Excel Sales Automation

Excel Sales Automation is a production-oriented Python CLI scaffold for a freelance paid test project. The first version is intentionally focused on local-only Excel automation: comparing two workbook files containing online sales or product data, matching rows by product code, calculating price-related metrics, and exporting a new Excel report.

The client requires full Python source code, not only a compiled executable, so this repository is organized as a maintainable package under `src/` with tests and configuration files.

## Current Scope

This initial scaffold provides the package structure, configuration defaults, CLI entrypoint placeholder, and smoke tests. It does **not** implement real scraping, browser automation, marketplace integrations, or API collection yet.

Future scraping, API, or browser automation is explicitly out of scope for the paid-test phase and is represented only by a stub module for later extension planning.

## Planned Workflow

1. Load a baseline Excel file.
2. Load a comparison Excel file.
3. Match rows by product code.
4. Calculate price difference and margin rate.
5. Export a formatted Excel report.

## Installation Placeholder

Use Python 3.11 or newer.

```bash
python -m venv .venv
source .venv/bin/activate
python -m pip install -e ".[dev]"
```

## Usage Placeholder

After installation, the CLI command will be available as:

```bash
excel-sales-compare --help
```

The comparison workflow is not implemented in this scaffold phase. The CLI currently exposes a safe placeholder entrypoint.

## Default Configuration

Default column mappings live in `config/default_config.yaml`:

- `baseline_product_code_column`: product code column in the baseline workbook
- `comparison_product_code_column`: product code column in the comparison workbook
- `baseline_price_column`: price column in the baseline workbook
- `comparison_price_column`: price column in the comparison workbook
- `cost_column`: cost column used for margin calculations
- `margin_formula`: formula selection mode
- `duplicate_policy`: duplicate product-code handling mode
- `rounding_decimals`: numeric rounding precision for report metrics

## Formula Placeholder

The intended default calculations are:

- Price difference: `comparison_price - baseline_price`
- Margin rate: formula to be confirmed with the client; default configuration uses `margin_formula: auto` until finalized.

Potential margin formula if cost data is available:

```text
margin_rate = (comparison_price - cost) / comparison_price
```

## Assumptions Placeholder

Current documented assumptions:

- Input files are local Excel workbooks.
- Product code is the primary matching key.
- Duplicate product-code behavior should be reported rather than silently ignored.
- No credentials are required for the paid-test version.
- Future secrets or external service credentials must be passed through environment variables, never hardcoded.
- Browser automation, scraping, and API integrations are not part of this phase.

## Development and Testing

Run tests with:

```bash
pytest
```

Run linting with:

```bash
ruff check .
```

Format code with:

```bash
ruff format .
```
