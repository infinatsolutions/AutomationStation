# Excel Sales Automation

Excel Sales Automation is a production-oriented Python CLI scaffold for a freelance paid test project. The first version is intentionally focused on local-only Excel automation: comparing two workbook files containing online sales or product data, matching rows by product code, calculating price-related metrics, and exporting a new Excel report.

The client requires full Python source code, not only a compiled executable, so this repository is organized as a maintainable package under `src/` with tests and configuration files.

## Current Scope

This initial scaffold provides the repository and package foundation for a local-only Excel automation tool. It does **not** perform the full Excel comparison workflow yet.

Future scraping, API, or browser automation is explicitly out of scope for the paid-test scaffold phase and is represented only by a stub module for later extension planning.

## Implemented in this scaffold

The current repository includes only the foundation needed for upcoming implementation phases:

- Package scaffold under `src/excel_sales_automation`.
- CLI placeholder exposed through the `excel-sales-compare` console script.
- Default configuration file at `config/default_config.yaml` with typed YAML loading and validation.
- Smoke test confirming the package imports and exposes a version string.
- Local-only Excel automation project structure for future comparison/reporting work.

## Planned in upcoming phases

The following items are planned but are **not implemented** in this scaffold:

- Excel input reading.
- Product-code normalization.
- Row matching.
- Price difference and margin-rate calculations.
- Formatted Excel report export.
- Future scraping/API/browser automation extension points through a stub only.

## Planned Workflow

The future workflow is expected to be:

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

The comparison workflow is not implemented in this scaffold phase. The CLI currently exposes a safe placeholder entrypoint only.

## Default Configuration

Default column mappings live in `config/default_config.yaml`. The CLI can load this default file or a user-provided YAML config path in later workflow phases. Use config overrides when the client workbook uses different column names or sheet names:

- `baseline_product_code_column`: product code column in the baseline workbook
- `comparison_product_code_column`: product code column in the comparison workbook
- `baseline_price_column`: price column in the baseline workbook
- `comparison_price_column`: price column in the comparison workbook
- `cost_column`: cost column used for margin calculations
- `margin_formula`: formula selection mode
- `duplicate_policy`: duplicate product-code handling mode
- `rounding_decimals`: numeric rounding precision for report metrics
- `baseline_sheet_name`: baseline workbook sheet name, sheet index, or `null`
- `comparison_sheet_name`: comparison workbook sheet name, sheet index, or `null`

Supported values:

- `margin_formula`: `auto`, `price_difference_over_comparison`, or `gross_margin`
- `duplicate_policy`: `report` or `fail`


Example override file:

```yaml
columns:
  baseline_product_code_column: SKU
  comparison_product_code_column: SKU
  baseline_price_column: Old Price
  comparison_price_column: New Price
  cost_column: Cost
formulas:
  margin_formula: gross_margin
  rounding_decimals: 2
runtime:
  baseline_sheet_name: Products
  comparison_sheet_name: Products
  duplicate_policy: report
```

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
