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

## CLI Usage

Validate configuration only:

```bash
excel-sales-compare check-config --config config/default_config.yaml
```

Run the full local Excel comparison workflow:

```bash
excel-sales-compare compare \
  --baseline-file path/to/baseline.xlsx \
  --comparison-file path/to/comparison.xlsx \
  --output-file path/to/report.xlsx \
  --config config/default_config.yaml
```

Optional sheet overrides are available when a workbook needs a specific worksheet:

```bash
excel-sales-compare compare \
  --baseline-file baseline.xlsx \
  --comparison-file comparison.xlsx \
  --output-file report.xlsx \
  --baseline-sheet Products \
  --comparison-sheet Products
```

The CLI runs locally only: it reads local Excel workbooks, applies configured column mappings, normalizes data, matches by product code, calculates metrics, and writes a local `.xlsx` report. It does not perform scraping, browser automation, API calls, or external credential handling.

## Default Configuration

Default column mappings live in `config/default_config.yaml`. The CLI already validates this default file or a user-provided YAML config path; the Excel comparison workflow itself remains planned for upcoming phases. Use config overrides when the client workbook uses different column names or sheet names:

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

## Input File Expectations

The current input layer supports local `.xlsx` and `.xlsm` workbooks through `pandas` and `openpyxl`. Legacy `.xls` files should be converted to `.xlsx` before processing.

Column names are matched exactly using the configured mappings. Product-code columns are read as text where possible so leading zeros are preserved when Excel stores the value as text. During normalization, blank product codes, `NaN`, `nan`, and `None` are treated as missing and tracked as invalid row issues.

Price and cost values may be numeric cells or strings containing commas and common currency symbols such as `$`. Invalid required prices are tracked as invalid row issues and converted to missing numeric values in the normalized copy. Optional cost values may be blank, but invalid non-blank cost values are also tracked.

Normalization validates that required product-code and price columns exist, then returns a cleaned DataFrame copy plus invalid row details; it does not mutate caller-provided DataFrames in place.

## Matching Behavior

Comparison uses normalized product-code values as exact join keys. Clean one-to-one product codes are matched with baseline columns suffixed as `_baseline` and comparison columns suffixed as `_comparison`, plus a canonical `product_code` output column.

Unmatched rows are preserved in separate report-ready tables: products present only in the comparison file are `missing_in_baseline`, and products present only in the baseline file are `missing_in_comparison`. Rows with missing product codes are excluded from matching and included in invalid-row output. Rows with valid product codes but invalid numeric fields may still match by product code; downstream calculation steps must exclude or flag them using the invalid-row output.

Duplicate product codes are ambiguous. With `duplicate_policy: fail`, comparison raises an error that lists duplicate source rows. With `duplicate_policy: report`, duplicate rows are reported and excluded from clean matched calculations instead of silently choosing one row.

## Financial Calculation Assumptions

Financial metrics are calculated only for matched product-code rows. The default assumptions are explicit because the paid-test prompt does not define a single required margin formula.

Calculated columns include `price_difference`, `price_difference_rate`, `margin_rate`, `margin_formula_used`, `calculation_status`, and `calculation_warning`.

Formulas:

- `price_difference = comparison_price - baseline_price`
- `price_difference_rate = price_difference / baseline_price`
- `price_difference_over_comparison`: `margin_rate = price_difference / comparison_price`
- `gross_margin`: `margin_rate = (comparison_price - cost) / comparison_price`
- `auto`: use `gross_margin` when a valid cost exists for the row; otherwise use `price_difference_over_comparison`

When both baseline and comparison cost values are present, comparison cost is used for gross margin. If comparison cost is missing, baseline cost is used as a fallback.

Division by zero and missing numeric inputs do not crash the application. They produce missing metric values with `calculation_status` set to `warning` or `invalid` and a human-readable `calculation_warning`. Numeric outputs are rounded with `rounding_decimals` from configuration.

## Output Workbook Structure

The final report writer creates a local `.xlsx` workbook for business review. Output directories are created automatically and existing files are overwritten by default.

The workbook contains these sheets:

- `Summary`: run timestamp, input file names, config path, row counts, formula definitions, assumptions, and calculation status counts
- `Matched`: matched product rows with calculated values such as `price_difference`, `price_difference_rate`, `margin_rate`, `calculation_status`, and `calculation_warning`
- `Missing_In_Baseline`: products present only in the comparison workbook
- `Missing_In_Comparison`: products present only in the baseline workbook
- `Duplicates`: duplicate product-code rows by source workbook
- `Invalid_Rows`: missing product codes and normalization/calculation issues that should be reviewed

The workbook uses readable formatting such as styled headers, frozen top rows, filters, practical column widths, and numeric formats when `openpyxl` is available. In constrained environments, a minimal standards-compliant `.xlsx` writer is used so report generation remains testable without external services.

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
