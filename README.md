# ODS Spreadsheet Test Fixtures

All records, people, merchants, businesses, and order identifiers in these files are fictional. Each sheet is labeled as test data.

## Fixtures

- `fixtures/fictional_monthly_budget.ods`: September 2026 personal budget ledger with dates, categories, merchants, income and expenses, plus `SUMIF` totals and net balance.
- `fixtures/fictional_order_register.ods`: October 2026 sample orders with quantities, unit prices, discounts, tax rates, per-line formulas, and batch totals.
- `fixtures/fictional_project_timesheet.ods`: September 2026 consulting entries with hours, hourly rates, billable flags, per-entry calculations, and billable summaries.

The files exercise text, dates, decimals, currency and percentage formats, styled headers and totals, and formula cells. `generate_fixtures.py` regenerates them with LibreOffice installed.