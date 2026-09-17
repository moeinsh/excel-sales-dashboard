# Excel Sales Dashboard Builder (Python + openpyxl) — sample project

Turns a plain CSV of sales transactions into a management-ready Excel
dashboard (`dashboard.xlsx`) with one command:

- **Dashboard sheet** — KPI summary built with real Excel formulas:
  total revenue (`SUM`), average monthly revenue, best product category
  and best month (`INDEX`/`MATCH`/`MAX`), so the numbers stay live if the
  data changes. Includes a native monthly-revenue **column chart** and a
  revenue-by-category **pie chart** (percentage labels).
- **Summary sheet** — monthly and per-category totals via `SUMIF`
  formulas; this is what feeds the KPIs and both charts.
- **Data sheet** — all 240 transactions as a formatted Excel Table
  (banded rows, filter headers) with a `Month` helper column (`TEXT`
  formula) powering the monthly rollups.

The demo run used the included `sales_data.csv`: 240 transactions across
12 months of 2025 (4 regions × 4 product categories), totalling
$338,511.28.

## Run it

```bash
pip install openpyxl
python build_dashboard.py --input sales_data.csv --output dashboard.xlsx
```

Open `dashboard.xlsx` in Excel / LibreOffice / Google Sheets — the KPIs
calculate, the charts render, the table filters.

## Point it at your own CSV

Your file needs the same four columns (header row required):

```
date,region,category,revenue
2025-03-14,North,Electronics,1299.99
```

Dates must be `YYYY-MM-DD`. Then:

```bash
python build_dashboard.py --input your_file.csv --output your_dashboard.xlsx
```

Category names are picked up automatically from the data (the four demo
categories in the Summary sheet are just labels — replace them with yours
or extend the list).

## Scope note (honest)

This is a demonstration sample showing my Excel-automation workflow end
to end — reading real data, writing formulas (not hard-coded values),
native charts, and table formatting with openpyxl. The sales figures are
synthetic demo data I generated for this sample (seeded, reproducible),
not real business results. No client, no testimonials, nothing invented.

It differs from my `excel-report-automation` sample on purpose: that one
**cleans** messy exports; this one **builds** a dashboard from clean data.

---

**Author:** Moein Shahidi — [@moeinsh](https://github.com/moeinsh)

© 2026 Moein Shahidi. Released under the MIT License.
