# Online Retail Excel Dashboard

Turns a transaction-level retail CSV (or a MySQL database) into a fully
formatted, multi-sheet Excel dashboard — KPI cards, customer segmentation,
sales trends, and top products/countries/customers — with zero manual
pivot tables or chart-building.

## What it builds

A single `.xlsx` file with these sheets:

| Sheet | Contents |
|---|---|
| **Dashboard** | 6 KPI cards + 5 charts (monthly trend, new vs returning revenue, RFM segments, revenue by country, top products) |
| Monthly Trend | Revenue, orders, customers, quantity by month |
| New vs Returning | Revenue split between new and returning customers, by month |
| RFM Segments | Customers grouped into Champions / Loyal / At Risk / Lost / etc., with definitions |
| Day-Hour Heatmap | Revenue by day of week × hour, color-scaled |
| Top Products | Top 15 products by revenue |
| Top Countries | Revenue, orders, customers by country |
| Top Customers | Top 15 customers by revenue |
| Notes | Data cleaning rules and methodology |

Charts are rendered with **matplotlib and embedded as images**, not native
Excel chart objects — so they display correctly in every viewer (Excel,
Google Sheets, LibreOffice, mobile apps, in-browser previewers), instead
of showing up blank in tools that don't support native Excel charts.

## Data expected

A transaction-level table (CSV or SQL table) with these columns:

```
Invoice, StockCode, Description, Quantity, InvoiceDate, Price, Customer ID, Country
```

(This matches the shape of the classic *Online Retail II* dataset.)

## Setup

```bash
pip install pandas openpyxl matplotlib
# only if using the MySQL version:
pip install mysql-connector-python sqlalchemy
```

## Usage

### From a CSV file

```bash
python build_dashboard_from_csv.py path/to/data.csv
python build_dashboard_from_csv.py path/to/data.csv --output my_dashboard.xlsx
```

### From a MySQL database

Edit the `DB_*` config and `TABLE_NAME` / `COLUMN_MAP` at the top of
`build_dashboard_from_mysql.py` to match your database, then:

```bash
python build_dashboard_from_mysql.py
python build_dashboard_from_mysql.py --output my_dashboard.xlsx
```

`build_dashboard_from_mysql.py` imports its aggregation, charting, and
workbook logic from `build_dashboard_from_csv.py` — keep both files in
the same folder.

## Notes on the numbers

- Cancelled invoices (Invoice number starting with `C`) and rows with
  `Quantity <= 0` or `Price <= 0` are excluded from all revenue/order
  KPIs and charts.
- The Top Products sheet additionally excludes non-merchandise stock
  codes (postage, manual entries, discounts, samples, bank charges).
- Most KPI cells on the Dashboard are **live formulas** referencing the
  data sheets — edit the underlying tables and they'll recalculate.
  A few (Unique Customers, Repeat Purchase Rate, Busiest Day/Hour) are
  plain computed values because they need the full row-level dataset,
  which isn't stored in the workbook — re-run the script to refresh them.

## Customizing the look

Colors and fonts are defined as constants near the top of
`build_dashboard_from_csv.py`:

```python
TEAL = "0F766E"
TEAL_DARK = "134E4A"
CHARCOAL = "1E293B"
AMBER = "F59E0B"
...
FONT_NAME = "Arial"
```

Change these and re-run to restyle the whole dashboard.
