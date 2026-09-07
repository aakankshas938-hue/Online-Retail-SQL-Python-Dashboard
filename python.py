
"""
Online Retail Excel Dashboard Builder — MySQL version
=======================================================
Connects to your MySQL database, pulls transaction-level retail data,
and produces the same formatted Excel dashboard as the CSV version
(KPI cards, RFM segmentation, new-vs-returning revenue, day/hour
heatmap, top products/countries/customers) — all charts embedded as
PICTURES so they display correctly in every viewer (Excel, Google
Sheets, LibreOffice, mobile apps, browser previewers).
 
Install dependencies:
    pip install pandas openpyxl matplotlib mysql-connector-python sqlalchemy --break-system-packages
 
Configure the DB_* constants and TABLE_NAME / COLUMN_MAP below to match
your database, then run:
    python build_dashboard_from_mysql.py
    python build_dashboard_from_mysql.py --output my_dashboard.xlsx
"""
 
import argparse
from pathlib import Path
 
import pandas as pd
from sqlalchemy import create_engine
 
# Reuse the exact same tested aggregation / chart / workbook logic as the
# CSV version so both scripts produce identical dashboards.
from build_dashboard_from_csv import (
    build_aggregates,
    build_workbook,
    make_chart_images,
)
 
# ============================================================
# MYSQL CONFIGURATION — edit these to match your setup
# ============================================================
DB_USER = "root"
DB_PASSWORD = "root"
DB_HOST = "127.0.0.1"
DB_PORT = "3306"
DB_NAME = "retail_sales"
 
# Name of the table holding transaction-level rows
TABLE_NAME = "sales"
 
# Map YOUR table's column names -> the names the dashboard code expects.
# Edit the right-hand side to match your actual column names.
COLUMN_MAP = {
    "Invoice": "Invoice",
    "StockCode": "StockCode",
    "Description": "Description",
    "Quantity": "Quantity",
    "InvoiceDate": "InvoiceDate",
    "Price": "Price",
    "Customer ID": "CustomerID",   # e.g. change to "customer_id" if that's your column
    "Country": "Country",
}
 
 
# ============================================================
# LOAD FROM MYSQL
# ============================================================
def load_from_mysql():
    conn_str = f"mysql+mysqlconnector://{DB_USER}:{DB_PASSWORD}@{DB_HOST}:{DB_PORT}/{DB_NAME}"
    engine = create_engine(conn_str)
 
    select_cols = ", ".join(f"`{v}` AS `{k}`" for k, v in COLUMN_MAP.items())
    query = f"SELECT {select_cols} FROM `{TABLE_NAME}`"
 
    print(f"Connecting to mysql://{DB_HOST}:{DB_PORT}/{DB_NAME} ...")
    df = pd.read_sql(query, engine)
    print(f"Fetched {len(df):,} rows from `{TABLE_NAME}`.")
 
    df["InvoiceDate"] = pd.to_datetime(df["InvoiceDate"])
    df["Invoice_str"] = df["Invoice"].astype(str)
    df["IsCancelled"] = df["Invoice_str"].str.startswith("C")
 
    cancelled = df[df["IsCancelled"]]
    cancel_orders = cancelled["Invoice"].nunique()
    cancel_value = (cancelled["Quantity"] * cancelled["Price"]).sum()
 
    clean = df[(~df["IsCancelled"]) & (df["Quantity"] > 0) & (df["Price"] > 0)].copy()
    clean["Revenue"] = clean["Quantity"] * clean["Price"]
    clean["Month"] = clean["InvoiceDate"].dt.to_period("M").astype(str)
    clean["Hour"] = clean["InvoiceDate"].dt.hour
    clean["DayOfWeek"] = clean["InvoiceDate"].dt.day_name()
 
    return clean, cancel_orders, cancel_value
 
 
# ============================================================
# MAIN
# ============================================================
def main():
    parser = argparse.ArgumentParser(description="Build an Online Retail Excel dashboard from a MySQL database.")
    parser.add_argument("--output", default="online_retail_dashboard.xlsx", help="Output .xlsx path")
    args = parser.parse_args()
 
    clean, cancel_orders, cancel_value = load_from_mysql()
 
    print("Building aggregates (KPIs, RFM, heatmap, etc.)...")
    agg = build_aggregates(clean, cancel_orders, cancel_value)
 
    print("Rendering chart images...")
    tmp_dir = Path(".dashboard_charts_tmp")
    tmp_dir.mkdir(exist_ok=True)
    chart_paths = make_chart_images(agg, tmp_dir)
 
    print("Building workbook...")
    build_workbook(agg, chart_paths, args.output)
 
 
if __name__ == "__main__":
    main()
 
