"""
Online Retail Excel Dashboard Builder — CSV version
=====================================================
Reads a transaction-level retail CSV (same shape as the classic
"Online Retail II" dataset: Invoice, StockCode, Description, Quantity,
InvoiceDate, Price, Customer ID, Country) and produces a formatted,
multi-sheet Excel dashboard with KPI cards, RFM customer segmentation,
new-vs-returning revenue, a day/hour sales heatmap, and top
products/countries/customers tables.
 
Charts are rendered with matplotlib and embedded as PICTURES (not native
Excel chart objects) so they display correctly in every viewer — Excel,
Google Sheets, LibreOffice, mobile apps, and in-browser previewers that
don't support native Excel charts.
 
Install dependencies:
    pip install pandas openpyxl matplotlib --break-system-packages
 
Usage:
    python build_dashboard_from_csv.py path/to/online_retail_II.csv
    python build_dashboard_from_csv.py path/to/data.csv --output my_dashboard.xlsx
"""
 
import argparse
import json
import sys
from pathlib import Path
 
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import column_index_from_string, coordinate_from_string
 
# ============================================================
# THEME — tweak these to restyle the whole dashboard
# ============================================================
TEAL = "0F766E"
TEAL_DARK = "134E4A"
CHARCOAL = "1E293B"
SLATE = "334155"
AMBER = "F59E0B"
AMBER_DARK = "B45309"
GREY_LIGHT = "94A3B8"
SLATE_LIGHT = "F1F5F9"
PANEL_BORDER = "CBD5E1"
WHITE = "FFFFFF"
FONT_NAME = "Arial"
 
NON_PRODUCT_CODES = [
    "POST", "DOT", "M", "m", "C2", "D", "S", "BANK CHARGES", "AMAZONFEE", "PADS", "CRUK",
]
 
 
# ============================================================
# 1. LOAD & CLEAN
# ============================================================
def load_and_clean(csv_path: str):
    df = pd.read_csv(csv_path, encoding="utf-8")
 
    required = ["Invoice", "StockCode", "Description", "Quantity", "InvoiceDate", "Price", "Customer ID", "Country"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"CSV is missing expected columns: {missing}")
 
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
# 2. AGGREGATE
# ============================================================
def build_aggregates(clean: pd.DataFrame, cancel_orders: int, cancel_value: float):
    agg = {}
 
    # ---- KPIs ----
    total_revenue = clean["Revenue"].sum()
    total_orders = clean["Invoice"].nunique()
    total_customers = clean["Customer ID"].nunique()
    total_quantity = int(clean["Quantity"].sum())
    avg_order_value = total_revenue / total_orders
 
    repeat_rate = (clean.dropna(subset=["Customer ID"]).groupby("Customer ID")["Invoice"]
                   .nunique().gt(1).mean())
 
    hour_rev = clean.groupby("Hour")["Revenue"].sum()
    day_rev = clean.groupby("DayOfWeek")["Revenue"].sum()
    peak_hour = int(hour_rev.idxmax())
    peak_day = day_rev.idxmax()
 
    agg["kpis"] = dict(
        total_revenue=float(total_revenue),
        total_orders=int(total_orders),
        total_customers=int(total_customers),
        total_quantity=total_quantity,
        avg_order_value=float(avg_order_value),
        n_countries=int(clean["Country"].nunique()),
        date_min=str(clean["InvoiceDate"].min().date()),
        date_max=str(clean["InvoiceDate"].max().date()),
        cancel_orders=int(cancel_orders),
        cancel_value=float(cancel_value),
        repeat_rate=float(repeat_rate),
        peak_hour=peak_hour,
        peak_day=peak_day,
    )
 
    # ---- Monthly trend ----
    monthly = (clean.groupby("Month")
               .agg(Revenue=("Revenue", "sum"), Orders=("Invoice", "nunique"),
                    Customers=("Customer ID", "nunique"), Quantity=("Quantity", "sum"))
               .reset_index().sort_values("Month"))
    agg["monthly"] = monthly
 
    # ---- Top products (merchandise only) ----
    products_only = clean[~clean["StockCode"].astype(str).str.upper().isin(
        [c.upper() for c in NON_PRODUCT_CODES])]
    top_products = (products_only.groupby(["StockCode", "Description"])
                     .agg(Revenue=("Revenue", "sum"), Quantity=("Quantity", "sum"),
                          Orders=("Invoice", "nunique"))
                     .reset_index().sort_values("Revenue", ascending=False).head(15))
    top_products["Description"] = top_products["Description"].fillna("").astype(str).str.strip()
    agg["top_products"] = top_products
 
    # ---- Top countries ----
    country = (clean.groupby("Country")
               .agg(Revenue=("Revenue", "sum"), Orders=("Invoice", "nunique"),
                    Customers=("Customer ID", "nunique"), Quantity=("Quantity", "sum"))
               .reset_index().sort_values("Revenue", ascending=False))
    agg["country"] = country
 
    # ---- Top customers ----
    top_customers = (clean.dropna(subset=["Customer ID"]).groupby("Customer ID")
                      .agg(Revenue=("Revenue", "sum"), Orders=("Invoice", "nunique"),
                           Quantity=("Quantity", "sum"))
                      .reset_index().sort_values("Revenue", ascending=False).head(15))
    agg["top_customers"] = top_customers
 
    # ---- New vs returning customer revenue per month ----
    cust_df = clean.dropna(subset=["Customer ID"]).copy()
    first_month = cust_df.groupby("Customer ID")["InvoiceDate"].min().dt.to_period("M").astype(str)
    first_month.name = "FirstMonth"
    cust_df = cust_df.merge(first_month, on="Customer ID")
    cust_df["CustType"] = np.where(cust_df["Month"] == cust_df["FirstMonth"], "New", "Returning")
    newret = (cust_df.groupby(["Month", "CustType"])["Revenue"].sum()
              .unstack(fill_value=0).reset_index().sort_values("Month"))
    for col in ["New", "Returning"]:
        if col not in newret.columns:
            newret[col] = 0.0
    agg["newret"] = newret[["Month", "New", "Returning"]]
 
    # ---- RFM segmentation ----
    snapshot_date = clean["InvoiceDate"].max() + pd.Timedelta(days=1)
    rfm = cust_df.groupby("Customer ID").agg(
        Recency=("InvoiceDate", lambda x: (snapshot_date - x.max()).days),
        Frequency=("Invoice", "nunique"),
        Monetary=("Revenue", "sum"),
    ).reset_index()
    rfm["R_Score"] = pd.qcut(rfm["Recency"], 4, labels=[4, 3, 2, 1]).astype(int)
    rfm["F_Score"] = pd.qcut(rfm["Frequency"].rank(method="first"), 4, labels=[1, 2, 3, 4]).astype(int)
    rfm["M_Score"] = pd.qcut(rfm["Monetary"], 4, labels=[1, 2, 3, 4]).astype(int)
 
    def segment(row):
        r, f, m = row["R_Score"], row["F_Score"], row["M_Score"]
        if r >= 4 and f >= 4 and m >= 4:
            return "Champions"
        if r >= 3 and f >= 3:
            return "Loyal Customers"
        if r >= 4 and f <= 2:
            return "New Customers"
        if r <= 2 and f >= 3:
            return "At Risk"
        if r <= 2 and f <= 2 and m <= 2:
            return "Lost"
        return "Potential Loyalist"
 
    rfm["Segment"] = rfm.apply(segment, axis=1)
    seg_summary = (rfm.groupby("Segment")
                   .agg(Customers=("Customer ID", "count"), Revenue=("Monetary", "sum"),
                        AvgRecency=("Recency", "mean"), AvgFrequency=("Frequency", "mean"))
                   .reset_index().sort_values("Revenue", ascending=False))
    agg["rfm_segments"] = seg_summary
 
    # ---- Day x Hour heatmap ----
    day_order = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
    heat = (clean.groupby(["DayOfWeek", "Hour"])["Revenue"].sum().reset_index()
            .pivot(index="DayOfWeek", columns="Hour", values="Revenue")
            .reindex(day_order).fillna(0))
    agg["heatmap"] = heat
 
    return agg
 
 
# ============================================================
# 3. CHART IMAGES (matplotlib -> PNG)
# ============================================================
def money_fmt(x, pos):
    return f"${x/1000:,.0f}K" if x < 1_000_000 else f"${x/1_000_000:,.1f}M"
 
 
def make_chart_images(agg, tmp_dir: Path):
    plt.rcParams["font.family"] = "DejaVu Sans"
    plt.rcParams["axes.edgecolor"] = PANEL_BORDER
    plt.rcParams["axes.linewidth"] = 0.8
 
    monthly = agg["monthly"]
    country = agg["country"].head(8)
    top_products = agg["top_products"].head(10)
    newret = agg["newret"]
    seg = agg["rfm_segments"]
 
    paths = {}
 
    # Monthly revenue trend
    fig, ax = plt.subplots(figsize=(13.2, 3.6), dpi=150)
    ax.plot(monthly["Month"], monthly["Revenue"], color=f"#{TEAL}", linewidth=2.5, marker="o", markersize=3.5)
    ax.fill_between(range(len(monthly)), monthly["Revenue"], color=f"#{TEAL}", alpha=0.08)
    ax.set_ylabel("Revenue", fontsize=9, color=f"#{SLATE}")
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(money_fmt))
    ax.tick_params(axis="x", rotation=45, labelsize=7.5, colors=f"#{SLATE}")
    ax.tick_params(axis="y", labelsize=8, colors=f"#{SLATE}")
    ax.spines[["top", "right"]].set_visible(False)
    ax.grid(axis="y", color=f"#{PANEL_BORDER}", linewidth=0.6, alpha=0.6)
    ax.set_axisbelow(True)
    plt.tight_layout()
    p = tmp_dir / "chart_monthly_trend.png"
    plt.savefig(p, transparent=True)
    plt.close()
    paths["monthly_trend"] = p
 
    # New vs returning
    fig, ax = plt.subplots(figsize=(6.4, 3.6), dpi=150)
    ax.bar(newret["Month"], newret["New"], color=f"#{AMBER}", label="New", width=0.65)
    ax.bar(newret["Month"], newret["Returning"], bottom=newret["New"], color=f"#{TEAL_DARK}", label="Returning", width=0.65)
    ax.set_ylabel("Revenue", fontsize=9, color=f"#{SLATE}")
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(money_fmt))
    ax.tick_params(axis="x", rotation=45, labelsize=6.5, colors=f"#{SLATE}")
    ax.tick_params(axis="y", labelsize=8, colors=f"#{SLATE}")
    ax.spines[["top", "right"]].set_visible(False)
    ax.grid(axis="y", color=f"#{PANEL_BORDER}", linewidth=0.6, alpha=0.6)
    ax.set_axisbelow(True)
    ax.legend(frameon=False, fontsize=8, loc="upper left")
    plt.tight_layout()
    p = tmp_dir / "chart_new_vs_returning.png"
    plt.savefig(p, transparent=True)
    plt.close()
    paths["new_vs_returning"] = p
 
    # RFM pie
    colors_map = [TEAL_DARK, TEAL, AMBER, AMBER_DARK, SLATE, GREY_LIGHT]
    fig, ax = plt.subplots(figsize=(6.4, 3.6), dpi=150)
    wedges, texts, autotexts = ax.pie(
        seg["Revenue"], labels=seg["Segment"], autopct="%1.0f%%",
        colors=[f"#{c}" for c in colors_map[:len(seg)]],
        textprops={"fontsize": 8, "color": f"#{SLATE}"}, pctdistance=0.75,
        wedgeprops={"edgecolor": "white", "linewidth": 1},
    )
    for at in autotexts:
        at.set_color("white")
        at.set_fontsize(8)
        at.set_fontweight("bold")
    plt.tight_layout()
    p = tmp_dir / "chart_rfm_pie.png"
    plt.savefig(p, transparent=True)
    plt.close()
    paths["rfm_pie"] = p
 
    # Revenue by country
    fig, ax = plt.subplots(figsize=(6.4, 3.6), dpi=150)
    c_sorted = country.sort_values("Revenue")
    ax.barh(c_sorted["Country"], c_sorted["Revenue"], color=f"#{TEAL_DARK}", height=0.6)
    ax.set_xlabel("Revenue", fontsize=9, color=f"#{SLATE}")
    ax.xaxis.set_major_formatter(mticker.FuncFormatter(money_fmt))
    ax.tick_params(axis="y", labelsize=8.5, colors=f"#{SLATE}")
    ax.tick_params(axis="x", labelsize=8, colors=f"#{SLATE}")
    ax.spines[["top", "right"]].set_visible(False)
    ax.grid(axis="x", color=f"#{PANEL_BORDER}", linewidth=0.6, alpha=0.6)
    ax.set_axisbelow(True)
    plt.tight_layout()
    p = tmp_dir / "chart_country.png"
    plt.savefig(p, transparent=True)
    plt.close()
    paths["country"] = p
 
    # Top products
    fig, ax = plt.subplots(figsize=(6.4, 3.6), dpi=150)
    p_sorted = top_products.sort_values("Revenue")
    ax.barh(p_sorted["Description"].str.title(), p_sorted["Revenue"], color=f"#{AMBER_DARK}", height=0.6)
    ax.set_xlabel("Revenue", fontsize=9, color=f"#{SLATE}")
    ax.xaxis.set_major_formatter(mticker.FuncFormatter(money_fmt))
    ax.tick_params(axis="y", labelsize=7.5, colors=f"#{SLATE}")
    ax.tick_params(axis="x", labelsize=8, colors=f"#{SLATE}")
    ax.spines[["top", "right"]].set_visible(False)
    ax.grid(axis="x", color=f"#{PANEL_BORDER}", linewidth=0.6, alpha=0.6)
    ax.set_axisbelow(True)
    plt.tight_layout()
    p = tmp_dir / "chart_top_products.png"
    plt.savefig(p, transparent=True)
    plt.close()
    paths["top_products"] = p
 
    return paths
 
 
# ============================================================
# 4. WORKBOOK HELPERS
# ============================================================
header_font = Font(name=FONT_NAME, size=10, bold=True, color=WHITE)
normal_font = Font(name=FONT_NAME, size=10, color="0F172A")
bold_font = Font(name=FONT_NAME, size=10, bold=True, color="0F172A")
charcoal_fill = PatternFill("solid", fgColor=CHARCOAL)
light_fill = PatternFill("solid", fgColor=SLATE_LIGHT)
thin = Side(style="thin", color=PANEL_BORDER)
med = Side(style="medium", color=PANEL_BORDER)
border = Border(left=thin, right=thin, top=thin, bottom=thin)
center = Alignment(horizontal="center", vertical="center")
 
 
def write_table(ws, headers, rows_df, start_row=1, currency_cols=None):
    currency_cols = currency_cols or []
    for j, h in enumerate(headers, start=1):
        c = ws.cell(row=start_row, column=j, value=h)
        c.font = header_font
        c.fill = charcoal_fill
        c.alignment = center
        c.border = border
    for i, (_, row) in enumerate(rows_df.iterrows()):
        r = start_row + 1 + i
        for j, col in enumerate(rows_df.columns, start=1):
            cell = ws.cell(row=r, column=j, value=row[col])
            cell.font = normal_font
            cell.border = border
            if j - 1 in currency_cols:
                cell.number_format = "$#,##0"
            if i % 2 == 1:
                cell.fill = light_fill
    return start_row + len(rows_df) + 1
 
 
def section_header(ws, cell_range, text, color=TEAL_DARK):
    ws.merge_cells(cell_range)
    top_left = cell_range.split(":")[0]
    ws[top_left] = text
    ws[top_left].font = Font(name=FONT_NAME, size=12, bold=True, color=color)
    ws[top_left].alignment = Alignment(horizontal="left", vertical="center")
 
 
def panel_border_only(ws, cell_range):
    first, last = cell_range.split(":")
    fc, fr = coordinate_from_string(first)
    lc, lr = coordinate_from_string(last)
    fci, lci = column_index_from_string(fc), column_index_from_string(lc)
    for r in range(fr, lr + 1):
        for c in range(fci, lci + 1):
            cell = ws.cell(row=r, column=c)
            top_ = med if r == fr else None
            bottom_ = med if r == lr else None
            left_ = med if c == fci else None
            right_ = med if c == lci else None
            if top_ or bottom_ or left_ or right_:
                cell.border = Border(top=top_, bottom=bottom_, left=left_, right=right_)
 
 
def add_image(ws, path, anchor_cell, width_px, height_px):
    img = XLImage(str(path))
    img.width = width_px
    img.height = height_px
    ws.add_image(img, anchor_cell)
 
 
def set_print_setup(ws, print_area):
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    if print_area:
        ws.print_area = print_area
 
 
# ============================================================
# 5. BUILD WORKBOOK
# ============================================================
def build_workbook(agg, chart_paths, output_path: str):
    kpis = agg["kpis"]
    monthly = agg["monthly"]
    n_month_rows = len(monthly)
 
    wb = Workbook()
 
    # ---- Monthly Trend ----
    ws_m = wb.active
    ws_m.title = "Monthly Trend"
    ws_m.sheet_view.showGridLines = False
    write_table(ws_m, ["Month", "Revenue", "Orders", "Customers", "Quantity"], monthly, currency_cols=[1])
    for j, w in enumerate([10, 16, 10, 12, 12], start=1):
        ws_m.column_dimensions[get_column_letter(j)].width = w
 
    # ---- New vs Returning ----
    ws_nr = wb.create_sheet("New vs Returning")
    ws_nr.sheet_view.showGridLines = False
    newret2 = agg["newret"].rename(columns={"New": "New Customer Revenue", "Returning": "Returning Customer Revenue"})
    write_table(ws_nr, ["Month", "New Customer Revenue", "Returning Customer Revenue"], newret2, currency_cols=[1, 2])
    for j, w in enumerate([10, 22, 24], start=1):
        ws_nr.column_dimensions[get_column_letter(j)].width = w
 
    # ---- RFM Segments ----
    ws_seg = wb.create_sheet("RFM Segments")
    ws_seg.sheet_view.showGridLines = False
    seg2 = agg["rfm_segments"].rename(columns={"AvgRecency": "Avg Recency (days)", "AvgFrequency": "Avg Frequency (orders)"})
    seg2["Avg Recency (days)"] = seg2["Avg Recency (days)"].round(0)
    seg2["Avg Frequency (orders)"] = seg2["Avg Frequency (orders)"].round(1)
    next_row = write_table(ws_seg, ["Segment", "Customers", "Revenue", "Avg Recency (days)", "Avg Frequency (orders)"], seg2, currency_cols=[2])
    for j, w in enumerate([20, 12, 16, 18, 20], start=1):
        ws_seg.column_dimensions[get_column_letter(j)].width = w
    ws_seg.cell(row=next_row + 2, column=1, value="Segment definitions").font = Font(name=FONT_NAME, size=11, bold=True, color=TEAL_DARK)
    defs = [
        "Champions — bought recently, buy often, spend the most.",
        "Loyal Customers — buy regularly with good recency and frequency.",
        "Potential Loyalist — recent customers with moderate frequency; room to grow.",
        "New Customers — very recent first-time buyers, frequency still low.",
        "At Risk — used to buy frequently but haven't purchased recently.",
        "Lost — low recency, frequency, and spend across the board.",
    ]
    for i, d in enumerate(defs, start=1):
        ws_seg.cell(row=next_row + 2 + i, column=1, value=d).font = normal_font
 
    # ---- Day-Hour Heatmap ----
    ws_h = wb.create_sheet("Day-Hour Heatmap")
    ws_h.sheet_view.showGridLines = False
    heat = agg["heatmap"]
    ws_h.cell(row=1, column=1, value="Day / Hour").font = header_font
    ws_h.cell(row=1, column=1).fill = charcoal_fill
    ws_h.cell(row=1, column=1).border = border
    for j, hcol in enumerate(heat.columns, start=2):
        c = ws_h.cell(row=1, column=j, value=f"{hcol}:00")
        c.font = header_font
        c.fill = charcoal_fill
        c.alignment = center
        c.border = border
    for i, (day, row) in enumerate(heat.iterrows()):
        r = i + 2
        c = ws_h.cell(row=r, column=1, value=day)
        c.font = bold_font
        c.border = border
        for j, hcol in enumerate(heat.columns, start=2):
            cell = ws_h.cell(row=r, column=j, value=float(row[hcol]))
            cell.number_format = "$#,##0"
            cell.font = Font(name=FONT_NAME, size=9)
            cell.border = border
    from openpyxl.formatting.rule import ColorScaleRule
    rule = ColorScaleRule(start_type="min", start_color="FFFFFF",
                           mid_type="percentile", mid_value=50, mid_color="5EEAD4",
                           end_type="max", end_color=TEAL_DARK)
    last_h_row = len(heat) + 1
    last_h_col = get_column_letter(len(heat.columns) + 1)
    ws_h.conditional_formatting.add(f"B2:{last_h_col}{last_h_row}", rule)
    ws_h.column_dimensions["A"].width = 14
    for j in range(2, len(heat.columns) + 2):
        ws_h.column_dimensions[get_column_letter(j)].width = 11
 
    # ---- Top Products ----
    ws_p = wb.create_sheet("Top Products")
    ws_p.sheet_view.showGridLines = False
    tp = agg["top_products"].rename(columns={"StockCode": "Stock Code", "Quantity": "Quantity Sold"})
    write_table(ws_p, ["Stock Code", "Description", "Revenue", "Quantity Sold", "Orders"], tp, currency_cols=[2])
    for j, w in enumerate([12, 38, 14, 14, 10], start=1):
        ws_p.column_dimensions[get_column_letter(j)].width = w
 
    # ---- Top Countries ----
    ws_c = wb.create_sheet("Top Countries")
    ws_c.sheet_view.showGridLines = False
    write_table(ws_c, ["Country", "Revenue", "Orders", "Customers", "Quantity"], agg["country"], currency_cols=[1])
    ws_c.column_dimensions["A"].width = 20
    for col in ["B", "C", "D", "E"]:
        ws_c.column_dimensions[col].width = 14
 
    # ---- Top Customers ----
    ws_cu = wb.create_sheet("Top Customers")
    ws_cu.sheet_view.showGridLines = False
    tc = agg["top_customers"].copy()
    tc["Customer ID"] = tc["Customer ID"].astype(int)
    write_table(ws_cu, ["Customer ID", "Revenue", "Orders", "Quantity"], tc, currency_cols=[1])
    ws_cu.column_dimensions["A"].width = 14
    for col in ["B", "C", "D"]:
        ws_cu.column_dimensions[col].width = 14
 
    # ---- Notes ----
    ws_n = wb.create_sheet("Notes")
    ws_n.sheet_view.showGridLines = False
    ws_n["A1"] = "Data Preparation & Methodology Notes"
    ws_n["A1"].font = Font(name=FONT_NAME, size=14, bold=True, color=TEAL_DARK)
    notes = [
        "",
        f"Date range: {kpis['date_min']} to {kpis['date_max']}.",
        "",
        "Rows excluded from all revenue/order KPIs and charts:",
        f"  - Cancelled invoices (Invoice starting with \"C\") — {kpis['cancel_orders']:,} orders, "
        f"{kpis['cancel_value']:,.2f} reversed revenue.",
        "  - Line items with Quantity <= 0 or Price <= 0.",
        "",
        "Top Products additionally excludes non-merchandise codes (postage, manual, discount, samples, fees).",
        "",
        "New vs Returning: a customer is \"New\" in the month of their first purchase in this dataset, "
        "\"Returning\" after. Requires a non-null Customer ID.",
        "",
        "RFM Segmentation: Recency = days since last purchase, Frequency = distinct orders, "
        "Monetary = total spend. Each split into quartiles (1-4, 4=best); segments assigned from combined scores.",
        "",
        "Day-Hour Heatmap: revenue summed by weekday and hour of the invoice timestamp.",
        "",
        "Dashboard KPI cells that are plain values (not live formulas) — Unique Customers, Repeat Purchase Rate, "
        "Busiest Day/Hour — are computed from the full cleaned dataset, which isn't stored row-by-row in this "
        "workbook. Re-run this script after refreshing your source data to update them.",
    ]
    for i, line in enumerate(notes, start=2):
        ws_n.cell(row=i, column=1, value=line).font = normal_font
    ws_n.column_dimensions["A"].width = 115
 
    # ============================================================
    # DASHBOARD SHEET
    # ============================================================
    ws = wb.create_sheet("Dashboard", 0)
    ws.sheet_view.showGridLines = False
 
    ws.merge_cells("A1:P2")
    ws["A1"] = "Online Retail — Customer & Sales Intelligence"
    ws["A1"].font = Font(name=FONT_NAME, size=20, bold=True, color=WHITE)
    ws["A1"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    for col in range(1, 17):
        ws.cell(row=1, column=col).fill = PatternFill("solid", fgColor=CHARCOAL)
        ws.cell(row=2, column=col).fill = PatternFill("solid", fgColor=CHARCOAL)
 
    ws.merge_cells("A3:P3")
    ws["A3"] = (f"=\"Period: \" & 'Monthly Trend'!A2 & \" to \" & 'Monthly Trend'!A{n_month_rows + 1} "
                f"& \" | Source data refreshed via script\"")
    ws["A3"].font = Font(name=FONT_NAME, size=10, italic=True, color=SLATE)
    ws.row_dimensions[3].height = 20
 
    kpi_defs = [
        ("TOTAL REVENUE", f"=SUM('Monthly Trend'!B2:B{n_month_rows + 1})", "$#,##0", TEAL_DARK),
        ("TOTAL ORDERS", f"=SUM('Monthly Trend'!C2:C{n_month_rows + 1})", "#,##0", TEAL),
        ("UNIQUE CUSTOMERS", kpis["total_customers"], "#,##0", TEAL),
        ("AVG ORDER VALUE", f"=SUM('Monthly Trend'!B2:B{n_month_rows + 1})/SUM('Monthly Trend'!C2:C{n_month_rows + 1})", "$#,##0.00", SLATE),
        ("REPEAT PURCHASE RATE", kpis["repeat_rate"], "0.0%", AMBER_DARK),
        ("BUSIEST DAY / HOUR", f"{kpis['peak_day']}, {kpis['peak_hour']}:00", "@", CHARCOAL),
    ]
 
    card_cols, card_span, gap_col, row_top = 3, 5, 1, 5
    kpi_value_cells = {}
    for idx, (label, value, fmt, color) in enumerate(kpi_defs):
        r, c = idx // card_cols, idx % card_cols
        c0 = 1 + c * (card_span + gap_col)
        c1 = c0 + card_span - 1
        top = row_top + r * 3
        col0, col1 = get_column_letter(c0), get_column_letter(c1)
 
        ws.merge_cells(f"{col0}{top}:{col1}{top}")
        lc = ws[f"{col0}{top}"]
        lc.value = label
        lc.font = Font(name=FONT_NAME, size=9, bold=True, color=WHITE)
        lc.fill = PatternFill("solid", fgColor=color)
        lc.alignment = Alignment(horizontal="left", vertical="center", indent=1)
 
        ws.merge_cells(f"{col0}{top+1}:{col1}{top+1}")
        vc = ws[f"{col0}{top+1}"]
        vc.value = value
        vc.number_format = fmt
        vc.font = Font(name=FONT_NAME, size=18, bold=True, color=color)
        vc.fill = PatternFill("solid", fgColor=SLATE_LIGHT)
        vc.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        kpi_value_cells[label] = f"{col0}{top+1}"
 
        for cc in range(c0, c1 + 1):
            ws.cell(row=top, column=cc).border = Border(top=med, left=med if cc == c0 else None, right=med if cc == c1 else None)
            ws.cell(row=top + 1, column=cc).border = Border(bottom=med, left=med if cc == c0 else None, right=med if cc == c1 else None)
        ws.row_dimensions[top].height = 16
        ws.row_dimensions[top + 1].height = 30
 
    kpi_block_bottom = row_top + 1 * 3 + 1
    for col in range(1, 17):
        ws.column_dimensions[get_column_letter(col)].width = 8.2
 
    # Document hardcoded KPI cells
    ws[kpi_value_cells["UNIQUE CUSTOMERS"]].comment = Comment(
        "Computed from the full cleaned dataset; not a live formula. Re-run the script to refresh.", "Dashboard build")
    ws[kpi_value_cells["REPEAT PURCHASE RATE"]].comment = Comment(
        "Share of customers with more than one order, from the full cleaned dataset.", "Dashboard build")
    ws[kpi_value_cells["BUSIEST DAY / HOUR"]].comment = Comment(
        "Day/hour with highest revenue, from the Day-Hour Heatmap sheet.", "Dashboard build")
 
    # ---- Chart panels ----
    chart_row1 = kpi_block_bottom + 2
    section_header(ws, f"A{chart_row1}:P{chart_row1}", "Monthly Revenue Trend")
    panel_top1 = chart_row1 + 1
    panel_bottom1 = panel_top1 + 16
    panel_border_only(ws, f"A{panel_top1}:P{panel_bottom1}")
    add_image(ws, chart_paths["monthly_trend"], f"A{panel_top1}", 990, 270)
 
    chart_row2 = panel_bottom1 + 2
    section_header(ws, f"A{chart_row2}:H{chart_row2}", "New vs Returning Customer Revenue")
    section_header(ws, f"I{chart_row2}:P{chart_row2}", "Customer Segments (RFM)")
    panel_top2 = chart_row2 + 1
    panel_bottom2 = panel_top2 + 16
    panel_border_only(ws, f"A{panel_top2}:H{panel_bottom2}")
    panel_border_only(ws, f"I{panel_top2}:P{panel_bottom2}")
    add_image(ws, chart_paths["new_vs_returning"], f"A{panel_top2}", 480, 270)
    add_image(ws, chart_paths["rfm_pie"], f"I{panel_top2}", 480, 270)
 
    chart_row3 = panel_bottom2 + 2
    section_header(ws, f"A{chart_row3}:H{chart_row3}", "Revenue by Country (Top 8)")
    section_header(ws, f"I{chart_row3}:P{chart_row3}", "Top 10 Products by Revenue")
    panel_top3 = chart_row3 + 1
    panel_bottom3 = panel_top3 + 16
    panel_border_only(ws, f"A{panel_top3}:H{panel_bottom3}")
    panel_border_only(ws, f"I{panel_top3}:P{panel_bottom3}")
    add_image(ws, chart_paths["country"], f"A{panel_top3}", 480, 270)
    add_image(ws, chart_paths["top_products"], f"I{panel_top3}", 480, 270)
 
    last_row = panel_bottom3 + 1
    ws.merge_cells(f"A{last_row+1}:P{last_row+1}")
    ws[f"A{last_row+1}"] = ("Full breakdowns and methodology: see Monthly Trend, RFM Segments, "
                             "Day-Hour Heatmap, Top Products, Top Countries, Top Customers, and Notes tabs.")
    ws[f"A{last_row+1}"].font = Font(name=FONT_NAME, size=9, italic=True, color=SLATE)
 
    set_print_setup(ws, f"A1:P{last_row+2}")
    for sheet_name in ["Monthly Trend", "New vs Returning", "RFM Segments", "Day-Hour Heatmap",
                        "Top Products", "Top Countries", "Top Customers", "Notes"]:
        set_print_setup(wb[sheet_name], None)
 
    wb.save(output_path)
    print(f"Saved dashboard to: {output_path}")
 
 
# ============================================================
# MAIN
# ============================================================
def main():
    parser = argparse.ArgumentParser(description="Build an Online Retail Excel dashboard from a CSV file.")
    parser.add_argument("csv_path", help="Path to the transaction-level CSV file")
    parser.add_argument("--output", default="online_retail_dashboard.xlsx", help="Output .xlsx path")
    args = parser.parse_args()
 
    if not Path(args.csv_path).exists():
        sys.exit(f"File not found: {args.csv_path}")
 
    print("Loading and cleaning data...")
    clean, cancel_orders, cancel_value = load_and_clean(args.csv_path)
 
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