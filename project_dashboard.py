import os
import re
import pandas as pd
import xlsxwriter
from sqlalchemy import create_engine
import urllib.parse

# ============================================================
# CONFIGURATION
# ============================================================

PROJECT_FOLDER = r"D:\Online Retail project"
OUTPUT_FILE = os.path.join(PROJECT_FOLDER, "Online_Retail_SQL_Dashboard.xlsx")

# MySQL Configuration
DB_USER = "root"
DB_PASSWORD = "root"
DB_HOST = "127.0.0.1"
DB_PORT = "3306"
DB_NAME = "retail_sales"

# Excel maximum rows
MAX_EXCEL_ROWS = 1_048_576
CHUNK_SIZE = 25_000

print("=" * 70)
print("ONLINE RETAIL II - SQL TO EXCEL DASHBOARD")
print("=" * 70)

# ============================================================
# 1. CONNECT TO SQL & LOAD DATA
# ============================================================

print("\nConnecting to MySQL database...")

try:
    # Using SQLAlchemy to avoid pandas UserWarning and improve performance
    db_url = f"mysql+pymysql://{DB_USER}:{urllib.parse.quote_plus(DB_PASSWORD)}@{DB_HOST}:{DB_PORT}/{DB_NAME}"
    engine = create_engine(db_url)
    
    print("MySQL connected successfully!")
    print("Reading data from 'clean_sales' table...")
    
    df = pd.read_sql("SELECT * FROM clean_sales", engine)
    
    print("MySQL connection closed.")

except Exception as e:
    print("\nERROR: Could not connect to MySQL or read data.")
    print(e)
    input("\nPress Enter to exit...")
    raise SystemExit

print("Rows loaded from SQL:", len(df))

# ============================================================
# 2. DATA CLEANING & PREPARATION (PYTHON)
# ============================================================

print("\nProcessing data in Python...")

df["InvoiceDate"] = pd.to_datetime(df["InvoiceDate"], errors="coerce")
df["Quantity"] = pd.to_numeric(df["Quantity"], errors="coerce")
df["Price"] = pd.to_numeric(df["Price"], errors="coerce")
df = df.dropna(subset=["Invoice", "StockCode", "InvoiceDate", "Quantity", "Price"]).copy()

df["Customer_ID"] = df["Customer_ID"].fillna("").astype(str).str.strip()
df["Sales_Amount"] = df["Quantity"] * df["Price"]
df = df[(df["Quantity"] > 0) & (df["Price"] > 0)].copy()
df = df.drop_duplicates().copy()

print("Clean rows ready for Dashboard:", len(df))

df["Year"] = df["InvoiceDate"].dt.year
df["Month"] = df["InvoiceDate"].dt.month
df["Year_Month"] = df["InvoiceDate"].dt.to_period("M").astype(str)

customer_df = df[df["Customer_ID"] != ""].copy()
product_df = df[df["StockCode"].astype(str).str.upper().ne("DOT")].copy()

# ============================================================
# 3. ANALYSIS & AGGREGATION (PYTHON)
# ============================================================

print("\nCalculating KPIs and Analytics...")

total_revenue = product_df["Sales_Amount"].sum()
total_units = product_df["Quantity"].sum()
total_orders = product_df["Invoice"].nunique()
total_customers = customer_df["Customer_ID"].nunique()
average_order_value = product_df.groupby("Invoice")["Sales_Amount"].sum().mean()

monthly_sales = (
    product_df.groupby("Year_Month", as_index=False)
    .agg(Total_Sales=("Sales_Amount", "sum"), Units_Sold=("Quantity", "sum"), Orders=("Invoice", "nunique"))
    .sort_values("Year_Month")
)
monthly_sales["Total_Sales"] = monthly_sales["Total_Sales"].round(2)

top_products = (
    product_df.groupby(["StockCode", "Description"], as_index=False)
    .agg(Units_Sold=("Quantity", "sum"), Total_Sales=("Sales_Amount", "sum"))
    .sort_values("Total_Sales", ascending=False).head(10)
)
top_products["Total_Sales"] = top_products["Total_Sales"].round(2)

country_sales = (
    product_df.groupby("Country", as_index=False)
    .agg(Total_Orders=("Invoice", "nunique"), Units_Sold=("Quantity", "sum"), Total_Sales=("Sales_Amount", "sum"))
    .sort_values("Total_Sales", ascending=False).head(10)
)
country_sales["Total_Sales"] = country_sales["Total_Sales"].round(2)

customer_sales = (
    customer_df.groupby("Customer_ID", as_index=False)
    .agg(Total_Orders=("Invoice", "nunique"), Units_Sold=("Quantity", "sum"), Total_Sales=("Sales_Amount", "sum"))
    .sort_values("Total_Sales", ascending=False).head(10)
)
customer_sales["Total_Sales"] = customer_sales["Total_Sales"].round(2)

yearly_sales = (
    product_df.groupby("Year", as_index=False)
    .agg(Total_Sales=("Sales_Amount", "sum"), Units_Sold=("Quantity", "sum"), Orders=("Invoice", "nunique"))
    .sort_values("Year")
)
yearly_sales["Total_Sales"] = yearly_sales["Total_Sales"].round(2)


# ============================================================
# 4. EXCEL WORKBOOK CREATION
# ============================================================

print("\nCreating Excel workbook...")

try:
    workbook = xlsxwriter.Workbook(
        OUTPUT_FILE,
        {"constant_memory": True, "strings_to_urls": False}
    )
except Exception as e:
    print("Could not create workbook:", e)
    input("\nPress Enter to exit...")
    raise SystemExit

# Colors
NAVY = "#17365D"; BLUE = "#4472C4"; GREEN = "#70AD47"; ORANGE = "#ED7D31"
PURPLE = "#8064A2"; TEAL = "#00A6A6"; RED = "#C00000"; WHITE = "#FFFFFF"; GREY = "#666666"

# Formats
title_format = workbook.add_format({"bold": True, "font_size": 24, "font_color": WHITE, "bg_color": NAVY, "align": "center", "valign": "vcenter"})
subtitle_format = workbook.add_format({"italic": True, "font_size": 11, "font_color": GREY, "align": "center", "valign": "vcenter"})
header_format = workbook.add_format({"bold": True, "font_color": WHITE, "bg_color": NAVY, "align": "center", "valign": "vcenter", "border": 1, "border_color": "#D9D9D9"})
money_format = workbook.add_format({"num_format": "£#,##0.00"})
number_format = workbook.add_format({"num_format": "#,##0"})
insight_cell_format = workbook.add_format({"border": 1})
insight_money_format = workbook.add_format({"border": 1, "num_format": "£#,##0.00"})
quality_title_format = workbook.add_format({"bold": True, "font_size": 14, "font_color": WHITE, "bg_color": RED})

kpi_label_formats = {
    "green": workbook.add_format({"bold": True, "font_color": WHITE, "bg_color": GREEN, "align": "center", "valign": "vcenter"}),
    "blue": workbook.add_format({"bold": True, "font_color": WHITE, "bg_color": BLUE, "align": "center", "valign": "vcenter"}),
    "orange": workbook.add_format({"bold": True, "font_color": WHITE, "bg_color": ORANGE, "align": "center", "valign": "vcenter"}),
    "purple": workbook.add_format({"bold": True, "font_color": WHITE, "bg_color": PURPLE, "align": "center", "valign": "vcenter"}),
    "teal": workbook.add_format({"bold": True, "font_color": WHITE, "bg_color": TEAL, "align": "center", "valign": "vcenter"})
}
kpi_value_formats = {
    "green": workbook.add_format({"bold": True, "font_size": 18, "font_color": GREEN, "bg_color": WHITE, "align": "center", "valign": "vcenter", "num_format": "£#,##0.00"}),
    "blue": workbook.add_format({"bold": True, "font_size": 18, "font_color": BLUE, "bg_color": WHITE, "align": "center", "valign": "vcenter", "num_format": "#,##0"}),
    "orange": workbook.add_format({"bold": True, "font_size": 18, "font_color": ORANGE, "bg_color": WHITE, "align": "center", "valign": "vcenter", "num_format": "#,##0"}),
    "purple": workbook.add_format({"bold": True, "font_size": 18, "font_color": PURPLE, "bg_color": WHITE, "align": "center", "valign": "vcenter", "num_format": "#,##0"}),
    "teal": workbook.add_format({"bold": True, "font_size": 18, "font_color": TEAL, "bg_color": WHITE, "align": "center", "valign": "vcenter", "num_format": "£#,##0.00"})
}
section_formats = {
    "blue": workbook.add_format({"bold": True, "font_size": 14, "font_color": WHITE, "bg_color": BLUE}),
    "green": workbook.add_format({"bold": True, "font_size": 14, "font_color": WHITE, "bg_color": GREEN}),
    "orange": workbook.add_format({"bold": True, "font_size": 14, "font_color": WHITE, "bg_color": ORANGE}),
    "purple": workbook.add_format({"bold": True, "font_size": 14, "font_color": WHITE, "bg_color": PURPLE}),
    "teal": workbook.add_format({"bold": True, "font_size": 14, "font_color": WHITE, "bg_color": TEAL})
}

def write_dataframe(worksheet, dataframe, start_row=0, start_col=0, header_fmt=None, formats=None):
    if dataframe is None or dataframe.empty: return
    formats = formats or {}
    for col_num, column_name in enumerate(dataframe.columns):
        worksheet.write(start_row, start_col + col_num, column_name, header_fmt)
    for row_num, row in enumerate(dataframe.itertuples(index=False), start_row + 1):
        for col_num, value in enumerate(row):
            if pd.isna(value): value = ""
            if isinstance(value, str):
                value = re.compile(r"[\x00-\x08\x0B\x0C\x0E-\x1F]").sub("", value)
            cell_format = formats.get(dataframe.columns[col_num])
            worksheet.write(row_num, start_col + col_num, value, cell_format)

# --- SHEET 1: DASHBOARD ---
dashboard = workbook.add_worksheet("Dashboard")
dashboard.hide_gridlines(2)
dashboard.set_tab_color(NAVY)
dashboard.set_zoom(90)
dashboard.merge_range("A1:L2", "ONLINE RETAIL SALES DASHBOARD", title_format)
dashboard.merge_range("A3:L3", "Online Retail II | Sales Performance Analysis", subtitle_format)

kpis = [
    ("TOTAL REVENUE", total_revenue, "green", "£#,##0.00"),
    ("TOTAL ORDERS", total_orders, "blue", "#,##0"),
    ("TOTAL UNITS", total_units, "orange", "#,##0"),
    ("CUSTOMERS", total_customers, "purple", "#,##0"),
    ("AVG ORDER VALUE", average_order_value, "teal", "£#,##0.00")
]
kpi_columns = [("A", "B"), ("C", "D"), ("E", "F"), ("G", "H"), ("I", "J")]

for kpi, position in zip(kpis, kpi_columns):
    label, value, color, num_format = kpi
    start_col, end_col = position
    dashboard.merge_range(f"{start_col}5:{end_col}5", label, kpi_label_formats[color])
    dashboard.merge_range(f"{start_col}6:{end_col}7", value, kpi_value_formats[color])

for col in range(12):
    dashboard.set_column(col, col, 14)

# --- SHEET 2: ANALYSIS ---
analysis = workbook.add_worksheet("Analysis")
analysis.hide_gridlines(2)
analysis.set_tab_color(BLUE)
analysis.freeze_panes(2, 0)

analysis.write("A1", "MONTHLY SALES", section_formats["blue"])
write_dataframe(analysis, monthly_sales, 1, 0, header_format, {"Total_Sales": money_format, "Units_Sold": number_format, "Orders": number_format})
analysis.write("F1", "TOP 10 PRODUCTS", section_formats["green"])
write_dataframe(analysis, top_products, 1, 5, header_format, {"Units_Sold": number_format, "Total_Sales": money_format})
analysis.write("K1", "TOP 10 COUNTRIES", section_formats["orange"])
write_dataframe(analysis, country_sales, 1, 10, header_format, {"Total_Orders": number_format, "Units_Sold": number_format, "Total_Sales": money_format})
analysis.write("Q1", "TOP 10 CUSTOMERS", section_formats["purple"])
write_dataframe(analysis, customer_sales, 1, 16, header_format, {"Total_Orders": number_format, "Units_Sold": number_format, "Total_Sales": money_format})
analysis.write("W1", "YEARLY SALES", section_formats["teal"])
write_dataframe(analysis, yearly_sales, 1, 22, header_format, {"Total_Sales": money_format, "Units_Sold": number_format, "Orders": number_format})

# --- CHARTS (Sizes reduced for better fit) ---
monthly_chart = workbook.add_chart({"type": "line"})
monthly_chart.add_series({
    "name": "Total Sales",
    "categories": ["Analysis", 2, 0, len(monthly_sales) + 1, 0],
    "values": ["Analysis", 2, 1, len(monthly_sales) + 1, 1],
    "line": {"color": BLUE, "width": 2.5},
    "marker": {"type": "circle", "size": 6, "border": {"color": BLUE}, "fill": {"color": BLUE}}
})
monthly_chart.set_title({"name": "Monthly Sales Trend"})
monthly_chart.set_x_axis({"name": "Month"})
monthly_chart.set_y_axis({"name": "Total Sales", "num_format": "£#,##0"})
monthly_chart.set_legend({"none": True})
monthly_chart.set_size({"width": 500, "height": 280})

product_chart = workbook.add_chart({"type": "bar"})
product_chart.add_series({
    "name": "Revenue",
    "categories": ["Analysis", 2, 6, len(top_products) + 1, 6],
    "values": ["Analysis", 2, 7, len(top_products) + 1, 7],
    "fill": {"color": BLUE}, "border": {"none": True}
})
product_chart.set_title({"name": "Top 10 Products by Revenue"})
product_chart.set_x_axis({"name": "Revenue", "num_format": "£#,##0"})
product_chart.set_y_axis({"name": "Product"})
product_chart.set_legend({"none": True})
product_chart.set_size({"width": 500, "height": 280})

country_chart = workbook.add_chart({"type": "column"})
country_chart.add_series({
    "name": "Revenue",
    "categories": ["Analysis", 2, 10, len(country_sales) + 1, 10],
    "values": ["Analysis", 2, 13, len(country_sales) + 1, 13],
    "fill": {"color": GREEN}, "border": {"none": True}
})
country_chart.set_title({"name": "Top 10 Countries by Revenue"})
country_chart.set_x_axis({"name": "Country"})
country_chart.set_y_axis({"name": "Revenue", "num_format": "£#,##0"})
country_chart.set_legend({"none": True})
country_chart.set_size({"width": 500, "height": 280})

customer_chart = workbook.add_chart({"type": "bar"})
customer_chart.add_series({
    "name": "Revenue",
    "categories": ["Analysis", 2, 16, len(customer_sales) + 1, 16],
    "values": ["Analysis", 2, 19, len(customer_sales) + 1, 19],
    "fill": {"color": PURPLE}, "border": {"none": True}
})
customer_chart.set_title({"name": "Top 10 Customers by Revenue"})
customer_chart.set_x_axis({"name": "Revenue", "num_format": "£#,##0"})
customer_chart.set_y_axis({"name": "Customer"})
customer_chart.set_legend({"none": True})
customer_chart.set_size({"width": 500, "height": 280})

yearly_chart = workbook.add_chart({"type": "column"})
yearly_chart.add_series({
    "name": "Revenue",
    "categories": ["Analysis", 2, 22, len(yearly_sales) + 1, 22],
    "values": ["Analysis", 2, 23, len(yearly_sales) + 1, 23],
    "fill": {"color": TEAL}, "border": {"none": True}
})
yearly_chart.set_title({"name": "Yearly Sales Performance"})
yearly_chart.set_x_axis({"name": "Year"})
yearly_chart.set_y_axis({"name": "Revenue", "num_format": "£#,##0"})
yearly_chart.set_legend({"none": True})
yearly_chart.set_size({"width": 500, "height": 280})

# Positions adjusted so they fit on screen without scrolling down too much
dashboard.insert_chart("A10", monthly_chart)
dashboard.insert_chart("G10", product_chart)
dashboard.insert_chart("A27", country_chart)
dashboard.insert_chart("G27", customer_chart)
dashboard.insert_chart("A44", yearly_chart)

# --- SHEET 3: CLEAN DATA ---
data_sheet = workbook.add_worksheet("Clean_Data")
data_sheet.hide_gridlines(2)
data_sheet.set_tab_color(GREEN)
data_sheet.freeze_panes(1, 0)

clean_money_format = workbook.add_format({"num_format": "£#,##0.00"})
clean_number_format = workbook.add_format({"num_format": "#,##0"})
clean_date_format = workbook.add_format({"num_format": "dd-mm-yyyy hh:mm"})

columns = df.columns.tolist()
export_df = df.iloc[:MAX_EXCEL_ROWS - 1] if len(df) > MAX_EXCEL_ROWS - 1 else df
total_rows = len(export_df)

print(f"\nWriting Clean_Data sheet: {total_rows} rows...")

for col_num, column_name in enumerate(columns):
    width = 40 if column_name == "Description" else 22 if column_name == "Country" else 20 if column_name == "InvoiceDate" else 16 if column_name in ["StockCode", "Customer_ID"] else 15 if column_name in ["Quantity", "Price", "Sales_Amount"] else 14
    data_sheet.set_column(col_num, col_num, width)

for start in range(0, total_rows, CHUNK_SIZE):
    end = min(start + CHUNK_SIZE, total_rows)
    chunk = export_df.iloc[start:end]
    for row_offset, row in enumerate(chunk.itertuples(index=False), start):
        excel_row = row_offset + 1
        for col_num, value in enumerate(row):
            column_name = columns[col_num]
            if pd.isna(value): value = ""
            if isinstance(value, str): value = re.compile(r"[\x00-\x08\x0B\x0C\x0E-\x1F]").sub("", value)
            elif isinstance(value, pd.Timestamp): value = value.to_pydatetime()
            
            if column_name == "InvoiceDate": data_sheet.write_datetime(excel_row, col_num, value, clean_date_format)
            elif column_name == "Sales_Amount": data_sheet.write_number(excel_row, col_num, float(value), clean_money_format)
            elif column_name == "Quantity": data_sheet.write_number(excel_row, col_num, float(value), clean_number_format)
            elif column_name == "Price": data_sheet.write_number(excel_row, col_num, float(value), clean_money_format)
            elif isinstance(value, (int, float)): data_sheet.write_number(excel_row, col_num, float(value))
            else: data_sheet.write(excel_row, col_num, value)
    print(f"  Written {end:,} / {total_rows:,} rows")

for col_num, column_name in enumerate(columns):
    data_sheet.write_number(0, col_num, 0, header_format)
    data_sheet.write_string(0, col_num, column_name, header_format)

data_sheet.autofilter(0, 0, len(export_df), len(columns) - 1)

# --- SHEET 4: BUSINESS INSIGHTS ---
insights = workbook.add_worksheet("Business_Insights")
insights.hide_gridlines(2)
insights.set_tab_color(PURPLE)
insights.merge_range("A1:F2", "BUSINESS INSIGHTS", title_format)

best_country = country_sales.iloc[0]
best_product = top_products.iloc[0]
best_customer = customer_sales.iloc[0]
best_month = monthly_sales.loc[monthly_sales["Total_Sales"].idxmax()]

insight_rows = [
    ("1", "Best Performing Country", str(best_country["Country"]), float(best_country["Total_Sales"])),
    ("2", "Best Performing Product", str(best_product["Description"]), float(best_product["Total_Sales"])),
    ("3", "Highest Value Customer", str(best_customer["Customer_ID"]), float(best_customer["Total_Sales"])),
    ("4", "Best Sales Month", str(best_month["Year_Month"]), float(best_month["Total_Sales"])),
    ("5", "Average Order Value", "Average revenue per order", float(average_order_value))
]

insight_header_format = workbook.add_format({"bold": True, "font_color": WHITE, "bg_color": NAVY, "align": "center", "border": 1})
headers = ["No.", "Metric", "Result", "Value"]
for col_num, header in enumerate(headers):
    insights.write(3, col_num, header, insight_header_format)

for row_num, row_data in enumerate(insight_rows, 4):
    for col_num, value in enumerate(row_data):
        if col_num == 3: insights.write_number(row_num, col_num, value, insight_money_format)
        else: insights.write(row_num, col_num, value, insight_cell_format)

insights.write(12, 0, "DATA QUALITY", quality_title_format)
blank_customers = int(df["Customer_ID"].astype(str).str.strip().eq("").sum())
quality_items = [
    ("Blank Customer IDs", blank_customers),
    ("Clean Dataset Rows", len(df)),
    ("Unique Products", df["StockCode"].nunique()),
    ("Unique Countries", df["Country"].nunique())
]

for row_num, item in enumerate(quality_items, 14):
    insights.write(row_num, 0, item[0], insight_cell_format)
    insights.write_number(row_num, 1, int(item[1]), insight_cell_format)

insights.set_column("A:A", 30); insights.set_column("B:B", 30); insights.set_column("C:C", 45); insights.set_column("D:D", 20)

# --- CLOSE WORKBOOK ---
print("\nSaving Excel dashboard...")
try:
    workbook.close()
except PermissionError:
    print("\nERROR: Excel file is currently open. Please close it.")
    input("\nPress Enter to exit...")
    raise SystemExit
except Exception as e:
    print("\nERROR while saving workbook:", e)
    input("\nPress Enter to exit...")
    raise SystemExit

print("\n" + "=" * 70)
print("SUCCESS! Excel Dashboard created successfully.")
print("Output:", OUTPUT_FILE)
print("=" * 70)

input("\nPress Enter to exit...")