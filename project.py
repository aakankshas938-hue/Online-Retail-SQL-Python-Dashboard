import pandas as pd
import mysql.connector
import matplotlib.pyplot as plt

# =====================================================
# MYSQL CONNECTION
# =====================================================

# MySQL connection
conn = mysql.connector.connect(
    host="127.0.0.1",
    port=3306,
    user="root",
    password="root",
    database="retail_sales"
)
print("MySQL connected successfully!")


# Load clean data
query = """
SELECT *
FROM clean_sales
"""

df = pd.read_sql(query, conn)

print("\nFirst 5 rows:")
print(df.head())

print("\nDataset Shape:")
print(df.shape)

print("\nColumn Information:")
print(df.info())

print("\nMissing Values:")
print(df.isnull().sum())

print("\nDuplicate Rows:")
print(df.duplicated().sum())

conn.close()

print("\nMySQL connection closed.")

# =====================================================
# PYTHON EDA
# =====================================================

print("\n===== BASIC STATISTICS =====")
print(df.describe())

print("\n===== UNIQUE VALUES =====")
print("Unique Invoices:", df["Invoice"].nunique())
print("Unique Products:", df["StockCode"].nunique())
print("Unique Customers:", df["Customer_ID"].nunique())
print("Unique Countries:", df["Country"].nunique())

print("\n===== SALES SUMMARY =====")
print("Total Revenue:", round(df["Sales_Amount"].sum(), 2))
print("Total Units Sold:", df["Quantity"].sum())
print("Total Orders:", df["Invoice"].nunique())
print("Total Customers:", df["Customer_ID"].nunique())

print("\n===== TOP 10 PRODUCTS BY REVENUE =====")

top_products = (
    df[df["StockCode"] != "DOT"]
    .groupby(["StockCode", "Description"], as_index=False)
    .agg(
        Units_Sold=("Quantity", "sum"),
        Total_Sales=("Sales_Amount", "sum")
    )
    .sort_values("Total_Sales", ascending=False)
    .head(10)
)

top_products["Total_Sales"] = top_products["Total_Sales"].round(2)

print(top_products)

print("\n===== TOP 10 COUNTRIES =====")

country_sales = (
    df[df["StockCode"] != "DOT"]
    .groupby("Country", as_index=False)
    .agg(
        Total_Orders=("Invoice", "nunique"),
        Units_Sold=("Quantity", "sum"),
        Total_Sales=("Sales_Amount", "sum")
    )
    .sort_values("Total_Sales", ascending=False)
    .head(10)
)

country_sales["Total_Sales"] = country_sales["Total_Sales"].round(2)

print(country_sales)


# =====================================================
# MONTHLY SALES TREND
# =====================================================

print("\n===== MONTHLY SALES TREND =====")

# Create Year-Month column
df["Year_Month"] = df["InvoiceDate"].dt.to_period("M").astype(str)

monthly_sales = (
    df.groupby("Year_Month", as_index=False)
      .agg(
          Total_Sales=("Sales_Amount", "sum"),
          Units_Sold=("Quantity", "sum"),
          Orders=("Invoice", "nunique")
      )
      .sort_values("Year_Month")
)

monthly_sales["Total_Sales"] = monthly_sales["Total_Sales"].round(2)

print(monthly_sales)


# =====================================================
# DUPLICATE ANALYSIS
# =====================================================

print("\n===== DUPLICATE ANALYSIS =====")

duplicates = df[df.duplicated(keep=False)]

print("Total duplicate rows:", len(duplicates))

print("\nSample duplicate rows:")
print(duplicates.head(10))

print("\nDuplicate rows by Invoice:")
print(
    duplicates.groupby("Invoice")
             .size()
             .sort_values(ascending=False)
             .head(10)
)


# =====================================================
# TOP 10 CUSTOMERS BY REVENUE
# =====================================================

print("\n===== TOP 10 CUSTOMERS BY REVENUE =====")

top_customers = (
    df.groupby("Customer_ID", as_index=False)
      .agg(
          Total_Orders=("Invoice", "nunique"),
          Units_Sold=("Quantity", "sum"),
          Total_Sales=("Sales_Amount", "sum")
      )
      .sort_values("Total_Sales", ascending=False)
      .head(10)
)

top_customers["Total_Sales"] = top_customers["Total_Sales"].round(2)

print(top_customers)

# =====================================================
# DATA QUALITY CHECK
# =====================================================

print("\n===== DATA QUALITY CHECK =====")

# Check blank Customer IDs
blank_customer = df["Customer_ID"].astype(str).str.strip().eq("").sum()

print("Blank Customer_ID:", blank_customer)


# Check negative quantity
negative_quantity = (df["Quantity"] < 0).sum()

print("Negative Quantity:", negative_quantity)


# Check zero quantity
zero_quantity = (df["Quantity"] == 0).sum()

print("Zero Quantity:", zero_quantity)


# Check zero price
zero_price = (df["Price"] == 0).sum()

print("Zero Price:", zero_price)


# Check negative sales
negative_sales = (df["Sales_Amount"] < 0).sum()

print("Negative Sales:", negative_sales)


# =====================================================
# ORIGINAL DUPLICATE CHECK
# =====================================================

original_columns = [
    "Invoice",
    "StockCode",
    "Description",
    "Quantity",
    "InvoiceDate",
    "Price",
    "Customer_ID",
    "Country",
    "Sales_Amount"
]

original_duplicates = df.duplicated(
    subset=original_columns
).sum()

print("Original Duplicate Rows:", original_duplicates)

# =====================================================
# CUSTOMER ANALYSIS
# =====================================================

print("\n===== CUSTOMER ANALYSIS =====")

# Remove blank Customer_ID only for customer-level analysis
customer_df = df[
    df["Customer_ID"].astype(str).str.strip() != ""
].copy()

print("Valid Customer Records:", len(customer_df))
print("Unique Valid Customers:", customer_df["Customer_ID"].nunique())


# Customer Revenue
customer_sales = (
    customer_df.groupby("Customer_ID", as_index=False)
    .agg(
        Total_Orders=("Invoice", "nunique"),
        Units_Sold=("Quantity", "sum"),
        Total_Sales=("Sales_Amount", "sum")
    )
    .sort_values("Total_Sales", ascending=False)
)

customer_sales["Total_Sales"] = customer_sales["Total_Sales"].round(2)


print("\n===== TOP 10 VALID CUSTOMERS =====")
print(customer_sales.head(10))


# Average Order Value
average_order_value = (
    customer_df.groupby("Invoice")["Sales_Amount"]
    .sum()
    .mean()
)

print("\nAverage Order Value:", round(average_order_value, 2))


# Average Customer Revenue
average_customer_revenue = customer_sales["Total_Sales"].mean()

print("Average Customer Revenue:", round(average_customer_revenue, 2))


# =====================================================
# DATA VISUALIZATION
# =====================================================

# 1. MONTHLY SALES TREND
plt.figure(figsize=(10, 5))
plt.plot(
    monthly_sales["Year_Month"],
    monthly_sales["Total_Sales"],
    marker="o",
    color="crimson"
)
plt.title("Monthly Sales Trend")
plt.xlabel("Month")
plt.ylabel("Total Sales")
plt.xticks(rotation=45)
plt.tight_layout()
plt.show()

# 2. TOP 10 PRODUCTS BY REVENUE
plt.figure(figsize=(10, 6))
plt.barh(
    top_products["Description"].astype(str),
    top_products["Total_Sales"],
    color="royalblue"
)
plt.title("Top 10 Products by Revenue")
plt.xlabel("Total Sales")
plt.ylabel("Product")
plt.gca().invert_yaxis()
plt.tight_layout()
plt.show()

# 3. TOP 10 COUNTRIES BY REVENUE
plt.figure(figsize=(10, 6))
plt.bar(
    country_sales["Country"],
    country_sales["Total_Sales"],
    color="darkorange"
)
plt.title("Top 10 Countries by Revenue")
plt.xlabel("Country")
plt.ylabel("Total Sales")
plt.xticks(rotation=45)
plt.tight_layout()
plt.show()

# 4. TOP 10 CUSTOMERS BY REVENUE
plt.figure(figsize=(10, 6))
top_customer_chart = customer_sales.head(10).sort_values("Total_Sales")
plt.barh(
    top_customer_chart["Customer_ID"].astype(str),
    top_customer_chart["Total_Sales"],
    color="seagreen"
)
plt.title("Top 10 Customers by Revenue")
plt.xlabel("Total Sales")
plt.ylabel("Customer ID")
plt.tight_layout()
plt.show()

# 5. SALES AMOUNT DISTRIBUTION
plt.figure(figsize=(10, 5))
plt.hist(
    df["Sales_Amount"],
    bins=50,
    color="mediumpurple",
    edgecolor="black"
)
plt.title("Sales Amount Distribution")
plt.xlabel("Sales Amount")
plt.ylabel("Frequency")
plt.tight_layout()
plt.show()


# =====================================================
# BUSINESS INSIGHTS
# =====================================================

print("\n===== BUSINESS INSIGHTS =====")

# 1. Best performing country
best_country = country_sales.iloc[0]

print("\n1. Best Performing Country:")
print(
    f"{best_country['Country']} generated "
    f"{best_country['Total_Sales']:.2f} in revenue."
)


# 2. Best performing product
best_product = top_products.iloc[0]

print("\n2. Best Performing Product:")
print(
    f"{best_product['Description']} generated "
    f"{best_product['Total_Sales']:.2f} in revenue."
)


# 3. Best customer
best_customer = customer_sales.iloc[0]

print("\n3. Highest Value Customer:")
print(
    f"Customer {best_customer['Customer_ID']} generated "
    f"{best_customer['Total_Sales']:.2f} in revenue."
)


# 4. Best sales month
best_month = monthly_sales.loc[
    monthly_sales["Total_Sales"].idxmax()
]

print("\n4. Best Sales Month:")
print(
    f"{best_month['Year_Month']} generated "
    f"{best_month['Total_Sales']:.2f} in revenue."
)


# 5. UK revenue contribution
uk_sales = df[
    (df["Country"] == "United Kingdom") &
    (df["StockCode"] != "DOT")
]["Sales_Amount"].sum()

total_sales = df[
    df["StockCode"] != "DOT"
]["Sales_Amount"].sum()

uk_percentage = (uk_sales / total_sales) * 100

print("\n5. UK Revenue Contribution:")
print(f"United Kingdom contributes {uk_percentage:.2f}% of total revenue.")


# 6. Average units per order
avg_units_order = (
    customer_df.groupby("Invoice")["Quantity"]
    .sum()
    .mean()
)

print("\n6. Average Units per Order:")
print(round(avg_units_order, 2))


# 7. Customer data quality
print("\n7. Customer Data Quality:")
print(
    f"{blank_customer} records have blank Customer_ID "
    "and were excluded from customer-level analysis."
)