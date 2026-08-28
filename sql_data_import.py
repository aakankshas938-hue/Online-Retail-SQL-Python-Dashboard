import pandas as pd
from sqlalchemy import create_engine
import time

# ============================================================
# CONFIGURATION
# ============================================================

CSV_FILE = r"D:\Online Retail project\online_retail_II.csv"

# MySQL Configuration
DB_USER = "root"
DB_PASSWORD = "root"
DB_HOST = "127.0.0.1"
DB_PORT = "3306"
DB_NAME = "retail_sales"

# ============================================================
# 1. READ CSV FILE
# ============================================================

print("Reading CSV file...")
try:
    # We use ISO-8859-1 encoding because standard UTF-8 fails on this dataset
    df = pd.read_csv(CSV_FILE, encoding="ISO-8859-1")
    print("CSV loaded successfully! Rows:", len(df))
except Exception as e:
    print("Error reading CSV:", e)
    exit()

# ============================================================
# 2. DATA CLEANING BEFORE SQL IMPORT
# ============================================================

print("\nProcessing data for SQL...")

# Remove spaces from column names (e.g., 'Customer ID' becomes 'Customer_ID')
df.columns = df.columns.str.strip().str.replace(' ', '_')

# Convert InvoiceDate to proper datetime format
df['InvoiceDate'] = pd.to_datetime(df['InvoiceDate'], errors='coerce')

# Drop rows where essential data is missing (to prevent SQL errors)
df = df.dropna(subset=['Invoice', 'StockCode', 'Description', 'Quantity', 'InvoiceDate', 'Price', 'Country'])

# Fill missing Customer_IDs with empty string (as your SQL table allows it)
df['Customer_ID'] = df['Customer_ID'].fillna("").astype(str)

print("Data processed. Ready to insert into MySQL.")

# ============================================================
# 3. CONNECT TO MYSQL & IMPORT DATA
# ============================================================

print("\nConnecting to MySQL...")

# Create connection string
db_url = f"mysql+pymysql://{DB_USER}:{DB_PASSWORD}@{DB_HOST}:{DB_PORT}/{DB_NAME}"

try:
    # Create SQLAlchemy engine (Best and fastest way to push large data to SQL)
    engine = create_engine(db_url)
    
    print(f"Importing {len(df)} rows into '{DB_NAME}.sales' table...")
    print("This may take a few minutes. Please wait...")
    
    start_time = time.time()
    
    # Insert data in chunks of 10,000 rows to avoid memory timeout
    df.to_sql(
        name='sales', 
        con=engine, 
        if_exists='append', # 'append' adds data to existing table. Use 'replace' if you want to delete old and start fresh.
        index=False, 
        chunksize=10000
    )
    
    end_time = time.time()
    
    print("\n" + "=" * 50)
    print("SUCCESS! Data imported to MySQL successfully.")
    print(f"Total time taken: {round((end_time - start_time) / 60, 2)} minutes")
    print("=" * 50)

except Exception as e:
    print("\nERROR: Could not import data to MySQL.")
    print("Details:", e)

