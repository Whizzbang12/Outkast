import requests
import pandas as pd
from dotenv import load_dotenv
from datetime import datetime, timedelta
import os

load_dotenv()   

# Get Token
token_response = requests.post(
    os.getenv("ZOEY_TOKEN_URL"),
    data={
        "grant_type": "client_credentials",
        "client_id": os.getenv("ZOEY_CLIENT_ID"),
        "client_secret": os.getenv("ZOEY_CLIENT_SECRET")
    }
)
access_token = token_response.json()["access_token"]
headers = {"Authorization": f"Bearer {access_token}", "Content-Type": "application/json"}
print("Token received successfully")

# Pull all the data till there are no more pages
def fetch_all_records(endpoint, headers):
    all_records = []
    page = 1
    while True:
        response = requests.get(
            f"{os.getenv('ZOEY_BASE_URL')}/{endpoint}",
            headers=headers,
            params={"limit": 100, "page": page}
        )
        data = response.json()
        records = list(data.values()) if isinstance(data, dict) else data
        if len(records) == 0:
            break
        all_records.extend(records)
        if len(records) < 100:
            break
        page += 1
    return all_records

# Raw data
print("Raw products, stock items, and orders:")
raw_products = fetch_all_records("products", headers)
raw_stock = fetch_all_records("stockitems", headers)
raw_orders = fetch_all_records("orders", headers)

# Tabled data
products_clean = []

for product in raw_products:
    products_clean.append(
        {
            "sku": product.get("sku"),
            "name": product.get("name"),
            "price": float(product.get("price") or 0),
            "cost": float(product.get("cost") or 0),
            "msrp": float(product.get("msrp") or 0)
        }
    )

df_products = pd.DataFrame(products_clean)
print(f"Shape: {df_products.shape}")
print(df_products)

inventory_clean = []

for stock_item in raw_stock:
    inventory_clean.append({
        "sku": stock_item.get("product_sku"),
        "qty_on_hand_b2b": float(stock_item.get("qty") or 0)
    })

df_inventory = pd.DataFrame(inventory_clean)
print(f"Shape: {df_inventory.shape}")
print(df_inventory)

VALID_STATUSES = ["complete", "processing", "pending_payment"]

sales_clean = []
for order in raw_orders:
    if order.get("status") not in VALID_STATUSES:
        continue
    order_date = order.get("created_at", "")[:10]  # Keep only the date part YYYY-MM-DD
    for item in order.get("order_items", []):
        qty_shipped = float(item.get("qty_shipped") or 0)
        if qty_shipped > 0:  # Only include items that were actually shipped
            sales_clean.append({
                "order_date": order_date,
                "sku":        item.get("sku"),
                "qty_shipped": qty_shipped,
                "price":      float(item.get("price") or 0)
            })

df_sales = pd.DataFrame(sales_clean)
df_sales["order_date"] = pd.to_datetime(df_sales["order_date"])
print(f"Shape: {df_sales.shape}")
print(df_sales)


print(f"\nDate range: {df_sales['order_date'].min()} to {df_sales['order_date'].max()}")
print(f"Unique SKUs with sales: {df_sales['sku'].nunique()}") # See how many different products were sold

# Filter 1 year data
cutoff_date = pd.Timestamp.today() - pd.Timedelta(weeks=52)
df_sales_52w = df_sales[df_sales["order_date"] >= cutoff_date].copy()

print(f"\n 1 year sales data:")
print(f"Cutoff date: {cutoff_date.date()}")
print(f"Total line items in last 52 weeks: {len(df_sales_52w)}")
print(f"Unique SKUs with sales in last 52 weeks: {df_sales_52w['sku'].nunique()}")
print(f"Date range: {df_sales_52w['order_date'].min()} to {df_sales_52w['order_date'].max()}")
