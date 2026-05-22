import requests
import pandas as pd
from dotenv import load_dotenv
import os

load_dotenv()

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

raw_products = fetch_all_records("products", headers)
raw_stock = fetch_all_records("stockitems", headers)
raw_orders = fetch_all_records("orders", headers)
print(f"Fetched {len(raw_products)} products, {len(raw_stock)} stock items, {len(raw_orders)} orders")

df_products = pd.DataFrame([{
    "sku":   product.get("sku"),
    "name":  product.get("name"),
    "price": float(product.get("price") or 0),
    "cost":  float(product.get("cost") or 0),
    "msrp":  float(product.get("msrp") or 0)
} for product in raw_products])

df_inventory = pd.DataFrame([{
    "sku":            stock_item.get("product_sku"),
    "qty_on_hand_b2b": float(stock_item.get("qty") or 0)
} for stock_item in raw_stock])

VALID_STATUSES = ["complete", "processing", "pending_payment"]
sales_clean = []
for order in raw_orders:
    if order.get("status") not in VALID_STATUSES:
        continue
    order_date = order.get("created_at", "")[:10]
    for item in order.get("order_items", []):
        qty_shipped = float(item.get("qty_shipped") or 0)
        if qty_shipped > 0:
            sales_clean.append({
                "order_date":  order_date,
                "sku":         item.get("sku"),
                "qty_shipped": qty_shipped
            })

df_sales = pd.DataFrame(sales_clean)
df_sales["order_date"] = pd.to_datetime(df_sales["order_date"])

today = pd.Timestamp.today()
cutoff_4w  = today - pd.Timedelta(weeks=4)
cutoff_13w = today - pd.Timedelta(weeks=13)

# Units sold in 4w and 32w
units_4w = (
    df_sales[df_sales["order_date"] >= cutoff_4w]
    .groupby("sku")["qty_shipped"]
    .sum()
    .reset_index()
    .rename(columns={"qty_shipped": "units_4w"})
)

units_13w = (
    df_sales[df_sales["order_date"] >= cutoff_13w]
    .groupby("sku")["qty_shipped"]
    .sum()
    .reset_index()
    .rename(columns={"qty_shipped": "units_13w"})
)

# joining the data
df_master = df_products.merge(df_inventory, on="sku", how="left")
df_master = df_master.merge(units_4w,  on="sku", how="left")
df_master = df_master.merge(units_13w, on="sku", how="left")

# Fill NaN with 0 for products with no sales in each window
df_master["units_4w"]  = df_master["units_4w"].fillna(0)
df_master["units_13w"] = df_master["units_13w"].fillna(0)
df_master["qty_on_hand_b2b"] = df_master["qty_on_hand_b2b"].fillna(0)

# velocity
# DailyVel = units sold in last 13 weeks divided by 91 days
# This tells us how many units sell per day on average
df_master["daily_vel"] = df_master["units_13w"] / 91

# days to cover
# DaysCover = current stock divided by daily velocity
# If daily velocity is 0, set days cover to 999 (infinite stock)
# This tells us how many days of stock we have left
df_master["days_cover"] = df_master.apply(
    lambda row: 999 if row["daily_vel"] == 0 else round(row["qty_on_hand_b2b"] / row["daily_vel"], 1),
    axis=1
)

# months to cover
# MonthsCover = DaysCover divided by 30
# This tells us how many months of stock we have left
df_master["months_cover"] = (df_master["days_cover"] / 30).round(1)

# reorder threshold
LEAD_TIME_DAYS = 180
SAFETY_DAYS = 20
REORDER_THRESHOLD = LEAD_TIME_DAYS + SAFETY_DAYS  # 200 days

# ReorderFlag = True if days cover is less than reorder threshold
# This means lee needs to order before stock runs out accounting for lead time
df_master["reorder_flag"] = df_master["days_cover"] < REORDER_THRESHOLD

# OverstockFlag = True if months cover is more than 12 months
df_master["overstock_flag"] = df_master["months_cover"] > 12

# suggested order quantity (might not need it for in this case)
TARGET_DAYS = 180
df_master["suggested_order_qty"] = df_master.apply(
    lambda row: max(0, round((row["daily_vel"] * TARGET_DAYS) - row["qty_on_hand_b2b"])),
    axis=1
)

print("\n METRICS TABLE")
print(f"Shape: {df_master.shape}")
print(df_master[["sku", "name", "qty_on_hand_b2b", "units_4w", "units_13w", "daily_vel", "days_cover", "months_cover", "reorder_flag"]].to_string())

print("\nPRODUCTS FLAGGED FOR REORDER")
reorder_df = df_master[
    (df_master["reorder_flag"] == True) &
    (df_master["daily_vel"] > 0)
].sort_values("days_cover")
print(f"Total products needing reorder: {len(reorder_df)}")
print(reorder_df[["sku", "name", "qty_on_hand_b2b", "daily_vel", "days_cover", "months_cover", "suggested_order_qty"]].to_string())

print("\nPRODUCTS WITH OVERSTOCK")
overstock_df = df_master[
    (df_master["overstock_flag"] == True) &
    (df_master["daily_vel"] > 0)
].sort_values("months_cover", ascending=False)
print(f"Total products with overstock: {len(overstock_df)}")
print(overstock_df[["sku", "name", "qty_on_hand_b2b", "daily_vel", "months_cover"]].to_string())