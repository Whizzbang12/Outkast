import requests
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

def fetch_all_records(endpoint, headers):
    """
    Loops through all pages of an endpoint until it gets an empty page.
    Returns a complete list of all records across all pages.
    """
    all_records = []
    page = 1

    while True:
        print(f"Fetching page {page}...")
        response = requests.get(
            f"{os.getenv('ZOEY_BASE_URL')}/{endpoint}",
            headers=headers,
            params={"limit": 100, "page": page}
        )

        if response.status_code != 200:
            print(f"Error on page {page}: {response.status_code}")
            break

        data = response.json()

        # If data is a dict, convert to list of values
        if isinstance(data, dict):
            records = list(data.values())
        else:
            records = data

        # If empty page, we have reached the end
        if len(records) == 0:
            print(f"Error: Empty page reached at page {page} - done")
            break

        all_records.extend(records)
        print(f"Got {len(records)} records (total so far: {len(all_records)})")

        # If we got less than 100, this was the last page
        if len(records) < 100:
            print(f"Last page reached")
            break

        page += 1

    return all_records

print("Products")
all_products = fetch_all_records("products", headers)
print(f"TOTAL PRODUCTS: {len(all_products)}\n")

print("Stock Items")
all_stock = fetch_all_records("stockitems", headers)
print(f"TOTAL STOCK ITEMS: {len(all_stock)}\n")

print("Orders")
all_orders = fetch_all_records("orders", headers)
print(f"TOTAL ORDERS: {len(all_orders)}\n")

# Count order statuses
from collections import Counter
statuses = [order.get("status") for order in all_orders]
status_counts = Counter(statuses)
print("\nOrder status breakdown:")
for status, count in status_counts.items():
    print(f"  {status}: {count}")