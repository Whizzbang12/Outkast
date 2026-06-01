import os
import requests
from dotenv import load_dotenv

load_dotenv()

def inspect_raw_shopify_skus():
    store = os.getenv("SHOPIFY_STORE")
    token = os.getenv("SHOPIFY_ACCESS_TOKEN")
    version = os.getenv("SHOPIFY_API_VERSION", "2024-10")

    # Limit to 5 just to get a quick snapshot of the data structure
    url = f"https://{store}/admin/api/{version}/products.json?limit=5"
    headers = {"X-Shopify-Access-Token": token}

    print(f"Hitting API: {url}...\n")
    
    response = requests.get(url, headers=headers)
    response.raise_for_status()
    
    products = response.json().get("products", [])
    
    for prod in products:
        print(f"📦 Product: {prod.get('title')}")
        for variant in prod.get("variants", []):
            # Using repr() to explicitly show if it's '' (empty string) or None
            raw_sku = variant.get("sku")
            print(f"  -> Variant ID: {variant.get('id')} | Raw 'sku' value: {repr(raw_sku)}")
        print("-" * 50)

if __name__ == "__main__":
    inspect_raw_shopify_skus()