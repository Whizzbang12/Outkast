import os
import requests
import pandas as pd
from config import ZOEY_VALID, SALES_START_DATE

def zoey_token():
    r = requests.post(
        os.getenv("ZOEY_TOKEN_URL"),
        data={"grant_type":    "client_credentials",
              "client_id":     os.getenv("ZOEY_CLIENT_ID"),
              "client_secret": os.getenv("ZOEY_CLIENT_SECRET")},
        timeout=30,
    )
    r.raise_for_status()
    return r.json()["access_token"]


def zoey_fetch_all(endpoint, headers, extra_params=None):
    """Page through a Zoey endpoint. extra_params are merged into every request."""
    out, page = [], 1
    base_params = dict(extra_params or {})
    while True:
        base_params["limit"] = 100
        base_params["page"]  = page
        r = requests.get(
            f"{os.getenv('ZOEY_BASE_URL')}/{endpoint}",
            headers=headers,
            params=base_params,
            timeout=60,
        )
        r.raise_for_status()
        data    = r.json()
        records = list(data.values()) if isinstance(data, dict) else data
        if not records:
            return out
        out.extend(records)
        if len(records) < 100:
            return out
        page += 1


def fetch_zoey():
    """Returns (products df, inventory df, sales df)."""
    headers = {"Authorization": f"Bearer {zoey_token()}",
               "Content-Type":  "application/json"}

    raw_p = zoey_fetch_all("products",   headers)
    raw_s = zoey_fetch_all("stockitems", headers)
    # Filter orders to SALES_START_DATE so we get Jan 1 2026 onward
    raw_o = zoey_fetch_all("orders", headers,
                           extra_params={"created_at_from": SALES_START_DATE})

    products = pd.DataFrame([{
        "sku":             p.get("sku"),
        "name":            p.get("name"),
        "cost":            float(p.get("cost")  or 0),
        "wholesale_price": float(p.get("price") or 0),
        "msrp":            float(p.get("msrp")  or 0),
    } for p in raw_p if p.get("sku")]).drop_duplicates("sku")

    inventory = pd.DataFrame([{
        "sku":      s.get("product_sku"),
        "qty_zoey": float(s.get("qty") or 0),
    } for s in raw_s if s.get("product_sku")])

    rows = []
    for o in raw_o:
        if o.get("status") not in ZOEY_VALID:
            continue
        date = (o.get("created_at") or "")[:10]
        if date < SALES_START_DATE:   # safety net
            continue
        for item in o.get("order_items", []):
            qty = float(item.get("qty_shipped") or 0)
            if item.get("sku") and qty > 0:
                subtotal = float(item.get("row_total")       or 0)
                discount = float(item.get("discount_amount") or 0)
                tax      = float(item.get("tax_amount")      or 0)
                cost     = float(item.get("base_cost")       or 0) * qty
                rows.append({
                    "order_date":   date,
                    "sku":          item["sku"],
                    "qty":          qty,
                    "subtotal":     subtotal,
                    "discount":     discount,
                    "net_rev":      subtotal - discount,
                    "tax":          tax,
                    "cost":         cost,
                    "gross_profit": (subtotal - discount) - cost,
                })
    sales = pd.DataFrame(rows)
    print(f"Zoey    : {len(products)} products  |  "
          f"{len(inventory)} stock rows  |  {len(sales)} sales lines  "
          f"(from {SALES_START_DATE})")
    return products, inventory, sales