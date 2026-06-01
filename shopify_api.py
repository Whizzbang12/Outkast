import os
import requests
import pandas as pd
from config import SHOPIFY_PAID, SALES_START_DATE

def shopify_get(path, params=None):
    store   = os.getenv("SHOPIFY_STORE")
    version = os.getenv("SHOPIFY_API_VERSION", "2024-10")
    url = (path if path.startswith("http")
           else f"https://{store}/admin/api/{version}/{path}")
    r = requests.get(
        url,
        headers={"X-Shopify-Access-Token": os.getenv("SHOPIFY_ACCESS_TOKEN")},
        params=params,
        timeout=60,
    )
    r.raise_for_status()
    nxt = None
    for part in r.headers.get("Link", "").split(","):
        if 'rel="next"' in part:
            nxt = part[part.find("<") + 1:part.find(">")]
    return r.json(), nxt


def shopify_fetch_all(path, key, params=None):
    out, p, nxt = [], dict(params or {}), None
    p.setdefault("limit", 250)
    while True:
        body, nxt = shopify_get(nxt or path, None if nxt else p)
        out.extend(body.get(key, []))
        if not nxt:
            return out


def fetch_shopify():
    """Returns (products df, inventory df, sales df)."""
    raw = shopify_fetch_all("products.json", "products")

    sku_to_item, prod_rows = {}, []
    skus_seen = set()
    skipped_blank = 0

    for prod in raw:
        title = prod.get("title", "")
        for v in prod.get("variants", []):
            iid = v.get("inventory_item_id")
            if not iid:
                continue  # no inventory item = nothing to track

            sku = (v.get("sku") or "").strip()
            if not sku:
                # Variant has no SKU set in Shopify — drop it entirely
                skipped_blank += 1
                continue

            # guard against duplicate SKUs across products
            if sku in skus_seen:
                continue
            skus_seen.add(sku)

            sku_to_item[sku] = iid
            prod_rows.append({
                "sku":             sku,
                "name":            title,
                "cost":            0.0,
                "wholesale_price": float(v.get("price") or 0),
                "msrp":            float(v.get("compare_at_price")
                                         or v.get("price") or 0),
            })

    if skipped_blank:
        print(f"Shopify : {skipped_blank} variants had no SKU "
              f"— DROPPED from the report entirely")

    products = pd.DataFrame(prod_rows).drop_duplicates("sku")

    item_ids = list(sku_to_item.values())

    # unit cost from inventory items
    cost_by_item = {}
    for i in range(0, len(item_ids), 100):
        chunk = item_ids[i:i + 100]
        body, _ = shopify_get("inventory_items.json",
                              {"ids": ",".join(map(str, chunk)), "limit": 250})
        for it in body.get("inventory_items", []):
            cost_by_item[it["id"]] = float(it.get("cost") or 0)
    for idx, row in products.iterrows():
        iid = sku_to_item.get(row["sku"])
        if iid in cost_by_item:
            products.at[idx, "cost"] = cost_by_item[iid]

    # inventory levels (summed across all locations)
    qty_by_item = {}
    for i in range(0, len(item_ids), 50):
        chunk = item_ids[i:i + 50]
        body, _ = shopify_get(
            "inventory_levels.json",
            {"inventory_item_ids": ",".join(map(str, chunk)), "limit": 250},
        )
        for lvl in body.get("inventory_levels", []):
            iid = lvl.get("inventory_item_id")
            qty_by_item[iid] = (qty_by_item.get(iid, 0)
                                + float(lvl.get("available") or 0))

    inventory = pd.DataFrame([
        {"sku": sku, "qty_shopify": qty_by_item.get(iid, 0.0)}
        for sku, iid in sku_to_item.items()
    ])

    # pull all orders back to SALES_START_DATE (Jan 1, 2026)
    cutoff = SALES_START_DATE
    raw_orders = shopify_fetch_all(
        "orders.json", "orders",
        {"status": "any", "created_at_min": cutoff},
    )
    rows = []
    for o in raw_orders:
        if (o.get("financial_status") not in SHOPIFY_PAID
                or o.get("cancelled_at")):
            continue
        date = (o.get("created_at") or "")[:10]
        if date < SALES_START_DATE:   # safety net
            continue
        # order-level tax — split proportionally across line items by revenue
        order_tax   = float(o.get("total_tax") or 0)
        order_rev   = float(o.get("subtotal_price") or 0)
        n_items     = len([i for i in o.get("line_items", [])
                           if float(i.get("quantity") or 0) > 0])
        for item in o.get("line_items", []):
            qty = float(item.get("quantity") or 0)
            sku = item.get("sku") or ""
            if not sku or qty <= 0:
                continue
            price    = float(item.get("price") or 0)
            subtotal = round(price * qty, 2)
            discount = float(item.get("total_discount") or 0)
            net_rev  = round(subtotal - discount, 2)
            # allocate order tax proportionally by line item net revenue
            tax = round(order_tax * (net_rev / order_rev), 2) \
                  if order_rev > 0 else 0.0
            # cost from sku_to_item lookup (fetched separately earlier)
            iid  = sku_to_item.get(sku, 0)
            cost = round(cost_by_item.get(iid, 0) * qty, 2)
            rows.append({
                "order_date":   date,
                "sku":          sku,
                "qty":          qty,
                "subtotal":     subtotal,
                "discount":     discount,
                "net_rev":      net_rev,
                "tax":          tax,
                "cost":         cost,
                "gross_profit": round(net_rev - cost, 2),
            })
    sales = pd.DataFrame(rows)
    print(f"Shopify : {len(products)} products  |  "
          f"{len(inventory)} stock rows  |  {len(sales)} sales lines  "
          f"(from {SALES_START_DATE})")
    return products, inventory, sales