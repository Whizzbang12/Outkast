import os
import datetime as dt
import pandas as pd


from config import REPORT_DIR
from zoey_api import fetch_zoey
from shopify_api import fetch_shopify
from analysis import build_sku_list, build_metrics, build_weekly_trend, build_daily_sales
from report import build_report, send_reorder_email

def main():
    import boto3
    run_date = dt.date.today()
    os.makedirs(REPORT_DIR, exist_ok=True)
    print("=" * 60)
    print(f"DAILY INVENTORY RUN  {run_date}")
    print("=" * 60)

    # --- Zoey ---
    try:
        z_products, z_inv, z_sales = fetch_zoey()
    except Exception as e:
        print(f"WARNING  Zoey failed: {e}")
        z_products = pd.DataFrame(columns=["sku","name","cost",
                                           "wholesale_price","msrp"])
        z_inv      = pd.DataFrame(columns=["sku","qty_zoey"])
        z_sales    = pd.DataFrame(columns=["order_date","sku","qty"])

    # --- Shopify ---
    try:
        s_products, s_inv, s_sales = fetch_shopify()
    except Exception as e:
        print(f"WARNING  Shopify failed: {e}")
        s_products = pd.DataFrame(columns=["sku","name","cost",
                                           "wholesale_price","msrp"])
        s_inv      = pd.DataFrame(columns=["sku","qty_shopify"])
        s_sales    = pd.DataFrame(columns=["order_date","sku","qty"])

    # --- Build & write ---
    sku_list = build_sku_list(z_products, s_products)
    df       = build_metrics(sku_list, z_inv, s_inv, z_sales, s_sales)
    trend    = build_weekly_trend(sku_list, z_sales, s_sales)
    daily    = build_daily_sales(sku_list, z_sales, s_sales)

    run_datetime = dt.datetime.now().strftime("%Y-%m-%d_%H-%M")
    out = os.path.join(REPORT_DIR, f"Inventory_Report_{run_datetime}.xlsx")
    build_report(df, trend, daily, run_date, out)
    df.to_csv(out.replace(".xlsx", ".csv"), index=False)

    s3_bucket = os.getenv("S3_BUCKET")
    if s3_bucket:
        s3 = boto3.client("s3")
        s3.upload_file(out, s3_bucket, f"reports/{os.path.basename(out)}")
        s3.upload_file(out.replace(".xlsx", ".csv"), s3_bucket, f"reports/{os.path.basename(out.replace('.xlsx', '.csv'))}")
        print(f"Uploaded reports to S3 bucket: {s3_bucket}")


    # --- email reorder alert (priority SKUs only) ---
    send_reorder_email(df, run_date, out)

    print(f"Done.  "
          f"{int((~df.is_discontinued & df.reorder_flag & (df.daily_vel>0)).sum())} need reorder  |  "
          f"{int((~df.is_discontinued & df.overstock_flag & (df.daily_vel>0)).sum())} overstocked  |  "
          f"{int(df.is_discontinued.sum())} discontinued (shown, no alerts).")

if __name__ == "__main__":
    main()