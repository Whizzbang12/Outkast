import pandas as pd
from config import (
    DISCONTINUED_SKUS,
    DAYS_4W, DAYS_13W, DAYS_52W,
    DEFAULT_LEAD_DAYS,
    SAFETY_DAYS, MONTHS_DIVISOR, OVERSTOCK_MONTHS,
    BLEND_MSRP, BLEND_WHOLESALE,
    WEEKLY_TREND_WEEKS,
    SALES_START_DATE, DAILY_DAYS
)

FIN_COLS = ["qty", "subtotal", "discount", "net_rev", "tax", "cost", "gross_profit"]

def _ensure_fin_cols(sales):
    """Guarantee all financial columns exist even if the API returned nothing."""
    for c in FIN_COLS:
        if c not in sales.columns:
            sales[c] = 0.0
    return sales

def _window(sales, days, col):
    if sales.empty:
        return pd.DataFrame(columns=["sku", col])
    cut = pd.Timestamp.today() - pd.Timedelta(days=days)
    return (sales[sales["order_date"] >= cut]
            .groupby("sku")["qty"].sum()
            .reset_index()
            .rename(columns={"qty": col}))


def _rev_window(sales, days, col):
    """Sum net_rev per SKU within the last `days` days."""
    if sales.empty or "net_rev" not in sales.columns:
        return pd.DataFrame(columns=["sku", col])
    cut = pd.Timestamp.today() - pd.Timedelta(days=days)
    s = sales.copy()
    s["order_date"] = pd.to_datetime(s["order_date"], errors="coerce")
    return (s[s["order_date"] >= cut]
            .groupby("sku")["net_rev"].sum()
            .reset_index()
            .rename(columns={"net_rev": col}))

def build_sku_list(z_products, s_products):
    """
    Combines products from both platforms.
    Zoey is preferred source for name/cost (B2B master catalogue).
    ALL SKUs included — discontinued ones are tagged so alert tabs
    can filter them out, but they remain visible in the full report.
    """
    all_skus = (pd.concat([z_products, s_products], ignore_index=True)
                .drop_duplicates("sku", keep="first"))

    all_skus["is_discontinued"] = all_skus["sku"].isin(DISCONTINUED_SKUS)
    n_disc = int(all_skus["is_discontinued"].sum())

    print(f"SKU list: {len(all_skus)} total products  "
          f"({n_disc} discontinued — included in report, "
          f"excluded from alerts)")
    return all_skus.reset_index(drop=True)

def build_metrics(sku_list, z_inv, s_inv, z_sales, s_sales):
    df = sku_list.copy()

    # ---- inventory ----
    df = (df.merge(z_inv, on="sku", how="left")
            .merge(s_inv, on="sku", how="left"))
    df["qty_zoey"]    = df["qty_zoey"].fillna(0)
    df["qty_shopify"] = df["qty_shopify"].fillna(0)
    df["on_hand"]     = df["qty_zoey"] + df["qty_shopify"]

    # ---- combined sales (de-dup same sku+date across platforms) ----
    all_sales = pd.concat([z_sales, s_sales], ignore_index=True)
    if not all_sales.empty:
        all_sales["order_date"] = pd.to_datetime(
            all_sales["order_date"], errors="coerce")
        all_sales = (all_sales
                     .dropna(subset=["order_date", "sku"])
                     .groupby(["order_date", "sku"])["qty"]
                     .sum().reset_index())

    # ---- sales windows ----
    df = df.merge(_window(all_sales, DAYS_4W,  "units_4w"),  on="sku", how="left")
    df = df.merge(_window(all_sales, DAYS_13W, "units_13w"), on="sku", how="left")
    df = df.merge(_window(all_sales, DAYS_52W, "units_52w"), on="sku", how="left")
    df["units_4w"]  = df["units_4w"].fillna(0)
    df["units_13w"] = df["units_13w"].fillna(0)
    df["units_52w"] = df["units_52w"].fillna(0)

    # ---- per-platform revenue (last 30 days) for Summary page ----
    z_sales_ts = z_sales.copy() if not z_sales.empty else pd.DataFrame(columns=["order_date","sku","qty","net_rev"])
    s_sales_ts = s_sales.copy() if not s_sales.empty else pd.DataFrame(columns=["order_date","sku","qty","net_rev"])
    if not z_sales_ts.empty:
        z_sales_ts["order_date"] = pd.to_datetime(z_sales_ts["order_date"], errors="coerce")
    if not s_sales_ts.empty:
        s_sales_ts["order_date"] = pd.to_datetime(s_sales_ts["order_date"], errors="coerce")
    df = df.merge(_rev_window(z_sales_ts, 30, "rev_zoey_30d"),    on="sku", how="left")
    df = df.merge(_rev_window(s_sales_ts, 30, "rev_shopify_30d"),  on="sku", how="left")
    df = df.merge(_rev_window(z_sales_ts,  7, "rev_zoey_7d"),     on="sku", how="left")
    df = df.merge(_rev_window(s_sales_ts,  7, "rev_shopify_7d"),   on="sku", how="left")
    for c in ("rev_zoey_30d","rev_shopify_30d","rev_zoey_7d","rev_shopify_7d"):
        df[c] = df[c].fillna(0).round(2)
    df["rev_total_30d"] = (df["rev_zoey_30d"] + df["rev_shopify_30d"]).round(2)
    df["rev_total_7d"]  = (df["rev_zoey_7d"]  + df["rev_shopify_7d"]).round(2)

    # ---- velocity: use 4W if recent sales exist, else fall back to 13W ----
    df["daily_vel_4w"]  = df["units_4w"]  / DAYS_4W
    df["daily_vel_13w"] = df["units_13w"] / DAYS_13W
    df["daily_vel_52w"] = df["units_52w"] / DAYS_52W
    df["daily_vel"] = df["daily_vel_52w"]  # changed to 52W velocity

    # ---- days / months cover ----
    df["days_cover"] = df.apply(
        lambda r: 999.0 if r["daily_vel"] == 0
        else round(r["on_hand"] / r["daily_vel"], 1), axis=1)
    df["months_cover"] = (df["days_cover"] / MONTHS_DIVISOR).round(1)

    # ---- reorder ----
    df["lead_time_days"]      = DEFAULT_LEAD_DAYS
    df["target_units"]        = (
        df["daily_vel"] * (DEFAULT_LEAD_DAYS + SAFETY_DAYS)).round(0)
    df["suggested_order_qty"] = (
        (df["target_units"] - df["on_hand"]).clip(lower=0).round(0))
    

    ###### FIX ROUNDED ORDER
    # df["rounded_order_qty"]   = df.apply(
    #     lambda r: 0 if r["suggested_order_qty"] <= 0
    #     else int(-(-r["suggested_order_qty"] // DEFAULT_MOQ) * DEFAULT_MOQ),
    #     axis=1)

    # ---- flags ----
    df["reorder_flag"]   = df["suggested_order_qty"].round(0) > 0
    df["overstock_flag"] = df["months_cover"] > OVERSTOCK_MONTHS
    # discontinued tag carried from sku_list (False for new products)
    if "is_discontinued" not in df.columns:
        df["is_discontinued"] = False
    df["is_discontinued"] = df["is_discontinued"].fillna(False)

    # ---- value columns ----
    df["inventory_value"]  = (df["cost"] * df["on_hand"]).round(2)
    df["excess_inv_value"] = (
        (df["on_hand"] - df["target_units"]).clip(lower=0)
        * df["cost"]).round(2)
    df["blended_price"] = (
        df["msrp"] * BLEND_MSRP
        + df["wholesale_price"] * BLEND_WHOLESALE).round(2)
    df["gp_per_sku"] = (
        (df["blended_price"] - df["cost"]) * df["on_hand"]).round(2)

    # ---- sort: lowest days-cover first (most urgent at top) ----
    return df.sort_values(
        ["reorder_flag", "days_cover"],
        ascending=[False, True]
    ).reset_index(drop=True)

def build_weekly_trend(sku_list, z_sales, s_sales):
    """
    Returns (DataFrame, week_labels).
    Per week columns: Units | Zoey Rev | Shopify Rev | Total Net Rev | GP
    Plus 13W totals for each.
    """
    all_sales = pd.concat([z_sales, s_sales], ignore_index=True)
    all_sales = _ensure_fin_cols(all_sales)
    base = sku_list[["sku", "name"]].copy()

    if all_sales.empty:
        return base, []

    # tag platform before merging so we can split revenue
    z_s = _ensure_fin_cols(z_sales.copy()) if not z_sales.empty \
          else pd.DataFrame(columns=["order_date","sku","qty","net_rev"])
    s_s = _ensure_fin_cols(s_sales.copy()) if not s_sales.empty \
          else pd.DataFrame(columns=["order_date","sku","qty","net_rev"])

    for s in [z_s, s_s, all_sales]:
        s["order_date"] = pd.to_datetime(s["order_date"], errors="coerce")

    def add_week(s):
        s = s.dropna(subset=["order_date","sku"]).copy()
        s["week"] = (
            s["order_date"]
            - pd.to_timedelta((s["order_date"].dt.weekday + 1) % 7, unit="D")
        ).dt.normalize()
        return s

    z_s      = add_week(z_s)
    s_s      = add_week(s_s)
    all_s    = add_week(all_sales.dropna(subset=["order_date","sku"]).copy())

    recent_weeks = sorted(all_s["week"].unique())[-WEEKLY_TREND_WEEKS:]
    z_s      = z_s[z_s["week"].isin(recent_weeks)]
    s_s      = s_s[s_s["week"].isin(recent_weeks)]
    all_s    = all_s[all_s["week"].isin(recent_weeks)]
    week_labels = [w.strftime("Wk %b %d/%Y") for w in recent_weeks]

    trend = base.copy()

    def pivot_metric(df_src, metric, suffix=""):
        if df_src.empty or metric not in df_src.columns:
            return pd.DataFrame({"sku": base["sku"].tolist()})
        piv = (df_src.pivot_table(index="sku", columns="week",
                                  values=metric, aggfunc="sum", fill_value=0)
                     .reset_index())
        piv.columns = (["sku"]
                       + [f"{w.strftime('Wk %b %d/%Y')}{suffix}"
                          for w in piv.columns[1:]])
        return piv

    # build all pivots and concat once per metric group
    new_cols = {}
    for wl in week_labels:
        new_cols[f"{wl} Units"]       = 0.0
        new_cols[f"{wl} Zoey Rev"]    = 0.0
        new_cols[f"{wl} Shopify Rev"] = 0.0
        new_cols[f"{wl} Net Rev"]     = 0.0
        new_cols[f"{wl} GP"]          = 0.0

    # units (combined)
    piv_u = pivot_metric(all_s, "qty")
    for wl, w in zip(week_labels, recent_weeks):
        col = wl   # pivot col name is just the week label
        for sku in base["sku"]:
            row = piv_u[piv_u["sku"] == sku]
            new_cols[f"{wl} Units"] = new_cols.get(f"{wl} Units", {})
    # easier: just merge pivots
    trend = base.copy()

    piv_units = pivot_metric(all_s, "qty")
    piv_zrev  = pivot_metric(z_s,   "net_rev", " zrev")
    piv_srev  = pivot_metric(s_s,   "net_rev", " srev")
    piv_rev   = pivot_metric(all_s, "net_rev", " rev")
    piv_gp    = pivot_metric(all_s, "gross_profit", " gp")

    for piv in [piv_units, piv_zrev, piv_srev, piv_rev, piv_gp]:
        trend = trend.merge(piv, on="sku", how="left")

    # rename and fill
    for wl, w in zip(week_labels, recent_weeks):
        ws = w.strftime("Wk %b %d/%Y")
        renames = {
            ws:          f"{wl} Units",
            f"{ws} zrev":f"{wl} Zoey Rev",
            f"{ws} srev":f"{wl} Shopify Rev",
            f"{ws} rev": f"{wl} Net Rev",
            f"{ws} gp":  f"{wl} GP",
        }
        trend = trend.rename(columns=renames)

    all_metric_cols = []
    for wl in week_labels:
        for m in ("Units","Zoey Rev","Shopify Rev","Net Rev","GP"):
            col = f"{wl} {m}"
            if col not in trend.columns:
                trend[col] = 0.0
            trend[col] = trend[col].fillna(0).round(2)
            all_metric_cols.append(col)

    # 13W totals
    for m in ("Units","Zoey Rev","Shopify Rev","Net Rev","GP"):
        mcols = [f"{wl} {m}" for wl in week_labels]
        trend[f"13W {m}"] = trend[mcols].sum(axis=1).round(2)

    return trend, week_labels

def build_daily_sales(sku_list, z_sales, s_sales):
    """
    Returns (df, summary) where df has:
      - SKU, Product
      - Per day: Units, Subtotal, Discount, Net Revenue, Cost, GP
      - 7D / 30D totals for each metric
      - Trend arrow (units, last 7 vs prior 7)
    summary = dict of per-day totals across all products.
    """
    today       = pd.Timestamp.today().normalize()
    dates       = [today - pd.Timedelta(days=i) for i in range(DAILY_DAYS-1, -1, -1)]
    date_labels = [d.strftime("%b %d, %Y") for d in dates]
    date_strs   = [d.strftime("%Y-%m-%d") for d in dates]
    base        = sku_list[["sku", "name"]].copy()

    # tag each sale with its platform so we can split qty by source
    z = _ensure_fin_cols(z_sales.copy()) if not z_sales.empty else \
        pd.DataFrame(columns=["order_date","sku","qty","net_rev"])
    s = _ensure_fin_cols(s_sales.copy()) if not s_sales.empty else \
        pd.DataFrame(columns=["order_date","sku","qty","net_rev"])
    z["qty_zoey"]    = z["qty"]
    z["qty_shopify"] = 0.0
    s["qty_zoey"]    = 0.0
    s["qty_shopify"] = s["qty"]

    all_sales = pd.concat([z, s], ignore_index=True)
    all_sales = _ensure_fin_cols(all_sales)
    all_sales["order_date"] = pd.to_datetime(all_sales["order_date"], errors="coerce")
    all_sales = (all_sales
                 .dropna(subset=["order_date", "sku"])
                 .copy())
    all_sales["date_str"] = all_sales["order_date"].dt.strftime("%Y-%m-%d")
    cutoff = pd.Timestamp(SALES_START_DATE)
    recent = all_sales[all_sales["order_date"] >= cutoff]

    # build one pivot per metric, then combine
    metric_map = {
        "qty":          "Units",
        "qty_zoey":     "Qty Zoey",
        "qty_shopify":  "Qty Shopify",
        "subtotal":     "Subtotal",
        "discount":     "Discount",
        "net_rev":      "Net Rev",
        "tax":          "Tax",
        "cost":         "COGS",
        "gross_profit": "GP",
    }

    # ── Build all day×metric columns at once (avoids fragmentation) ────
    daily_summaries = {ds: {m: 0.0 for m in metric_map} for ds in date_strs}
    new_cols = {}   # col_name -> list of values aligned to base["sku"]

    for metric, label in metric_map.items():
        if recent.empty or metric not in recent.columns:
            pivot_data = {}
        else:
            piv = (recent.pivot_table(index="sku", columns="date_str",
                                      values=metric, aggfunc="sum")
                         .reset_index())
            piv.columns.name = None
            pivot_data = {row["sku"]: row for _, row in piv.iterrows()}

        for ds, dl in zip(date_strs, date_labels):
            col  = f"{dl}|{metric}"
            vals = []
            for sku in base["sku"]:
                row = pivot_data.get(sku)
                v   = float(row[ds]) if (row is not None and ds in row
                                         and pd.notna(row[ds])) else 0.0
                vals.append(round(v, 2))
            new_cols[col] = vals
            daily_summaries[ds][metric] = round(
                recent[recent["date_str"] == ds][metric].sum()
                if not recent.empty and metric in recent.columns else 0, 2)

    # concat all new columns in one shot — no fragmentation
    df = pd.concat([base, pd.DataFrame(new_cols, index=base.index)], axis=1)

    # ── 7D / 30D totals per metric ──────────────────────────────────────
    last7_strs  = date_strs[-7:]
    prior7_strs = date_strs[-14:-7]
    total_cols  = {}
    for metric, label in metric_map.items():
        last7_cols = [f"{dl}|{metric}" for dl, ds in zip(date_labels, date_strs)
                      if ds in last7_strs]
        all30_cols = [f"{dl}|{metric}" for dl in date_labels]
        total_cols[f"7D {label}"]  = df[last7_cols].sum(axis=1).round(2).values
        total_cols[f"30D {label}"] = df[all30_cols].sum(axis=1).round(2).values

    df = pd.concat([df, pd.DataFrame(total_cols, index=df.index)], axis=1)

    # ── Trend (units: last 7 vs prior 7) ───────────────────────────────
    prior7_unit_cols = [f"{dl}|qty" for dl, ds in zip(date_labels, date_strs)
                        if ds in prior7_strs]
    last7_unit_cols  = [f"{dl}|qty" for dl, ds in zip(date_labels, date_strs)
                        if ds in last7_strs]
    def trend_arrow(r):
        l7 = r[last7_unit_cols].sum()
        p7 = r[prior7_unit_cols].sum()
        if p7 == 0 and l7 == 0: return "—"
        if p7 == 0:              return "↑ New"
        pct = (l7 - p7) / p7 * 100
        if pct >  10: return f"↑ {pct:+.0f}%"
        if pct < -10: return f"↓ {pct:+.0f}%"
        return f"→ {pct:+.0f}%"
    df["Trend"] = df.apply(trend_arrow, axis=1)

    # sort by 30D Net Rev desc
    df = df.sort_values("30D Net Rev", ascending=False).reset_index(drop=True)

    # ── Summary row ─────────────────────────────────────────────────────
    summary = {}
    for ds, dl in zip(date_strs, date_labels):
        summary[dl] = daily_summaries[ds]   # dict of metric→value

    return df, summary, date_labels, date_strs, metric_map