"""
Automated Data Cleaning and Reporting Pipeline
===============================================

Objective
---------
Take a messy sales dataset (CSV or Excel) and, with one command:

  1. Clean it   -> missing values, duplicates, inconsistent labels, bad dates,
                   invalid numbers, outliers
  2. Audit it   -> before/after data-quality scorecard + step-by-step cleaning log
  3. Report it  -> Excel workbook (KPIs, native charts, summary tables) and an
                   HTML report with visual summaries

Usage
-----
    python data_cleaning_automation.py                       # runs on built-in sample data
    python data_cleaning_automation.py --input sales.csv     # your own CSV / Excel file
    python data_cleaning_automation.py --input data_folder/  # batch: every CSV/Excel in a folder
    python data_cleaning_automation.py --outdir reports

Expected columns (names are normalised, e.g. "Order ID" -> order_id):
    required : order_id, order_date, region, category, quantity, unit_price
    optional : customer_name, email, product, discount_pct, payment_method, status
"""

import argparse
import base64
import difflib
import html
import io
import os
import re
import time
import warnings
from datetime import datetime

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import seaborn as sns

warnings.filterwarnings("ignore")
sns.set_theme(style="whitegrid", palette="deep")

SAMPLE_PATH = os.path.join("sample_data", "messy_sales.csv")
RANDOM_STATE = 42

REQUIRED = ["order_id", "order_date", "region", "category", "quantity", "unit_price"]
OPTIONAL_TEXT = ["customer_name", "email", "product", "payment_method", "status"]

PLACEHOLDERS = {"", "n/a", "na", "null", "none", "nan", "-", "--", "?", "missing", "#n/a", "unknown"}
COLUMN_ALIASES = {"qty": "quantity", "price": "unit_price", "date": "order_date",
                  "discount": "discount_pct", "discount_": "discount_pct", "payment": "payment_method"}

CANONICAL = {
    "region": (["North", "South", "East", "West", "Central"],
               {"n": "North", "s": "South", "e": "East", "w": "West", "c": "Central"}),
    "category": (["Electronics", "Fashion", "Home & Kitchen", "Beauty", "Sports"], {}),
    "payment_method": (["Credit Card", "Debit Card", "UPI", "Cash", "Net Banking"],
                       {"cc": "Credit Card", "dc": "Debit Card", "nb": "Net Banking"}),
    "status": (["Delivered", "Shipped", "Cancelled", "Pending"], {"canceled": "Cancelled"}),
}


# ======================================================================
# SAMPLE DATA (deliberately dirty)
# ======================================================================
def generate_messy_sample(path: str, n: int = 1500, seed: int = RANDOM_STATE) -> None:
    """Create a realistic but very messy sales file so the pipeline has real work to do."""
    rng = np.random.default_rng(seed)
    catalog = {
        "Electronics": {"Laptop": 55000, "Smartphone": 25000, "Headphones": 3000, "Tablet": 22000},
        "Fashion": {"T-Shirt": 800, "Jeans": 2200, "Sneakers": 3500, "Jacket": 4500},
        "Home & Kitchen": {"Blender": 3200, "Cookware Set": 5000, "Vacuum Cleaner": 9000, "Air Fryer": 7500},
        "Beauty": {"Perfume": 2800, "Moisturizer": 900, "Lipstick": 650, "Shampoo": 450},
        "Sports": {"Yoga Mat": 1200, "Dumbbells": 2500, "Football": 1100, "Running Shoes": 4200},
    }
    cat_variants = {
        "Electronics": ["Electronics", "electronics", "ELECTRONICS ", "Electonics"],
        "Fashion": ["Fashion", "fashion", "Fashon", " FASHION"],
        "Home & Kitchen": ["Home & Kitchen", "Home and Kitchen", "home&kitchen", "HOME & KITCHEN"],
        "Beauty": ["Beauty", "beauty", "BEAUTY "],
        "Sports": ["Sports", "sport", "SPORTS", "Sports "],
    }
    region_variants = {
        "North": ["North", "north", "N", "NORTH ", "Nrth"], "South": ["South", "south", "S", "Sth"],
        "East": ["East", "east", "E", "EAST "], "West": ["West", "west", "W", "Wset"],
        "Central": ["Central", "central", "Centrl"],
    }
    pay_variants = ["Credit Card", "credit card", "CC", "Debit Card", "UPI", "upi", "Cash", "cash",
                    "Net Banking", "netbanking"]
    status_variants = ["Delivered"] * 6 + ["delivered", "DELIVERED", "Shipped", "shipped", "Pending",
                                           "Cancelled", "canceled"]
    first = ["Aarav", "Priya", "Rohan", "Sneha", "Vikram", "Ananya", "Karan", "Meera", "Arjun", "Isha",
             "Rahul", "Neha", "Siddharth", "Pooja", "Aditya", "Kavya", "Manish", "Divya", "Nikhil", "Riya"]
    last = ["Sharma", "Verma", "Gupta", "Singh", "Mehta", "Iyer", "Reddy", "Kapoor", "Nair", "Joshi"]
    customers = [f"{f} {l}" for f in first for l in last][:120]
    date_formats = ["%Y-%m-%d", "%m/%d/%Y", "%b %d, %Y", "%d %B %Y"]

    rows = []
    for i in range(1, n + 1):
        cat = rng.choice(list(catalog))
        prod = rng.choice(list(catalog[cat]))
        price = catalog[cat][prod] * rng.uniform(0.92, 1.08)
        date = pd.Timestamp("2023-01-01") + pd.Timedelta(days=int(rng.integers(0, 730)))
        month_boost = 1 if date.month in (10, 11, 12) else 0   # seasonality: more orders late in year
        name = str(rng.choice(customers))
        email = name.lower().replace(" ", ".") + "@example.com"
        region = rng.choice(list(region_variants))
        rows.append({
            "Order ID": f"ORD-{i:05d}",
            "Order Date": date.strftime(str(rng.choice(date_formats))),
            "Customer Name": rng.choice([name, name.upper(), name.lower(), "  " + name + "  "]),
            "Email": rng.choice([email, email.upper(), email.replace("@", ""), " " + email]),
            "Region": rng.choice(region_variants[region]),
            "Category": rng.choice(cat_variants[cat]),
            "Product": rng.choice([prod, prod.lower(), prod.upper()]),
            "Quantity": int(rng.integers(1, 4 + month_boost * 2)),
            "Unit Price": rng.choice([f"{price:.2f}", f"₹{price:,.2f}", f"{price:.0f}"]),
            "Discount %": rng.choice([0, 5, 10, 15, 20, "10%", "5%"], p=[.3, .2, .2, .1, .1, .05, .05]),
            "Payment Method": rng.choice(pay_variants),
            "Status": rng.choice(status_variants),
        })
    df = pd.DataFrame(rows).astype(object)

    # ---- inject dirt ----
    def poke(col, k, value):
        df.loc[rng.choice(len(df), k, replace=False), col] = value

    for col, k in [("Customer Name", 40), ("Email", 60), ("Region", 35), ("Category", 25), ("Product", 30),
                   ("Quantity", 45), ("Unit Price", 50), ("Discount %", 70), ("Payment Method", 40),
                   ("Status", 20), ("Order Date", 18)]:
        poke(col, k, rng.choice([np.nan, "", "N/A", "null", "-"]))
    poke("Quantity", 12, -2)                      # invalid negatives
    poke("Quantity", 5, 999)                      # extreme outliers (typos)
    poke("Unit Price", 6, "9999999")              # extreme price outliers
    poke("Order Date", 5, "2031-05-17")           # future dates
    poke("Order Date", 4, "not a date")           # unparseable dates
    df = pd.concat([df, df.sample(70, random_state=seed)])           # exact duplicates
    updated = df.sample(25, random_state=seed + 1).copy()            # same order_id, changed value (re-entered)
    updated["Quantity"] = 2
    df = pd.concat([df, updated, pd.DataFrame([{}] * 6)])            # + fully empty rows
    df = df.sample(frac=1, random_state=seed).reset_index(drop=True)

    os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
    df.to_csv(path, index=False)
    print(f"      Sample dataset written to {path} ({len(df):,} rows)")


# ======================================================================
# CLEANING UTILITIES
# ======================================================================
class CleaningLog:
    """Records every cleaning action so the process is auditable."""

    def __init__(self):
        self.entries = []

    def add(self, step: str, issue: str, rows: int, action: str) -> None:
        self.entries.append({"Step": step, "Issue found": issue, "Rows / cells affected": int(rows), "Action taken": action})
        print(f"      [{step}] {issue}: {int(rows):,} -> {action}")

    def frame(self) -> pd.DataFrame:
        return pd.DataFrame(self.entries)


def is_texty(s: pd.Series) -> bool:
    return s.dtype == object or pd.api.types.is_string_dtype(s)


def to_number(s: pd.Series) -> pd.Series:
    """'₹1,200.50' / '10%' / 25 -> float (invalid -> NaN)."""
    cleaned = s.astype(str).str.replace(r"[^\d.\-]", "", regex=True).replace("", np.nan)
    return pd.to_numeric(cleaned, errors="coerce")


def collapse_ws(s: pd.Series) -> pd.Series:
    return s.where(s.isna(), s.astype(str).str.strip().str.replace(r"\s+", " ", regex=True))


def _norm(x: str) -> str:
    return re.sub(r"\s+", " ", str(x).lower().replace("&", " and ")).strip()


def standardize_labels(s: pd.Series, canonical: list, aliases: dict, cutoff: float = 0.72) -> pd.Series:
    """Map spelling/case/abbreviation variants onto canonical labels (exact -> alias -> fuzzy match)."""
    canon = {_norm(c): c for c in canonical}
    alias = {_norm(k): v for k, v in aliases.items()}
    cache = {}

    def fix(v):
        if pd.isna(v):
            return np.nan
        key = _norm(v)
        if key not in cache:
            if key in canon:
                cache[key] = canon[key]
            elif key in alias:
                cache[key] = alias[key]
            else:
                m = difflib.get_close_matches(key, list(canon), n=1, cutoff=cutoff)
                cache[key] = canon[m[0]] if m else "Other"
        return cache[key]

    return s.map(fix)


EMAIL_RE = re.compile(r"^[\w.+-]+@[\w-]+\.[\w.-]+$")


def profile(df: pd.DataFrame) -> dict:
    """Data-quality scorecard used for the before/after comparison."""
    cells = df.size or 1
    missing = int(df.isna().sum().sum())
    p = {
        "Rows": len(df),
        "Missing cells": missing,
        "Completeness %": round(100 * (1 - missing / cells), 2),
        "Exact duplicate rows": int(df.duplicated().sum()),
        "Duplicate order IDs": int(df["order_id"].dropna().duplicated().sum()),
        "Invalid quantity / price values": int((to_number(df["quantity"]) <= 0).sum()
                                               + (to_number(df["unit_price"]) <= 0).sum()),
        "Invalid emails": int(df["email"].dropna().map(
            lambda e: str(e).strip().lower() != "unknown" and not EMAIL_RE.match(str(e).strip().lower())).sum())
        if "email" in df else 0,
    }
    for col in ["region", "category", "payment_method", "status"]:
        if col in df:
            p[f"Distinct {col} labels"] = int(df[col].nunique())
    return p


# ======================================================================
# THE CLEANING PIPELINE
# ======================================================================
def clean_data(raw: pd.DataFrame, log: CleaningLog):
    df = raw.copy()

    # --- column names ---
    df.columns = [re.sub(r"[^a-z0-9]+", "_", str(c).strip().lower()).strip("_") for c in df.columns]
    df = df.rename(columns={c: COLUMN_ALIASES.get(c, c) for c in df.columns})
    missing_cols = [c for c in REQUIRED if c not in df.columns]
    if missing_cols:
        raise ValueError(f"Input is missing required column(s): {missing_cols}. Found: {list(df.columns)}")
    for col in OPTIONAL_TEXT:
        if col not in df:
            df[col] = np.nan
            log.add("Schema", f"Optional column '{col}' absent", len(df), "added empty column")
    if "discount_pct" not in df:
        df["discount_pct"] = np.nan

    # --- placeholders -> real NaN ---
    n_ph = 0
    for c in df.columns:
        if is_texty(df[c]):
            mask = df[c].notna() & df[c].astype(str).str.strip().str.lower().isin(PLACEHOLDERS)
            n_ph += int(mask.sum())
            df.loc[mask, c] = np.nan
    log.add("Missing", "Placeholder text ('N/A', 'null', '-', blank...)", n_ph, "converted to proper missing values")

    # --- fully empty rows ---
    empty = df.isna().all(axis=1)
    df = df[~empty]
    log.add("Rows", "Completely empty rows", empty.sum(), "dropped")

    before = profile(df)
    missing_before = df.isna().sum()

    # --- whitespace ---
    n_ws = 0
    for c in df.columns:
        if is_texty(df[c]):
            tidy = collapse_ws(df[c])
            n_ws += int(((tidy != df[c]) & df[c].notna()).sum())
            df[c] = tidy
    log.add("Text", "Leading/trailing/double whitespace", n_ws, "trimmed in all text columns")

    # --- label standardisation ---
    for col, (canonical, aliases) in CANONICAL.items():
        orig = df[col].copy()
        df[col] = standardize_labels(df[col], canonical, aliases)
        changed = int(((orig != df[col]) & orig.notna()).sum())
        log.add("Consistency", f"Inconsistent '{col}' labels ({orig.nunique()} variants)", changed,
                f"mapped to {len(canonical)} standard labels (exact, alias, fuzzy match)")

    for col in ["customer_name", "product"]:
        orig = df[col].copy()
        df[col] = df[col].where(df[col].isna(), df[col].astype(str).str.title())
        log.add("Consistency", f"Mixed capitalisation in '{col}'", ((orig != df[col]) & orig.notna()).sum(), "title-cased")

    em = df["email"].where(df["email"].isna(), df["email"].astype(str).str.lower().str.replace(" ", "", regex=False))
    bad_email = em.notna() & ~em.fillna("").map(lambda e: bool(EMAIL_RE.match(e)))
    em[bad_email] = np.nan
    df["email"] = em
    log.add("Validation", "Invalid email addresses", bad_email.sum(), "set to missing (then 'Unknown')")

    # --- dates ---
    parsed = pd.to_datetime(df["order_date"], format="mixed", errors="coerce")
    unparseable = int((parsed.isna() & df["order_date"].notna()).sum())
    future = parsed > pd.Timestamp.today().normalize()
    parsed[future] = pd.NaT
    log.add("Dates", "Unparseable dates", unparseable, "parsed 4+ formats; failures treated as missing")
    log.add("Dates", "Future-dated orders", future.sum(), "treated as invalid")
    df["order_date"] = parsed
    bad_dates = df["order_date"].isna()
    df = df[~bad_dates]
    log.add("Dates", "Rows without a valid order date", bad_dates.sum(), "dropped (a date cannot be guessed)")

    no_id = df["order_id"].isna()
    df = df[~no_id]
    log.add("Rows", "Rows without an order ID", no_id.sum(), "dropped")

    # --- numeric parsing ---
    for col in ["quantity", "unit_price", "discount_pct"]:
        df[col] = to_number(df[col])
    invalid_q = df["quantity"] <= 0
    invalid_p = df["unit_price"] <= 0
    df.loc[invalid_q, "quantity"] = np.nan
    df.loc[invalid_p, "unit_price"] = np.nan
    log.add("Validation", "Zero/negative quantities", invalid_q.sum(), "set to missing for imputation")
    log.add("Validation", "Zero/negative prices", invalid_p.sum(), "set to missing for imputation")
    bad_disc = (df["discount_pct"] < 0) | (df["discount_pct"] > 100)
    df.loc[bad_disc, "discount_pct"] = np.nan
    log.add("Validation", "Discount outside 0-100%", bad_disc.sum(), "set to missing")

    # --- duplicates ---
    exact = df.duplicated(keep="first")
    df = df[~exact]
    log.add("Duplicates", "Exact duplicate rows (after standardisation)", exact.sum(), "removed, first kept")
    dup_id = df["order_id"].duplicated(keep="last")
    df = df[~dup_id]
    log.add("Duplicates", "Same order ID entered twice with different values", dup_id.sum(),
            "removed, latest entry kept")

    # --- missing value imputation ---
    for col in ["customer_name", "email", "product", "region", "category", "payment_method", "status"]:
        n = int(df[col].isna().sum())
        df[col] = df[col].fillna("Unknown")
        log.add("Missing", f"Missing '{col}'", n, "filled with 'Unknown'")

    n = int(df["quantity"].isna().sum())
    df["quantity"] = df["quantity"].fillna(df.groupby("category")["quantity"].transform("median"))
    df["quantity"] = df["quantity"].fillna(df["quantity"].median()).round().astype(int)
    log.add("Missing", "Missing/invalid quantity", n, "imputed with category median")

    n = int(df["unit_price"].isna().sum())
    df["unit_price"] = df["unit_price"].fillna(df.groupby("product")["unit_price"].transform("median"))
    df["unit_price"] = df["unit_price"].fillna(df.groupby("category")["unit_price"].transform("median"))
    df["unit_price"] = df["unit_price"].fillna(df["unit_price"].median())
    log.add("Missing", "Missing/invalid unit price", n, "imputed with product median (category fallback)")

    n = int(df["discount_pct"].isna().sum())
    df["discount_pct"] = df["discount_pct"].fillna(0)
    log.add("Missing", "Missing discount", n, "assumed no discount (0%)")

    # --- outliers (IQR x3, winsorised) ---
    q1, q3 = df["quantity"].quantile([.25, .75])
    cap_q = q3 + 3 * (q3 - q1)
    out_q = df["quantity"] > cap_q
    df.loc[out_q, "quantity"] = int(cap_q)
    log.add("Outliers", f"Extreme quantities (> {cap_q:.0f})", out_q.sum(), f"capped at {int(cap_q)}")

    g = df.groupby("product")["unit_price"]
    cap_p = g.transform(lambda x: x.quantile(.75) + 3 * (x.quantile(.75) - x.quantile(.25)))
    out_p = (df["unit_price"] > cap_p) & (df["product"] != "Unknown")
    df.loc[out_p, "unit_price"] = cap_p[out_p]
    log.add("Outliers", "Extreme unit prices within a product", out_p.sum(), "capped at product-level IQR limit")

    # --- derived fields ---
    df["revenue"] = (df["quantity"] * df["unit_price"] * (1 - df["discount_pct"] / 100)).round(2)
    df["order_month"] = df["order_date"].dt.to_period("M").astype(str)
    df["unit_price"] = df["unit_price"].round(2)

    cols = ["order_id", "order_date", "order_month", "customer_name", "email", "region", "category", "product",
            "quantity", "unit_price", "discount_pct", "revenue", "payment_method", "status"]
    df = df[cols].sort_values("order_date").reset_index(drop=True)
    return df, before, missing_before


# ======================================================================
# ANALYSIS
# ======================================================================
def build_summaries(df: pd.DataFrame) -> dict:
    net = df[df["status"] != "Cancelled"]
    agg = dict(orders=("order_id", "nunique"), units=("quantity", "sum"), revenue=("revenue", "sum"))

    def table(col):
        t = net.groupby(col).agg(**agg).reset_index().sort_values("revenue", ascending=False)
        t["revenue_share_%"] = (t["revenue"] / t["revenue"].sum() * 100).round(1)
        t["revenue"] = t["revenue"].round(2)
        return t

    monthly = net.groupby("order_month").agg(**agg).reset_index()
    monthly["revenue"] = monthly["revenue"].round(2)
    top_cust = (net[net["customer_name"] != "Unknown"].groupby("customer_name")
                .agg(**agg).reset_index().sort_values("revenue", ascending=False).head(10))
    top_cust["revenue"] = top_cust["revenue"].round(2)
    status = df["status"].value_counts().rename_axis("status").reset_index(name="orders")

    by_cat, by_reg = table("category"), table("region")
    kpis = {
        "Orders (excl. cancelled)": f"{net['order_id'].nunique():,}",
        "Net revenue": f"{net['revenue'].sum():,.0f}",
        "Average order value": f"{net['revenue'].mean():,.0f}",
        "Units sold": f"{int(net['quantity'].sum()):,}",
        "Unique customers": f"{net.loc[net['customer_name'] != 'Unknown', 'customer_name'].nunique():,}",
        "Cancellation rate": f"{(df['status'] == 'Cancelled').mean() * 100:.1f}%",
        "Top category": str(by_cat.iloc[0]["category"]),
        "Top region": str(by_reg.iloc[0]["region"]),
        "Period covered": f"{df['order_date'].min():%d %b %Y} - {df['order_date'].max():%d %b %Y}",
    }
    return {"category": by_cat, "region": by_reg, "monthly": monthly, "customers": top_cust,
            "payment": table("payment_method"), "status": status, "kpis": kpis}


# ======================================================================
# REPORTING
# ======================================================================
def make_charts(df, summ, missing_before, charts_dir) -> dict:
    os.makedirs(charts_dir, exist_ok=True)
    charts = {}

    def finish(fig, name, title):
        path = os.path.join(charts_dir, f"{name}.png")
        fig.savefig(path, dpi=130, bbox_inches="tight")
        buf = io.BytesIO()
        fig.savefig(buf, format="png", dpi=110, bbox_inches="tight")
        plt.close(fig)
        charts[name] = (title, base64.b64encode(buf.getvalue()).decode())

    mb = missing_before[missing_before > 0].sort_values()
    fig, ax = plt.subplots(figsize=(7, 4))
    ax.barh(mb.index, mb.values, color="tab:red", label="Before cleaning")
    ax.barh(mb.index, [0] * len(mb), color="tab:green", label="After cleaning (0)")
    ax.set(title="Missing Values per Column", xlabel="Missing cells")
    ax.legend(loc="lower right")
    finish(fig, "missing_values", "Missing values: before vs after")

    t = summ["category"]
    fig, ax = plt.subplots(figsize=(7, 4))
    sns.barplot(data=t, x="revenue", y="category", ax=ax)
    ax.set(title="Net Revenue by Category", xlabel="Revenue", ylabel="")
    finish(fig, "revenue_by_category", "Revenue by category")

    m = summ["monthly"]
    fig, ax = plt.subplots(figsize=(8, 4))
    ax.plot(m["order_month"], m["revenue"], marker="o", lw=2)
    ax.set(title="Monthly Net Revenue", xlabel="", ylabel="Revenue")
    ax.set_xticks(range(0, len(m), 3))
    ax.set_xticklabels(m["order_month"].iloc[::3], rotation=45, ha="right")
    finish(fig, "monthly_revenue", "Monthly revenue trend")

    r = summ["region"]
    fig, ax = plt.subplots(figsize=(5.5, 4))
    ax.pie(r["revenue"], labels=r["region"], autopct="%1.0f%%", startangle=90,
           colors=sns.color_palette("deep", len(r)))
    ax.set_title("Revenue Share by Region")
    finish(fig, "revenue_by_region", "Revenue share by region")
    return charts


def write_excel(path, df, summ, before, after, log, missing_before):
    print("      Building Excel workbook")
    with pd.ExcelWriter(path, engine="xlsxwriter", datetime_format="yyyy-mm-dd") as writer:
        wb = writer.book
        title = wb.add_format({"bold": True, "font_size": 18, "font_color": "#1F3864"})
        sub = wb.add_format({"italic": True, "font_color": "#595959"})
        kl = wb.add_format({"bold": True, "bg_color": "#DDEBF7", "border": 1})
        kv = wb.add_format({"border": 1, "align": "right"})
        hdr = wb.add_format({"bold": True, "bg_color": "#1F3864", "font_color": "white", "border": 1})
        money = wb.add_format({"num_format": "#,##0.00"})
        wrap = wb.add_format({"text_wrap": True, "valign": "top"})

        # ---------- Dashboard ----------
        ws = wb.add_worksheet("Dashboard")
        ws.hide_gridlines(2)
        ws.write("A1", "Sales Data Report", title)
        ws.write("A2", f"Generated {datetime.now():%d %b %Y %H:%M} | {len(df):,} clean records", sub)
        ws.write("A4", "Key metric", hdr)
        ws.write("B4", "Value", hdr)
        for i, (k, v) in enumerate(summ["kpis"].items(), start=4):
            ws.write(i, 0, k, kl)
            ws.write(i, 1, v, kv)
        ws.set_column("A:A", 26)
        ws.set_column("B:B", 24)

        # ---------- summary tables ----------
        def sheet(name, frame, money_cols=()):
            frame.to_excel(writer, sheet_name=name, index=False)
            s = writer.sheets[name]
            for j, c in enumerate(frame.columns):
                s.write(0, j, c, hdr)
                w = min(max(len(str(c)), frame[c].astype(str).str.len().max() if len(frame) else 0) + 3, 42)
                s.set_column(j, j, w, money if c in money_cols else None)
            s.freeze_panes(1, 0)
            return s

        sheet("By Category", summ["category"], ["revenue"])
        sheet("By Region", summ["region"], ["revenue"])
        sheet("Monthly Trend", summ["monthly"], ["revenue"])
        sheet("Top Customers", summ["customers"], ["revenue"])
        sheet("Payment Mix", summ["payment"], ["revenue"])

        # ---------- charts on dashboard ----------
        n_c, n_r, n_m, n_p = (len(summ[k]) for k in ["category", "region", "monthly", "payment"])
        c1 = wb.add_chart({"type": "column"})
        c1.add_series({"name": "Revenue", "categories": ["By Category", 1, 0, n_c, 0],
                       "values": ["By Category", 1, 3, n_c, 3], "fill": {"color": "#2E75B6"}})
        c1.set_title({"name": "Net Revenue by Category"})
        c1.set_legend({"none": True})
        c2 = wb.add_chart({"type": "line"})
        c2.add_series({"name": "Revenue", "categories": ["Monthly Trend", 1, 0, n_m, 0],
                       "values": ["Monthly Trend", 1, 3, n_m, 3], "line": {"color": "#C55A11", "width": 2.25}})
        c2.set_title({"name": "Monthly Net Revenue"})
        c2.set_legend({"none": True})
        c2.set_x_axis({"num_font": {"rotation": -45}})
        c3 = wb.add_chart({"type": "pie"})
        c3.add_series({"name": "Region", "categories": ["By Region", 1, 0, n_r, 0],
                       "values": ["By Region", 1, 3, n_r, 3], "data_labels": {"percentage": True}})
        c3.set_title({"name": "Revenue Share by Region"})
        c4 = wb.add_chart({"type": "bar"})
        c4.add_series({"name": "Orders", "categories": ["Payment Mix", 1, 0, n_p, 0],
                       "values": ["Payment Mix", 1, 1, n_p, 1], "fill": {"color": "#548235"}})
        c4.set_title({"name": "Orders by Payment Method"})
        c4.set_legend({"none": True})
        for ch, cell in [(c1, "D4"), (c2, "L4"), (c3, "D21"), (c4, "L21")]:
            ch.set_size({"width": 560, "height": 320})
            ws.insert_chart(cell, ch)

        # ---------- data quality ----------
        dq = pd.DataFrame({"Metric": list(before), "Before": list(before.values()),
                           "After": [after.get(k, "") for k in before]})
        sheet("Data Quality", dq)
        mc = pd.DataFrame({"Column": missing_before.index, "Missing before": missing_before.values,
                           "Missing after": 0})
        mc.to_excel(writer, sheet_name="Data Quality", index=False, startrow=len(dq) + 3)
        s = writer.sheets["Data Quality"]
        for j, c in enumerate(mc.columns):
            s.write(len(dq) + 3, j, c, hdr)

        # ---------- cleaning log & data ----------
        ls = sheet("Cleaning Log", log.frame())
        ls.set_column("B:B", 55, wrap)
        ls.set_column("D:D", 62, wrap)

        export = df.copy()
        ds = sheet("Clean Data", export, ["unit_price", "revenue"])
        ds.autofilter(0, 0, len(export), len(export.columns) - 1)


def write_html(path, summ, before, after, log, charts, runtime, source):
    def tbl(frame):
        return frame.to_html(index=False, border=0, classes="t", escape=True)

    kpis = "".join(f"<div class='kpi'><span>{html.escape(k)}</span><b>{html.escape(str(v))}</b></div>"
                   for k, v in summ["kpis"].items())
    dq = pd.DataFrame({"Metric": list(before), "Before": list(before.values()),
                       "After": [after.get(k, "") for k in before]})
    imgs = "".join(f"<figure><img src='data:image/png;base64,{b64}' alt='{html.escape(t)}'>"
                   f"<figcaption>{html.escape(t)}</figcaption></figure>" for t, b64 in charts.values())
    page = f"""<!doctype html><html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1"><title>Data Cleaning &amp; Sales Report</title>
<style>
body{{font-family:Segoe UI,Arial,sans-serif;margin:0;background:#f4f6fa;color:#222}}
header{{background:#1F3864;color:#fff;padding:24px 40px}} main{{padding:24px 40px;max-width:1200px;margin:auto}}
h2{{border-left:5px solid #2E75B6;padding-left:10px;margin-top:36px}}
.kpis{{display:grid;grid-template-columns:repeat(auto-fit,minmax(190px,1fr));gap:12px}}
.kpi{{background:#fff;border-radius:8px;padding:14px;box-shadow:0 1px 3px #0002}}
.kpi span{{display:block;font-size:12px;color:#666}} .kpi b{{font-size:20px;color:#1F3864}}
.charts{{display:grid;grid-template-columns:repeat(auto-fit,minmax(420px,1fr));gap:16px}}
figure{{margin:0;background:#fff;border-radius:8px;padding:10px;box-shadow:0 1px 3px #0002}} img{{max-width:100%}}
figcaption{{text-align:center;font-size:12px;color:#666}}
table.t{{border-collapse:collapse;width:100%;background:#fff;box-shadow:0 1px 3px #0002;font-size:13px}}
.t th{{background:#1F3864;color:#fff;text-align:left;padding:8px}} .t td{{padding:7px 8px;border-bottom:1px solid #eee}}
</style></head><body>
<header><h1>Automated Data Cleaning &amp; Sales Report</h1>
<div>Source: {html.escape(source)} | Generated {datetime.now():%d %b %Y %H:%M} | Pipeline runtime {runtime:.1f}s</div></header>
<main>
<h2>Key Metrics</h2><div class="kpis">{kpis}</div>
<h2>Visual Summary</h2><div class="charts">{imgs}</div>
<h2>Data Quality: Before vs After</h2>{tbl(dq)}
<h2>Cleaning Log</h2>{tbl(log.frame())}
<h2>Revenue by Category</h2>{tbl(summ['category'])}
<h2>Revenue by Region</h2>{tbl(summ['region'])}
<h2>Top 10 Customers</h2>{tbl(summ['customers'])}
</main></body></html>"""
    with open(path, "w", encoding="utf-8") as f:
        f.write(page)


# ======================================================================
# ORCHESTRATION
# ======================================================================
def read_input(path: str) -> pd.DataFrame:
    if path.lower().endswith((".xlsx", ".xlsm", ".xls")):
        return pd.read_excel(path, dtype=str)
    return pd.read_csv(path, dtype=str, keep_default_na=False, na_values=[""], encoding_errors="replace")


def process_file(path: str, outdir: str) -> None:
    start = time.time()
    stem = os.path.splitext(os.path.basename(path))[0]
    print(f"\n=== Processing {path} ===")
    raw = read_input(path)
    print(f"[1/4] Loaded {len(raw):,} rows x {raw.shape[1]} columns")

    print("[2/4] Cleaning")
    log = CleaningLog()
    df, before, missing_before = clean_data(raw, log)
    after = profile(df)
    print(f"      {len(raw):,} raw rows -> {len(df):,} clean rows")

    print("[3/4] Analysing and charting")
    summ = build_summaries(df)
    os.makedirs(outdir, exist_ok=True)
    charts = make_charts(df, summ, missing_before, os.path.join(outdir, "charts", stem))

    print("[4/4] Writing reports")
    stamp = datetime.now().strftime("%Y%m%d")
    df.to_csv(os.path.join(outdir, f"{stem}_cleaned.csv"), index=False)
    write_excel(os.path.join(outdir, f"{stem}_report_{stamp}.xlsx"), df, summ, before, after, log, missing_before)
    write_html(os.path.join(outdir, f"{stem}_report_{stamp}.html"), summ, before, after, log, charts,
               time.time() - start, path)

    print(f"\n      Data quality: completeness {before['Completeness %']}% -> {after['Completeness %']}% | "
          f"duplicates {before['Exact duplicate rows'] + before['Duplicate order IDs']} -> "
          f"{after['Exact duplicate rows'] + after['Duplicate order IDs']}")
    print(f"      Done in {time.time() - start:.1f}s. Reports saved in ./{outdir}/")


def main() -> None:
    parser = argparse.ArgumentParser(description="Automated data cleaning and reporting pipeline")
    parser.add_argument("--input", default=None, help="CSV/Excel file, or a folder of files (batch mode)")
    parser.add_argument("--outdir", default="reports", help="Folder for reports (default: reports)")
    args = parser.parse_args()

    if args.input is None:
        print("[0/4] No input given -> using built-in messy sample data")
        if not os.path.exists(SAMPLE_PATH):
            generate_messy_sample(SAMPLE_PATH)
        args.input = SAMPLE_PATH

    if os.path.isdir(args.input):
        files = sorted(os.path.join(args.input, f) for f in os.listdir(args.input)
                       if f.lower().endswith((".csv", ".xlsx", ".xls", ".xlsm")))
        if not files:
            raise SystemExit(f"No CSV/Excel files found in {args.input}")
        failed = []
        for f in files:
            try:
                process_file(f, args.outdir)
            except Exception as exc:  # one bad file must not stop the whole batch
                print(f"      !! Skipped {f}: {exc}")
                failed.append(f)
        print(f"\nBatch finished: {len(files) - len(failed)} succeeded, {len(failed)} failed")
        if failed:
            print("Failed files: " + ", ".join(failed))
    else:
        try:
            process_file(args.input, args.outdir)
        except (ValueError, FileNotFoundError) as exc:
            raise SystemExit(f"Error: {exc}")


if __name__ == "__main__":
    main()
