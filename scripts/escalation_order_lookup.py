"""Look up order details for the escalation tracker (Product Feedback Tracker sheet).

Usage: python scripts/escalation_order_lookup.py IND-442901 IND-442755 ...

Reads the Shopify "Sales Data Report with Stylist Name" CSVs in Downloads (the
long Jan-01 file plus the newest month-to-date file) and prints one JSON object
per order: store, channel, stylist and product lines. MTM vs RTW is inferred
from the variant title: made-to-measure lines have no size variant.
"""
import glob
import json
import os
import re
import sys

import pandas as pd

DOWNLOADS = os.path.expanduser(r"~\Downloads")
COLS = ["Order name", "POS location name", "Sales channel", "Line type",
        "Product title at time of sale", "Product variant title at time of sale",
        "Product type", "Assisting staff member name", "Day"]


def report_files():
    files = glob.glob(os.path.join(DOWNLOADS, "Sales Data Report with Stylist Name - *.csv"))
    files = [f for f in files if not re.search(r"\(\d+\)\.csv$", f)]
    # Newest end date first; keep the longest-range file and the newest file.
    def end(f):
        return re.findall(r"\d{4}-\d{2}-\d{2}", os.path.basename(f))[-1]
    def start(f):
        return re.findall(r"\d{4}-\d{2}-\d{2}", os.path.basename(f))[0]
    files.sort(key=end, reverse=True)
    picked = files[:1]
    longest = min(files, key=start) if files else None
    if longest and longest not in picked:
        picked.append(longest)
    return picked


def product_label(ptype, variant):
    ptype = ptype or ""
    rtw = bool(variant) and bool(re.search(r"/\s*\S+$|^\s*(XS|S|M|L|XL|XXL|\d{2})\s*$", variant))
    if ptype == "Shirt":
        return "RTW SHIRT" if rtw else "MTM SHIRT"
    if "T-Shirt" in ptype:
        return "LONG SLEEVES T-SHIRT" if "Long Sleeve" in ptype else "T-SHIRT"
    if "Jeans" in ptype or "Denim" in ptype:
        return "JEANS"
    if "Blazer" in ptype or "Jacket" in ptype or ptype == "Suits":
        return "BLAZER"
    if "Pants" in ptype or ptype in ("Chinos", "Smartpants", "Joggers"):
        return "PANTS"
    return ptype.upper()


def main(orders):
    wanted = {re.sub(r"[^0-9]", "", o) for o in orders}
    frames = []
    for f in report_files():
        df = pd.read_csv(f, dtype=str, usecols=COLS)
        df["num"] = df["Order name"].str.replace(r"[^0-9]", "", regex=True)
        frames.append(df[df["num"].isin(wanted)])
    df = pd.concat(frames).drop_duplicates() if frames else pd.DataFrame(columns=COLS + ["num"])
    for o in orders:
        num = re.sub(r"[^0-9]", "", o)
        rows = df[df["num"] == num]
        if rows.empty:
            print(json.dumps({"order": o, "found": False}))
            continue
        first = rows.iloc[0]
        lines = rows[(rows["Line type"] == "product") & rows["Product title at time of sale"].notna()]
        lines = lines.drop_duplicates(["Product title at time of sale", "Product variant title at time of sale"])
        store = first["POS location name"] if isinstance(first["POS location name"], str) else "Online"
        stylists = sorted({s for s in lines["Assisting staff member name"].dropna()})
        print(json.dumps({
            "order": "IND-" + num,
            "found": True,
            "date": first["Day"],
            "store": store,
            "channel": first["Sales channel"],
            "stylist": ", ".join(stylists) if store != "Online" else "Online",
            "products": [
                {"label": product_label(r["Product type"], r["Product variant title at time of sale"]
                                        if isinstance(r["Product variant title at time of sale"], str) else ""),
                 "name": r["Product title at time of sale"]
                 + (" - " + r["Product variant title at time of sale"]
                    if isinstance(r["Product variant title at time of sale"], str) else "")}
                for _, r in lines.iterrows()
            ],
        }, ensure_ascii=False))


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main(sys.argv[1:])
