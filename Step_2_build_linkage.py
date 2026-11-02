"""
Step 2 - Build the linkage rows from the prices downloaded in Step 1.

For each candidate group with one old and one new trading item:
  - Date  = first price date of the new trading item if it is a real new listing,
            otherwise first date where the new trading item price differs from the old one
            (first new price after the last date where old and new prices are equal);
            if old and new prices are never equal in the window, day after old last price,
  - price = close price of the new trading item at Date,
  - share = shr of the new row in Link-candidates.xlsx,
  - assetid = root asset id of the old trading item (chains keep the same asset id),
  - Link  = asset id of the new trading item.
Groups are set aside (problems file) when:
  - Treated is False (rejected) or True (already done) -> skipped, not a problem,
  - more than one old or new trading item,
  - old and new trading items share the same asset id,
  - no price for the new trading item in the window,
  - gap between last common price date and Date greater than MAX_GAP_BDAYS business days,
  - no linkage: more than MAX_GAP_DAYS calendar days between the last price of the old trading item
    (full CIQ history) and the first price of the new trading item at Date,
  - price jump: price differs by more than MAX_PRICE_JUMP_PCT from the last close of the old trading item
    before Date (or no close of the old trading item before Date in the window); checked by hand.

Outputs (folder "output"):
  - proposed_linkage.csv
  - linkage_problems.csv
"""
import os

import numpy as np
import pandas as pd

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
MAX_GAP_BDAYS = 5
MAX_PRICE_JUMP_PCT = 20
MAX_GAP_DAYS = 15
STAMP = "AUC v2"


def load_inputs():
    groups = pd.read_csv(os.path.join(OUTPUT_DIR, "linkage_groups.csv"),
                         parse_dates=["date", "first_price_date", "last_price_date", "linkage_date"])
    prices = pd.read_csv(os.path.join(OUTPUT_DIR, "prices_around_linkage.csv"), parse_dates=["pricingdate"])
    return groups, prices


def build_row(group, prices):
    """Return (row, None) or (None, problem reason) for one candidate group."""
    old = group[group["role"] == "old"]
    new = group[group["role"] == "new"]
    if len(old) != 1 or len(new) != 1:
        return None, f"{len(old)} old / {len(new)} new trading items"
    old, new = old.iloc[0], new.iloc[0]
    if old["asset"] == new["asset"]:
        return None, "old and new share the same asset id"

    new_prices = prices[prices["tradingitemid"] == new["tradingitemid"]].sort_values("pricingdate")
    old_prices = prices[prices["tradingitemid"] == old["tradingitemid"]].sort_values("pricingdate")
    if new_prices.empty:
        return None, "no price for new trading item in window"

    old_last = old["last_price_date"]
    real_new_listing = new["first_price_date"] > old["date"] - pd.Timedelta(days=30)
    if real_new_listing:
        first = new_prices[new_prices["pricingdate"] >= new["first_price_date"]]
        reference = old_last
    else:
        both = new_prices.merge(old_prices, on="pricingdate", suffixes=("", "_old"))
        equal = both[(both["priceclose"] - both["priceclose_old"]).abs() < 1e-6]
        reference = equal["pricingdate"].max() if not equal.empty else old_last
        first = new_prices[new_prices["pricingdate"] > reference]
    if first.empty:
        return None, "no price for new trading item after last common price in window"
    first = first.iloc[0]

    gap = int(np.busday_count(reference.date(), first["pricingdate"].date())) if pd.notna(reference) else None
    if not real_new_listing and (gap is None or gap > MAX_GAP_BDAYS):
        return None, f"gap of {gap} business days between last common price {reference} and new price"

    gap_days = (first["pricingdate"] - old_last).days if pd.notna(old_last) else None
    if gap_days is not None and gap_days > MAX_GAP_DAYS:
        return None, (f"no linkage: gap of {gap_days} days between old last price on {old_last:%Y-%m-%d} and "
                      f"new first price {first['priceclose']:g} on {first['pricingdate']:%Y-%m-%d}")

    old_before = old_prices[old_prices["pricingdate"] < first["pricingdate"]]
    if old_before.empty:
        return None, f"price jump not checked: no price for old trading item before {first['pricingdate']:%Y-%m-%d}"
    old_last_price = old_before.iloc[-1]
    jump = (first["priceclose"] / old_last_price["priceclose"] - 1) * 100
    if abs(jump) > MAX_PRICE_JUMP_PCT:
        return None, (f"price jump of {jump:+.1f}%: old {old_last_price['priceclose']:g} on "
                      f"{old_last_price['pricingdate']:%Y-%m-%d}, new {first['priceclose']:g} on "
                      f"{first['pricingdate']:%Y-%m-%d}")

    return {
        "group": group["group"].iloc[0],
        "old_asset": old["asset"],
        "companyid": new["companyid"],
        "securityid": new["securityid"],
        "tradingitemid": new["tradingitemid"],
        "company": new["company"],
        "security": new["security"],
        "Link": new["asset"],
        "share": new["shr"],
        "price": first["priceclose"],
        "Date": first["pricingdate"],
        "old_last_price_date": old_last,
        "gap_bdays": gap,
        "stamp": STAMP,
        "note": "-",
    }, None


def resolve_root_assets(rows, known_links=None):
    """Chains keep the first asset id: if old asset was itself linked, use its root."""
    root = dict(known_links or {})
    rows = rows.sort_values("Date")
    assetids = []
    for _, r in rows.iterrows():
        a = r["old_asset"]
        while a in root:
            a = root[a]
        root[r["Link"]] = a
        assetids.append(a)
    rows.insert(0, "assetid", assetids)
    return rows


def build(groups, prices, known_links=None):
    todo = groups[groups["Treated"].isna()] if known_links is not None else groups
    rows, problems = [], []
    for _, group in todo.groupby("group"):
        row, reason = build_row(group, prices)
        if row:
            rows.append(row)
        else:
            problems.append({"group": group["group"].iloc[0], "company": group["company"].iloc[0],
                             "rows": ", ".join(group["row"].astype(str)), "reason": reason})
    return resolve_root_assets(pd.DataFrame(rows), known_links), pd.DataFrame(problems)


if __name__ == "__main__":
    groups, prices = load_inputs()
    historical = pd.read_excel(os.path.join(BASE_DIR, "input", "Historical_DocLinkage.xlsx"), header=1)
    known_links = dict(zip(historical["Link"], historical["assetid"]))
    rows, problems = build(groups, prices, known_links)
    rows.to_csv(os.path.join(OUTPUT_DIR, "proposed_linkage.csv"), index=False)
    problems.to_csv(os.path.join(OUTPUT_DIR, "linkage_problems.csv"), index=False)
    print(f"{len(rows)} linkages proposed, {len(problems)} groups set aside")
    print(problems["reason"].str.split(" ").str[:3].str.join(" ").value_counts().to_string())
