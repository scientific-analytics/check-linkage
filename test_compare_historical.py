"""Compare Step 2 output (without seeding from the historical) to the existing Historical_DocLinkage.xlsx."""
import os

import pandas as pd

from Step_2_build_linkage import BASE_DIR, OUTPUT_DIR, build, load_inputs

pd.set_option("display.width", 250)

groups, prices = load_inputs()
rows, problems = build(groups, prices)

historical = pd.read_excel(os.path.join(BASE_DIR, "input", "Historical_DocLinkage.xlsx"), header=1)
historical = historical.dropna(subset=["Link"])
historical["Date"] = pd.to_datetime(historical["Date"])

cmp = historical.merge(rows, on="Link", how="left", suffixes=("_hist", "_mine"))
cmp["tradingitem_ok"] = cmp["tradingitemid_hist"].astype(str).str.strip() == cmp["tradingitemid_mine"].astype("Int64").astype(str)
cmp["assetid_ok"] = cmp["assetid_hist"] == cmp["assetid_mine"]
cmp["date_diff_days"] = (cmp["Date_mine"] - cmp["Date_hist"]).dt.days
cmp["price_diff_pct"] = (cmp["price_mine"] / cmp["price_hist"].astype(float) - 1) * 100
cmp["share_diff_pct"] = (cmp["share_mine"] / cmp["share_hist"].astype(float) - 1) * 100

cols = ["company_hist", "Link", "assetid_hist", "assetid_mine", "tradingitem_ok",
        "Date_hist", "Date_mine", "date_diff_days", "price_hist", "price_mine", "price_diff_pct",
        "share_hist", "share_mine", "share_diff_pct"]
print(cmp[cols].round(2).to_string(index=False))

missing = cmp[cmp["Date_mine"].isna()]
if not missing.empty:
    linked_groups = groups[groups["asset"].isin(missing["Link"])]["group"].unique()
    print("\nHistorical rows not produced by Step 2:")
    print(problems[problems["group"].isin(linked_groups)].to_string(index=False))

print(f"\nassetid ok: {cmp['assetid_ok'].sum()}/{len(cmp)}"
      f" | Date exact: {(cmp['date_diff_days'] == 0).sum()}/{len(cmp)}"
      f" | price within 1%: {(cmp['price_diff_pct'].abs() < 1).sum()}/{len(cmp)}"
      f" | share within 1%: {(cmp['share_diff_pct'].abs() < 1).sum()}/{len(cmp)}")
