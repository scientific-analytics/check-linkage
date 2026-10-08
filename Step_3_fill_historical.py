"""
Step 3 - Write the results of Step 2.

  - output/Historical_DocLinkage_new.xlsx : copy of Historical_DocLinkage.xlsx with the proposed linkages appended
    (rows whose Link is already in the file are skipped, layout and formats are kept).
  - output/linkages/<assetid>_<Link>/ : one folder per proposed linkage with
        prices.csv  : daily close of each security of the group around the linkage (about one month)
        linkage.csv : the row appended to the historical file
  - output/problems/<group>_<company>/ : same prices.csv and the reason for groups set aside.
The source Historical_DocLinkage.xlsx is not modified.
"""
import os
import re
from copy import copy

import openpyxl
import pandas as pd

from Step_2_build_linkage import BASE_DIR, OUTPUT_DIR, load_inputs

HISTORICAL_FILE = os.path.join(BASE_DIR, "Historical_DocLinkage.xlsx")
NEW_HISTORICAL_FILE = os.path.join(OUTPUT_DIR, "Historical_DocLinkage_new.xlsx")
LINKAGES_DIR = os.path.join(OUTPUT_DIR, "linkages")
PROBLEMS_DIR = os.path.join(OUTPUT_DIR, "problems")
HEADER_ROW = 2
COLUMNS = ["assetid", "companyid", "securityid", "tradingitemid", "company", "security",
           "Link", "share", "price", "Date", "stamp", "note"]


def safe_name(text):
    return re.sub(r"[^\w.-]+", "_", str(text).strip()).strip("_")


def group_prices(group, prices):
    """Daily close of each security of the group, one column per security."""
    group = group.sort_values(["role", "row"], ascending=[False, True])
    labels = {r["tradingitemid"]: f"{r['role']} {r['asset']} {r['security']} ({r['tradingitemid']})"
              for _, r in group.iterrows()}
    data = prices[prices["tradingitemid"].isin(labels)]
    wide = data.pivot(index="pricingdate", columns="tradingitemid", values="priceclose")
    wide = wide.reindex(columns=[t for t in labels if t in wide.columns]).rename(columns=labels)
    return wide.sort_index()


def write_folders(groups, prices, proposed, problems):
    for _, row in proposed.iterrows():
        folder = os.path.join(LINKAGES_DIR, f"{safe_name(row['assetid'])}_{safe_name(row['Link'])}")
        os.makedirs(folder, exist_ok=True)
        group_prices(groups[groups["group"] == row["group"]], prices).to_csv(os.path.join(folder, "prices.csv"))
        row[["group"] + COLUMNS].to_frame().T.to_csv(os.path.join(folder, "linkage.csv"), index=False)
    for _, row in problems.iterrows():
        folder = os.path.join(PROBLEMS_DIR, f"{row['group']}_{safe_name(row['company'])}")
        os.makedirs(folder, exist_ok=True)
        group_prices(groups[groups["group"] == row["group"]], prices).to_csv(os.path.join(folder, "prices.csv"))
        row.to_frame().T.to_csv(os.path.join(folder, "problem.csv"), index=False)


def write_new_historical(proposed):
    wb = openpyxl.load_workbook(HISTORICAL_FILE)
    ws = wb.active
    header = {ws.cell(HEADER_ROW, c).value: c for c in range(1, ws.max_column + 1)}
    last_row = max(r for r in range(HEADER_ROW + 1, ws.max_row + 1) if ws.cell(r, header["assetid"]).value)
    existing_links = {ws.cell(r, header["Link"]).value for r in range(HEADER_ROW + 1, last_row + 1)}

    to_add = proposed[~proposed["Link"].isin(existing_links)].sort_values(["company", "Date"])
    for i, (_, row) in enumerate(to_add.iterrows(), start=1):
        values = row.to_dict()
        values["share"] = round(values["share"], 3)
        values["price"] = round(values["price"], 4)
        values["Date"] = values["Date"].to_pydatetime()
        for col in COLUMNS:
            cell = ws.cell(last_row + i, header[col], values[col])
            cell.number_format = ws.cell(last_row, header[col]).number_format
            cell.font = copy(ws.cell(last_row, header[col]).font)
    wb.save(NEW_HISTORICAL_FILE)
    return len(to_add), len(proposed) - len(to_add)


if __name__ == "__main__":
    groups, prices = load_inputs()
    proposed = pd.read_csv(os.path.join(OUTPUT_DIR, "proposed_linkage.csv"), parse_dates=["Date"])
    problems = pd.read_csv(os.path.join(OUTPUT_DIR, "linkage_problems.csv"))

    write_folders(groups, prices, proposed, problems)
    print(f"{len(proposed)} linkage folders in {LINKAGES_DIR}, {len(problems)} problem folders in {PROBLEMS_DIR}")

    added, skipped = write_new_historical(proposed)
    print(f"{NEW_HISTORICAL_FILE}: {added} rows appended, {skipped} already present")
