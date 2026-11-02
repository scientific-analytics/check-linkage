"""Output checks of the linkage pipeline, run from the project root (see SKILL.md)."""
import glob
import os
import re
import sys

import pandas as pd

ROOT = os.getcwd()
sys.path.insert(0, ROOT)
from Step_2_build_linkage import MAX_GAP_DAYS, MAX_PRICE_JUMP_PCT  # noqa: E402

V2 = os.path.join(ROOT, "output", "Historical_DocLinkage_v2.xlsx")
V1 = os.path.join(ROOT, "output", "v1.csv")
PRICES = os.path.join(ROOT, "output", "prices_around_linkage.csv")
COLUMNS = ["assetid", "companyid", "securityid", "tradingitemid", "company", "security",
           "Link", "share", "price", "Date", "stamp", "note"]
INITIAL_FILE_COMMENT = "déjà présent dans le fichier initial"
EXCLUDED = ["price jump", "no linkage", "linkage removed", "circular linkage not validated"]

results = []


def check(name, bad, show=None):
    """bad: DataFrame (or list) of failing rows; empty means OK."""
    n = len(bad)
    results.append((name, n == 0))
    print(f"{'OK  ' if n == 0 else 'FAIL'} {name}" + ("" if n == 0 else f" ({n})"))
    if n:
        print(bad[show].to_string() if isinstance(bad, pd.DataFrame) and show else bad, "\n")


doc = pd.read_excel(V2, sheet_name="DocLinkage", header=1).iloc[:, 1:].dropna(subset=["Link"])
dec = pd.read_excel(V2, sheet_name="Decisions", header=1)
v1 = pd.read_csv(V1, dtype=str)
for d in (doc, dec):
    d["Link"] = d["Link"].astype(str).str.strip()

# v1.csv format
fmt = []
if list(v1.columns) != COLUMNS:
    fmt.append(f"columns {list(v1.columns)}")
spaces = v1.apply(lambda c: c.str.contains(r"^\s|\s$|\t", na=False)).any(axis=1)
fmt += [f"space/tab row {i}" for i in v1.index[spaces]]
for col in ["companyid", "securityid", "tradingitemid"]:
    fmt += [f"{col} not integer row {i}" for i in v1.index[~v1[col].str.fullmatch(r"\d+", na=False)]]
fmt += [f"Date row {i}" for i in v1.index[~v1["Date"].str.fullmatch(r"\d{4}-\d{2}-\d{2}", na=False)]]
check("v1.csv format", fmt)

# v1.csv = DocLinkage
key = lambda d: set(zip(d["tradingitemid"].astype(str).str.strip().str.split(".").str[0], d["Link"].str.strip()))
check("v1.csv = DocLinkage", sorted(key(v1) ^ key(doc)))

check("unique Link", doc[doc["Link"].duplicated(keep=False)], ["assetid", "Link", "Date"])

jump = pd.to_numeric(doc["jump_pct"], errors="coerce")
check(f"price jump <= {MAX_PRICE_JUMP_PCT}%", doc[jump.abs() > MAX_PRICE_JUMP_PCT],
      ["company", "assetid", "Link", "Date", "price", "old_price", "jump_pct"])

old_date = pd.to_datetime(doc["old_price_date"], errors="coerce")
gap = (pd.to_datetime(doc["Date"]) - old_date).dt.days
check(f"price gap <= {MAX_GAP_DAYS} days", doc[gap > MAX_GAP_DAYS],
      ["company", "assetid", "Link", "Date", "old_price_date"])

no_price = doc[pd.to_numeric(doc["old_price"], errors="coerce").isna()
               & (doc["old_price"].astype(str) != INITIAL_FILE_COMMENT)]
check("old price filled", no_price, ["company", "assetid", "Link", "Date", "old_price"])

excluded = dec[dec["decision"].isin(EXCLUDED)]
check("excluded decisions not in DocLinkage", excluded[excluded["Link"].isin(doc["Link"])],
      ["group", "company", "decision", "Link"])
manual = dec[dec["decision"] == "manual linkage added"]
check("manual linkages in DocLinkage", manual[~manual["Link"].isin(doc["Link"])],
      ["group", "company", "Link"])

# new trading item priced at Date, checked only when local prices cover Date
frames = [pd.read_csv(PRICES, parse_dates=["pricingdate"])[["tradingitemid", "pricingdate"]]]
for path in glob.glob(os.path.join(ROOT, "output", "auc_timeseries", "*", "full_timeseries.csv")):
    wide = pd.read_csv(path, parse_dates=["pricingdate"]).set_index("pricingdate")
    for col in wide.columns:
        ti = int(re.search(r"\((\d+)\)$", col).group(1))
        frames.append(pd.DataFrame({"tradingitemid": ti, "pricingdate": wide[col].dropna().index}))
prices = pd.concat(frames)
by_ti = {ti: set(d["pricingdate"]) for ti, d in prices.groupby("tradingitemid")}
doc["Date"] = pd.to_datetime(doc["Date"])
covered, not_priced = 0, []
for i, (ti, date) in enumerate(zip(doc["tradingitemid"].astype("int64"), doc["Date"])):
    dates = by_ti.get(ti)
    if not dates or not min(dates) <= date <= max(dates):
        continue
    covered += 1
    if date not in dates:
        not_priced.append(i)
print(f"     ({covered}/{len(doc)} rows with local prices around Date)")
check("new trading item priced at Date", doc.iloc[not_priced], ["company", "tradingitemid", "Link", "Date"])

ok = sum(r for _, r in results)
print(f"\n{ok} OK / {len(results) - ok} FAIL")
