"""
Download the CIQ prices of every trading item of a company, only while the trading item is in AUC.

For each company id of COMPANIES (or of the company ids given on the command line):
  - AUC periods = consecutive reviews of data_gov/<date>/_universe/auc.csv where the trading item is present,
    a trading item being in AUC from its review date until the next review,
  - CIQ prices are downloaded on these periods only,
  - auc_timeseries.csv      : close price per date, one column per trading item, NaN when not in AUC,
  - auc_timeseries_long.csv : raw CIQ rows (currency, close, adjustment factor, volume),
  - auc_periods.csv         : AUC periods of each trading item,
  - full_timeseries.csv / full_timeseries_long.csv : same files with the full CIQ history (AUC or not).

Outputs (folder "output/auc_timeseries/<companyid>_<company>").
"""
import os
import sys

import pandas as pd
import psycopg2
from dotenv import load_dotenv

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(BASE_DIR, "output", "auc_timeseries")
DATA_GOV = "/home/rudy/PycharmProjects/check_referential/data_gov"
COMPANIES = [24766, 24915054, 387895, 5856173, 874082, 4481284, 884372, 704861, 875425, 663731, 874650, 879755]

load_dotenv(os.path.join(BASE_DIR, ".env"))

db_params = {
    "host": os.getenv("DB_HOST"),
    "database": os.getenv("DB_NAME"),
    "user": os.getenv("DB_USER"),
    "password": os.getenv("DB_PASSWORD"),
    "port": os.getenv("DB_PORT")
}

PRICE_QUERY = """
WITH tmp_window AS (
    SELECT unnest(%s) AS tradingitemid, unnest(%s)::date AS start_date, unnest(%s)::date AS end_date
)
SELECT p.tradingitemid,
       p.pricingdate,
       cur.isocode AS currency,
       p.priceclose,
       p.adjustmentfactor,
       p.volume
FROM tmp_window w
    JOIN public.ciqpriceequity p ON p.tradingitemid = w.tradingitemid
        AND p.pricingdate >= w.start_date AND p.pricingdate < w.end_date
    JOIN public.ciqtradingitem t ON t.tradingitemid = p.tradingitemid
    LEFT JOIN public.ciqcurrency cur ON cur.currencyid = t.currencyid
ORDER BY p.tradingitemid, p.pricingdate;
"""

FULL_PRICE_QUERY = """
SELECT p.tradingitemid,
       p.pricingdate,
       cur.isocode AS currency,
       p.priceclose,
       p.adjustmentfactor,
       p.volume
FROM public.ciqpriceequity p
    JOIN public.ciqtradingitem t ON t.tradingitemid = p.tradingitemid
    LEFT JOIN public.ciqcurrency cur ON cur.currencyid = t.currencyid
WHERE p.tradingitemid = ANY(%s)
ORDER BY p.tradingitemid, p.pricingdate;
"""


def fetch(cursor, query, params, assets):
    cursor.execute(query, params)
    prices = pd.DataFrame(cursor.fetchall(), columns=[c[0] for c in cursor.description]).drop_duplicates()
    prices["pricingdate"] = pd.to_datetime(prices["pricingdate"])
    prices.insert(1, "asset", prices["tradingitemid"].map(assets))
    return prices


def write_timeseries(prices, names, folder, prefix):
    wide = prices.pivot(index="pricingdate", columns="tradingitemid", values="priceclose")
    wide = wide.reindex(columns=[t for t in names if t in wide.columns]).rename(columns=names)
    wide.to_csv(os.path.join(folder, f"{prefix}_timeseries.csv"))
    prices.to_csv(os.path.join(folder, f"{prefix}_timeseries_long.csv"), index=False)


def review_dates():
    return sorted(d for d in os.listdir(DATA_GOV) if os.path.isfile(os.path.join(DATA_GOV, d, "_universe", "auc.csv")))


def load_auc_membership(companies):
    """One row per (review, trading item) of the companies, with the review validity [start, end)."""
    reviews = review_dates()
    rows = []
    for i, d in enumerate(reviews):
        auc = pd.read_csv(os.path.join(DATA_GOV, d, "_universe", "auc.csv"),
                          usecols=["companyid", "tradingitemid", "company", "security"])
        auc = auc[auc["companyid"].isin(companies)].copy()
        auc["review"] = d
        auc["start"] = pd.Timestamp(d)
        auc["end"] = pd.Timestamp(reviews[i + 1] if i + 1 < len(reviews) else "2099-12-31")
        rows.append(auc)
    return pd.concat(rows, ignore_index=True), reviews


def load_asset_ids(reviews):
    path = os.path.join(DATA_GOV, reviews[-1], "_universe", "auc_asset.csv")
    assets = pd.read_csv(path, usecols=["tradingitemid", "asset"]).drop_duplicates("tradingitemid")
    return dict(zip(assets["tradingitemid"], assets["asset"]))


def auc_periods(membership):
    """Merge consecutive reviews of a trading item into periods."""
    m = membership.sort_values(["tradingitemid", "start"]).copy()
    m["period"] = (m["start"] != m.groupby("tradingitemid")["end"].shift()).cumsum()
    return (m.groupby(["tradingitemid", "period"])
            .agg(first_review=("review", "first"), last_review=("review", "last"),
                 start=("start", "min"), end=("end", "max"))
            .reset_index().drop(columns="period").sort_values("start"))


def column_names(membership, assets):
    last = membership.sort_values("start").drop_duplicates("tradingitemid", keep="last")
    return {t: f"{assets.get(t, '?')} {s} ({t})" for t, s in zip(last["tradingitemid"], last["security"])}


if __name__ == "__main__":
    companies = [int(c) for c in sys.argv[1:]] or COMPANIES
    membership, reviews = load_auc_membership(companies)
    assets = load_asset_ids(reviews)
    conn = psycopg2.connect(**db_params)
    cursor = conn.cursor()
    try:
        for companyid, company in membership.groupby("companyid"):
            name = company.sort_values("start")["company"].iloc[-1]
            periods = auc_periods(company)
            periods.insert(1, "asset", periods["tradingitemid"].map(assets))
            names = column_names(company, assets)

            prices = fetch(cursor, PRICE_QUERY, (
                [int(t) for t in periods["tradingitemid"]],
                periods["start"].dt.strftime("%Y-%m-%d").tolist(),
                periods["end"].dt.strftime("%Y-%m-%d").tolist(),
            ), assets)
            full = fetch(cursor, FULL_PRICE_QUERY, ([int(t) for t in names],), assets)

            folder = os.path.join(OUTPUT_DIR, f"{companyid}_{name.replace(' ', '_').replace('.', '')}")
            os.makedirs(folder, exist_ok=True)
            write_timeseries(prices, names, folder, "auc")
            write_timeseries(full, names, folder, "full")
            periods.to_csv(os.path.join(folder, "auc_periods.csv"), index=False)
            print(f"\n=== {name} ({companyid}) -> {folder}")
            print(periods.to_string(index=False))
            for label, data in (("AUC", prices), ("full", full)):
                print(f"{label} prices:")
                print(data.groupby(["tradingitemid", "asset", "currency"])["pricingdate"].agg(["size", "min", "max"]).to_string())
    finally:
        conn.close()
