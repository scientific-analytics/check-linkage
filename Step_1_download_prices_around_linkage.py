"""
Step 1 - Download CIQ prices around each linkage candidate.

For each candidate group of Link-candidates.xlsx (rows "N.1", "N.2", ...):
  - "new" trading items = rows with the most recent review date of the group,
  - "old" trading items = the other rows.
The linkage date is estimated from CIQ price history:
  - first price date of the new trading items when they start after the old review,
  - otherwise last price date of the old trading items.
Then prices of every trading item of the group are downloaded on
[linkage date - WINDOW_DAYS, linkage date + WINDOW_DAYS].

Outputs (folder "output"):
  - linkage_groups.csv   : one row per trading item with its group, role, first/last price date, window and
                           the currencies of the old and new trading items of its group (old_currency, new_currency)
  - prices_around_linkage.csv : daily prices in the window
"""
import logging
import os
from datetime import datetime, timedelta

import pandas as pd
import psycopg2
from dotenv import load_dotenv

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
CANDIDATES_FILE = os.path.join(BASE_DIR, "input", "Link-candidates.xlsx")
DATA_GOV = "/home/rudy/PycharmProjects/check_referential/data_gov"
WINDOW_DAYS = 15

load_dotenv(os.path.join(BASE_DIR, ".env"))

db_params = {
    "host": os.getenv("DB_HOST"),
    "database": os.getenv("DB_NAME"),
    "user": os.getenv("DB_USER"),
    "password": os.getenv("DB_PASSWORD"),
    "port": os.getenv("DB_PORT")
}


def setup_logging():
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_file = os.path.join(OUTPUT_DIR, f"download_prices_{timestamp}.log")
    logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s",
                        handlers=[logging.FileHandler(log_file), logging.StreamHandler()])
    return logging.getLogger("download_prices")


def load_candidates():
    df = pd.read_excel(CANDIDATES_FILE, header=1)
    df = df.dropna(subset=["row", "tradingitemid"])
    df["group"] = df["row"].astype(str).str.split(".").str[0].astype(int)
    df["date"] = pd.to_datetime(df["date"])
    for col in ["companyid", "securityid", "tradingitemid"]:
        df[col] = df[col].astype(str).str.strip().astype("int64")
    df["role"] = "old"
    df.loc[df["date"] == df.groupby("group")["date"].transform("max"), "role"] = "new"
    return df


def load_asset_ids():
    """Asset id per trading item from the most recent data_gov/<date>/_universe/auc_asset.csv."""
    folders = sorted(f for f in os.listdir(DATA_GOV)
                     if os.path.isfile(os.path.join(DATA_GOV, f, "_universe", "auc_asset.csv")))
    path = os.path.join(DATA_GOV, folders[-1], "_universe", "auc_asset.csv")
    assets = pd.read_csv(path, usecols=["tradingitemid", "asset"])
    return assets.drop_duplicates("tradingitemid"), path


def create_price_range_query():
    return """
    WITH tmp_ti_list AS (SELECT unnest(%s) AS tradingitemid)
    SELECT p.tradingitemid,
           MIN(p.pricingdate) AS first_price_date,
           MAX(p.pricingdate) AS last_price_date
    FROM public.ciqpriceequity p
    WHERE p.tradingitemid IN (SELECT tradingitemid FROM tmp_ti_list)
    GROUP BY p.tradingitemid;
    """


def create_price_query():
    return """
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
            AND p.pricingdate BETWEEN w.start_date AND w.end_date
        JOIN public.ciqtradingitem t ON t.tradingitemid = p.tradingitemid
        LEFT JOIN public.ciqcurrency cur ON cur.currencyid = t.currencyid
    ORDER BY p.tradingitemid, p.pricingdate;
    """


def estimate_linkage_date(group):
    old = group[group["role"] == "old"]
    new = group[group["role"] == "new"]
    new_first = new["first_price_date"].min()
    old_last = old["last_price_date"].max()
    old_review = old["date"].max()
    if pd.notna(new_first) and new_first > old_review - timedelta(days=30):
        return new_first
    if pd.notna(old_last):
        return old_last
    return new["date"].max()


def group_currencies(candidates):
    """old_currency and new_currency of each group (currencies joined with "/" when the old trading items differ)."""
    currencies = (candidates.dropna(subset=["currency"])
                  .groupby(["group", "role"])["currency"]
                  .agg(lambda c: "/".join(sorted(c.astype(str).unique())))
                  .unstack("role")
                  .reindex(columns=["old", "new"]))
    return currencies.add_suffix("_currency")


if __name__ == "__main__":
    logger = setup_logging()
    logger.info("=== Starting Download Prices Around Linkage ===")

    try:
        candidates = load_candidates()
        logger.info(f"Loaded {len(candidates)} rows in {candidates['group'].nunique()} groups from {CANDIDATES_FILE}")

        assets, assets_file = load_asset_ids()
        candidates = candidates.merge(assets, on="tradingitemid", how="left")
        logger.info(f"Asset ids loaded from {assets_file}: "
                    f"{candidates['asset'].isna().sum()} candidate rows without asset id")

        logger.info("Connecting to database...")
        conn = psycopg2.connect(**db_params)
        cursor = conn.cursor()
        logger.info("Database connection established successfully")

        tradingitem_ids = [int(t) for t in candidates["tradingitemid"].unique()]
        logger.info(f"Fetching first/last price date for {len(tradingitem_ids)} trading items...")
        cursor.execute(create_price_range_query(), (tradingitem_ids,))
        price_range = pd.DataFrame(cursor.fetchall(), columns=[d[0] for d in cursor.description])
        price_range[["first_price_date", "last_price_date"]] = price_range[
            ["first_price_date", "last_price_date"]].apply(pd.to_datetime)
        candidates = candidates.merge(price_range, on="tradingitemid", how="left")
        no_price = candidates[candidates["first_price_date"].isna()]
        if not no_price.empty:
            logger.warning(f"{len(no_price)} trading items without any price: {no_price['row'].tolist()}")

        linkage_dates = candidates.groupby("group").apply(estimate_linkage_date).rename("linkage_date")
        candidates = candidates.merge(linkage_dates, left_on="group", right_index=True)
        candidates["start_date"] = candidates["linkage_date"] - timedelta(days=WINDOW_DAYS)
        candidates["end_date"] = candidates["linkage_date"] + timedelta(days=WINDOW_DAYS)

        currencies = group_currencies(candidates)
        candidates = candidates.merge(currencies, left_on="group", right_index=True, how="left")
        changes = currencies[currencies["old_currency"].notna() & currencies["new_currency"].notna()
                             & (currencies["old_currency"] != currencies["new_currency"])]
        logger.info(f"{len(changes)} groups with a currency change between old and new trading items")

        groups_file = os.path.join(OUTPUT_DIR, "linkage_groups.csv")
        candidates.to_csv(groups_file, index=False)
        logger.info(f"Groups saved to: {groups_file}")

        windows = candidates[["tradingitemid", "start_date", "end_date"]].drop_duplicates()
        logger.info(f"Downloading prices for {len(windows)} (trading item, window) pairs...")
        cursor.execute(create_price_query(), (
            [int(t) for t in windows["tradingitemid"]],
            windows["start_date"].dt.strftime("%Y-%m-%d").tolist(),
            windows["end_date"].dt.strftime("%Y-%m-%d").tolist(),
        ))
        prices = pd.DataFrame(cursor.fetchall(), columns=[d[0] for d in cursor.description]).drop_duplicates()
        logger.info(f"Query executed successfully: {len(prices)} price records retrieved")

        prices_file = os.path.join(OUTPUT_DIR, "prices_around_linkage.csv")
        prices.to_csv(prices_file, index=False)
        logger.info(f"Export CSV terminé : {prices_file}")

    except Exception as e:
        logger.error(f"Erreur: {e}")
        raise

    finally:
        if 'conn' in locals():
            conn.close()
            logger.info("Database connection closed")
        logger.info("=== Download Prices Around Linkage Completed ===")
