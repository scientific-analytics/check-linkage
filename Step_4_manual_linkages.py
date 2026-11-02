"""
Step 4 - Add the manually reviewed linkages to the result of Step 3.

  - output/Historical_DocLinkage_v2.xlsx : copy of output/Historical_DocLinkage_new.xlsx with
        DocLinkage : MANUAL_LINKAGES appended at the end (rows whose Link is already in the file are skipped),
                     plus columns old_price, old_price_date (last close of the old trading item before Date)
                     jump_pct (price / old_price - 1, in %) and old_share (shr of the old trading item) after note; every row (Historical_DocLinkage.xlsx,
                     Step 2, manual) with |jump_pct| > MAX_PRICE_JUMP_PCT ("price jump") or more than MAX_GAP_DAYS
                     days between old_price_date and Date ("no linkage") is moved to Decisions,
        Problems   : groups of linkage_problems.csv not solved by a manual linkage, one row per trading item,
                     note filled from problem_notes.csv (old row: note of its linkage, new row: note of the
                     linkage to it),
                     plus CIRCULAR_PROBLEMS (circular linkages left aside, with their reason),
        Decisions  : one row per forced decision (manual linkage added, new security without linkage,
                     circular linkage not validated or left aside), plus the linkages set aside by Step 2
                     because of a price jump above MAX_PRICE_JUMP_PCT (decision "price jump", to check by hand)
                     or of a gap of more than MAX_GAP_DAYS days between old and new prices (decision "no linkage").
  - output/v1.csv : final file, DocLinkage sheet only with the 12 columns of Historical_DocLinkage.xlsx
    (assetid ... note), strings stripped, Date as YYYY-MM-DD,
  - output/manual_linkages_<timestamp>.log

Manual linkages:
  - circular linkages: at the AUC review 2008-12-19 the main line was replaced by another line (Old / Stapled shares)
    and comes back later (return linkage already in DocLinkage); only kept when both lines have the same price,
  - groups with 1 old / 2 new trading items: "Correct linkage" row kept, "New security" row has no linkage.
The source Historical_DocLinkage_new.xlsx is not modified.
"""
import glob
import logging
import os
import re
from copy import copy
from datetime import datetime

import openpyxl
import pandas as pd
import psycopg2
from dotenv import load_dotenv

from Step_2_build_linkage import MAX_GAP_DAYS, MAX_PRICE_JUMP_PCT

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
SOURCE_FILE = os.path.join(OUTPUT_DIR, "Historical_DocLinkage_new.xlsx")
TARGET_FILE = os.path.join(OUTPUT_DIR, "Historical_DocLinkage_v2.xlsx")
FINAL_FILE = os.path.join(OUTPUT_DIR, "v1.csv")
NOTES_FILE = os.path.join(BASE_DIR, "problem_notes.csv")
HEADER_ROW = 2
COLUMNS = ["assetid", "companyid", "securityid", "tradingitemid", "company", "security",
           "Link", "share", "price", "Date", "stamp", "note"]
EXTRA = ["group", "role", "first_price_date", "last_price_date", "reason"]
CIRCULAR_DATE = datetime(2008, 11, 3)
OLD_PRICES_FILE = os.path.join(OUTPUT_DIR, "old_prices.csv")
DATA_GOV = "/home/rudy/PycharmProjects/check_referential/data_gov"
HISTORICAL_FILE = os.path.join(BASE_DIR, "input", "Historical_DocLinkage.xlsx")
INITIAL_FILE_COMMENT = "déjà présent dans le fichier initial"

# group: (assetid, companyid, securityid, tradingitemid, company, security, Link, share, price, Date, comment)
MANUAL_LINKAGES = {
    26: ("3928:FOR", 24766, 2589358, 2589359, "American Tower Corporation", "Class A Common Shares", "1208:VEP",
         392, 32.61, CIRCULAR_DATE, "circular, return linkage on 2012-01-04"),
    40: ("1269:HYB", 24915054, 20096027, 31274260, "AusNet Pty Ltd", "Stapled Securities", "3888:HAS",
         2092.68, 1.08, CIRCULAR_DATE, "circular, return linkage on 2015-06-09"),
    46: ("7410:WEV", 387895, 20003063, 20186118, "Bank of Ireland Group plc", "Old Ordinary Stock", "1329:QAS",
         33.474, 63, CIRCULAR_DATE, "circular, return linkage on 2017-07-10"),
    50: ("1534:SUP", 5856173, 226206415, 226206416, "Beazley plc", "Old Ordinary Shares", "6540:QIP",
         355.8, 1.06, CIRCULAR_DATE, "circular, return linkages on 2009-06-09 and 2016-04-14"),
    85: ("1422:JYM", 874082, 20064357, 20152776, "Cape plc", "Old Ordinary Shares", "5568:BUJ",
         113.838, 0.703, CIRCULAR_DATE, "circular, return linkage on 2011-06-27; AIM line vs Main Market line, "
                                        "prices differ slightly (0.7075 vs 0.7025 on 2008-11-03), accepted"),
    199: ("1380:CEF", 704861, 224398358, 38443856, "Man Group Plc", "Old Ordinary Shares", "4561:TEM",
          1715.269, 3.608, CIRCULAR_DATE, "circular, no return linkage: 4561:TEM priced until 2012-11-09, "
                                          "Man Group absent for 5 reviews, back with 1380:CEF on 2014-06-20; accepted"),
    290: ("6325:WAB", 874650, 20096618, 20200058, "Swiss Re AG", "Common Stock", "5431:JYJ",
          332.648, 49.16, datetime(2008, 12, 19), "circular, linkage at the AUC switch (6325:WAB last in AUC "
                                                   "2008-12-18 at 48.68); no return linkage: both lines in AUC on 2011-09-16 and "
                                         "2011-12-16 (overlap); accepted"),
    91: ("3953:NEG", 4481284, 225701882, 20236041, "Centamin Limited", "Old Ordinary Shares", "7339:QUQ",
         879.019, 0.29, CIRCULAR_DATE, "circular, return linkage on 2011-12-30"),
    98: ("3972:PAV", 884372, 22033321, 22033632, "Coast Entertainment Holdings Limited", "Stapled Securities",
         "1542:VIQ", 236.632, 1.25, CIRCULAR_DATE, "circular, return linkage on 2018-11-30"),
    203: ("4522:PIS", 875425, 20093942, 20219295, "Mebuki Financial GroupInc.", "Old Common Stock", "2701:VYT",
          917.403, 411.966, CIRCULAR_DATE, "circular, return linkage on 2016-10-03"),
    56: ("6196:BAN", 263424, 172042562, 172047428, "Biglari Holdings Inc.", "Class B Common Stock", "5371:PYR",
         2.068, 242.69, datetime(2018, 5, 1), "correct linkage, 1553:TYQ Class A is a new security (no linkage)"),
    143: ("1094:QIW", 22986276, 266052764, 266673936, "Fidelity National Financial Inc.", "FNF Group Common Stock",
          "2291:FYR", 276, 32.76, datetime(2014, 6, 30), "correct linkage, 3806:KIF FNFV is a new security (no linkage)"),
    157: ("1119:MAK", 671670, 83094656, 83100652, "H. Lundbeck A/S", "B Shares", "3414:HIW",
          794.272, 32.5, datetime(2022, 6, 10), "correct linkage, 6922:SIK A Shares is a new security (no linkage)"),
    287: ("3897:BOD", 11420997, 143092165, 143092166, "SunPower Corporation", "Common Stock", "7045:MEM",
          119.076, 7.50, datetime(2011, 11, 17), "correct linkage, 1-for-1 reclassification of Class A and Class B into "
                                            "a single Common Stock class (SEC-confirmed), linked via Class A "
                                            "(most liquid); 7965:KUG Class B not linked (no linkage)"),
    341: ("7378:NOD", 22666093, 1776257809, 1776948460, "Warner Bros. Discovery Inc.", "New Ordinary Shares",
          "7463:MON", 2426.844, 24.43, datetime(2022, 4, 8), "correct linkage, 1-for-1 reclassification of Series A, B "
                                                        "and C into New Ordinary Shares following the WarnerMedia/"
                                                        "Discovery merger (SEC 10-K confirmed), linked via Series A "
                                                        "(most liquid, price identical to new line every day in "
                                                        "the overlap window); 6251:CEV Series C not linked (no "
                                                        "linkage)"),
}

# group: (company, tradingitemid, asset, security) of the new trading items that are new securities, not linkages
NEW_SECURITIES = {
    56: ("Biglari Holdings Inc.", 285880067, "1553:TYQ", "Class A common stock"),
    143: ("Fidelity National Financial Inc.", 266673935, "3806:KIF", "FNFV Group Common Stock"),
    157: ("H. Lundbeck A/S", 83100651, "6922:SIK", "A Shares"),
}
# group: (company, [(tradingitemid, asset, security) of the new trading items], comment) of groups without linkage
NO_LINKAGE = {
    19: ("Alphabet Inc.", [(11311662, "2788:VUH", "Class A Common Stock"), (260152426, "3742:JAS", "Class C Capital Stock")],
         "2 old / 2 new trading items (Old Class A 5093:DOB, Old Class C 3723:KAH, last price 2015-10-05): no linkage"),
    129: ("Esprit Energy Trust", [(22941855, "7133:CAR", "Trust Units")],
          "2 old / 1 new trading items (Class A 5761:MYL, Class B 1018:KUT merged into Trust Units on 2005-07-05): "
          "not validated because of a price adjustment factor (Common Shares 7149:VOQ exchanged ~4 for 1 trust unit "
          "+ ProspEx shares on 2004-10-05, 4.03 -> ~13, not adjusted in CIQ)"),
    189: ("Liberty Global Ltd.", [(22743070, "1899:QIC", "Liberty Global Class A Common Stock"),
                                  (23708238, "4888:GIK", "Liberty Global Class C Common Stock")],
          "2 old / 2 new trading items (Old Series A 1615:COJ, Old Series C 3163:SOD, last price 2013-06-07): "
          "no linkage, 1:1 share exchange Liberty Global Inc -> Liberty Global plc following Virgin Media "
          "acquisition but price and share count different"),
    193: ("Lionsgate Studios Corp.", [(1941153615, "6501:LUN", "Common Shares")],
          "2 old / 1 new trading items (Class A Voting 4271:VYN, Class B Non-Voting 4538:PEJ, last price "
          "2025-05-06): no linkage, Class A and Class B merged into Common Shares on 2025-05-06, non 1:1 ratio"),
    194: ("Lionsgate Studios Corp.", [(412297487, "4271:VYN", "Class A Voting Shares"),
                                      (412297486, "4538:PEJ", "Class B Non-Voting Shares")],
          "1 old / 2 new trading items (Old Common Shares 3138:MOG, last price 2016-12-08): no linkage, split "
          "into Class A Voting and Class B Non-Voting Shares on 2016-12-09, non 1:1 ratio"),
    204: ("Mediclinic Group Limited", [(241673056, "4650:POQ", "Ordinary Shares")],
          "1 old / 1 new trading items (Old Ordinary Shares 1404:PYF, Johannesburg ZAR, last price 2016-02-05): "
          "no linkage, merger with Al Noor Hospitals creating Mediclinic International plc, listed in London GBP "
          "from 2016-02-16 after a trading halt, currency change and non 1:1 share exchange ratio"),
    230: ("Odfjell SE", [(22260586, "1367:QIR", "Series A Shares")],
          "1 old / 1 new trading items (Series B Shares 7195:FAB): no linkage, Series A and Series B are two "
          "permanent share classes both fully listed since 1992, overlapping in AUC (both present 2005-03-18 and "
          "2005-06-17), AUC just alternates which one is the main line"),
    241: ("Paramount Skydance Corporation", [(24993622, "2719:LYV", "Old Class B Common Stock"),
                                              (25082106, "6685:TYP", "Old Class A Common Stock")],
          "2 old / 2 new trading items (Former Viacom Class B 3442:HAM, Former Viacom Class A 2776:SIP): no "
          "linkage, Old Viacom split into CBS Corporation (this companyid) and New Viacom Inc. (companyid "
          "25403751) on 2006-01-03, non-uniform share ratio (0.47 for B, 0.50 for A), no overlap in AUC between "
          "2005-12-16 and 2006-03-17 reviews"),
    240: ("Paramount Skydance Corporation", [(1955787307, "4450:RAB", "Class B Common Stock")],
          "2 old / 1 new trading items (Old Class A 6685:TYP, Old Class B 2719:LYV, last price 2025-08-06): no "
          "linkage, consolidation into new Class B Common Stock (Skydance Corp.) following Skydance's acquisition "
          "of Paramount Global, non-uniform ratio (Class A 1.5:1, Class B 1:1), no overlap in AUC between "
          "2025-06-20 and 2025-09-19 reviews"),
    288: ("Sunshine Lake Pharma Co. Ltd.", [(1949996720, "2966:WUC", "Foreign Shares-Foreign Listed")],
          "1 old / 1 new trading items (Old Foreign Shares-Foreign Listed 3095:BEG, last price 13.02 on "
          "2025-07-29): no linkage, 7 business day gap and price jump to 56.25 (~4.3x), too large to accept as a "
          "continuation"),
    299: ("Telecom Italia S.p.A.", [(20221167, "2009:RYC", "Ordinary Shares"),
                                    (65931182, "6795:NYB", "Saving Shares")],
          "1 old / 2 new trading items (Old Ordinary Shares 3486:VYB, last price 1.014 on 2003-08-01): no linkage, "
          "reverse merger between Telecom Italia and parent Olivetti S.p.A. in 2003, Olivetti issuing new Ordinary "
          "(2.245) and Saving (1.483) shares to former Telecom Italia/Olivetti holders, price and share count not "
          "consistent between old and new (not a clean ratio)"),
    351: ("Wilh. Wilhelmsen Holding ASA", [(108503363, "3396:MAM", "Class B Shares")],
          "1 old / 1 new trading items (Class A Shares 3211:TYJ): no linkage, Class A and Class B are two "
          "permanent share classes both fully listed since 1992, company only sporadically in AUC (12 of 98 "
          "reviews, with ~7y and ~6.75y gaps), not a corporate event"),
    342: ("Warner Bros. Discovery Inc.", [(49031559, "7378:NOD", "Series A Common Stock"),
                                           (49031561, "6251:CEV", "Series C Common Stock")],
          "1 old / 2 new trading items (Old Series A Common Stock 7483:SIM, last price 18.30 on 2008-09-17): no "
          "linkage, Discovery Holding Company restructured into Discovery Communications on 2008-09-18, each old "
          "share converted into 0.5 new Series A + 0.5 new Series C share (SEC S-4 confirmed), non-1:1 ratio"),
}
DECISION_COLUMNS =["group", "company", "decision", "assetid", "tradingitemid", "security", "Link",
                    "share", "price", "Date", "comment", "decided_on"]

# group: (assetid, companyid, securityid, tradingitemid, company, security, Link, share, price, reason)
CIRCULAR_PROBLEMS = {}

# Link: (group, comment) of linkages proposed by Step 2 (already in Historical_DocLinkage_new.xlsx) not validated,
# removed from DocLinkage
REMOVED_LINKAGES = {
    "6441:MUT": (232, "Old Mutual return linkage 1247:WOG (London GBP) -> 6441:MUT (Johannesburg ZAR) on 2018-06-26, "
                      "not validated with the circular linkage of group 233"),
}

# group: (assetid, companyid, securityid, tradingitemid, company, security, Link, share, price, reason)
# circular linkages reviewed and not validated: in Decisions only, neither in DocLinkage nor in Problems
NOT_VALIDATED = {
    233: ("6441:MUT", 663731, 20098673, 20130168, "Old Mutual Limited", "Old Ordinary Shares", "1247:WOG", 4407.094,
          0.661, "Different exchanges and currencies: 6441:MUT Johannesburg ZAR (8.54 on 2008-12-18) vs 1247:WOG "
                 "London GBP (0.575 on 2008-12-19); return linkage on 2018-06-26 (GBP -> ZAR)"),
    334: ("6109:HAJ", 879755, 243589439, 241958866, "WCT Holdings Berhad", "Old Ordinary Shares", "6327:BYK", 884.246,
          1.209, "WCT in AUC only for 3 reviews (2008-03-21, 2008-09-19 on 6109:HAJ, 2008-12-19 on 6327:BYK) "
                 "and never comes back; prices differ at 6th decimal"),
}


def setup_logging():
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_file = os.path.join(OUTPUT_DIR, f"manual_linkages_{timestamp}.log")
    logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s",
                        handlers=[logging.FileHandler(log_file), logging.StreamHandler()])
    return logging.getLogger("manual_linkages")


def copy_style(cell, template):
    cell.number_format = template.number_format
    cell.font = copy(template.font)


def remove_linkages(ws, header, logger):
    """Delete the rows of REMOVED_LINKAGES from DocLinkage, return their values for the Decisions sheet."""
    removed = []
    for r in range(ws.max_row, HEADER_ROW, -1):
        link = ws.cell(r, header["Link"]).value
        if link in REMOVED_LINKAGES:
            values = {col: ws.cell(r, header[col]).value for col in COLUMNS}
            group, comment = REMOVED_LINKAGES[link]
            removed.append((group, values, comment))
            ws.delete_rows(r)
            logger.info(f"group {group} {values['company']}: removed {values['assetid']} -> {link} "
                        f"price {values['price']} on {values['Date']:%Y-%m-%d} ({comment})")
    groups = pd.read_csv(os.path.join(OUTPUT_DIR, "linkage_groups.csv"))
    problems = pd.read_csv(os.path.join(OUTPUT_DIR, "linkage_problems.csv")).set_index("group")["reason"]
    prices = pd.read_csv(os.path.join(OUTPUT_DIR, "prices_around_linkage.csv"), parse_dates=["pricingdate"])
    for link in set(REMOVED_LINKAGES) - {v["Link"] for _, v, _ in removed}:
        group, comment = REMOVED_LINKAGES[link]
        company = groups.loc[groups["group"] == group, "company"].iloc[0]
        values = {**dict.fromkeys(COLUMNS), "company": company, "Link": link}
        if group in problems.index and problems[group].startswith(("price jump of", "no linkage: gap")):
            # linkage set aside by Step 2: values of the linkage Step 2 would have proposed
            row = step2_decision(group, None, problems[group], groups, prices)
            values.update(zip(["assetid", "tradingitemid", "security", "Link", "share", "price", "Date"], row[3:10]))
            comment = f"{comment}; {problems[group]}"
        removed.append((group, values, comment))
        logger.info(f"group {group} {company}: linkage to {link} not in DocLinkage (set aside by Step 2), "
                    f"kept as removed in Decisions ({comment})")
    return removed


def append_manual_linkages(ws, header, logger):
    last_row = max(r for r in range(HEADER_ROW + 1, ws.max_row + 1) if ws.cell(r, header["assetid"]).value)
    existing_links = {ws.cell(r, header["Link"]).value for r in range(HEADER_ROW + 1, last_row + 1)}
    logger.info(f"DocLinkage: {last_row - HEADER_ROW} rows in {SOURCE_FILE}")
    added = 0
    for group, row in MANUAL_LINKAGES.items():
        *values, comment = row
        if values[6] in existing_links:
            logger.warning(f"group {group} {values[4]}: Link {values[6]} already in DocLinkage, skipped")
            continue
        added += 1
        for col, value in zip(COLUMNS, values + ["AUC v2", "-"]):
            copy_style(ws.cell(last_row + added, header[col], value), ws.cell(last_row, header[col]))
        logger.info(f"group {group} {values[4]}: added {values[0]} -> {values[6]} share {values[7]} "
                    f"price {values[8]} on {values[9]:%Y-%m-%d} ({comment})")
    logger.info(f"DocLinkage: {added} manual linkages appended at row {last_row + 1}")
    return added


def load_prices():
    """Daily close by trading item: prices_around_linkage.csv and output/auc_timeseries/*/full_timeseries.csv."""
    prices = pd.read_csv(os.path.join(OUTPUT_DIR, "prices_around_linkage.csv"), parse_dates=["pricingdate"])
    frames = [prices[["tradingitemid", "pricingdate", "priceclose"]]]
    for path in glob.glob(os.path.join(OUTPUT_DIR, "auc_timeseries", "*", "full_timeseries.csv")):
        wide = pd.read_csv(path, parse_dates=["pricingdate"]).set_index("pricingdate")
        wide.columns = [int(re.search(r"\((\d+)\)$", c).group(1)) for c in wide.columns]
        frames.append(wide.stack().rename("priceclose").reset_index().rename(columns={"level_1": "tradingitemid"}))
    if os.path.exists(OLD_PRICES_FILE):
        frames.append(pd.read_csv(OLD_PRICES_FILE, parse_dates=["pricingdate"])[["tradingitemid", "pricingdate",
                                                                                  "priceclose"]])
    data = pd.concat(frames).drop_duplicates(["tradingitemid", "pricingdate"]).sort_values("pricingdate")
    return {ti: d.set_index("pricingdate")["priceclose"] for ti, d in data.groupby("tradingitemid")}


def load_auc_assets():
    """Trading items of each asset with their period (since, until), from the latest data_gov auc_asset.csv."""
    latest = sorted(glob.glob(os.path.join(DATA_GOV, "2*", "_universe", "auc_asset.csv")))[-1]
    assets = pd.read_csv(latest, parse_dates=["since"])
    assets["until"] = pd.to_datetime(assets["until"], errors="coerce").fillna(pd.Timestamp.max)
    return assets


def chain_old_assets(ws, header):
    """(tradingitemid, Link) -> old asset: Link of the previous linkage of the same assetid chain (by Date),
    else assetid."""
    rows = [(str(ws.cell(r, header["assetid"]).value).strip(), ws.cell(r, header["Date"]).value,
             int(ws.cell(r, header["tradingitemid"]).value), str(ws.cell(r, header["Link"]).value).strip())
            for r in range(HEADER_ROW + 1, ws.max_row + 1) if ws.cell(r, header["assetid"]).value]
    old, previous = {}, {}
    for assetid, _, ti, link in sorted(rows, key=lambda x: (x[0], x[1])):
        old[(ti, link)] = previous.get(assetid, assetid)
        previous[assetid] = link
    return old


def auc_shares(tradingitemids):
    """shr of each trading item at its last AUC review, from data_gov/<date>/_universe/auc.csv."""
    frames = []
    for path in sorted(glob.glob(os.path.join(DATA_GOV, "2*", "_universe", "auc.csv"))):
        auc = pd.read_csv(path, usecols=["date", "tradingitemid", "shr"])
        frames.append(auc[auc["tradingitemid"].isin(tradingitemids)])
    data = pd.concat(frames)
    return data.sort_values("date").groupby("tradingitemid")["shr"].last()


def download_old_prices(requests, logger):
    """Last CIQ close strictly before Date of each (tradingitemid, Date), appended to OLD_PRICES_FILE.
    Skipped with a warning when CIQ is not reachable."""
    load_dotenv(os.path.join(BASE_DIR, ".env"))
    try:
        conn = psycopg2.connect(host=os.getenv("DB_HOST"), database=os.getenv("DB_NAME"), user=os.getenv("DB_USER"),
                                password=os.getenv("DB_PASSWORD"), port=os.getenv("DB_PORT"), connect_timeout=15)
    except psycopg2.OperationalError as e:
        logger.warning(f"CIQ not reachable, {len(requests)} old prices not downloaded: {str(e).splitlines()[0]}")
        return False
    query = """
    WITH w AS (SELECT unnest(%s::bigint[]) AS tradingitemid, unnest(%s::date[]) AS linkage_date)
    SELECT w.tradingitemid, w.linkage_date, p.pricingdate, p.priceclose
    FROM w CROSS JOIN LATERAL (
        SELECT pricingdate, priceclose FROM public.ciqpriceequity p
        WHERE p.tradingitemid = w.tradingitemid AND p.pricingdate < w.linkage_date
        ORDER BY pricingdate DESC LIMIT 1) p
    """
    with conn, conn.cursor() as cur:
        cur.execute(query, ([int(t) for t, _ in requests], [d.date() for _, d in requests]))
        data = pd.DataFrame(cur.fetchall(), columns=["tradingitemid", "linkage_date", "pricingdate", "priceclose"])
    conn.close()
    data["pricingdate"] = pd.to_datetime(data["pricingdate"])
    data["priceclose"] = data["priceclose"].astype(float)
    if os.path.exists(OLD_PRICES_FILE):
        data = pd.concat([pd.read_csv(OLD_PRICES_FILE, parse_dates=["pricingdate"]), data])
    data.drop_duplicates(["tradingitemid", "linkage_date"], keep="last") \
        .sort_values(["tradingitemid", "pricingdate"]).to_csv(OLD_PRICES_FILE, index=False)
    logger.info(f"CIQ: last old price before Date downloaded for {len(requests)} linkages into {OLD_PRICES_FILE}")
    return True


def old_trading_item(new_ti, assetid, link, groups, old_assets, chain_assets, auc_assets, date):
    """Old trading item of a linkage, in the candidate group of the new trading item: the one of old_asset of
    Step 2; otherwise (manual linkages, rows of Historical_DocLinkage.xlsx) the trading item of the old asset of the
    chain (previous Link of the same assetid), searched in the candidate group, then in all groups, then in
    auc_asset.csv (alive the day before Date); last resort the only old trading item of the group where it is new."""
    in_groups = groups.loc[groups["tradingitemid"] == new_ti, "group"]
    others = groups[groups["group"].isin(in_groups) & (groups["tradingitemid"] != new_ti)]
    if (new_ti, link) in old_assets:
        match = others[others["asset"] == old_assets[(new_ti, link)]]
        if not match.empty:
            return int(match["tradingitemid"].iloc[0])
    old_asset = chain_assets.get((new_ti, link), assetid)
    match = others[others["asset"] == old_asset]
    if not match.empty:
        return int(match["tradingitemid"].iloc[0])
    match = groups[(groups["asset"] == old_asset) & (groups["tradingitemid"] != new_ti)].sort_values("date")
    if not match.empty:
        return int(match["tradingitemid"].iloc[-1])
    day_before = date - pd.Timedelta(days=1)
    match = auc_assets[(auc_assets["asset"] == old_asset) & (auc_assets["tradingitemid"] != new_ti)]
    alive = match[(match["since"] <= day_before) & (match["until"] >= day_before)]
    match = alive if not alive.empty else match
    if not match.empty:
        return int(match["tradingitemid"].iloc[-1])
    as_new = groups.loc[(groups["tradingitemid"] == new_ti) & (groups["role"] == "new"), "group"]
    olds = others.loc[others["group"].isin(as_new) & (others["role"] == "old"), "tradingitemid"].unique()
    return int(olds[0]) if len(olds) == 1 else None


def add_old_price(ws, header, logger):
    """Columns after note: old_price and old_price_date (last close of the old trading item before Date; downloaded from CIQ into OLD_PRICES_FILE when not in the local prices or older than
    MAX_GAP_DAYS), jump_pct (price / old_price - 1, in %) and old_share (shr of the old
    trading item at its last review in Link-candidates.xlsx, else in data_gov auc.csv).
    No old price for a row already in HISTORICAL_FILE: old_price = INITIAL_FILE_COMMENT.
    Rows with a price jump above MAX_PRICE_JUMP_PCT or a gap above MAX_GAP_DAYS are removed from DocLinkage,
    returned as Decisions rows ("price jump" / "no linkage")."""
    manual_groups = {row[6]: group for group, row in MANUAL_LINKAGES.items()}
    groups = pd.read_csv(os.path.join(OUTPUT_DIR, "linkage_groups.csv"))
    proposed = pd.read_csv(os.path.join(OUTPUT_DIR, "proposed_linkage.csv"))
    old_assets = dict(zip(zip(proposed["tradingitemid"], proposed["Link"]), proposed["old_asset"]))
    initial = pd.read_excel(HISTORICAL_FILE, header=HEADER_ROW - 1).dropna(subset=["Link"])
    initial_rows = set(zip(initial["tradingitemid"].astype(int), initial["Link"].astype(str).str.strip()))
    chain_assets = chain_old_assets(ws, header)
    auc_assets = load_auc_assets()
    linkages = []
    for r in range(HEADER_ROW + 1, ws.max_row + 1):
        values = {c: ws.cell(r, header[c]).value for c in COLUMNS}
        if not values["assetid"]:
            continue
        date = pd.Timestamp(values["Date"])
        key = (int(values["tradingitemid"]), str(values["Link"]).strip())
        old_ti = old_trading_item(key[0], str(values["assetid"]).strip(), key[1], groups, old_assets,
                                  chain_assets, auc_assets, date)
        linkages.append((r, values, date, key, old_ti))
    prices = load_prices()

    def last_close(old_ti, date):
        old = prices.get(old_ti)
        old = old[old.index < date] if old is not None else None
        return None if old is None or old.empty else (old.index[-1], float(old.iloc[-1]))

    downloaded = pd.read_csv(OLD_PRICES_FILE, parse_dates=["linkage_date"]) if os.path.exists(OLD_PRICES_FILE) \
        else pd.DataFrame(columns=["tradingitemid", "linkage_date"])
    done = set(zip(downloaded["tradingitemid"], downloaded["linkage_date"]))
    requests = [(old_ti, date) for _, _, date, _, old_ti in linkages if old_ti is not None
                and (old_ti, date) not in done
                and (last_close(old_ti, date) is None or (date - last_close(old_ti, date)[0]).days > MAX_GAP_DAYS)]
    if requests and download_old_prices(requests, logger):
        prices = load_prices()
    cols = {name: header["note"] + i
            for i, name in enumerate(["old_price", "old_price_date", "jump_pct", "old_share"], start=1)}
    old_shares = groups.sort_values("date").groupby("tradingitemid")["shr"].last()
    not_in_groups = {old_ti for *_, old_ti in linkages if old_ti is not None} - set(old_shares.index)
    if not_in_groups:
        old_shares = pd.concat([old_shares, auc_shares(not_in_groups)])
    src = ws.cell(HEADER_ROW, header["note"])
    for name, col in cols.items():
        cell = ws.cell(HEADER_ROW, col, name)
        cell.font, cell.fill = copy(src.font), copy(src.fill)
        cell.alignment, cell.border = copy(src.alignment), copy(src.border)
        template = header["Date"] if name == "old_price_date" else header["share" if name == "old_share" else "price"]
        ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = \
            ws.column_dimensions[openpyxl.utils.get_column_letter(template)].width
    missing, moved, to_delete = 0, [], []
    for r, values, date, key, old_ti in linkages:
        if old_ti in old_shares.index:
            copy_style(ws.cell(r, cols["old_share"], round(float(old_shares[old_ti]), 3)), ws.cell(r, header["share"]))
        old = last_close(old_ti, date)
        if old is None:
            missing += 1
            in_initial = key in initial_rows
            if in_initial:
                cell = ws.cell(r, cols["old_price"], INITIAL_FILE_COMMENT)
                cell.font = copy(ws.cell(r, header["note"]).font)
            logger.warning(f"{str(values['company']).strip()} {values['assetid']} -> {values['Link']} on "
                           f"{date:%Y-%m-%d}: no price of the old trading item {old_ti} before Date, old_price "
                           + (f"'{INITIAL_FILE_COMMENT}'" if in_initial else "left empty"))
            continue
        old_date, old_price = old
        copy_style(ws.cell(r, cols["old_price"], round(old_price, 4)), ws.cell(r, header["price"]))
        copy_style(ws.cell(r, cols["old_price_date"], old_date.to_pydatetime()), ws.cell(r, header["Date"]))
        jump = (values["price"] / old_price - 1) * 100
        cell = ws.cell(r, cols["jump_pct"], round(jump, 2))
        cell.font = copy(ws.cell(r, header["price"]).font)
        cell.number_format = "0.00"
        gap_days = (date - old_date).days
        origin = ("initial file" if key in initial_rows else
                  "manual linkage" if key[1] in manual_groups else "Step 2")
        if gap_days > MAX_GAP_DAYS:
            decision = "no linkage"
            comment = (f"no linkage: gap of {gap_days} days between old last price {old_price:g} on "
                       f"{old_date:%Y-%m-%d} and new price {values['price']:g} on {date:%Y-%m-%d} ({origin})")
        elif abs(jump) > MAX_PRICE_JUMP_PCT:
            decision = "price jump"
            comment = (f"price jump of {jump:+.1f}%: old {old_price:g} on {old_date:%Y-%m-%d}, new "
                       f"{values['price']:g} on {date:%Y-%m-%d} ({origin})")
        else:
            continue
        moved.append([manual_groups.get(key[1]), str(values["company"]).strip(), decision,
                      str(values["assetid"]).strip(), values["tradingitemid"], values["security"], key[1],
                      values["share"], values["price"], values["Date"], comment])
        to_delete.append(r)
        logger.info(f"{str(values['company']).strip()} {values['assetid']} -> {key[1]}: moved from DocLinkage to "
                    f"Decisions ({comment})")
    for r in reversed(to_delete):
        ws.delete_rows(r)
    letters = ", ".join(openpyxl.utils.get_column_letter(c) for c in cols.values())
    logger.info(f"DocLinkage: old_price, old_price_date, jump_pct, old_share added in columns {letters}, "
                f"{missing} rows without old price, {len(moved)} rows moved to Decisions")
    return moved


def step2_decision(group, decision, reason, groups, prices):
    """Decisions row of a group set aside by Step 2 for a price jump or a price gap: the linkage Step 2 would have
    proposed, Date and price of the new trading item read at the end of the reason ("... X on YYYY-MM-DD")."""
    old = groups[(groups["group"] == group) & (groups["role"] == "old")].iloc[0]
    new = groups[(groups["group"] == group) & (groups["role"] == "new")].iloc[0]
    date = pd.Timestamp(re.search(r"on (\d{4}-\d{2}-\d{2})$", reason).group(1))
    close = prices[(prices["tradingitemid"] == new["tradingitemid"]) & (prices["pricingdate"] == date)]["priceclose"]
    return [group, new["company"], decision, old["asset"], new["tradingitemid"], new["security"], new["asset"],
            round(new["shr"], 3), round(float(close.iloc[0]), 4), date.to_pydatetime(), reason]


def problem_rows(logger):
    """Groups of linkage_problems.csv not solved by a manual linkage, one row per trading item, and the Decisions
    rows of the groups set aside by Step 2 for a price jump."""
    groups = pd.read_csv(os.path.join(OUTPUT_DIR, "linkage_groups.csv"),
                         parse_dates=["first_price_date", "last_price_date"])
    problems = pd.read_csv(os.path.join(OUTPUT_DIR, "linkage_problems.csv"))
    prices = pd.read_csv(os.path.join(OUTPUT_DIR, "prices_around_linkage.csv"), parse_dates=["pricingdate"])
    rows, jumps = [], []
    for _, p in problems.iterrows():
        g = p["group"]
        if g in MANUAL_LINKAGES:
            logger.info(f"group {g} {p['company']}: solved by a manual linkage, not in Problems")
            continue
        if g in NO_LINKAGE:
            logger.info(f"group {g} {p['company']}: no linkage decided, not in Problems ({NO_LINKAGE[g][2]})")
            continue
        removed = [link for link, (group, _) in REMOVED_LINKAGES.items() if group == g]
        if removed:
            logger.info(f"group {g} {p['company']}: linkage to {removed[0]} not validated, not in Problems "
                        f"({REMOVED_LINKAGES[removed[0]][1]})")
            continue
        if p["reason"].startswith(("price jump of", "no linkage: gap")):
            decision = "price jump" if p["reason"].startswith("price jump") else "no linkage"
            jumps.append(step2_decision(g, decision, p["reason"], groups, prices))
            logger.info(f"group {g} {p['company']}: in Decisions, not in Problems ({p['reason']})")
            continue
        if g in NOT_VALIDATED:
            logger.info(f"group {g} {p['company']}: circular linkage not validated, not in Problems "
                        f"({NOT_VALIDATED[g][-1]})")
            continue
        if g in CIRCULAR_PROBLEMS:
            *values, reason = CIRCULAR_PROBLEMS[g]
            new = groups[(groups["group"] == g) & (groups["role"] == "new")].iloc[0]
            rows.append({**dict(zip(COLUMNS, values + [CIRCULAR_DATE, "AUC v2", "-"])), "group": g, "role": "new",
                         "first_price_date": new["first_price_date"], "last_price_date": new["last_price_date"],
                         "reason": "Circular linkage: " + reason})
            logger.info(f"group {g} {p['company']}: circular linkage left aside ({reason})")
            continue
        grp = groups[groups["group"] == g].sort_values(["role", "row"], ascending=[False, True])
        old_assets = grp.loc[grp["role"] == "old", "asset"].tolist()
        for _, t in grp.iterrows():
            is_old = t["role"] == "old"
            rows.append({"assetid": t["asset"] if is_old else (old_assets[0] if len(old_assets) == 1 else None),
                         "companyid": t["companyid"], "securityid": t["securityid"],
                         "tradingitemid": t["tradingitemid"], "company": t["company"], "security": t["security"],
                         "Link": None if is_old else t["asset"], "share": round(t["shr"], 3), "price": None,
                         "Date": None, "stamp": "AUC v2", "note": "-", "group": g, "role": t["role"],
                         "first_price_date": t["first_price_date"], "last_price_date": t["last_price_date"],
                         "reason": p["reason"]})
        logger.info(f"group {g} {p['company']}: {len(grp)} trading items in Problems ({p['reason']})")
    return rows, jumps


def apply_problem_notes(rows, logger):
    """Note of each Problems row from NOTES_FILE, matched within the group on (assetid, Link)."""
    notes = pd.read_csv(NOTES_FILE)
    used = set()
    for g in {r["group"] for r in rows}:
        group_rows = [r for r in rows if r["group"] == g]
        olds = {r["assetid"] for r in group_rows if r["role"] == "old"}
        news = {r["Link"] for r in group_rows if r["role"] == "new"}
        for r in group_rows:
            if r["role"] == "old":
                match = notes[(notes["assetid"] == r["assetid"]) & notes["Link"].isin(news)]
            else:
                match = notes[(notes["Link"] == r["Link"]) & notes["assetid"].isin(olds)]
            if match.empty:
                continue
            used.update(match.index)
            r["note"] = " | ".join(dict.fromkeys(match["note"]))
    logger.info(f"Problems: {len(used)}/{len(notes)} notes of {NOTES_FILE} used")
    for _, n in notes.drop(index=list(used)).iterrows():
        logger.warning(f"note not used, {n['company']} {n['assetid']} -> {n['Link']} not in Problems")
    return rows


def write_problems_sheet(wb, src, header, rows):
    ws = wb.create_sheet("Problems")
    ncol = src.max_column
    for r in (1, HEADER_ROW):
        for c in range(1, ncol + 1):
            s, cell = src.cell(r, c), ws.cell(r, c, src.cell(r, c).value)
            cell.font, cell.fill, cell.alignment, cell.border = copy(s.font), copy(s.fill), copy(s.alignment), copy(s.border)
    for c in range(1, ncol + 1):
        letter = openpyxl.utils.get_column_letter(c)
        ws.column_dimensions[letter].width = src.column_dimensions[letter].width
    note_header = src.cell(HEADER_ROW, header["note"])
    extra_cols = {name: ncol + i for i, name in enumerate(EXTRA, start=1)}
    for name, c in extra_cols.items():
        cell = ws.cell(HEADER_ROW, c, name)
        cell.font, cell.fill = copy(note_header.font), copy(note_header.fill)
        cell.alignment, cell.border = copy(note_header.alignment), copy(note_header.border)
        ws.column_dimensions[openpyxl.utils.get_column_letter(c)].width = 110 if name == "reason" else 14

    template_row = HEADER_ROW + 1
    for i, row in enumerate(rows, start=template_row):
        for col in COLUMNS:
            value = row[col].to_pydatetime() if isinstance(row[col], pd.Timestamp) else row[col]
            copy_style(ws.cell(i, header[col], value), src.cell(template_row, header[col]))
        for name, c in extra_cols.items():
            value = row[name]
            value = value.to_pydatetime() if isinstance(value, pd.Timestamp) else (None if pd.isna(value) else value)
            cell = ws.cell(i, c, value)
            cell.font = copy(src.cell(template_row, header["note"]).font)
            if name.endswith("_date"):
                cell.number_format = src.cell(template_row, header["Date"]).number_format
    ws.freeze_panes = src.freeze_panes


def decision_rows(removed, jumps, moved):
    rows = list(jumps) + list(moved)
    moved_links = {r[6] for r in moved}
    for group, v, comment in removed:
        rows.append([group, v["company"], "linkage removed", v["assetid"], v["tradingitemid"], v["security"], v["Link"],
                     v["share"], v["price"], v["Date"], comment])
    for group, (assetid, _, _, ti, company, security, link, share, price, date, comment) in MANUAL_LINKAGES.items():
        if link in moved_links:
            continue
        rows.append([group, company, "manual linkage added", assetid, ti, security, link, share, price, date, comment])
    for group, (company, ti, asset, security) in NEW_SECURITIES.items():
        rows.append([group, company, "new security, no linkage", None, ti, security, asset, None, None, None,
                     f"new trading item of the group with {MANUAL_LINKAGES[group][6]}"])
    for group, (company, items, comment) in NO_LINKAGE.items():
        for ti, asset, security in items:
            rows.append([group, company, "no linkage", None, ti, security, asset, None, None, None, comment])
    for group, (assetid, _, _, ti, company, security, link, share, price, reason) in NOT_VALIDATED.items():
        rows.append([group, company, "circular linkage not validated", assetid, ti, security, link, share, price,
                     CIRCULAR_DATE, reason])
    for group, (assetid, _, _, ti, company, security, link, share, price, reason) in CIRCULAR_PROBLEMS.items():
        rows.append([group, company, "circular linkage left aside", assetid, ti, security, link, share, price,
                     CIRCULAR_DATE, reason])
    today = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
    return sorted((r + [today] for r in rows), key=lambda r: (r[1], -1 if r[0] is None else r[0], r[2]))


def write_decisions_sheet(wb, src, header, rows):
    ws = wb.create_sheet("Decisions")
    header_cell = src.cell(HEADER_ROW, header["note"])
    template = HEADER_ROW + 1
    widths = {"company": 36, "decision": 28, "security": 26, "comment": 110}
    for c, name in enumerate(DECISION_COLUMNS, start=1):
        cell = ws.cell(HEADER_ROW, c, name)
        cell.font, cell.fill = copy(header_cell.font), copy(header_cell.fill)
        cell.alignment, cell.border = copy(header_cell.alignment), copy(header_cell.border)
        ws.column_dimensions[openpyxl.utils.get_column_letter(c)].width = widths.get(name, 14)
    for i, row in enumerate(rows, start=template):
        for c, (name, value) in enumerate(zip(DECISION_COLUMNS, row), start=1):
            cell = ws.cell(i, c, value)
            cell.font = copy(src.cell(template, header["note"]).font)
            if name in header:
                cell.number_format = src.cell(template, header[name]).number_format
            elif name == "decided_on":
                cell.number_format = src.cell(template, header["Date"]).number_format
    ws.freeze_panes = ws.cell(template, 1)


if __name__ == "__main__":
    logger = setup_logging()
    logger.info("=== Starting Manual Linkages ===")
    wb = openpyxl.load_workbook(SOURCE_FILE)
    ws = wb["DocLinkage"]
    header = {ws.cell(HEADER_ROW, c).value: c for c in range(1, ws.max_column + 1)}

    removed = remove_linkages(ws, header, logger)
    added = append_manual_linkages(ws, header, logger)
    moved = add_old_price(ws, header, logger)
    rows, jumps = problem_rows(logger)
    rows = apply_problem_notes(rows, logger)
    write_problems_sheet(wb, ws, header, rows)
    decisions = decision_rows(removed, jumps, moved)
    write_decisions_sheet(wb, ws, header, decisions)
    for d in decisions:
        if d[2] == "new security, no linkage":
            logger.info(f"group {d[0]} {d[1]}: {d[6]} {d[5]} is a new security, no linkage")
    wb.save(TARGET_FILE)

    doc = pd.read_excel(TARGET_FILE, sheet_name="DocLinkage", header=HEADER_ROW - 1).dropna(subset=["Link"])
    duplicated = doc["Link"].duplicated().sum()
    if duplicated:
        logger.warning(f"{duplicated} duplicated Link in DocLinkage")
    logger.info(f"{TARGET_FILE}: DocLinkage {len(doc)} rows ({added - sum(r[0] in MANUAL_LINKAGES for r in moved)} manual), "
                f"Problems {len(rows)} rows in {len({r['group'] for r in rows})} groups, Decisions {len(decisions)} rows")
    final = doc[COLUMNS].copy()
    for col in ["assetid", "company", "security", "Link", "stamp", "note"]:
        final[col] = final[col].astype(str).str.strip()
    for col in ["companyid", "securityid", "tradingitemid"]:
        final[col] = final[col].astype(str).str.strip().astype("int64")
    final["Date"] = pd.to_datetime(final["Date"]).dt.strftime("%Y-%m-%d")
    final.to_csv(FINAL_FILE, index=False)
    logger.info(f"{FINAL_FILE}: {len(final)} rows, columns {', '.join(COLUMNS)}")
    logger.info("=== Manual Linkages Completed ===")
