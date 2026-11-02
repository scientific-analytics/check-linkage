"""
Step 4b - Linkages decided by rules (Step_4_manual_linkages.py is kept unchanged).

Inputs:
  - output/linkage_groups.csv, output/prices_around_linkage.csv (Step 1),
  - input/Historical_DocLinkage.xlsx (initial linkages),
  - input/manual_decisions.csv: decisions of the manual review, one row per decision (manual linkage added,
    new security, no linkage, circular linkage not validated, linkage removed), with their decided_on date,
  - CIQ for the old prices missing in the Step 1 window (cached in output/old_prices.csv and
    output/price_ranges.csv). output/auc_timeseries/ is not used.

Rules, applied to every candidate group of Link-candidates.xlsx (Treated empty) and to every row of
Historical_DocLinkage.xlsx. Date and price of a group are computed as in Step 2 (first price of a real new listing,
else first new price after the last price common to old and new trading items).
  Problems (to check by hand):
    - several old or several new trading items (several securities alive at the same time),
    - old and new trading items with the same asset id.
  Linkage with a note (column note): old trading item still priced more than MAX_GAP_DAYS days after Date (both
  alive: circular or parallel lines), when the Decisions rules below are respected.
  Decisions (out of DocLinkage):
    - "no price": no price of the new trading item, or of the old one before Date,
    - "currency change": old and new trading items in different currencies,
    - "no linkage": more than MAX_GAP_DAYS days between the last price of the old trading item and Date,
    - "price jump": |price / last old price - 1| > MAX_PRICE_JUMP_PCT, on priceclose (adjustmentfactor not applied).
  Otherwise the linkage is kept in DocLinkage.
Manual decisions are kept and replace the rules for their groups; a manual linkage still goes through the Decisions
rules (no price, currency change, no linkage, price jump).

Output: output/Historical_DocLinkage_v4.xlsx
  - DocLinkage : initial rows, then rows of the candidate groups (by company and Date), then manual linkages,
                 with columns old_price, old_price_date, jump_pct, old_share after note,
  - Problems   : one row per trading item of the groups (or initial rows) to check by hand, with the reason,
  - Decisions  : manual decisions and rows moved out by the rules, with the reason.
"""
import glob
import logging
import os
from copy import copy
from datetime import datetime

import openpyxl
import pandas as pd
import psycopg2
from dotenv import load_dotenv

import Step_4_manual_linkages as manual
from Step_2_build_linkage import MAX_GAP_DAYS, MAX_PRICE_JUMP_PCT, STAMP, load_inputs

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
HISTORICAL_FILE = os.path.join(BASE_DIR, "input", "Historical_DocLinkage.xlsx")
TARGET_FILE = os.path.join(OUTPUT_DIR, "Historical_DocLinkage_v4.xlsx")
PRICE_RANGES_FILE = os.path.join(OUTPUT_DIR, "price_ranges.csv")
MANUAL_FILE = os.path.join(BASE_DIR, "input", "manual_decisions.csv")
DECISION_COLUMNS = ["group", "company", "decision", "assetid", "tradingitemid", "security", "Link", "share", "price",
                    "Date", "comment"]
HEADER_ROW = 2
COLUMNS = manual.COLUMNS
INFO_COLUMNS = ["old_price", "old_price_date", "jump_pct", "old_share"]
CANDIDATES = "Link-candidates"  # origin of a linkage computed from a group of input/Link-candidates.xlsx

logging.basicConfig(level=logging.INFO, format="%(levelname)s - %(message)s")
logger = logging.getLogger("step_4b")


# ---------------------------------------------------------------- data of trading items not in the candidate groups

def load_prices():
    """Daily close by trading item: Step 1 window (prices_around_linkage.csv) and CIQ cache (old_prices.csv)."""
    frames = [pd.read_csv(os.path.join(OUTPUT_DIR, "prices_around_linkage.csv"), parse_dates=["pricingdate"])]
    if os.path.exists(manual.OLD_PRICES_FILE):
        frames.append(pd.read_csv(manual.OLD_PRICES_FILE, parse_dates=["pricingdate"]))
    data = pd.concat(f[["tradingitemid", "pricingdate", "priceclose"]] for f in frames)
    data = data.drop_duplicates(["tradingitemid", "pricingdate"]).sort_values("pricingdate")
    return {ti: d.set_index("pricingdate")["priceclose"] for ti, d in data.groupby("tradingitemid")}


def auc_values(tradingitemids):
    """currency and shr of each trading item at its last review in data_gov/<date>/_universe/auc.csv."""
    frames = []
    for path in sorted(glob.glob(os.path.join(manual.DATA_GOV, "2*", "_universe", "auc.csv"))):
        auc = pd.read_csv(path, usecols=["date", "tradingitemid", "currency", "shr"])
        frames.append(auc[auc["tradingitemid"].isin(tradingitemids)])
    if not frames:
        return pd.DataFrame(columns=["currency", "shr"])
    return pd.concat(frames).sort_values("date").groupby("tradingitemid")[["currency", "shr"]].last()


def ciq_price_ranges(tradingitemids):
    """First and last CIQ price date of each trading item, cached in PRICE_RANGES_FILE."""
    cache = pd.read_csv(PRICE_RANGES_FILE, parse_dates=["first_price_date", "last_price_date"]) \
        if os.path.exists(PRICE_RANGES_FILE) else pd.DataFrame(columns=["tradingitemid", "first_price_date",
                                                                        "last_price_date"])
    missing = sorted(set(tradingitemids) - set(cache["tradingitemid"]))
    if missing:
        load_dotenv(os.path.join(BASE_DIR, ".env"))
        try:
            conn = psycopg2.connect(host=os.getenv("DB_HOST"), database=os.getenv("DB_NAME"),
                                    user=os.getenv("DB_USER"), password=os.getenv("DB_PASSWORD"),
                                    port=os.getenv("DB_PORT"), connect_timeout=15)
        except psycopg2.OperationalError as e:
            logger.warning(f"CIQ not reachable, price range of {len(missing)} trading items not downloaded: "
                           f"{str(e).splitlines()[0]}")
        else:
            with conn, conn.cursor() as cur:
                cur.execute("""SELECT tradingitemid, MIN(pricingdate), MAX(pricingdate) FROM public.ciqpriceequity
                               WHERE tradingitemid = ANY(%s) GROUP BY tradingitemid""", ([int(t) for t in missing],))
                new = pd.DataFrame(cur.fetchall(), columns=["tradingitemid", "first_price_date", "last_price_date"])
            conn.close()
            new[["first_price_date", "last_price_date"]] = new[["first_price_date", "last_price_date"]].apply(
                pd.to_datetime)
            cache = pd.concat([cache, new])
            cache.to_csv(PRICE_RANGES_FILE, index=False)
            logger.info(f"CIQ: price range of {len(new)} trading items saved in {PRICE_RANGES_FILE}")
    return cache.set_index("tradingitemid")


def load_manual():
    """Manual decisions, one row per decision (input/manual_decisions.csv)."""
    df = pd.read_csv(MANUAL_FILE, parse_dates=["Date", "decided_on"])
    for col in ["group", "companyid", "securityid", "tradingitemid"]:
        df[col] = df[col].astype("Int64")
    return df


def decision_rows(manual_df, removed, rule_decisions, moved_links):
    """Decisions sheet: rule decisions (decided today), manual decisions (decided_on of the CSV), manual
    linkages moved out by the rules replaced by their rule decision."""
    today = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
    rows = [d + [today] for d in rule_decisions]
    for g, v, comment, decided_on in removed:
        rows.append([g, v["company"], "linkage removed", v["assetid"], v["tradingitemid"], v["security"], v["Link"],
                     v["share"], v["price"], v["Date"], comment, decided_on])
    for _, r in manual_df[manual_df["decision"] != "linkage removed"].iterrows():
        if r["decision"] == "manual linkage added" and r["Link"] in moved_links:
            continue
        rows.append([None if pd.isna(r[c]) else (int(r[c]) if c in ("group", "tradingitemid") else
                     (r[c].to_pydatetime() if isinstance(r[c], pd.Timestamp) else r[c])) for c in DECISION_COLUMNS]
                    + [r["decided_on"].to_pydatetime()])
    return sorted(rows, key=lambda r: (str(r[1]), -1 if r[0] is None else r[0], r[2]))


# ---------------------------------------------------------------------------------------------- rules on one row

def check_row(values, old_ti, old_last, old_currency, new_currency, prices, check_alive=True):
    """Apply the rules to one linkage, Decisions rules first (no price, currency change, gap, price jump), then
    the Problems rule (old trading item still priced after Date). Return (destination, decision or Problems
    reason, info): destination in "linkage", "decision", "problem"."""
    date = pd.Timestamp(values["Date"])
    info = dict.fromkeys(INFO_COLUMNS)
    old = prices.get(old_ti)
    old = old[old.index < date] if old is not None else None
    if old is not None and not old.empty:
        info.update(old_price=round(float(old.iloc[-1]), 4), old_price_date=old.index[-1].to_pydatetime(),
                    jump_pct=round((values["price"] / float(old.iloc[-1]) - 1) * 100, 2))
    if old is None or old.empty:
        return "decision", ("no price", f"no price of the old trading item {old_ti} before {date:%Y-%m-%d}"), info
    if pd.notna(old_currency) and pd.notna(new_currency) and old_currency != new_currency:
        return "decision", ("currency change", f"currency change {old_currency} -> {new_currency}: old "
                                               f"{info['old_price']:g} on {old.index[-1]:%Y-%m-%d}, new "
                                               f"{values['price']:g} on {date:%Y-%m-%d}"), info
    gap = (date - old.index[-1]).days
    if gap > MAX_GAP_DAYS:
        return "decision", ("no linkage", f"gap of {gap} days between old last price {info['old_price']:g} on "
                                          f"{old.index[-1]:%Y-%m-%d} and new price {values['price']:g} on "
                                          f"{date:%Y-%m-%d}"), info
    if abs(info["jump_pct"]) > MAX_PRICE_JUMP_PCT:
        return "decision", ("price jump", f"price jump of {info['jump_pct']:+.1f}%: old {info['old_price']:g} on "
                                          f"{old.index[-1]:%Y-%m-%d}, new {values['price']:g} on "
                                          f"{date:%Y-%m-%d}"), info
    if check_alive and pd.notna(old_last) and (old_last - date).days > MAX_GAP_DAYS:
        return "linkage", (f"old and new trading items both alive: old trading item {old_ti} priced until "
                           f"{old_last:%Y-%m-%d}, new trading item from {date:%Y-%m-%d}"), info
    return "linkage", None, info


def switch(old, new, prices):
    """Date and price of the new trading item, as in Step 2. Return (date, price) or None."""
    new_prices = prices[prices["tradingitemid"] == new["tradingitemid"]].sort_values("pricingdate")
    if new_prices.empty:
        return None
    if new["first_price_date"] > old["date"] - pd.Timedelta(days=30):
        first = new_prices[new_prices["pricingdate"] >= new["first_price_date"]]
    else:
        old_prices = prices[prices["tradingitemid"] == old["tradingitemid"]]
        both = new_prices.merge(old_prices, on="pricingdate", suffixes=("", "_old"))
        equal = both[(both["priceclose"] - both["priceclose_old"]).abs() < 1e-6]
        reference = equal["pricingdate"].max() if not equal.empty else old["last_price_date"]
        first = new_prices[new_prices["pricingdate"] > reference] if pd.notna(reference) else new_prices
    if first.empty:
        return None
    return first["pricingdate"].iloc[0], float(first["priceclose"].iloc[0])


def problem_rows_of_group(group, reason):
    rows = []
    olds = group.loc[group["role"] == "old", "asset"].tolist()
    for _, t in group.sort_values(["role", "row"], ascending=[False, True]).iterrows():
        is_old = t["role"] == "old"
        rows.append({"assetid": t["asset"] if is_old else (olds[0] if len(olds) == 1 else None),
                     "companyid": t["companyid"], "securityid": t["securityid"], "tradingitemid": t["tradingitemid"],
                     "company": t["company"], "security": t["security"], "Link": None if is_old else t["asset"],
                     "share": round(t["shr"], 3), "price": None, "Date": None, "stamp": STAMP, "note": "-",
                     "group": t["group"], "role": t["role"], "first_price_date": t["first_price_date"],
                     "last_price_date": t["last_price_date"], "reason": reason})
    return rows


def resolve_root(rows, known):
    """assetid = root asset of the chain (Link -> assetid of the linkages already known), loops guarded."""
    root = dict(known)
    for row in sorted(rows, key=lambda r: r["Date"]):
        a, seen = row["old_asset"], set()
        while a in root and a not in seen:
            seen.add(a)
            a = root[a]
        row["assetid"] = a
        root[row["Link"]] = a


# ------------------------------------------------------------------------------------------------------------ main

def main():
    logger.info("=== Step 4b: linkages by rules ===")
    groups, window_prices = load_inputs()
    groups = groups[groups["Treated"].isna()]
    initial = pd.read_excel(HISTORICAL_FILE, header=HEADER_ROW - 1).iloc[:, 1:].dropna(subset=["Link"])
    for col in ["assetid", "company", "security", "Link", "stamp", "note"]:
        initial[col] = initial[col].astype(str).str.strip()
    for col in ["companyid", "securityid", "tradingitemid"]:
        initial[col] = initial[col].astype(str).str.strip().astype("int64")
    auc_assets = manual.load_auc_assets()
    manual_df = load_manual()
    manual_groups = set(manual_df.loc[manual_df["decision"] != "linkage removed", "group"].dropna().astype(int))
    removed_groups = {int(r["group"]): (r["Link"], r["comment"], r["decided_on"].to_pydatetime())
                      for _, r in manual_df[manual_df["decision"] == "linkage removed"].iterrows()}

    # 1. candidate rows: (origin, group, values, old_ti, old_last, old_currency, new_currency)
    candidates, problems = [], []
    for g, group in groups.groupby("group"):
        if g in manual_groups:
            continue
        old, new = group[group["role"] == "old"], group[group["role"] == "new"]
        if len(old) != 1 or len(new) != 1:
            problems += problem_rows_of_group(group, f"several securities alive at the same time: {len(old)} old / "
                                                     f"{len(new)} new trading items")
            continue
        old, new = old.iloc[0], new.iloc[0]
        if old["asset"] == new["asset"]:
            problems += problem_rows_of_group(group, "old and new trading items share the same asset id")
            continue
        found = switch(old, new, window_prices)
        if found is None:
            candidates.append((CANDIDATES, g, {"company": new["company"], "assetid": old["asset"], "old_asset": old["asset"],
                                            "companyid": new["companyid"], "securityid": new["securityid"],
                                            "tradingitemid": new["tradingitemid"], "security": new["security"],
                                            "Link": new["asset"], "share": round(new["shr"], 3), "price": None,
                                            "Date": None, "stamp": STAMP, "note": "-"},
                               None, None, None, None))
            continue
        date, price = found
        values = {"old_asset": old["asset"], "assetid": old["asset"], "companyid": new["companyid"],
                  "securityid": new["securityid"], "tradingitemid": new["tradingitemid"], "company": new["company"],
                  "security": new["security"], "Link": new["asset"], "share": round(new["shr"], 3),
                  "price": round(price, 4), "Date": date.to_pydatetime(), "stamp": STAMP, "note": "-"}
        candidates.append((CANDIDATES, g, values, int(old["tradingitemid"]), old["last_price_date"], old["currency"],
                           new["currency"]))

    # initial rows: old trading item of the chain (previous Link of the same assetid), else of assetid
    chain, previous = {}, {}
    for _, r in initial.sort_values(["assetid", "Date"]).iterrows():
        chain[(r["tradingitemid"], r["Link"])] = previous.get(r["assetid"], r["assetid"])
        previous[r["assetid"]] = r["Link"]
    for _, r in initial.iterrows():
        values = {c: r[c] for c in COLUMNS}
        values["Date"] = pd.Timestamp(r["Date"]).to_pydatetime()
        old_ti = manual.old_trading_item(r["tradingitemid"], r["assetid"], r["Link"], groups, {}, chain, auc_assets,
                                         pd.Timestamp(r["Date"]))
        candidates.append(("initial file", None, values, old_ti, None, None, None))

    # manual linkages
    for _, r in manual_df[manual_df["decision"] == "manual linkage added"].iterrows():
        g = int(r["group"])
        values = {c: r[c] for c in COLUMNS if c in r.index}
        values.update(stamp=STAMP, note="-", Date=r["Date"].to_pydatetime(),
                      **{c: int(r[c]) for c in ("companyid", "securityid", "tradingitemid")})
        old_ti = manual.old_trading_item(values["tradingitemid"], values["assetid"], values["Link"], groups, {}, {},
                                         auc_assets, pd.Timestamp(values["Date"]))
        candidates.append(("manual linkage", g, values, old_ti, None, None, None))

    # 2. data of the old and new trading items not in the candidate groups (auc.csv, CIQ)
    in_groups = groups.drop_duplicates("tradingitemid").set_index("tradingitemid")
    tis = {c[3] for c in candidates if c[3] is not None} | {int(c[2]["tradingitemid"]) for c in candidates}
    auc = auc_values(tis - set(in_groups.index))
    ranges = ciq_price_ranges({c[3] for c in candidates if c[3] is not None and c[4] is None
                               and c[3] not in in_groups.index})

    def currency(ti):
        if ti in in_groups.index:
            return in_groups.at[ti, "currency"]
        return auc.at[ti, "currency"] if ti in auc.index else None

    def shares(ti):
        if ti in in_groups.index:
            return in_groups.at[ti, "shr"]
        return auc.at[ti, "shr"] if ti in auc.index else None

    def last_price_date(ti):
        if ti in in_groups.index:
            return in_groups.at[ti, "last_price_date"]
        return ranges.at[ti, "last_price_date"] if ti in ranges.index else pd.NaT

    # 3. old prices: Step 1 window and CIQ cache, else last CIQ close before Date
    prices = load_prices()
    done = set()
    if os.path.exists(manual.OLD_PRICES_FILE):
        cache = pd.read_csv(manual.OLD_PRICES_FILE, parse_dates=["linkage_date"])
        done = set(zip(cache["tradingitemid"], cache["linkage_date"]))
    requests = []
    for _, _, values, old_ti, *_ in candidates:
        if old_ti is None or values["Date"] is None or (old_ti, pd.Timestamp(values["Date"])) in done:
            continue
        old = prices.get(old_ti)
        old = old[old.index < values["Date"]] if old is not None else None
        if old is None or old.empty or (pd.Timestamp(values["Date"]) - old.index[-1]).days > MAX_GAP_DAYS:
            requests.append((old_ti, pd.Timestamp(values["Date"])))
    if requests and manual.download_old_prices(sorted(set(requests)), logger):
        prices = load_prices()

    # old trading items of the groups in Problems: last close in the window and old_share
    for p in problems:
        if p["role"] != "old":
            continue
        ti = int(p["tradingitemid"])
        old = prices.get(ti)
        old = old[old.index <= p["last_price_date"]] if old is not None and pd.notna(p["last_price_date"]) else old
        if old is not None and not old.empty:
            p["old_price"], p["old_price_date"] = round(float(old.iloc[-1]), 4), old.index[-1].to_pydatetime()
        p["old_share"] = p["share"]

    # 4. rules
    kept, decisions, removed = [], [], []
    for origin, g, values, old_ti, old_last, old_cur, new_cur in candidates:
        label = f"{origin} {'' if g is None else f'group {g} '}{values['company']} {values['assetid']} -> " \
                f"{values['Link']}"
        if values["Date"] is None:
            destination, result, info = "decision", ("no price", "no price of the new trading item in the window"), \
                dict.fromkeys(INFO_COLUMNS)
        else:
            new_ti = int(values["tradingitemid"])
            destination, result, info = check_row(
                values, old_ti, old_last if old_last is not None else last_price_date(old_ti),
                old_cur if old_cur is not None else currency(old_ti),
                new_cur if new_cur is not None else currency(new_ti), prices,
                check_alive=origin != "manual linkage")
        info["old_share"] = None if old_ti is None or pd.isna(shares(old_ti)) else round(float(shares(old_ti)), 3)
        values.update(info)
        if g in removed_groups:
            link, comment, decided_on = removed_groups[g]
            reason = result[1] if isinstance(result, tuple) else result
            removed.append((g, values, comment + ("" if reason is None else f"; {reason}"), decided_on))
            logger.info(f"{label}: linkage removed by manual decision ({comment})")
            continue
        if destination == "linkage":
            if result:
                values["note"] = result
                logger.info(f"{label}: linkage with note ({result})")
            kept.append((origin, values))
        elif destination == "problem":
            problems.append({**{c: values.get(c) for c in COLUMNS + INFO_COLUMNS}, "group": g, "role": origin,
                             "first_price_date": None, "last_price_date": None, "reason": result})
            logger.info(f"{label}: Problems ({result})")
        else:
            decision, reason = result
            decisions.append([g, str(values["company"]).strip(), decision, values["assetid"], values["tradingitemid"],
                              values["security"], values["Link"], values["share"], values["price"], values["Date"],
                              f"{reason} ({origin})"])
            logger.info(f"{label}: {decision} ({reason})")

    # 5. DocLinkage rows: initial, groups (root assetid, no Link already in the initial file), manual
    initial_rows = [v for o, v in kept if o == "initial file"]
    group_rows = [v for o, v in kept if o == CANDIDATES]
    manual_rows = [v for o, v in kept if o == "manual linkage"]
    resolve_root(group_rows, dict(zip(initial["Link"], initial["assetid"])))
    initial_links = set(initial["Link"])
    skipped = [v for v in group_rows if v["Link"] in initial_links]
    for v in skipped:
        logger.info(f"group linkage {v['assetid']} -> {v['Link']}: Link already in the initial file, skipped")
    group_rows = sorted((v for v in group_rows if v["Link"] not in initial_links),
                        key=lambda v: (str(v["company"]).strip(), v["Date"]))
    doc_rows = initial_rows + group_rows + manual_rows

    moved_links = {d[6] for d in decisions if d[-1].endswith("(manual linkage)")}
    decision_list = decision_rows(manual_df, removed, decisions, moved_links)
    counts = {"doc": len(doc_rows), "initial": len(initial_rows), "groups": len(group_rows), "manual": len(manual_rows),
              "problems": len({(p["group"], p["Link"]) if p["group"] is None else p["group"] for p in problems}),
              "decisions": pd.Series([d[2] for d in decision_list]).value_counts().to_dict()}
    write_workbook(doc_rows, problems, decision_list, counts)

    links = pd.Series([v["Link"] for v in doc_rows])
    if links.duplicated().any():
        logger.warning(f"{links.duplicated().sum()} duplicated Link in DocLinkage")
    logger.info(f"{TARGET_FILE}: DocLinkage {len(doc_rows)} rows ({len(initial_rows)} initial file, "
                f"{len(group_rows)} groups, {len(manual_rows)} manual), Problems {len(problems)} rows, "
                f"Decisions {len(decision_list)} rows")
    for decision, n in pd.Series([d[2] for d in decision_list]).value_counts().items():
        logger.info(f"  Decisions {decision}: {n}")


def write_note_sheet(wb, counts):
    """Sheet Méthodo: short summary of the method, thresholds and counts of this run."""
    d = counts["decisions"]
    lines = [
        ("Méthodologie du linkage", "title"),
        (f"Fichier généré par Step_4b_rules.py le {datetime.now():%Y-%m-%d}.", None),
        ("", None),
        ("Objectif", "head"),
        ("Quand une société change de ligne de cotation (fusion, changement de bourse, nouvelle classe d'actions...), "
         "relier l'ancien asset au nouveau titre pour garder un historique continu.", None),
        ("Une ligne de DocLinkage : assetid = ancien asset (racine de la chaîne), Link = nouvel asset ; companyid, "
         "securityid, tradingitemid, security, share et price sont ceux du nouveau titre à la Date.", None),
        ("", None),
        ("Sources", "head"),
        ("Link-candidates.xlsx : groupes de candidats repérés dans AUC (titres de la dernière review = nouveaux, "
         "autres = anciens).", None),
        ("Historical_DocLinkage.xlsx : linkages déjà existantes (fichier initial).", None),
        ("Prix CIQ : fenêtre de ±15 jours autour de la bascule (Step 1), complétée par CIQ si besoin.", None),
        ("", None),
        ("Date et prix", "head"),
        ("Date = premier jour où le nouveau titre cote pour lui-même : premier prix d'une nouvelle cotation, ou jour "
         "qui suit le dernier prix commun quand CIQ a recopié l'historique de l'ancien titre.", None),
        ("price = clôture du nouveau titre à la Date ; share = nombre d'actions du nouveau titre (AUC).", None),
        ("", None),
        ("Règles (groupes, fichier initial et linkages manuelles)", "head"),
        (f"Linkage gardée : 1 ancien titre s'arrête, 1 nouveau reprend, même devise, interruption <= {MAX_GAP_DAYS} "
         f"jours, saut de prix <= {MAX_PRICE_JUMP_PCT} %.", None),
        (f"Decisions « price jump » : saut > {MAX_PRICE_JUMP_PCT} % entre le dernier prix de l'ancien titre et le "
         f"prix de linkage (priceclose, adjustmentfactor non appliqué).", None),
        ("Decisions « currency change » : devise différente entre l'ancien et le nouveau titre.", None),
        (f"Decisions « no linkage » : plus de {MAX_GAP_DAYS} jours entre le dernier prix de l'ancien titre et la "
         f"Date.", None),
        ("Decisions « no price » : pas de prix de l'ancien ou du nouveau titre.", None),
        ("Problems (à vérifier à la main) : plusieurs titres vivants en même temps (plusieurs anciens ou plusieurs "
         "nouveaux titres), même asset id.", None),
        (f"Linkage avec note (colonne note) : l'ancien titre cote encore plus de {MAX_GAP_DAYS} jours après la Date "
         f"(circulaire, cotation parallèle) mais les règles de saut, devise, interruption et prix sont respectées.",
         None),
        ("Ordre : les règles Decisions passent en premier (un gros saut va toujours dans Decisions).", None),
        ("Décisions manuelles (circulaires, séparations de classes, no linkage) reprises de la revue, dans "
         "input/manual_decisions.csv ; les "
         "linkages manuelles passent aussi les règles de saut, devise, interruption et prix.", None),
        ("Colonnes d'information : old_price, old_price_date, jump_pct, old_share (ancien titre) ; old_share "
         "n'est pas un critère.", None),
        ("", None),
        ("Origine des linkages", "head"),
        ("« initial file » : la linkage existait déjà, complète, dans Historical_DocLinkage.xlsx. Date, price et "
         "share sont repris du fichier ; le script vérifie seulement les règles.", None),
        ("« Link-candidates » : la linkage n'existait pas. Link-candidates.xlsx donne seulement une paire ancien / "
         "nouveau titre ; le script calcule Date, price et share à partir des prix CIQ, puis vérifie les règles.",
         None),
        ("« manual linkage » : groupe de Link-candidates.xlsx dont Date, price et share ont été fixés à la main "
         "pendant la revue (circulaires, séparations de classes) ; le script vérifie les règles.", None),
        ("", None),
        ("Lecture des onglets", "head"),
        ("Problems, colonne source : « old » / « new » = rôle du trading item dans son groupe de Link-candidates.xlsx "
         "(groupe à plusieurs titres, une ligne par titre) ; sinon origine de la linkage : « Link-candidates » = "
         "groupe de candidats de Link-candidates.xlsx (date, prix et share calculés par le script), « initial file » = ligne de Historical_DocLinkage.xlsx, "
         "« manual linkage » = linkage manuelle.", None),
        ("Problems, colonne reason : règle qui a envoyé le cas dans Problems.", None),
        ("Decisions, colonne comment : raison chiffrée de la décision ; l'origine (Link-candidates, initial file, "
         "manual linkage) est indiquée entre parenthèses à la fin.", None),
        ("", None),
        ("Résultat", "head"),
        (f"DocLinkage : {counts['doc']} linkages ({counts['initial']} du fichier initial, {counts['groups']} des "
         f"groupes, {counts['manual']} manuelles).", None),
        (f"Problems : {counts['problems']} cas à vérifier.", None),
        ("Decisions : " + ", ".join(f"{k} {v}" for k, v in d.items()) + ".", None),
    ]
    ws = wb.create_sheet("Méthodo")
    ws.column_dimensions["A"].width = 140
    for i, (text, kind) in enumerate(lines, start=1):
        cell = ws.cell(i, 1, text)
        cell.alignment = openpyxl.styles.Alignment(wrap_text=True, vertical="top")
        if kind == "title":
            cell.font = openpyxl.styles.Font(bold=True, size=14)
        elif kind == "head":
            cell.font = openpyxl.styles.Font(bold=True)


def write_workbook(doc_rows, problems, decisions, counts):
    wb = openpyxl.load_workbook(HISTORICAL_FILE)
    ws = wb.active
    header = {ws.cell(HEADER_ROW, c).value: c for c in range(1, ws.max_column + 1) if ws.cell(HEADER_ROW, c).value}
    template = HEADER_ROW + 1
    styles = {col: (copy(ws.cell(template, header[col]).font), ws.cell(template, header[col]).number_format)
              for col in COLUMNS}
    ws.delete_rows(template, ws.max_row - HEADER_ROW)

    note_header = ws.cell(HEADER_ROW, header["note"])
    for i, name in enumerate(INFO_COLUMNS, start=1):
        c = header["note"] + i
        header[name] = c
        cell = ws.cell(HEADER_ROW, c, name)
        cell.font, cell.fill = copy(note_header.font), copy(note_header.fill)
        cell.alignment, cell.border = copy(note_header.alignment), copy(note_header.border)
        width_of = "Date" if name == "old_price_date" else ("share" if name == "old_share" else "price")
        ws.column_dimensions[openpyxl.utils.get_column_letter(c)].width = \
            ws.column_dimensions[openpyxl.utils.get_column_letter(header[width_of])].width
    styles.update({"old_price": styles["price"], "old_price_date": styles["Date"], "old_share": styles["share"],
                   "jump_pct": (styles["price"][0], "0.00")})

    for i, values in enumerate(doc_rows, start=template):
        for col in COLUMNS + INFO_COLUMNS:
            value = values.get(col)
            if isinstance(value, pd.Timestamp):
                value = value.to_pydatetime()
            cell = ws.cell(i, header[col], None if value is not None and pd.isna(value) else value)
            cell.font, cell.number_format = copy(styles[col][0]), styles[col][1]

    manual.write_problems_sheet(wb, ws, {k: header[k] for k in COLUMNS}, problems)
    pb = wb["Problems"]
    for c in range(1, pb.max_column + 1):
        if pb.cell(HEADER_ROW, c).value == "role":
            pb.cell(HEADER_ROW, c).value = "source"
    for i, row in enumerate(problems, start=template):
        for col in INFO_COLUMNS:
            value = row.get(col)
            if value is None or (not isinstance(value, datetime) and pd.isna(value)):
                continue
            cell = pb.cell(i, header[col], value.to_pydatetime() if isinstance(value, pd.Timestamp) else value)
            cell.font, cell.number_format = copy(styles[col][0]), styles[col][1]
    manual.write_decisions_sheet(wb, ws, {k: header[k] for k in COLUMNS}, decisions)
    write_note_sheet(wb, counts)
    wb.save(TARGET_FILE)


if __name__ == "__main__":
    main()
