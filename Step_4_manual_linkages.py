"""
Step 4 - Add the manually reviewed linkages to the result of Step 3.

  - output/Historical_DocLinkage_v2.xlsx : copy of output/Historical_DocLinkage_new.xlsx with
        DocLinkage : MANUAL_LINKAGES appended at the end (rows whose Link is already in the file are skipped),
        Problems   : groups of linkage_problems.csv not solved by a manual linkage, one row per trading item,
                     note filled from problem_notes.csv (old row: note of its linkage, new row: note of the
                     linkage to it),
                     plus CIRCULAR_PROBLEMS (circular linkages left aside, with their reason),
        Decisions  : one row per forced decision (manual linkage added, new security without linkage,
                     circular linkage not validated or left aside).
  - output/manual_linkages_<timestamp>.log

Manual linkages:
  - circular linkages: at the AUC review 2008-12-19 the main line was replaced by another line (Old / Stapled shares)
    and comes back later (return linkage already in DocLinkage); only kept when both lines have the same price,
  - groups with 1 old / 2 new trading items: "Correct linkage" row kept, "New security" row has no linkage.
The source Historical_DocLinkage_new.xlsx is not modified.
"""
import logging
import os
from copy import copy
from datetime import datetime

import openpyxl
import pandas as pd

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
SOURCE_FILE = os.path.join(OUTPUT_DIR, "Historical_DocLinkage_new.xlsx")
TARGET_FILE = os.path.join(OUTPUT_DIR, "Historical_DocLinkage_v2.xlsx")
NOTES_FILE = os.path.join(BASE_DIR, "problem_notes.csv")
HEADER_ROW = 2
COLUMNS = ["assetid", "companyid", "securityid", "tradingitemid", "company", "security",
           "Link", "share", "price", "Date", "stamp", "note"]
EXTRA = ["group", "role", "first_price_date", "last_price_date", "reason"]
CIRCULAR_DATE = datetime(2008, 11, 3)

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
    for link in set(REMOVED_LINKAGES) - {v["Link"] for _, v, _ in removed}:
        logger.warning(f"Link {link} to remove not found in DocLinkage")
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


def problem_rows(logger):
    """Groups of linkage_problems.csv not solved by a manual linkage, one row per trading item."""
    groups = pd.read_csv(os.path.join(OUTPUT_DIR, "linkage_groups.csv"),
                         parse_dates=["first_price_date", "last_price_date"])
    problems = pd.read_csv(os.path.join(OUTPUT_DIR, "linkage_problems.csv"))
    rows = []
    for _, p in problems.iterrows():
        g = p["group"]
        if g in MANUAL_LINKAGES:
            logger.info(f"group {g} {p['company']}: solved by a manual linkage, not in Problems")
            continue
        if g in NO_LINKAGE:
            logger.info(f"group {g} {p['company']}: no linkage decided, not in Problems ({NO_LINKAGE[g][2]})")
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
    return rows


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


def decision_rows(removed):
    rows = []
    for group, v, comment in removed:
        rows.append([group, v["company"], "linkage removed", v["assetid"], v["tradingitemid"], v["security"], v["Link"],
                     v["share"], v["price"], v["Date"], comment])
    for group, (assetid, _, _, ti, company, security, link, share, price, date, comment) in MANUAL_LINKAGES.items():
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
    return sorted((r + [today] for r in rows), key=lambda r: (r[1], r[0], r[2]))


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
    rows = apply_problem_notes(problem_rows(logger), logger)
    write_problems_sheet(wb, ws, header, rows)
    decisions = decision_rows(removed)
    write_decisions_sheet(wb, ws, header, decisions)
    for d in decisions:
        if d[2] == "new security, no linkage":
            logger.info(f"group {d[0]} {d[1]}: {d[6]} {d[5]} is a new security, no linkage")
    wb.save(TARGET_FILE)

    doc = pd.read_excel(TARGET_FILE, sheet_name="DocLinkage", header=HEADER_ROW - 1).dropna(subset=["Link"])
    duplicated = doc["Link"].duplicated().sum()
    if duplicated:
        logger.warning(f"{duplicated} duplicated Link in DocLinkage")
    logger.info(f"{TARGET_FILE}: DocLinkage {len(doc)} rows ({added} manual), "
                f"Problems {len(rows)} rows in {len({r['group'] for r in rows})} groups, Decisions {len(decisions)} rows")
    logger.info("=== Manual Linkages Completed ===")
