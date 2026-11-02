---
name: audit-code
description: Audit the linkage pipeline (Step_1 to Step_4 scripts) and its outputs (output/Historical_DocLinkage_v2.xlsx, output/v1.csv). Use when the user asks to audit, check or review the code, or to verify the final linkage file before delivery.
---

# Audit of the linkage pipeline

Two parts: run the output checks, then review the code. Report findings in the user's language, most severe first, with `file:line` and a concrete failing case. Do not modify code or outputs during the audit; propose fixes and wait for the user.

## 1. Output checks

Run from the project root:

```bash
.venv/bin/python .claude/skills/audit-code/check_outputs.py
```

The script checks the invariants below and prints `OK` or `FAIL` per check, with the offending rows. Every `FAIL` is a finding. Investigate it in the code (part 2) before reporting.

| Check | Rule |
|---|---|
| v1.csv format | exactly the 12 columns of `Historical_DocLinkage.xlsx` (assetid ... note), no tab or leading/trailing space, ids are integers, Date is `YYYY-MM-DD` |
| v1.csv = DocLinkage | same rows as the DocLinkage sheet of v2 |
| unique Link | no Link twice in DocLinkage |
| price jump | every DocLinkage row has `abs(jump_pct) <= MAX_PRICE_JUMP_PCT` (Step 2 constant) |
| price gap | every DocLinkage row has `Date - old_price_date <= MAX_GAP_DAYS` days |
| old price | every DocLinkage row has old_price, or the comment for rows of the initial file |
| Decisions vs DocLinkage | no Link of a "price jump", "no linkage", "linkage removed" or "circular linkage not validated" decision in DocLinkage; every "manual linkage added" Link in DocLinkage |
| chains | per assetid, Links sorted by Date form one chain; the new trading item of each row is priced at Date |

## 2. Code review

Read `Step_1_download_prices_around_linkage.py`, `Step_2_build_linkage.py`, `Step_3_fill_historical.py`, `Step_4_manual_linkages.py`. Check each point:

- **Row convention.** In DocLinkage, assetid = old (root) asset, Link = new asset, and companyid, securityid, tradingitemid, security, share and price are those of the **new** trading item. Check every hard-coded row (`MANUAL_LINKAGES`, `NOT_VALIDATED`, `CIRCULAR_PROBLEMS`) against `output/linkage_groups.csv` (role new). Known past bug: SunPower 287 and WBD 341 carried the old trading item and `share = 1`.
- **Rules.** Thresholds are defined once (`MAX_PRICE_JUMP_PCT`, `MAX_GAP_DAYS`, `MAX_GAP_BDAYS` in Step 2) and imported, never duplicated. The same rules apply to Step 2 rows, manual rows and rows of the initial file.
- **Old trading item.** The old trading item of a chain link is the previous Link of the chain, not the root asset. Check `old_trading_item` in Step 4 and `resolve_root_assets` in Step 2.
- **Infinite loop.** `resolve_root_assets` (`while a in root`) loops forever on a circular chain A -> B -> A if v2 is reused as `Historical_DocLinkage.xlsx`. Flag it while there is no guard.
- **Price windows.** Step 1 downloads only +-15 days around the estimated date. Any check that reads prices outside this window must fall back to CIQ or `output/auc_timeseries/`, or mark the row as not checked. It must never silently skip it.
- **Excel writes.** `ws.delete_rows` loops go bottom-up; the styles of the copied template rows are kept; dates are written as `datetime`, not as strings.
- **Data quality of inputs.** Tabs and spaces in cells of `Historical_DocLinkage.xlsx`; rows of the initial file with swapped Date / price / share (known case: ARRIS 5962:MIW and 2884:GAP).
- **Secrets and paths.** DB credentials come only from `.env` (in `.gitignore`); absolute paths (`DATA_GOV`) are constants at the top of the file.
- **Reproducibility.** Running Step 2, Step 3 then Step 4 twice gives the same v2 and v1.csv (except `decided_on` and the log name).

## Report

For each finding: severity (bug / data / robustness / cleanup), `file:line`, one-sentence description, failing case (row, group or Link), proposed fix. End with the output-check summary (`n OK / n FAIL`).
