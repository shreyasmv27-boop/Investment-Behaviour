"""Add NAV history (for the chart) to funds.xlsx.

Run on your own computer, in the same folder as funds.xlsx (needs internet):
    pip install pandas openpyxl requests
    python fetch_nav.py

For every fund it:
  1. finds the AMFI scheme code on mfapi.in by fund name (Direct - Growth plan, or the ETF itself),
     unless you have already typed a scheme_code in that row
  2. downloads the full daily NAV history and keeps one point per week
  3. writes scheme_code + the matched AMFI name (column 'matched_name') into the funds sheet,
     and all history into a 'nav_history' sheet

Your NAV, returns, AUM and every other figure are left exactly as you typed them.
Set UPDATE_FIGURES = True only if you want NAV / 1Y / 3Y / 5Y / launch year recomputed from AMFI data.

After it runs, open funds.xlsx and check the 'matched_name' column: every row should be the right fund.
If one is wrong, type the correct scheme_code in that row (find it on mfapi.in or amfiindia.com) and run again.
"""

import re
import sys
import time

import pandas as pd
import requests

FILE = "funds.xlsx"
AS_OF = None             # e.g. "2026-11-21" to end every chart on the same date; None = latest
UPDATE_FIGURES = False
API = "https://api.mfapi.in/mf"
SKIP = ("idcw", "dividend", "bonus", "segregated", "payout", "reinvest", "regular")


def words(s):
    return set(re.findall(r"[a-z0-9]+", s.lower())) - {"fund", "the", "of", "and", "plan", "option", "mutual"}


def find_code(name, plan):
    base = re.sub(r"\(.*?\)", "", name).strip()
    queries = [base]
    inner = re.findall(r"\((.*?)\)", name)
    if inner:                                   # renamed funds: also try the old name, e.g. "ICICI Prudential Bluechip"
        queries.append(" ".join(base.split()[:2]) + " " + inner[0])
    is_etf = "etf" in str(plan).lower() and "fof" not in name.lower()
    best = None
    for q in queries:
        try:
            hits = requests.get(f"{API}/search", params={"q": q}, timeout=30).json()
        except Exception:
            continue
        for h in hits:
            n = h["schemeName"].lower()
            if any(k in n for k in SKIP):
                continue
            if not is_etf and not ("direct" in n and "growth" in n):
                continue
            score = len(words(q) & words(n)) - 0.01 * len(n)
            if best is None or score > best[0]:
                best = (score, h["schemeCode"], h["schemeName"])
        time.sleep(0.2)
    return (best[1], best[2]) if best else (None, None)


def cagr(h, years):
    end = h["date"].iloc[-1]
    start = h[h["date"] <= end - pd.DateOffset(years=years)]
    if start.empty:
        return None
    return round(((h["nav"].iloc[-1] / start["nav"].iloc[-1]) ** (1 / years) - 1) * 100, 2)


def main():
    funds = pd.read_excel(FILE, sheet_name=0)
    for c in ("scheme_code", "matched_name"):
        if c not in funds.columns:
            funds[c] = pd.NA
    funds["scheme_code"] = funds["scheme_code"].astype("object")
    funds["matched_name"] = funds["matched_name"].astype("object")
    all_hist = []
    for i, r in funds.iterrows():
        code = r["scheme_code"]
        if pd.isna(code) or str(code).strip() == "":
            code, matched = find_code(r["fund_name"], r.get("plan", ""))
            if not code:
                print(f"  NOT FOUND: {r['fund_name']} (type its scheme_code by hand)")
                continue
        else:
            code, matched = str(code).split(".")[0], None
        try:
            data = requests.get(f"{API}/{code}", timeout=30).json()
        except Exception as e:
            print(f"  {r['fund_name']}: download failed ({e})")
            continue
        if not data.get("data"):
            print(f"  {r['fund_name']}: no data for scheme code {code}")
            continue
        h = pd.DataFrame(data["data"])
        h["date"] = pd.to_datetime(h["date"], format="%d-%m-%Y")
        h["nav"] = pd.to_numeric(h["nav"], errors="coerce")
        h = h.dropna().sort_values("date")
        if AS_OF:
            h = h[h["date"] <= pd.Timestamp(AS_OF)]
        if h.empty:
            continue
        funds.loc[i, "scheme_code"] = str(code)
        funds.loc[i, "matched_name"] = matched or data["meta"].get("scheme_name", "")
        if UPDATE_FIGURES:
            funds.loc[i, "nav"] = round(h["nav"].iloc[-1], 4)
            for y in (1, 3, 5):
                funds.loc[i, f"return_{y}y"] = cagr(h, y)
            funds.loc[i, "launch_year"] = h["date"].iloc[0].year
        weekly = h.set_index("date").resample("W-FRI").last().dropna().reset_index()
        if len(weekly) < 30:                     # very new fund: keep daily points so the chart isn't empty
            weekly = h
        weekly.insert(0, "scheme_code", str(code))
        all_hist.append(weekly[["scheme_code", "date", "nav"]])
        print(f"  OK  {r['fund_name']}  ->  {funds.loc[i, 'matched_name']}  ({len(h)} NAVs, to {h['date'].iloc[-1]:%d %b %Y})")
        time.sleep(0.3)
    if not all_hist:
        sys.exit("Nothing downloaded. Check your internet connection.")
    nav = pd.concat(all_hist)
    nav["date"] = nav["date"].dt.date
    with pd.ExcelWriter(FILE) as w:
        funds.to_excel(w, sheet_name="funds", index=False)
        nav.to_excel(w, sheet_name="nav_history", index=False)
    print(f"\nDone: {len(all_hist)} of {len(funds)} funds have a chart. Now check the 'matched_name' column.")


if __name__ == "__main__":
    main()
