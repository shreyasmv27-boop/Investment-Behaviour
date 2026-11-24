"""Fill funds.xlsx with real NAV history and CAGR returns from AMFI scheme codes.

Run on your own computer (needs internet):
    pip install pandas openpyxl requests
    python fetch_nav.py

What it does, for every row in the 'funds' sheet that has a scheme_code:
  - downloads the full daily NAV history from mfapi.in (free, sourced from AMFI)
  - sets nav (latest), return_1y / return_3y / return_5y (point-to-point CAGR) and launch_year (first NAV date)
  - writes all history to a 'nav_history' sheet (weekly points, to keep the file small)
Everything else in the sheet (AUM, expense ratio, exit load, risk, fund manager, minimums) is left as you typed it.
All figures are frozen as of AS_OF, so every respondent sees the same numbers.
"""

import sys
import time

import pandas as pd
import requests

FILE = "funds.xlsx"
AS_OF = None          # e.g. "2026-11-21" to freeze data at a date; None = latest available
WEEKLY = True         # keep one NAV per week in nav_history (set False for daily)


def cagr(hist, years):
    end_date = hist["date"].iloc[-1]
    start = hist[hist["date"] <= end_date - pd.DateOffset(years=years)]
    if start.empty:
        return None
    v0, v1 = start["nav"].iloc[-1], hist["nav"].iloc[-1]
    return round(((v1 / v0) ** (1 / years) - 1) * 100, 2)


def main():
    funds = pd.read_excel(FILE, sheet_name=0)
    if "scheme_code" not in funds.columns:
        sys.exit("Add a 'scheme_code' column to the funds sheet first.")
    all_hist = []
    for i, r in funds.iterrows():
        code = r["scheme_code"]
        if pd.isna(code):
            continue
        code = str(int(code))
        try:
            data = requests.get(f"https://api.mfapi.in/mf/{code}", timeout=30).json()
        except Exception as e:
            print(f"  {code}: download failed ({e})")
            continue
        if not data.get("data"):
            print(f"  {code}: no data, check the scheme code")
            continue
        h = pd.DataFrame(data["data"])
        h["date"] = pd.to_datetime(h["date"], format="%d-%m-%Y")
        h["nav"] = pd.to_numeric(h["nav"], errors="coerce")
        h = h.dropna().sort_values("date")
        if AS_OF:
            h = h[h["date"] <= pd.Timestamp(AS_OF)]
        if h.empty:
            continue
        funds.loc[i, "nav"] = round(h["nav"].iloc[-1], 4)
        for y in (1, 3, 5):
            funds.loc[i, f"return_{y}y"] = cagr(h, y)
        funds.loc[i, "launch_year"] = h["date"].iloc[0].year
        keep = h.set_index("date").resample("W-FRI").last().dropna().reset_index() if WEEKLY else h
        keep.insert(0, "scheme_code", code)
        all_hist.append(keep[["scheme_code", "date", "nav"]])
        print(f"  {code}: {data['meta'].get('scheme_name', '')} — {len(h)} NAVs, latest {h['date'].iloc[-1]:%d %b %Y}")
        time.sleep(0.3)
    if not all_hist:
        sys.exit("Nothing downloaded.")
    nav = pd.concat(all_hist)
    nav["date"] = nav["date"].dt.date
    with pd.ExcelWriter(FILE) as w:
        funds.to_excel(w, sheet_name="funds", index=False)
        nav.to_excel(w, sheet_name="nav_history", index=False)
    print(f"Done: {len(all_hist)} funds updated, {len(nav)} NAV rows written to {FILE}.")


if __name__ == "__main__":
    main()
