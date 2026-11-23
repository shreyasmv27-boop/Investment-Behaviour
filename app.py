"""Mutual fund allocation simulator, built for the Capitalmind 'Crossing the Chasm' study.

Flow: start page (name + app used) -> invest page (allocate Rs 1,00,000) -> done page (link to Google Form).
Fund data is read from funds.xlsx. Responses go to Google Sheets if configured, otherwise to local files.
"""

import html
import json
import random
import uuid
from datetime import datetime, timezone
from pathlib import Path

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Invest", page_icon="₹", layout="centered", initial_sidebar_state="collapsed")

# ---------------------------------------------------------------- settings
TOTAL = 100_000              # investable amount per respondent
STEP = 500                   # amount input step
SHUFFLE_ORDER = True         # randomise fund order per respondent (logged, so position effects can be measured)
DATA_FILE = Path(__file__).parent / "funds.xlsx"
LOCAL_DIR = Path(__file__).parent / "responses"

# Accent colours are approximate and only set the feel. No logos or wordmarks are used.
APPS = {
    "Groww":        dict(accent="#00B386", soft="#E8F8F3", ink="#44475B", layout="cards"),
    "Zerodha Coin": dict(accent="#387ED1", soft="#EEF4FC", ink="#424242", layout="rows"),
    "Angel One":    dict(accent="#1F4FD8", soft="#EDF2FD", ink="#1B2340", layout="cards"),
    "PhonePe":      dict(accent="#5F259F", soft="#F3EDFA", ink="#2A1640", layout="cards"),
    "Paytm Money":  dict(accent="#00A3E0", soft="#E9F7FD", ink="#0B2A5B", layout="rows"),
    "Other / I don't invest yet": dict(accent="#111111", soft="#F2F2F2", ink="#222222", layout="cards"),
}

RISK_LEVELS = ["Low", "Low to Moderate", "Moderate", "Moderately High", "High", "Very High"]
RISK_COLOURS = ["#1E9E5A", "#6CB33F", "#C9A227", "#E58A1F", "#E05A2B", "#C62828"]

REQUIRED = ["fund_name", "amc", "category", "sub_category", "nav", "return_1y", "return_3y", "return_5y",
            "risk", "min_lumpsum", "exit_load", "expense_ratio", "fund_manager"]
OPTIONAL = ["min_sip", "aum_cr", "launch_year"]


def secret(key, default=None):
    try:
        return st.secrets[key]
    except Exception:
        return default


# ---------------------------------------------------------------- helpers
def inr(n):
    """Indian digit grouping: 100000 -> 1,00,000."""
    n = int(round(n))
    s, sign = str(abs(n)), "-" if n < 0 else ""
    if len(s) <= 3:
        return sign + s
    head, tail = s[:-3], s[-3:]
    parts = []
    while len(head) > 2:
        parts.insert(0, head[-2:])
        head = head[:-2]
    if head:
        parts.insert(0, head)
    return sign + ",".join(parts) + "," + tail


def pct(v):
    return "–" if pd.isna(v) else f"{v:.1f}%"


def er(v):
    return "–" if pd.isna(v) else f"{v:.2f}%"


def now():
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def esc(v):
    return html.escape("" if pd.isna(v) else str(v))


@st.cache_data
def load_funds(path: str):
    df = pd.read_excel(path)
    df.columns = [str(c).strip().lower().replace(" ", "_") for c in df.columns]
    missing = [c for c in REQUIRED if c not in df.columns]
    if missing:
        return None, f"funds.xlsx is missing these columns: {', '.join(missing)}"
    for c in OPTIONAL:
        if c not in df.columns:
            df[c] = pd.NA
    df = df.dropna(subset=["fund_name"]).reset_index(drop=True)
    num = ["nav", "return_1y", "return_3y", "return_5y", "min_lumpsum", "min_sip", "expense_ratio", "aum_cr"]
    for c in num:
        df[c] = pd.to_numeric(df[c], errors="coerce")
    # Accept values typed as fractions (0.125) or percents (12.5).
    # Returns: all values within +-1 means fractions. Expense ratio: below 0.05 means a fraction (0.0092 = 0.92%).
    for c in ["return_1y", "return_3y", "return_5y"]:
        if df[c].abs().max() <= 1.0:
            df[c] = df[c] * 100
    if df["expense_ratio"].max() < 0.05:
        df["expense_ratio"] = df["expense_ratio"] * 100
    df["min_lumpsum"] = df["min_lumpsum"].fillna(0)
    df["risk"] = df["risk"].astype(str).str.strip()
    df["fid"] = [f"f{i:03d}" for i in range(len(df))]
    return df, None


def init_state():
    ss = st.session_state
    if "rid" not in ss:
        ss.rid = "R-" + uuid.uuid4().hex[:8].upper()
        ss.step = "start"
        ss.alloc = {}
        ss.events = []
        ss.open_details = set()
        ss.started_at = now()


def log(event, fund="", value=""):
    st.session_state.events.append(
        {"rid": st.session_state.rid, "ts": now(), "event": event, "fund": fund, "value": str(value)})


# ---------------------------------------------------------------- styling
def base_css(t):
    st.markdown(f"""
<style>
#MainMenu, footer, header [data-testid="stToolbar"] {{visibility: hidden;}}
.block-container {{padding-top: 1.5rem; max-width: 760px;}}
html, body, [class*="css"] {{color: {t['ink']};}}
.stButton > button[kind="primary"], .stFormSubmitButton > button {{
  background: {t['accent']}; border-color: {t['accent']}; color: #fff; border-radius: 10px;}}
.stButton > button[kind="primary"]:hover {{filter: brightness(0.95); color: #fff;}}
.stButton > button[kind="secondary"] {{border-radius: 10px;}}
.sim-banner {{font-size: 12px; color: #6b6b6b; background: #f6f6f6; border-radius: 8px;
  padding: 6px 10px; margin-bottom: 14px;}}
.hero {{background: {t['soft']}; border-radius: 16px; padding: 18px 20px; margin-bottom: 14px;}}
.hero .lbl {{font-size: 13px; opacity: .75;}}
.hero .big {{font-size: 30px; font-weight: 600; color: {t['ink']};}}
.hero .row {{display: flex; gap: 28px; margin-top: 10px; flex-wrap: wrap;}}
.hero .row div span {{display: block; font-size: 12px; opacity: .7;}}
.hero .row div b {{font-size: 16px;}}
.fund-name {{font-size: 16px; font-weight: 600; color: {t['ink']}; line-height: 1.3;}}
.fund-meta {{font-size: 13px; opacity: .7; margin-top: 2px;}}
.ret {{text-align: right;}}
.ret b {{font-size: 18px; color: {t['accent']};}}
.ret span {{display: block; font-size: 11px; opacity: .65;}}
.chips {{margin-top: 8px; display: flex; gap: 6px; flex-wrap: wrap;}}
.chip {{font-size: 11.5px; padding: 3px 8px; border-radius: 999px; background: #f2f2f2;}}
.risk {{color: #fff;}}
.rowret {{display: grid; grid-template-columns: repeat(3, 1fr); gap: 6px; text-align: right; font-size: 14px;}}
.rowret span {{display: block; font-size: 11px; opacity: .6;}}
.det {{display: grid; grid-template-columns: 1fr 1fr; gap: 10px 18px; background: {t['soft']};
  border-radius: 12px; padding: 12px 14px; margin: 4px 0 6px; font-size: 13.5px;}}
.det span {{display: block; font-size: 11.5px; opacity: .65;}}
.bar {{height: 8px; border-radius: 4px; background: #eee; overflow: hidden; margin: 3px 0 10px;}}
.bar > div {{height: 100%; background: {t['accent']};}}
[data-testid="stVerticalBlockBorderWrapper"] {{border-radius: 14px;}}
</style>""", unsafe_allow_html=True)


def risk_chip(level):
    i = RISK_LEVELS.index(level) if level in RISK_LEVELS else 3
    return f'<span class="chip risk" style="background:{RISK_COLOURS[i]}">{esc(level)} risk</span>'


# ---------------------------------------------------------------- pages
def page_start():
    base_css(APPS["Other / I don't invest yet"])
    st.markdown("### Invest ₹1,00,000")
    st.write("A short investing exercise for a GIM research study. It takes about 5 minutes. "
             "No real money is involved and nothing is bought.")
    name = st.text_input("Your name", key="in_name", placeholder="First name is fine")
    app = st.selectbox("Which app do you mainly use to invest in mutual funds?", list(APPS.keys()),
                       index=None, placeholder="Select an app", key="in_app")
    if st.button("Continue", type="primary", use_container_width=True):
        if not (name or "").strip() or not app:
            st.error("Please enter your name and choose an app.")
            return
        ss = st.session_state
        ss.name, ss.app = name.strip(), app
        log("start", value=app)
        ss.step = "invest"
        st.rerun()


def on_amount(fid, fund_name):
    v = st.session_state.get(f"amt_{fid}") or 0
    st.session_state.alloc[fid] = int(v)
    log("set_amount", fund_name, int(v))


def on_toggle(fid, fund_name):
    s = st.session_state.open_details
    if fid in s:
        s.discard(fid)
        log("close_details", fund_name)
    else:
        s.add(fid)
        log("open_details", fund_name)


def fund_card(f, t, position):
    fid, name = f["fid"], f["fund_name"]
    with st.container(border=True):
        if t["layout"] == "cards":
            c1, c2 = st.columns([3, 1])
            c1.markdown(f'<div class="fund-name">{esc(name)}</div>'
                        f'<div class="fund-meta">{esc(f["amc"])} · {esc(f["sub_category"])}</div>',
                        unsafe_allow_html=True)
            c2.markdown(f'<div class="ret"><b>{pct(f["return_3y"])}</b><span>3Y annualised</span></div>',
                        unsafe_allow_html=True)
        else:
            c1, c2 = st.columns([2, 2])
            c1.markdown(f'<div class="fund-name">{esc(name)}</div>'
                        f'<div class="fund-meta">{esc(f["amc"])} · {esc(f["sub_category"])}</div>',
                        unsafe_allow_html=True)
            c2.markdown(f'<div class="rowret"><div><span>1Y</span>{pct(f["return_1y"])}</div>'
                        f'<div><span>3Y</span>{pct(f["return_3y"])}</div>'
                        f'<div><span>5Y</span>{pct(f["return_5y"])}</div></div>', unsafe_allow_html=True)
        st.markdown(f'<div class="chips">{risk_chip(f["risk"])}'
                    f'<span class="chip">Expense ratio {er(f["expense_ratio"])}</span>'
                    f'<span class="chip">Min ₹{inr(f["min_lumpsum"])}</span></div>', unsafe_allow_html=True)

        if fid in st.session_state.open_details:
            aum = "–" if pd.isna(f["aum_cr"]) else f"₹{inr(f['aum_cr'])} Cr"
            since = "–" if pd.isna(f["launch_year"]) else esc(int(f["launch_year"]))
            st.markdown(f"""<div class="det">
<div><span>NAV</span>₹{f['nav']:,.2f}</div>
<div><span>Risk</span>{esc(f['risk'])}</div>
<div><span>Returns 1Y / 3Y / 5Y</span>{pct(f['return_1y'])} / {pct(f['return_3y'])} / {pct(f['return_5y'])}</div>
<div><span>Expense ratio</span>{er(f['expense_ratio'])}</div>
<div><span>Min investment</span>₹{inr(f['min_lumpsum'])}</div>
<div><span>Exit load</span>{esc(f['exit_load'])}</div>
<div><span>Fund manager</span>{esc(f['fund_manager'])}</div>
<div><span>Fund size (AUM)</span>{aum}</div>
<div><span>Launched</span>{since}</div>
<div><span>Category</span>{esc(f['category'])} · {esc(f['sub_category'])}</div>
</div>""", unsafe_allow_html=True)

        b1, b2 = st.columns([1, 1.4], vertical_alignment="bottom")
        open_ = fid in st.session_state.open_details
        b1.button("Hide details" if open_ else "View details", key=f"det_{fid}",
                  on_click=on_toggle, args=(fid, name), use_container_width=True)
        b2.number_input("Amount to invest (₹)", min_value=0, max_value=TOTAL, step=STEP,
                        value=st.session_state.alloc.get(fid, 0), key=f"amt_{fid}",
                        on_change=on_amount, args=(fid, name))
        amt = st.session_state.alloc.get(fid, 0)
        if 0 < amt < f["min_lumpsum"]:
            st.caption(f":red[Minimum investment for this fund is ₹{inr(f['min_lumpsum'])}.]")


def page_invest(df):
    ss = st.session_state
    t = APPS[ss.app]
    base_css(t)

    if "order" not in ss:
        ids = df["fid"].tolist()
        if SHUFFLE_ORDER:
            random.Random(ss.rid).shuffle(ids)
        ss.order = ids
        log("list_order", value=json.dumps([df.set_index("fid").loc[i, "fund_name"] for i in ids]))

    invested = sum(ss.alloc.values())
    left = TOTAL - invested
    n = sum(1 for v in ss.alloc.values() if v > 0)

    st.markdown(f'<div class="sim-banner">Research simulation styled to feel like {esc(ss.app)}. '
                f'Not affiliated with any app or fund house. No real money is used.</div>',
                unsafe_allow_html=True)
    st.markdown(f"""<div class="hero"><div class="lbl">Hi {esc(ss.name)}, you have</div>
<div class="big">₹{inr(TOTAL)}</div>
<div class="row"><div><span>Invested</span><b>₹{inr(invested)}</b></div>
<div><span>Left as cash</span><b>₹{inr(left)}</b></div>
<div><span>Funds chosen</span><b>{n}</b></div></div></div>""", unsafe_allow_html=True)
    st.write("Treat this as your own money. Put it into any funds you like, in any amounts. "
             "Anything you don't invest stays as cash.")
    if left < 0:
        st.error(f"You've gone over ₹1,00,000 by ₹{inr(-left)}. Reduce an amount to continue.")

    cats = ["All"] + sorted(df["category"].dropna().unique().tolist())
    st.radio("Category", cats, horizontal=True, key="flt", label_visibility="collapsed",
             on_change=lambda: log("filter", value=st.session_state.flt))
    s1, s2 = st.columns([1.3, 1])
    s1.text_input("Search", key="q", placeholder="Search funds or fund houses", label_visibility="collapsed",
                  on_change=lambda: log("search", value=st.session_state.q))
    s2.selectbox("Sort", ["Default", "3Y returns: high to low", "Expense ratio: low to high",
                          "Risk: low to high", "Fund size: large to small"],
                 key="srt", label_visibility="collapsed", on_change=lambda: log("sort", value=st.session_state.srt))

    view = df.set_index("fid").loc[ss.order].reset_index()
    view["position"] = range(1, len(view) + 1)
    if ss.flt != "All":
        view = view[view["category"] == ss.flt]
    if ss.q:
        q = ss.q.lower()
        view = view[view["fund_name"].str.lower().str.contains(q) | view["amc"].str.lower().str.contains(q)]
    srt = ss.srt
    if srt.startswith("3Y"):
        view = view.sort_values("return_3y", ascending=False, na_position="last")
    elif srt.startswith("Expense"):
        view = view.sort_values("expense_ratio", na_position="last")
    elif srt.startswith("Risk"):
        view = view.assign(_r=view["risk"].map({r: i for i, r in enumerate(RISK_LEVELS)})).sort_values("_r")
    elif srt.startswith("Fund size"):
        view = view.sort_values("aum_cr", ascending=False, na_position="last")

    if view.empty:
        st.info("No funds match. Try another category or search.")
    for _, f in view.iterrows():
        fund_card(f, t, int(f["position"]))

    # ---- portfolio summary
    st.markdown("#### Your portfolio")
    chosen = df[df["fid"].map(lambda i: ss.alloc.get(i, 0) > 0)].copy()
    if chosen.empty:
        st.caption("You haven't invested in any fund yet. Everything is currently cash.")
    else:
        chosen["amount"] = chosen["fid"].map(ss.alloc)
        for _, r in chosen.sort_values("amount", ascending=False).iterrows():
            share = r["amount"] / TOTAL
            st.markdown(f'<div style="display:flex;justify-content:space-between;font-size:14px">'
                        f'<span>{esc(r["fund_name"])}</span><b>₹{inr(r["amount"])} · {share:.0%}</b></div>'
                        f'<div class="bar"><div style="width:{min(share, 1) * 100:.1f}%"></div></div>',
                        unsafe_allow_html=True)
    if left > 0:
        st.markdown(f'<div style="display:flex;justify-content:space-between;font-size:14px;opacity:.7">'
                    f'<span>Cash (not invested)</span><b>₹{inr(left)} · {left / TOTAL:.0%}</b></div>',
                    unsafe_allow_html=True)

    below_min = [r["fund_name"] for _, r in df.iterrows() if 0 < ss.alloc.get(r["fid"], 0) < r["min_lumpsum"]]
    if below_min:
        st.warning("Some amounts are below the fund's minimum: " + ", ".join(below_min))
    ok = st.checkbox("I'm happy with this portfolio", key="confirm")
    if st.button("Confirm and submit", type="primary", use_container_width=True,
                 disabled=(left < 0 or bool(below_min) or not ok)):
        submit(df)


def submit(df):
    ss = st.session_state
    ss.submitted_at = now()
    log("submit", value=sum(ss.alloc.values()))
    pos = {fid: i + 1 for i, fid in enumerate(ss.order)}
    allocs = []
    for _, r in df.iterrows():
        a = ss.alloc.get(r["fid"], 0)
        if a > 0:
            allocs.append({"rid": ss.rid, "fund_name": r["fund_name"], "amc": r["amc"], "category": r["category"],
                           "sub_category": r["sub_category"], "risk": r["risk"], "amount": a,
                           "share": round(a / TOTAL, 4), "list_position": pos[r["fid"]]})
    invested = sum(a["amount"] for a in allocs)
    duration = (datetime.fromisoformat(ss.submitted_at) - datetime.fromisoformat(ss.started_at)).total_seconds()
    resp = {"rid": ss.rid, "name": ss.name, "app": ss.app, "started_at": ss.started_at,
            "submitted_at": ss.submitted_at, "seconds_taken": int(duration), "invested": invested,
            "cash": TOTAL - invested, "n_funds": len(allocs),
            "largest_share": max((a["share"] for a in allocs), default=0),
            "details_opened": sum(1 for e in ss.events if e["event"] == "open_details")}
    try:
        save(resp, allocs, ss.events)
        ss.saved = True
    except Exception as e:  # keep the respondent moving even if saving fails
        ss.saved = False
        ss.save_error = str(e)
    ss.step = "done"
    st.rerun()


def save(resp, allocs, events):
    if secret("gcp_service_account") and secret("sheet_id"):
        import gspread
        from google.oauth2.service_account import Credentials
        creds = Credentials.from_service_account_info(
            dict(secret("gcp_service_account")), scopes=["https://www.googleapis.com/auth/spreadsheets"])
        sh = gspread.authorize(creds).open_by_key(secret("sheet_id"))

        def tab(name, header):
            try:
                w = sh.worksheet(name)
            except gspread.WorksheetNotFound:
                w = sh.add_worksheet(name, rows=2000, cols=len(header))
            if not w.row_values(1):
                w.append_row(header)
            return w

        tab("responses", list(resp)).append_row([str(v) for v in resp.values()])
        if allocs:
            tab("allocations", list(allocs[0])).append_rows([[str(v) for v in a.values()] for a in allocs])
        tab("events", list(events[0])).append_rows([[str(v) for v in e.values()] for e in events])
    else:
        LOCAL_DIR.mkdir(exist_ok=True)
        for name, rows in [("responses", [resp]), ("allocations", allocs), ("events", events)]:
            if rows:
                p = LOCAL_DIR / f"{name}.csv"
                pd.DataFrame(rows).to_csv(p, mode="a", header=not p.exists(), index=False)


def page_done():
    ss = st.session_state
    base_css(APPS[ss.app])
    st.markdown("### Thank you")
    st.write("Your portfolio has been recorded. One last step: a short questionnaire (about 4 minutes).")
    st.markdown(f"Your respondent ID is **{ss.rid}**.")
    form_url, id_field = secret("form_url"), secret("form_id_field")
    if form_url:
        link = f"{form_url}?usp=pp_url&{id_field}={ss.rid}" if id_field else form_url
        st.link_button("Open the questionnaire", link, type="primary", use_container_width=True)
        if not id_field:
            st.caption("Please enter your respondent ID at the start of the form.")
    else:
        st.info("Questionnaire link not set yet (add form_url in the app's secrets).")
    if not ss.get("saved", True):
        st.error(f"Your answers could not be saved: {ss.get('save_error', '')}. "
                 "Please share your respondent ID with the research team.")


# ---------------------------------------------------------------- main
init_state()
if st.session_state.step == "start":
    page_start()
else:
    if not DATA_FILE.exists():
        st.error("funds.xlsx not found next to app.py.")
        st.stop()
    funds, err = load_funds(str(DATA_FILE))
    if err:
        st.error(err)
        st.stop()
    page_invest(funds) if st.session_state.step == "invest" else page_done()
