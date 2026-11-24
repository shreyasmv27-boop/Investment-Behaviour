"""Mutual fund allocation simulator, built for the Capitalmind 'Crossing the Chasm' study.

Flow: start page (name + app used) -> explore list -> fund detail page (invest here) -> submit -> Google Form link.
Fund data is read from funds.xlsx. Responses go to Google Sheets if configured, otherwise to local CSV files.
"""

import hashlib
import html
import json
import random
import uuid
from datetime import datetime, timezone
from pathlib import Path

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Invest", page_icon="₹", layout="wide", initial_sidebar_state="collapsed")

# ---------------------------------------------------------------- settings
TOTAL = 100_000              # investable amount per respondent
STEP = 500                   # amount input step
SHUFFLE_ORDER = True         # random fund order per respondent; the order is logged to test position effects
DATA_FILE = Path(__file__).parent / "funds.xlsx"
LOCAL_DIR = Path(__file__).parent / "responses"

# Accent colours are approximate and only set the feel. No real logos or wordmarks are used.
APPS = {
    "Groww":        dict(accent="#00B386", soft="#E8F8F3", ink="#44475B"),
    "Zerodha Coin": dict(accent="#387ED1", soft="#EEF4FC", ink="#2B2B2B"),
    "Angel One":    dict(accent="#1F4FD8", soft="#EDF2FD", ink="#1B2340"),
    "PhonePe":      dict(accent="#5F259F", soft="#F3EDFA", ink="#2A1640"),
    "Paytm Money":  dict(accent="#00A3E0", soft="#E9F7FD", ink="#0B2A5B"),
    "Other / I don't invest yet": dict(accent="#3B5BDB", soft="#EEF2FF", ink="#222222"),
}

RISK_LEVELS = ["Low", "Low to Moderate", "Moderate", "Moderately High", "High", "Very High"]

# Colour pairs for generated fund-house logos (background, mountain, sun)
LOGO_PALETTES = [
    ("#FDE7B0", "#5FB37A", "#F5C342"), ("#D6E6FF", "#2F5BD3", "#E2483D"), ("#FFE1D6", "#E4572E", "#F6AE2D"),
    ("#E3DBFF", "#6741D9", "#F06595"), ("#D3F4EE", "#0CA678", "#4DABF7"), ("#FFE3EC", "#D6336C", "#FCC419"),
    ("#E7F5FF", "#1C7ED6", "#51CF66"), ("#FFF3BF", "#F08C00", "#E64980"),
]

REQUIRED = ["fund_name", "amc", "category", "sub_category", "nav", "return_1y", "return_3y", "return_5y",
            "risk", "min_lumpsum", "exit_load", "expense_ratio", "fund_manager", "aum_cr"]
OPTIONAL = ["plan", "min_sip", "launch_year"]


def secret(key, default=None):
    try:
        return st.secrets[key]
    except Exception:
        return default


# ---------------------------------------------------------------- formatting
def inr(n, dec=0):
    """Indian digit grouping: 146784.53 -> 1,46,784.53"""
    neg = n < 0
    n = abs(n)
    whole = int(n)
    frac = f"{n - whole:.{dec}f}"[1:] if dec else ""
    s = str(whole)
    if len(s) > 3:
        head, tail = s[:-3], s[-3:]
        parts = []
        while len(head) > 2:
            parts.insert(0, head[-2:])
            head = head[:-2]
        if head:
            parts.insert(0, head)
        s = ",".join(parts) + "," + tail
    return ("-" if neg else "") + s + frac


def pct(v, d=1):
    return "–" if pd.isna(v) else f"{v:.{d}f}%"


def esc(v):
    return html.escape("" if pd.isna(v) else str(v))


def now():
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


# ---------------------------------------------------------------- visuals
def logo(amc, size=44):
    """A colourful generated logo per fund house (deterministic, not a real brand mark)."""
    i = int(hashlib.md5(str(amc).encode()).hexdigest(), 16) % len(LOGO_PALETTES)
    bg, mt, sun = LOGO_PALETTES[i]
    return (f'<svg width="{size}" height="{size}" viewBox="0 0 44 44" aria-hidden="true">'
            f'<defs><clipPath id="c{i}"><circle cx="22" cy="22" r="21"/></clipPath></defs>'
            f'<g clip-path="url(#c{i})"><rect width="44" height="44" fill="{bg}"/>'
            f'<circle cx="29" cy="14" r="6" fill="{sun}"/>'
            f'<path d="M-2 40 L14 18 L24 30 L31 22 L46 40 Z" fill="{mt}"/>'
            f'<path d="M-2 44 L14 26 L30 44 Z" fill="{mt}" opacity=".7"/></g></svg>')


def risk_gauge(level):
    """Semicircle riskometer from green (low) to red (very high) with a needle."""
    import math
    i = RISK_LEVELS.index(level) if level in RISK_LEVELS else 3
    ang = math.radians(180 - i * 36)            # Low = left (180 deg), Very High = right (0 deg)
    cx, cy, r = 110, 108, 64
    x2, y2 = cx + r * math.cos(ang), cy - r * math.sin(ang)
    return f"""<div class="gauge"><svg width="220" height="126" viewBox="0 0 220 126" role="img"
aria-label="Risk level {esc(level)}"><defs><linearGradient id="rg" x1="0" x2="1" y1="0" y2="0">
<stop offset="0" stop-color="#7CC47F"/><stop offset=".45" stop-color="#F6EE6E"/>
<stop offset=".75" stop-color="#F3A15C"/><stop offset="1" stop-color="#E0554A"/></linearGradient></defs>
<path d="M22 108 A88 88 0 0 1 198 108" fill="none" stroke="url(#rg)" stroke-width="16" stroke-linecap="round"/>
<line x1="{cx}" y1="{cy}" x2="{x2:.1f}" y2="{y2:.1f}" stroke="#222" stroke-width="5" stroke-linecap="round"/>
<circle cx="{cx}" cy="{cy}" r="10" fill="#222"/></svg>
<div><div class="g-l">Risk involved</div><div class="g-v">{esc(level)}</div></div></div>"""


def base_css(t):
    st.markdown(f"""
<style>
#MainMenu, footer, [data-testid="stToolbar"] {{visibility: hidden;}}
.block-container {{padding-top: 1.2rem; max-width: 1120px;}}
.stButton > button[kind="primary"], [data-testid="stFormSubmitButton"] > button {{
  background: {t['accent']}; border-color: {t['accent']}; color: #fff; border-radius: 10px;}}
.stButton > button[kind="primary"]:hover {{filter: brightness(.95); color: #fff;}}
.stButton > button {{border-radius: 10px;}}
.topbar {{display:flex; align-items:center; gap:12px; padding:4px 0 14px; border-bottom:1px solid #eee; margin-bottom:16px;}}
.dot {{width:30px; height:30px; border-radius:9px; background:{t['accent']}; color:#fff; display:flex;
  align-items:center; justify-content:center; font-weight:700;}}
.topbar b {{font-size:18px; color:{t['ink']};}}
.topbar span {{font-size:12px; color:#8a8a8a; margin-left:auto;}}
.wallet {{background:linear-gradient(120deg, {t['accent']}, {t['accent']}CC); color:#fff; border-radius:18px;
  padding:18px 22px; margin-bottom:16px;}}
.wallet .l {{font-size:13px; opacity:.85;}}
.wallet .big {{font-size:30px; font-weight:700;}}
.wallet .row {{display:flex; gap:32px; margin-top:8px; flex-wrap:wrap;}}
.wallet .row span {{display:block; font-size:12px; opacity:.8;}}
.wallet .row b {{font-size:17px;}}
.fhead {{display:flex; gap:14px; align-items:center;}}
.fname {{font-size:17px; color:{t['ink']}; font-weight:500; line-height:1.3;}}
.fmeta {{font-size:14px; color:#8a8a8a; margin-top:3px; display:flex; gap:16px; flex-wrap:wrap;}}
.fmeta b {{color:#555; font-weight:500;}}
.inv {{display:inline-block; font-size:12px; background:{t['soft']}; color:{t['accent']}; border-radius:999px;
  padding:2px 10px; font-weight:600;}}
.pill {{display:inline-block; font-size:13px; background:{t['soft']}; color:{t['accent']}; border-radius:999px;
  padding:4px 12px; margin-right:6px;}}
.gauge {{display:flex; align-items:flex-end; gap:22px; padding:6px 0 4px;}}
.g-l {{font-size:15px; color:#8a8a8a;}}
.g-v {{font-size:22px; font-weight:600; color:#222;}}
.stats {{display:grid; grid-template-columns:repeat(3, 1fr); gap:0; border:1px solid #eee; border-radius:14px; overflow:hidden;}}
.stats div {{padding:12px 16px; border-bottom:1px solid #f0f0f0; font-size:15px; color:#222;}}
.stats span {{display:block; font-size:12px; color:#8a8a8a; margin-bottom:2px;}}
.rets {{display:flex; gap:12px; margin:6px 0 4px;}}
.rets div {{flex:1; background:{t['soft']}; border-radius:12px; padding:10px 14px;}}
.rets span {{display:block; font-size:12px; color:#777;}}
.rets b {{font-size:20px; color:{t['accent']};}}
.bar {{height:8px; border-radius:4px; background:#eee; overflow:hidden; margin:3px 0 12px;}}
.bar > div {{height:100%; background:{t['accent']};}}
.note {{font-size:12px; color:#8a8a8a; background:#f7f7f7; border-radius:8px; padding:6px 10px; margin-top:16px;}}
[data-testid="stVerticalBlockBorderWrapper"] {{border-radius:14px;}}
@media (max-width: 640px) {{ .stats {{grid-template-columns:1fr 1fr;}} }}
</style>""", unsafe_allow_html=True)


# ---------------------------------------------------------------- data + state
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
    for c in ["nav", "return_1y", "return_3y", "return_5y", "min_lumpsum", "min_sip", "expense_ratio", "aum_cr"]:
        df[c] = pd.to_numeric(df[c], errors="coerce")
    # Values typed as fractions (0.125) are converted to percents (12.5)
    for c in ["return_1y", "return_3y", "return_5y"]:
        if df[c].abs().max() <= 1.0:
            df[c] = df[c] * 100
    if df["expense_ratio"].max() < 0.05:
        df["expense_ratio"] = df["expense_ratio"] * 100
    df["plan"] = df["plan"].fillna("Growth")
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
        ss.view = None
        ss.started_at = now()


def log(event, fund="", value=""):
    st.session_state.events.append(
        {"rid": st.session_state.rid, "ts": now(), "event": event, "fund": fund, "value": str(value)})


def open_fund(fid, name, position):
    st.session_state.view = fid
    log("open_fund", name, position)


def close_fund(name):
    st.session_state.view = None
    log("close_fund", name)


# ---------------------------------------------------------------- pages
def page_start():
    base_css(APPS["Other / I don't invest yet"])
    _, mid, _ = st.columns([1, 1.6, 1])
    with mid:
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


def topbar(t):
    ss = st.session_state
    st.markdown(f'<div class="topbar"><div class="dot">₹</div><b>Mutual funds</b>'
                f'<span>Research simulation in a {esc(ss.app)}-style layout · not affiliated with any app or fund house · no real money</span></div>',
                unsafe_allow_html=True)


def wallet(t):
    ss = st.session_state
    invested = sum(ss.alloc.values())
    n = sum(1 for v in ss.alloc.values() if v > 0)
    st.markdown(f"""<div class="wallet"><div class="l">Hi {esc(ss.name)}, your investable amount</div>
<div class="big">₹{inr(TOTAL)}</div><div class="row">
<div><span>Invested</span><b>₹{inr(invested)}</b></div>
<div><span>Left as cash</span><b>₹{inr(TOTAL - invested)}</b></div>
<div><span>Funds chosen</span><b>{n}</b></div></div></div>""", unsafe_allow_html=True)


def page_list(df, t):
    ss = st.session_state
    topbar(t)
    wallet(t)
    st.write("Treat this as your own money. Open any fund to see its details and invest. "
             "Anything you don't invest stays as cash.")

    left, right = st.columns([1, 3], gap="large")
    with left:
        cats = ["All"] + sorted(df["category"].dropna().unique().tolist())
        st.markdown("**Category**")
        st.radio("Category", cats, key="flt", label_visibility="collapsed",
                 on_change=lambda: log("filter_category", value=st.session_state.flt))
        pool = df if ss.get("flt", "All") == "All" else df[df["category"] == ss.flt]
        st.markdown("**Sub category**")
        st.multiselect("Sub category", sorted(pool["sub_category"].dropna().unique()), key="subs",
                       label_visibility="collapsed", placeholder="Any",
                       on_change=lambda: log("filter_sub", value="|".join(st.session_state.subs)))
        st.markdown("**Sort by**")
        st.selectbox("Sort", ["Default", "Fund size: large to small", "Fund size: small to large", "Name: A to Z"],
                     key="srt", label_visibility="collapsed",
                     on_change=lambda: log("sort", value=st.session_state.srt))

    with right:
        st.text_input("Search", key="q", placeholder="Search for mutual funds or fund houses",
                      label_visibility="collapsed", on_change=lambda: log("search", value=st.session_state.q))

        view = df.set_index("fid").loc[ss.order].reset_index()
        view["position"] = range(1, len(view) + 1)
        if ss.flt != "All":
            view = view[view["category"] == ss.flt]
        if ss.get("subs"):
            view = view[view["sub_category"].isin(ss.subs)]
        if ss.q:
            q = ss.q.lower()
            view = view[view["fund_name"].str.lower().str.contains(q, regex=False)
                        | view["amc"].str.lower().str.contains(q, regex=False)]
        if ss.srt == "Fund size: large to small":
            view = view.sort_values("aum_cr", ascending=False, na_position="last")
        elif ss.srt == "Fund size: small to large":
            view = view.sort_values("aum_cr", na_position="last")
        elif ss.srt == "Name: A to Z":
            view = view.sort_values("fund_name")

        chips = "".join(f'<span class="pill">{esc(x)}</span>'
                        for x in ([ss.flt] if ss.flt != "All" else []) + list(ss.get("subs", [])))
        st.markdown(f'<div style="display:flex;justify-content:space-between;align-items:center;margin:2px 0 10px">'
                    f'<div>{chips}</div><span style="color:#8a8a8a;font-size:14px">{len(view)} results</span></div>',
                    unsafe_allow_html=True)
        if view.empty:
            st.info("No funds match. Try another category or search.")
        for _, f in view.iterrows():
            amt = ss.alloc.get(f["fid"], 0)
            with st.container(border=True):
                c1, c2 = st.columns([5, 1.1], vertical_alignment="center")
                badge = f' <span class="inv">₹{inr(amt)} invested</span>' if amt else ""
                aum = "–" if pd.isna(f["aum_cr"]) else f"₹{inr(f['aum_cr'], 2)} Cr."
                c1.markdown(f'<div class="fhead">{logo(f["amc"])}<div><div class="fname">{esc(f["fund_name"])}{badge}</div>'
                            f'<div class="fmeta"><span>AUM <b>{aum}</b></span><span>{esc(f["plan"])}</span>'
                            f'<span>{esc(f["category"])}</span><span>{esc(f["sub_category"])}</span></div></div></div>',
                            unsafe_allow_html=True)
                c2.button("Edit" if amt else "View", key=f"open_{f['fid']}", use_container_width=True,
                          type="secondary", on_click=open_fund, args=(f["fid"], f["fund_name"], int(f["position"])))

        portfolio(df, t)


def page_fund(df, t):
    ss = st.session_state
    f = df.set_index("fid").loc[ss.view]
    name = f["fund_name"]
    current = ss.alloc.get(ss.view, 0)
    others = sum(v for k, v in ss.alloc.items() if k != ss.view)
    room = TOTAL - others
    topbar(t)
    st.button("← Back to all funds", on_click=close_fund, args=(name,))

    aum = "–" if pd.isna(f["aum_cr"]) else f"₹{inr(f['aum_cr'], 2)} Cr."
    st.markdown(f'<div class="fhead" style="margin:8px 0 6px">{logo(f["amc"], 60)}<div>'
                f'<div class="fname" style="font-size:24px">{esc(name)}</div>'
                f'<div class="fmeta"><span>{esc(f["amc"])}</span><span>{esc(f["plan"])}</span>'
                f'<span>{esc(f["category"])}</span><span>{esc(f["sub_category"])}</span></div></div></div>',
                unsafe_allow_html=True)

    a, b = st.columns([1.25, 1], gap="large")
    with a:
        st.markdown(f"""<div class="rets"><div><span>1Y returns</span><b>{pct(f['return_1y'])}</b></div>
<div><span>3Y returns (p.a.)</span><b>{pct(f['return_3y'])}</b></div>
<div><span>5Y returns (p.a.)</span><b>{pct(f['return_5y'])}</b></div></div>""", unsafe_allow_html=True)
        since = "–" if pd.isna(f["launch_year"]) else esc(int(f["launch_year"]))
        min_sip = "–" if pd.isna(f["min_sip"]) else f"₹{inr(f['min_sip'])}"
        st.markdown(f"""<div class="stats">
<div><span>NAV</span>₹{inr(f['nav'], 2)}</div>
<div><span>Fund size (AUM)</span>{aum}</div>
<div><span>Expense ratio</span>{pct(f['expense_ratio'], 2)}</div>
<div><span>Min. lump sum</span>₹{inr(f['min_lumpsum'])}</div>
<div><span>Min. SIP</span>{min_sip}</div>
<div><span>Launched</span>{since}</div>
<div style="grid-column:1 / -1"><span>Exit load</span>{esc(f['exit_load'])}</div>
<div style="grid-column:1 / -1;border-bottom:none"><span>Fund manager</span>{esc(f['fund_manager'])}</div>
</div>""", unsafe_allow_html=True)
    with b:
        st.markdown(risk_gauge(f["risk"]), unsafe_allow_html=True)
        with st.container(border=True):
            st.markdown(f"**Invest in this fund**  \nYou can put up to ₹{inr(room)} here.")
            with st.form(f"form_{ss.view}", border=False):
                amt = st.number_input("Amount (₹)", min_value=0, max_value=TOTAL, step=STEP, value=current)
                go = st.form_submit_button("Update investment" if current else "Invest", use_container_width=True)
            if go:
                if amt > room:
                    st.error(f"That's more than you have left. You can invest up to ₹{inr(room)} here.")
                elif 0 < amt < f["min_lumpsum"]:
                    st.error(f"Minimum investment for this fund is ₹{inr(f['min_lumpsum'])}.")
                else:
                    ss.alloc[ss.view] = int(amt)
                    log("set_amount", name, int(amt))
                    ss.view = None
                    log("close_fund", name)
                    st.rerun()
            if current:
                if st.button("Remove from portfolio", use_container_width=True):
                    ss.alloc[ss.view] = 0
                    log("set_amount", name, 0)
                    ss.view = None
                    st.rerun()


def portfolio(df, t):
    ss = st.session_state
    invested = sum(ss.alloc.values())
    left = TOTAL - invested
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
        st.markdown(f'<div style="display:flex;justify-content:space-between;font-size:14px;color:#8a8a8a">'
                    f'<span>Cash (not invested)</span><b>₹{inr(left)} · {left / TOTAL:.0%}</b></div>',
                    unsafe_allow_html=True)
    ok = st.checkbox("I'm happy with this portfolio", key="confirm")
    if st.button("Confirm and submit", type="primary", use_container_width=True, disabled=not ok):
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
            "funds_opened": len({e["fund"] for e in ss.events if e["event"] == "open_fund"})}
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


def page_done(t):
    ss = st.session_state
    _, mid, _ = st.columns([1, 1.6, 1])
    with mid:
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
ss = st.session_state
if ss.step == "start":
    page_start()
else:
    if not DATA_FILE.exists():
        st.error("funds.xlsx not found next to app.py.")
        st.stop()
    funds, err = load_funds(str(DATA_FILE))
    if err:
        st.error(err)
        st.stop()
    theme = APPS[ss.app]
    base_css(theme)
    if "order" not in ss:
        ids = funds["fid"].tolist()
        if SHUFFLE_ORDER:
            random.Random(ss.rid).shuffle(ids)
        ss.order = ids
        log("list_order", value=json.dumps([funds.set_index("fid").loc[i, "fund_name"] for i in ids]))
    if ss.step == "done":
        page_done(theme)
    elif ss.view:
        page_fund(funds, theme)
    else:
        page_list(funds, theme)
