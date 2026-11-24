"""Mutual fund allocation simulator, built for the Capitalmind 'Crossing the Chasm' study.

Flow: start page (name + app used) -> explore list -> fund page (Buy / SIP) -> portfolio tab.
The portfolio is saved after every change; a Questionnaire link in the side menu opens the Google Form.
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
SIP_MONTHS = 12            # an SIP commits monthly amount x SIP_MONTHS out of the Rs 1,00,000
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
#MainMenu, footer, header[data-testid="stHeader"], [data-testid="stToolbar"] {{display: none !important;}}
.block-container {{padding-top: 2rem !important; max-width: 1180px;}}
.stButton button[kind="tertiary"] {{padding: 0; min-height: 0; justify-content: flex-start; text-align: left;
  color: {t['ink']}; border: none;}}
.stButton button[kind="tertiary"] p {{font-size: 17px; font-weight: 500;}}
.stButton button[kind="tertiary"]:hover {{color: {t['accent']}; text-decoration: underline;}}
.st-key-mode_buy button, .st-key-mode_sip button {{border-radius: 999px; height: 46px;}}
.st-key-mode_buy button p, .st-key-mode_sip button p {{font-size: 17px; font-weight: 600;}}
.st-key-nav_explore button, .st-key-nav_portfolio button {{flex-direction: column; height: 84px; border: none;
  background: transparent; border-radius: 14px;}}
.st-key-nav_explore button p, .st-key-nav_portfolio button p {{font-size: 14px; white-space: normal;
  overflow: visible; text-overflow: clip; line-height: 1.2;}}
.st-key-nav_explore button div, .st-key-nav_portfolio button div {{flex-direction: column; gap: 6px;}}
.st-key-nav_explore button:hover, .st-key-nav_portfolio button:hover {{background: {t['soft']};}}
.st-key-nav_explore [data-testid="stIconMaterial"], .st-key-nav_portfolio [data-testid="stIconMaterial"] {{
  font-size: 30px; color: {t['accent']};}}
.st-key-nav_{st.session_state.get('tab', 'explore')} button {{background: {t['soft']};}}
.hrow {{font-size: 13px; color: #8a8a8a;}}
[data-testid="stExpander"] details {{border: none; border-bottom: 1px solid #eee; border-radius: 0;}}
[data-testid="stExpander"] summary {{padding: 14px 4px;}}
[data-testid="stExpander"] summary p {{font-size: 16px; color: #6b6b6b;}}
[data-testid="stExpander"] summary:hover p {{color: {t['accent']};}}
.cell {{font-size: 15px; color: #222; text-align: right;}}
.cell small {{display: block; font-size: 12px; color: #8a8a8a;}}
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
        ss.tab = "explore"
        ss.mode = None
        ss.started_at = now()


def log(event, fund="", value=""):
    st.session_state.events.append(
        {"rid": st.session_state.rid, "ts": now(), "event": event, "fund": fund, "value": str(value)})


def committed(h):
    return 0 if not h else h["amount"] * (SIP_MONTHS if h["mode"] == "sip" else 1)


def total_committed(exclude=None):
    return sum(committed(h) for k, h in st.session_state.alloc.items() if k != exclude)


def open_fund(fid, name, position):
    ss = st.session_state
    ss.view = fid
    ss.mode = None
    log("open_fund", name, f"{ss.tab}:{position}")


def set_tab(tab):
    st.session_state.tab = tab
    st.session_state.view = None
    log("tab", value=tab)


def set_mode(mode, name):
    st.session_state.mode = mode
    log("choose_mode", name, mode)


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


def topbar():
    ss = st.session_state
    st.markdown('<div class="topbar"><div class="dot">₹</div><b>Mutual funds</b></div>', unsafe_allow_html=True)


def wallet():
    ss = st.session_state
    invested = total_committed()
    n = sum(1 for h in ss.alloc.values() if h)
    st.markdown(f"""<div class="wallet"><div class="l">Hi {esc(ss.name)}, available to invest</div>
<div class="big">₹{inr(TOTAL - invested)}</div><div class="row">
<div><span>Starting amount</span><b>₹{inr(TOTAL)}</b></div>
<div><span>Invested</span><b>₹{inr(invested)}</b></div>
<div><span>Funds chosen</span><b>{n}</b></div></div></div>""", unsafe_allow_html=True)


def sidenav():
    st.button("Mutual funds", key="nav_explore", icon=":material/work:", use_container_width=True,
              on_click=set_tab, args=("explore",))
    n = sum(1 for h in st.session_state.alloc.values() if h)
    st.button(f"Portfolio ({n})" if n else "Portfolio", key="nav_portfolio", icon=":material/pie_chart:",
              use_container_width=True, on_click=set_tab, args=("portfolio",))
    form_url, id_field = secret("form_url"), secret("form_id_field")
    if form_url:
        rid = st.session_state.rid
        link = f"{form_url}?usp=pp_url&{id_field}={rid}" if id_field else form_url
        st.markdown("<br>", unsafe_allow_html=True)
        st.link_button("Questionnaire", link, icon=":material/assignment:", use_container_width=True)


def holding_text(h):
    if not h:
        return ""
    if h["mode"] == "sip":
        return f"SIP ₹{inr(h['amount'])}/month"
    return f"Bought ₹{inr(h['amount'])}"


SORTS = {"Default": (None, True), "Fund size: large to small": ("aum_cr", False),
         "Fund size: small to large": ("aum_cr", True), "3Y CAGR: high to low": ("return_3y", False),
         "Expense ratio: low to high": ("expense_ratio", True), "Name: A to Z": ("fund_name", True)}
F_MIN = {"Any": None, "Up to ₹500": 500, "Up to ₹1,000": 1000, "Up to ₹5,000": 5000}
F_ER = {"Any": None, "Below 0.5%": (0, 0.5), "0.5% to 1%": (0.5, 1.0), "Above 1%": (1.0, 99)}
F_AUM = {"Any": None, "Below ₹5,000 Cr": (0, 5000), "₹5,000 to 20,000 Cr": (5000, 20000),
         "Above ₹20,000 Cr": (20000, 1e9)}
F_EXIT = ["Any", "No exit load", "Has exit load"]
F_CAGR = {"Any": None, "Above 10%": 10, "Above 15%": 15, "Above 20%": 20}
FILTER_KEYS = ["f_amc", "flt", "subs", "f_plan", "f_min", "f_er", "f_aum", "f_exit", "f_cp", "f_cg"]


def logf(key):
    return lambda: log("filter_" + key, value=str(st.session_state.get(key)))


def reset_filters():
    for k in FILTER_KEYS:
        st.session_state.pop(k, None)
    log("filter_reset")


def filter_panel(df):
    ss = st.session_state
    with st.expander("AMC"):
        st.multiselect("AMC", sorted(df["amc"].dropna().unique()), key="f_amc", label_visibility="collapsed",
                       placeholder="All fund houses", on_change=logf("f_amc"))
    with st.expander("Category"):
        cats = ["All"] + sorted(df["category"].dropna().unique().tolist())
        st.radio("Category", cats, key="flt", label_visibility="collapsed", on_change=logf("flt"))
        pool = df if ss.get("flt", "All") == "All" else df[df["category"] == ss.flt]
        st.multiselect("Sub category", sorted(pool["sub_category"].dropna().unique()), key="subs",
                       placeholder="Any sub category", on_change=logf("subs"))
    with st.expander("Dividend type"):
        st.multiselect("Dividend type", sorted(df["plan"].dropna().astype(str).unique()), key="f_plan",
                       label_visibility="collapsed", placeholder="Any", on_change=logf("f_plan"))
    with st.expander("Min. purchase amount"):
        st.radio("Min. purchase amount", list(F_MIN), key="f_min", label_visibility="collapsed",
                 on_change=logf("f_min"))
    with st.expander("Expense ratio"):
        st.radio("Expense ratio", list(F_ER), key="f_er", label_visibility="collapsed", on_change=logf("f_er"))
    with st.expander("AUM"):
        st.radio("AUM", list(F_AUM), key="f_aum", label_visibility="collapsed", on_change=logf("f_aum"))
    with st.expander("Exit load"):
        st.radio("Exit load", F_EXIT, key="f_exit", label_visibility="collapsed", on_change=logf("f_exit"))
    with st.expander("CAGR"):
        st.radio("Period", ["1Y", "3Y", "5Y"], key="f_cp", horizontal=True, index=1, on_change=logf("f_cp"))
        st.radio("Minimum CAGR", list(F_CAGR), key="f_cg", on_change=logf("f_cg"))
    st.button("Reset filters", on_click=reset_filters, use_container_width=True)


def apply_filters(v):
    """Returns the filtered list and labels for the active filters."""
    ss, active = st.session_state, []
    if ss.get("f_amc"):
        v = v[v["amc"].isin(ss.f_amc)]
        active += ss.f_amc
    if ss.get("flt", "All") != "All":
        v = v[v["category"] == ss.flt]
        active.append(ss.flt)
    if ss.get("subs"):
        v = v[v["sub_category"].isin(ss.subs)]
        active += ss.subs
    if ss.get("f_plan"):
        v = v[v["plan"].astype(str).isin(ss.f_plan)]
        active += ss.f_plan
    if F_MIN.get(ss.get("f_min", "Any")):
        v = v[v["min_lumpsum"] <= F_MIN[ss.f_min]]
        active.append("Min " + ss.f_min.lower())
    if F_ER.get(ss.get("f_er", "Any")):
        lo, hi = F_ER[ss.f_er]
        v = v[(v["expense_ratio"] >= lo) & (v["expense_ratio"] < hi)]
        active.append("Expense ratio " + ss.f_er.lower())
    if F_AUM.get(ss.get("f_aum", "Any")):
        lo, hi = F_AUM[ss.f_aum]
        v = v[(v["aum_cr"] >= lo) & (v["aum_cr"] < hi)]
        active.append("AUM " + ss.f_aum.lower())
    if ss.get("f_exit", "Any") != "Any":
        nil = v["exit_load"].astype(str).str.strip().str.lower().str.startswith(("nil", "0", "none"))
        v = v[nil] if ss.f_exit == "No exit load" else v[~nil]
        active.append(ss.f_exit)
    if F_CAGR.get(ss.get("f_cg", "Any")):
        p = ss.get("f_cp", "3Y")
        v = v[v[f"return_{p.lower()}"] > F_CAGR[ss.f_cg]]
        active.append(f"{p} CAGR {ss.f_cg.lower()}")
    return v, active


def page_explore(df):
    ss = st.session_state
    st.write("Treat this as your own money. Click any fund to see its details and invest. "
             "Anything you don't invest stays as cash.")
    left, right = st.columns([1, 3], gap="large")
    with left:
        filter_panel(df)
    with right:
        s1, s2 = st.columns([2.2, 1])
        s1.text_input("Search", key="q", placeholder="Search for mutual funds or fund houses",
                      label_visibility="collapsed", on_change=lambda: log("search", value=st.session_state.q))
        s2.selectbox("Sort", list(SORTS), key="srt", label_visibility="collapsed",
                     on_change=lambda: log("sort", value=st.session_state.srt))
        view = df.set_index("fid").loc[ss.order].reset_index()
        view["position"] = range(1, len(view) + 1)
        view, active = apply_filters(view)
        if ss.q:
            q = ss.q.lower()
            view = view[view["fund_name"].str.lower().str.contains(q, regex=False)
                        | view["amc"].str.lower().str.contains(q, regex=False)]
        col, asc = SORTS[ss.srt]
        if col:
            view = view.sort_values(col, ascending=asc, na_position="last")

        chips = "".join(f'<span class="pill">{esc(x)}</span>' for x in active)
        st.markdown(f'<div style="display:flex;justify-content:space-between;align-items:center;margin:2px 0 10px">'
                    f'<div>{chips}</div><span style="color:#8a8a8a;font-size:14px">{len(view)} results</span></div>',
                    unsafe_allow_html=True)
        if view.empty:
            st.info("No funds match. Try another category or search.")
        for _, f in view.iterrows():
            h = ss.alloc.get(f["fid"])
            with st.container(border=True):
                c0, c1 = st.columns([0.09, 0.91], vertical_alignment="center")
                c0.markdown(logo(f["amc"]), unsafe_allow_html=True)
                with c1:
                    st.button(f["fund_name"], key=f"open_{f['fid']}", type="tertiary",
                              on_click=open_fund, args=(f["fid"], f["fund_name"], int(f["position"])))
                    aum = "–" if pd.isna(f["aum_cr"]) else f"₹{inr(f['aum_cr'], 2)} Cr."
                    badge = f'<span class="inv">{holding_text(h)}</span>' if h else ""
                    st.markdown(f'<div class="fmeta" style="margin-top:-6px"><span>AUM <b>{aum}</b></span>'
                                f'<span>{esc(f["plan"])}</span><span>{esc(f["category"])}</span>'
                                f'<span>{esc(f["sub_category"])}</span>{badge}</div>', unsafe_allow_html=True)


def page_portfolio(df):
    ss = st.session_state
    held = [(fid, h) for fid, h in ss.alloc.items() if h]
    if not held:
        st.markdown("#### Your portfolio")
        st.info("You haven't invested in any fund yet. Everything is currently cash.")
        st.button("Explore mutual funds", type="primary", on_click=set_tab, args=("explore",))
        return
    lump = sum(committed(h) for _, h in held if h["mode"] == "buy")
    sip = sum(committed(h) for _, h in held if h["mode"] == "sip")
    st.markdown(f"""<div class="stats" style="grid-template-columns:repeat(4,1fr);margin-bottom:18px">
<div><span>Lump sum</span>₹{inr(lump)}</div><div><span>SIPs ({SIP_MONTHS} months)</span>₹{inr(sip)}</div>
<div><span>Cash</span>₹{inr(TOTAL - lump - sip)}</div><div><span>Funds</span>{len(held)}</div></div>""",
                unsafe_allow_html=True)

    widths = [0.06, 0.4, 0.16, 0.12, 0.14, 0.12]
    hdr = st.columns(widths)
    for c, label in zip(hdr[1:], ["Fund", "Type", "NAV", "Amount", "Allocation"]):
        c.markdown(f'<div class="hrow" style="text-align:{"left" if label == "Fund" else "right"}">{label}</div>',
                   unsafe_allow_html=True)
    fx = df.set_index("fid")
    for fid, h in sorted(held, key=lambda x: -committed(x[1])):
        f = fx.loc[fid]
        st.markdown('<hr style="margin:4px 0 10px;border:none;border-top:1px solid #eee">', unsafe_allow_html=True)
        c = st.columns(widths, vertical_alignment="center")
        c[0].markdown(logo(f["amc"], 36), unsafe_allow_html=True)
        with c[1]:
            st.button(f["fund_name"], key=f"pf_{fid}", type="tertiary", on_click=open_fund,
                      args=(fid, f["fund_name"], ss.order.index(fid) + 1))
            st.markdown(f'<div class="fmeta" style="margin-top:-6px;font-size:13px"><span>{esc(f["plan"])}</span>'
                        f'<span>{esc(f["category"])}</span><span>{esc(f["sub_category"])}</span></div>',
                        unsafe_allow_html=True)
        kind = (f'SIP<small>₹{inr(h["amount"])}/month × {SIP_MONTHS}</small>' if h["mode"] == "sip"
                else f'Lump sum<small>{h["amount"] / f["nav"]:,.3f} units</small>')
        c[2].markdown(f'<div class="cell">{kind}</div>', unsafe_allow_html=True)
        c[3].markdown(f'<div class="cell">₹{inr(f["nav"], 2)}</div>', unsafe_allow_html=True)
        c[4].markdown(f'<div class="cell">₹{inr(committed(h))}</div>', unsafe_allow_html=True)
        c[5].markdown(f'<div class="cell">{committed(h) / TOTAL:.0%}</div>', unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    st.caption("Your portfolio is saved automatically after every change.")


def page_fund(df):
    ss = st.session_state
    fid = ss.view
    f = df.set_index("fid").loc[fid]
    name = f["fund_name"]
    h = ss.alloc.get(fid)
    room = TOTAL - total_committed(exclude=fid)
    sip_ok = not pd.isna(f["min_sip"])

    st.button("← Back to " + ("portfolio" if ss.tab == "portfolio" else "all funds"), on_click=close_fund, args=(name,))
    aum = "–" if pd.isna(f["aum_cr"]) else f"₹{inr(f['aum_cr'], 2)} Cr."
    st.markdown(f'<div class="fhead" style="margin:8px 0 6px">{logo(f["amc"], 60)}<div>'
                f'<div class="fname" style="font-size:24px">{esc(name)}</div>'
                f'<div class="fmeta"><span>{esc(f["amc"])}</span><span>{esc(f["plan"])}</span>'
                f'<span>{esc(f["category"])}</span><span>{esc(f["sub_category"])}</span></div></div></div>',
                unsafe_allow_html=True)

    a, b = st.columns([1.25, 1], gap="large")
    with a:
        st.markdown(f"""<div class="rets"><div><span>1Y CAGR return</span><b>{pct(f['return_1y'])}</b></div>
<div><span>3Y CAGR return</span><b>{pct(f['return_3y'])}</b></div>
<div><span>5Y CAGR return</span><b>{pct(f['return_5y'])}</b></div></div>""", unsafe_allow_html=True)
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
        if h:
            st.markdown(f'<div class="inv" style="font-size:14px;margin-bottom:8px">You have: {holding_text(h)}</div>',
                        unsafe_allow_html=True)
        c1, c2 = st.columns(2)
        c1.button("Buy", key="mode_buy", use_container_width=True,
                  type="primary" if ss.mode == "buy" else "secondary", on_click=set_mode, args=("buy", name))
        c2.button("SIP", key="mode_sip", use_container_width=True, disabled=not sip_ok,
                  type="primary" if ss.mode == "sip" else "secondary", on_click=set_mode, args=("sip", name))
        if not sip_ok:
            st.caption("SIP is not available for this fund.")

        if ss.mode:
            sip = ss.mode == "sip"
            with st.container(border=True):
                if sip:
                    st.markdown(f"**Start a monthly SIP**  \nRuns for {SIP_MONTHS} months. "
                                f"You can commit up to ₹{inr(room)} in total, so at most ₹{inr(room // SIP_MONTHS)}/month.")
                else:
                    st.markdown(f"**One-time purchase**  \nYou can invest up to ₹{inr(room)} here.")
                start = h["amount"] if h and h["mode"] == ss.mode else 0
                with st.form(f"form_{fid}_{ss.mode}", border=False):
                    amt = st.number_input("Monthly amount (₹)" if sip else "Amount (₹)", min_value=0,
                                          max_value=TOTAL, step=100 if sip else STEP, value=start)
                    go = st.form_submit_button("Start SIP" if sip else "Buy now", use_container_width=True)
                if go:
                    minimum = f["min_sip"] if sip else f["min_lumpsum"]
                    total = amt * (SIP_MONTHS if sip else 1)
                    if amt <= 0:
                        st.error("Enter an amount.")
                    elif total > room:
                        st.error(f"That's ₹{inr(total)} in total, more than the ₹{inr(room)} you have left.")
                    elif amt < minimum:
                        st.error(f"Minimum {'SIP' if sip else 'investment'} for this fund is ₹{inr(minimum)}.")
                    else:
                        ss.alloc[fid] = {"mode": ss.mode, "amount": int(amt)}
                        log("set_amount", name, f"{ss.mode}:{int(amt)}")
                        persist(df)
                        close_fund(name)
                        st.rerun()
        if h and st.button("Remove from portfolio", use_container_width=True):
            ss.alloc[fid] = None
            log("set_amount", name, "removed")
            persist(df)
            close_fund(name)
            st.rerun()


def persist(df):
    """Save a snapshot of the current portfolio plus any new events. Called after every portfolio change."""
    ss = st.session_state
    ss.snapshot = ss.get("snapshot", 0) + 1
    ts = now()
    pos = {fid: i + 1 for i, fid in enumerate(ss.order)}
    allocs = []
    for _, r in df.iterrows():
        h = ss.alloc.get(r["fid"])
        if h:
            allocs.append({"rid": ss.rid, "snapshot": ss.snapshot, "fund_name": r["fund_name"], "amc": r["amc"],
                           "category": r["category"], "sub_category": r["sub_category"], "risk": r["risk"],
                           "mode": h["mode"], "amount": h["amount"], "committed": committed(h),
                           "share": round(committed(h) / TOTAL, 4), "list_position": pos[r["fid"]]})
    invested = sum(a["committed"] for a in allocs)
    duration = (datetime.fromisoformat(ts) - datetime.fromisoformat(ss.started_at)).total_seconds()
    resp = {"rid": ss.rid, "snapshot": ss.snapshot, "name": ss.name, "app": ss.app,
            "started_at": ss.started_at, "saved_at": ts, "seconds_taken": int(duration), "invested": invested,
            "cash": TOTAL - invested, "n_funds": len(allocs),
            "n_sips": sum(1 for a in allocs if a["mode"] == "sip"),
            "largest_share": max((a["share"] for a in allocs), default=0),
            "funds_opened": len({e["fund"] for e in ss.events if e["event"] == "open_fund"})}
    new_events = ss.events[ss.get("events_saved", 0):]
    try:
        save(resp, allocs, new_events)
        ss.events_saved = len(ss.events)
        ss.save_error = None
    except Exception as e:  # never block the respondent because saving failed
        ss.save_error = str(e)


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
        if events:
            tab("events", list(events[0])).append_rows([[str(v) for v in e.values()] for e in events])
    else:
        LOCAL_DIR.mkdir(exist_ok=True)
        for name, rows in [("responses", [resp]), ("allocations", allocs), ("events", events)]:
            if rows:
                p = LOCAL_DIR / f"{name}.csv"
                pd.DataFrame(rows).to_csv(p, mode="a", header=not p.exists(), index=False)


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
    if True:
        topbar()
        nav, body = st.columns([0.13, 0.87], gap="medium")
        with nav:
            sidenav()
        with body:
            if ss.view:
                page_fund(funds)
            else:
                wallet()
                if ss.tab == "portfolio":
                    page_portfolio(funds)
                else:
                    page_explore(funds)
