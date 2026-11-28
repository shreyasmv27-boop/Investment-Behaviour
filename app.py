"""Mutual fund allocation simulator, built for the Capitalmind 'Crossing the Chasm' study.

Flow: start page (name + app used) -> fund list -> fund page (Buy / SIP) -> portfolio -> Completed -> thank you.
Every tap is logged. Data goes to Google Sheets if configured (tabs: investors, purchases, portfolio, clicks),
otherwise to CSV files in responses/. The researcher dashboard is at  <app url>/?admin=1  (password in Secrets).
Fund data: funds.xlsx. NAV history for the charts: nav_history.csv.
"""

import hashlib
import html
import json
import queue
import random
import threading
import time
import urllib.parse
import uuid
from datetime import datetime, timedelta, timezone
from pathlib import Path

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Invest", page_icon="₹", layout="wide", initial_sidebar_state="collapsed")

# ---------------------------------------------------------------- settings
TOTAL = 100_000              # investable amount per respondent
STEP = 500                   # amount input step
SIP_MONTHS = 12            # an SIP commits monthly amount x SIP_MONTHS out of the Rs 1,00,000
SHUFFLE_ORDER = False        # False: every participant sees funds in the row order of funds.xlsx (display_order)
DATA_FILE = Path(__file__).parent / "funds.xlsx"
NAV_FILE = Path(__file__).parent / "nav_history.csv"
HOLDINGS_FILE = Path(__file__).parent / "holdings.xlsx"

# ---- Questionnaire shown after "Completed". Every question is mandatory.
# To add a question, add a dict here. "id" becomes the column name in the Google Sheet (keep ids unique and
# add new questions at the end, so existing columns don't shift).
#   type "single": pick one option        type "multi": pick one or more ("only" = can't be combined, "max" = most picks)
#   "other": an option that opens a box to type the answer; saved as "Others: <what they typed>"
#   type "top_bottom": pick exactly `top` best and `bottom` worst from the same list
#   "correct": for quiz questions, the right option (saved as a separate *_correct column)
QUESTIONS = [
    {"id": "experience", "type": "single", "text": "How long have you been investing in mutual funds?",
     "options": ["Less than 2 years", "2 to 6 years", "More than 6 years"]},
    {"id": "quiz_turnover", "type": "single",
     "text": "A mutual fund shows a portfolio turnover of 100% in a year. What does that mean?",
     "options": ["The fund's value doubled in a year", "The fund replaced roughly all its holdings over the year",
                 "100% investment in that year, no cash left", "I'm not sure"],
     "correct": "The fund replaced roughly all its holdings over the year"},
    {"id": "heard_from", "type": "single", "text": "Where did you last hear about a fund you bought?",
     "options": ["Instagram", "YouTube", "Friend or advisor", "Newspaper", "Others"], "other": "Others"},
    {"id": "portfolio_similar", "type": "single", "text": "Is the portfolio you built here similar to your real one?",
     "options": ["Yes", "Partly", "No"]},
    {"id": "philosophy", "type": "top_bottom", "top": 3, "bottom": 2,
     "text": "Based on your investment philosophy, which factors matter most and least when you pick a fund?",
     "options": ["AMC brand", "CAGR return (1Y/3Y/5Y)", "Expense ratio", "Exit load", "Fund manager", "Fund size",
                 "Sector allocation", "Portfolio allocation", "Fund age"]},
    {"id": "new_vs_established", "type": "single",
     "text": "Which would you choose: a new fund (₹400 Cr AUM, 10% 3Y CAGR) or an established fund "
             "(₹10,000 Cr AUM, 8% 3Y CAGR)?",
     "options": ["New fund", "Established fund"]},
    {"id": "visible_manager", "type": "single",
     "text": "Two fund managers have the same 5-year returns. Manager A is often on TV and social media explaining "
             "the market. Manager B rarely speaks in public. Which would you invest with?",
     "options": ["Manager A, the visible one", "Manager B, the quiet one", "Indifferent"]},
    {"id": "new_fund_growth", "type": "multi", "max": 3,
     "text": "Imagine a new fund with almost no advertising. In your view, what is most likely to help it grow? "
             "Pick up to 3.",
     "options": ["Word of mouth", "Better returns than others", "Lower costs", "A clear investment philosophy",
                 "Advisors and distributors recommending it", "The manager's public credibility",
                 "Doing well in a market fall"]},
    {"id": "global_view", "type": "single",
     "text": "For long-term investing, how do you feel about putting some money into US or global markets, "
             "through funds that invest abroad?",
     "options": ["I already do this", "I would consider it", "I'm unsure", "I wouldn't, I prefer India"]},
    {"id": "abandoned_fund", "type": "multi", "only": "No, never",
     "text": "Have you ever stopped investing in a fund or sold it for any of these reasons? Select all that apply.",
     "options": ["Underperformance", "Panic during a market fall", "Fund manager change", "Found a better fund",
                 "Costs (expense ratio or exit load)", "Its sector fell out of favour", "No, never"]},
]


AGES = ["Under 30 years", "30 to 50 years", "Above 50 years"]
CITIES = ["Bengaluru", "Delhi", "Mumbai", "Gurgaon", "Chennai", "Other"]

FLUSH_EVERY = 6              # send clicks to the database after this many taps (plus at every purchase / completion)
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
OPTIONAL = ["plan", "min_sip", "launch_year", "scheme_code"]


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


IST = timezone(timedelta(hours=5, minutes=30))


def now():
    return datetime.now(IST).strftime("%Y-%m-%d %H:%M:%S")


# ---------------------------------------------------------------- visuals
def run_js(code):
    """Run a small script in the page (scrolling, redirect). Uses st.iframe on new Streamlit, components.html on old."""
    doc = f"<script>{code}</script>"
    if hasattr(st, "iframe"):
        st.iframe(doc, height=1)
    else:
        import streamlit.components.v1 as components
        components.html(doc, height=0)


def scroll_top():
    """Streamlit keeps the scroll position between pages; jump back to the top (needed on phones)."""
    n = st.session_state.get("nav_count", 0)
    with st.container(key="scrolltop"):
        run_js(f"""/* {n} */
const w = window.parent, d = w.document;
const go = () => {{
  ['[data-testid="stAppScrollToBottomContainer"]', '[data-testid="stMain"]', '[data-testid="stAppViewContainer"]',
   'section.main', '.stMain', '.main']
    .forEach(q => d.querySelectorAll(q).forEach(e => {{ if (e.scrollTop) e.scrollTop = 0; }}));
  if (d.scrollingElement) d.scrollingElement.scrollTop = 0;
  d.documentElement.scrollTop = 0; d.body.scrollTop = 0; w.scrollTo(0, 0); }};
/* Keep pinning to the top while the page (chart, holdings) finishes drawing on slow phones,
   but stop the moment the person touches or scrolls. */
let stop = false;
const quit = () => {{ stop = true; }};
['touchstart', 'wheel', 'keydown'].forEach(t => w.addEventListener(t, quit, {{once: true, passive: true}}));
go();
[50, 150, 300, 500, 800, 1200, 1700, 2300].forEach(ms => setTimeout(() => {{ if (!stop) go(); }}, ms));""")


def scroll_to(key, n):
    """Bring the buy/SIP box into view after Buy or SIP is tapped."""
    with st.container(key="scrollto"):
        run_js(f"""/* {n} */
setTimeout(() => {{ const e = window.parent.document.querySelector('.st-key-{key}');
  if (e) e.scrollIntoView({{behavior: 'smooth', block: 'center'}}); }}, 250);""")


def manager_links(names, fund):
    """Each fund manager opens a Google search in a new tab. The link passes through the app first so the tap is logged."""
    ss = st.session_state
    out = []
    for n in [x.strip() for x in str(names).split(",") if x.strip()]:
        q = urllib.parse.urlencode({"go": "manager", "rid": ss.rid, "who": ss.get("name", ""), "m": n, "f": fund})
        out.append(f'<a class="mgr" href="?{q}" target="_blank" rel="noopener">{esc(n)}</a>')
    return ", ".join(out) if out else "–"


def page_redirect(qp):
    """Opened in a new tab from a fund manager link: log the tap, then go to Google."""
    m, fund = qp.get("m", ""), qp.get("f", "")
    target = "https://www.google.com/search?" + urllib.parse.urlencode({"q": f'"{m}" fund manager {fund}'})
    t = now()
    writer()[0].put({"investor": None, "clicks": [{"rid": qp.get("rid", ""), "name": qp.get("who", ""), "time": t,
                     "seconds_from_start": "", "screen": "fund page", "action": ACTIONS["manager_click"],
                     "fund": fund, "detail": m, "code": "manager_click"}]})
    st.markdown(f"Opening Google search for **{esc(m)}**…  \n[Tap here if nothing happens]({target})")
    url = json.dumps(json.dumps(target)).replace("</", "<\\/")
    run_js(f"""const s = window.parent.document.createElement('script');
s.text = "window.location.replace(" + {url} + ")";
window.parent.document.body.appendChild(s);""")


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
    return f"""<div class="gauge"><svg width="150" height="86" viewBox="0 0 220 126" role="img"
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
.block-container {{max-width: 820px;}}
.st-key-scrolltop, .st-key-scrollto {{position:absolute; height:0; overflow:hidden;}}
[data-testid="stLayoutWrapper"]:has(> .st-key-backbar), .st-key-backbar {{position:-webkit-sticky; position:sticky;
  top:0; z-index:1000; background:#fff;}}
.st-key-backbar {{padding:8px 0; box-shadow:0 6px 8px -8px rgba(0,0,0,.25);}}
.st-key-back_btn button {{min-height:42px; border-radius:999px; padding:0 16px;}}
.st-key-back_btn button p {{font-size:15px; font-weight:600;}}
/* keep chosen rows side by side on phones (Streamlit stacks columns below 640px) */
.st-key-tabs [data-testid="stHorizontalBlock"], .st-key-tools [data-testid="stHorizontalBlock"],
[class*="st-key-row_"] [data-testid="stHorizontalBlock"], [class*="st-key-prow_"] [data-testid="stHorizontalBlock"],
.st-key-actions [data-testid="stHorizontalBlock"], .st-key-actions_open [data-testid="stHorizontalBlock"] {{
  flex-wrap:nowrap !important;}}
.st-key-tabs [data-testid="stColumn"], .st-key-tools [data-testid="stColumn"],
.st-key-actions [data-testid="stColumn"], .st-key-actions_open [data-testid="stColumn"] {{
  min-width:0 !important; flex:1 1 0 !important; width:auto !important;}}
[class*="st-key-row_"] [data-testid="stColumn"]:first-child, [class*="st-key-prow_"] [data-testid="stColumn"]:first-child {{
  min-width:0 !important; flex:0 0 46px !important; width:46px !important;}}
[class*="st-key-row_"] [data-testid="stColumn"]:last-child, [class*="st-key-prow_"] [data-testid="stColumn"]:last-child {{
  min-width:0 !important; flex:1 1 0 !important; width:auto !important;}}
.st-key-tabs button {{height:46px; border-radius:12px; border:1px solid #eee;}}
.st-key-tabs button p {{font-size:15px; font-weight:500;}}
.st-key-nav_{st.session_state.get('tab', 'explore')} button {{background:{t['soft']}; border-color:{t['soft']};
  color:{t['accent']};}}
.st-key-tools button {{height:40px;}}
.intro {{font-size:15px; color:#555; margin:2px 0 8px;}}
.chips {{display:flex; flex-wrap:wrap; gap:6px; align-items:center; margin:4px 0 10px;}}
.chips .pill {{margin:0;}}
.chips .count {{margin-left:auto; color:#8a8a8a; font-size:14px;}}
.stButton button[kind="tertiary"] {{text-align:left; white-space:normal; height:auto;}}
.stButton button[kind="tertiary"] p, .stButton button[kind="tertiary"] div {{white-space:normal !important;
  overflow:visible !important; text-overflow:clip !important; line-height:1.3;}}
[data-testid="stFormSubmitButton"] button {{background:{t['accent']} !important; border-color:{t['accent']} !important;
  color:#fff !important; border-radius:10px; height:46px;}}
[data-testid="stFormSubmitButton"] button p {{color:#fff !important; font-weight:600; font-size:16px;}}
a.mgr {{color:{t['accent']}; text-decoration:none; border-bottom:1px dashed {t['accent']}88;}}
.stats.four {{grid-template-columns:repeat(4,1fr);}}
.hold {{margin-top:4px;}}
.hold > div:not(.bar) {{display:flex; justify-content:space-between; gap:10px; font-size:15px; padding:3px 0;}}
.hold span {{color:#8a8a8a; font-size:13px;}}
.hold .tot {{border-top:1px solid #f0f0f0; margin-top:6px; padding-top:8px !important;}}
.hold .bar {{margin:4px 0 0;}}
.st-key-actions {{margin:6px 0 14px;}}
.st-key-actions button {{height:50px; border-radius:12px;}}
.st-key-actions button p {{font-size:17px; font-weight:600;}}
.st-key-actions button[kind="secondary"] {{border:1.5px solid {t['accent']}; color:{t['accent']}; background:#fff;}}
.st-key-actions button[kind="secondary"]:disabled {{border-color:#ddd; color:#aaa;}}
.st-key-complete_btn button {{height:48px; border-radius:12px; margin-top:-4px;}}
.st-key-complete_btn button p {{font-weight:600; font-size:16px;}}
.st-key-retbox {{background:{t['soft']}; border-radius:14px; padding:10px 14px; margin:8px 0 4px;}}
.st-key-retbox [data-testid="stHorizontalBlock"] {{flex-wrap:nowrap !important;}}
.st-key-retbox [data-testid="stColumn"] {{min-width:0 !important; flex:1 1 0 !important; width:auto !important;}}
.retl {{font-size:13px; color:#777; margin-bottom:2px;}}
.retv {{font-size:26px; font-weight:700; color:{t['accent']}; text-align:right;}}
.retn {{font-size:12px; color:#8a8a8a; text-align:right;}}
.htitle {{font-size:17px; font-weight:600; color:#222; margin:22px 0 8px;}}
.hcard {{border:1px solid #eee; border-radius:14px; padding:12px 16px 4px;}}
.hb-t {{display:flex; justify-content:space-between; gap:12px; font-size:14px; color:#222;}}
.hb-t small {{display:block; color:#8a8a8a; font-size:12px;}}
.hb .bar {{height:6px; margin:4px 0 10px;}}
.hsub {{font-size:12px; color:#8a8a8a; margin:6px 0 6px; text-transform:uppercase; letter-spacing:.03em;}}
.qhead h3 {{margin:0 0 4px;}} .qhead p {{color:#666; font-size:15px;}}
.qt {{font-size:16px; font-weight:500; color:#222; line-height:1.4; margin-bottom:6px;}}
.qn {{display:inline-block; min-width:24px; height:24px; border-radius:50%; background:{t['soft']}; color:{t['accent']};
  font-size:13px; text-align:center; line-height:24px; margin-right:8px; font-weight:600;}}
.qmiss {{color:#C0392B; font-size:13px; font-weight:600; margin-left:6px;}}
.qsub {{font-size:14px; color:#555; margin:6px 0 2px;}}
.st-key-survey_done_btn button {{height:50px; border-radius:12px;}}
.st-key-survey_done_btn button p {{font-weight:600; font-size:17px;}}
.sipcalc {{font-size:14px; color:#555; margin:-6px 0 10px;}}
.sipcalc.over {{color:#C0392B;}}
.st-key-go_btn button {{height:48px; border-radius:10px;}}
.st-key-go_btn button p {{font-weight:600; font-size:16px;}}
.chg {{font-size:20px; font-weight:600; margin:6px 0 -6px;}}
.chg span {{font-size:13px; font-weight:400; color:#8a8a8a;}}
.thanks {{text-align:center; padding:12vh 8px 24px;}}
.thanks .tick {{width:64px; height:64px; border-radius:50%; background:{t['accent']}; color:#fff; font-size:34px;
  line-height:64px; margin:0 auto 18px;}}
.thanks h2 {{font-size:26px; margin:0 0 8px;}}
.thanks p {{color:#777;}}
iframe[title="st_keyup.st_keyup"] {{margin-bottom:-6px;}}
@media (max-width: 640px) {{
  .block-container {{padding: 1.2rem 1rem 3rem !important;}}
  .stats {{grid-template-columns:1fr 1fr;}}
  .stats.four {{grid-template-columns:1fr 1fr;}}
  .stats div {{padding:10px 12px; font-size:14px;}}
  .rets {{gap:8px;}} .rets div {{padding:8px 10px;}} .rets b {{font-size:17px;}} .rets span {{font-size:11px;}}
  .wallet {{padding:16px 18px;}} .wallet .big {{font-size:28px;}} .wallet .row {{gap:18px;}}
  .fname {{font-size:16px;}}
  .fmeta {{font-size:13px; gap:10px;}}
  .gauge {{gap:14px;}}
}}
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
    if "display_order" in df.columns:
        df = df.sort_values("display_order", kind="stable").reset_index(drop=True)
    df["fid"] = [f"f{i:03d}" for i in range(len(df))]
    return df, None


@st.cache_data
def load_nav(path: str):
    """nav_history.csv with columns fund_name, date, nav (one row per day)."""
    try:
        nav = pd.read_csv(path)
    except Exception:
        return None
    nav["date"] = pd.to_datetime(nav["date"], errors="coerce")
    nav["nav"] = pd.to_numeric(nav["nav"], errors="coerce")
    nav = nav.dropna(subset=["date", "nav"]).sort_values("date")
    return {k: g[["date", "nav"]].reset_index(drop=True) for k, g in nav.groupby("fund_name")}


@st.cache_data
def load_holdings(path: str):
    """holdings.xlsx: sheet 'holdings' (top 5 per fund) and sheet 'sectors' (sector split per fund)."""
    try:
        h = pd.read_excel(path, sheet_name="holdings")
        s = pd.read_excel(path, sheet_name="sectors")
    except Exception:
        return {}, {}
    s["asset_class"] = s["asset_class"].fillna("")
    return ({k: g.sort_values("rank") for k, g in h.groupby("fund_name")},
            {k: g for k, g in s.groupby("fund_name")})


def bars(rows, label_col, sub_col=None):
    out = []
    for _, r in rows.iterrows():
        sub = f'<small>{esc(r[sub_col])}</small>' if sub_col and str(r.get(sub_col, "")) not in ("", "nan") else ""
        w = max(float(r["weight"]), 0) * 100
        out.append(f'<div class="hb"><div class="hb-t"><span>{esc(r[label_col])}{sub}</span><b>{w:.1f}%</b></div>'
                   f'<div class="bar"><div style="width:{min(w, 100):.1f}%"></div></div></div>')
    return "".join(out)


def holdings_section(f):
    hold, sect = load_holdings(str(HOLDINGS_FILE)) if HOLDINGS_FILE.exists() else ({}, {})
    name = f["fund_name"]
    if name not in hold and name not in sect:
        return
    if name in sect:
        g = sect[name]
        html = ""
        classes = [a for a in dict.fromkeys(g["asset_class"]) if a]
        if len(classes) > 1:
            for a in classes:
                part = g[g["asset_class"] == a].sort_values("weight", ascending=False)
                html += f'<div class="hsub">Within the {esc(a.lower())} part of the fund</div>' + bars(part, "sector")
        else:
            html = bars(g.sort_values("weight", ascending=False), "sector")
        st.markdown(f'<div class="htitle">Sector allocation</div><div class="hcard">{html}</div>',
                    unsafe_allow_html=True)
    if name in hold:
        st.markdown(f'<div class="htitle">Top 5 holdings</div>'
                    f'<div class="hcard">{bars(hold[name], "holding", "instrument")}</div>', unsafe_allow_html=True)


def data_as_of():
    nav = load_nav(str(NAV_FILE)) if NAV_FILE.exists() else None
    if not nav:
        return None
    last = pd.Series([g["date"].iloc[-1] for g in nav.values()]).mode().iloc[0]
    return f"{last:%d %b %Y}"


PERIODS = {"1M": 1, "6M": 6, "1Y": 12, "3Y": 36, "5Y": 60, "All": None}


def nav_chart(f):
    nav = load_nav(str(NAV_FILE)) if NAV_FILE.exists() else None
    h = nav.get(f["fund_name"]) if nav else None
    if h is None or len(h) < 2:
        return
    import altair as alt
    span = (h["date"].iloc[-1] - h["date"].iloc[0]).days / 30.4
    options = [p for p, m in PERIODS.items() if m is None or m <= span + 0.5]
    key = f"chart_{f['fund_name']}"
    if st.session_state.get(key) not in options:
        st.session_state[key] = "3Y" if "3Y" in options else "All"
    sel = st.session_state[key]
    months = PERIODS[sel]
    view = h if months is None else h[h["date"] >= h["date"].iloc[-1] - pd.DateOffset(months=months)]
    up = view["nav"].iloc[-1] >= view["nav"].iloc[0]
    colour = "#2E9E5B" if up else "#D64545"
    chg = (view["nav"].iloc[-1] / view["nav"].iloc[0] - 1) * 100
    label = "since launch" if sel == "All" else f"in {sel}"
    st.markdown(f'<div class="chg" style="color:{colour}">{chg:+.2f}% <span>{label}</span></div>',
                unsafe_allow_html=True)
    lo, hi = view["nav"].min(), view["nav"].max()
    pad = (hi - lo) * 0.08 or hi * 0.01
    base = alt.Chart(view).encode(
        x=alt.X("date:T", title=None, axis=alt.Axis(grid=False, labelColor="#9a9a9a", tickCount=4, format="%b %y")),
        y=alt.Y("nav:Q", title=None, scale=alt.Scale(domain=[lo - pad, hi + pad], zero=False),
                axis=alt.Axis(labelColor="#9a9a9a", gridColor="#f2f2f2", tickCount=4)),
        tooltip=[alt.Tooltip("date:T", title="Date", format="%d %b %Y"), alt.Tooltip("nav:Q", title="NAV", format=",.2f")])
    chart = (base.mark_area(color=colour, opacity=0.08, clip=True) + base.mark_line(color=colour, strokeWidth=2))
    st.altair_chart(chart.properties(height=210).configure_view(strokeWidth=0), width="stretch")
    with st.container(key="periods"):
        st.segmented_control("Period", options, key=key, label_visibility="collapsed",
                             on_change=lambda: log("chart_period", f["fund_name"], st.session_state.get(key)))


# ---- resume after a refresh or a locked phone
# Streamlit forgets a session about 2 minutes after the phone sleeps or the page is refreshed. Each person's
# progress is kept here (in server memory) under the ID in their page address (?r=...), so they carry on where
# they left off instead of starting again.
KEEP = ["rid", "step", "alloc", "events", "events_saved", "view", "tab", "mode", "started_at", "name", "app", "age",
        "location", "status", "order", "completed_at", "survey_started", "nav_count", "survey_missing"]
KEEP_HOURS = 12


@st.cache_resource
def saved_sessions():
    return {}


def remember():
    ss = st.session_state
    if ss.get("step", "start") == "start":
        return
    snap = {k: ss[k] for k in KEEP if k in ss}
    snap.update({k: ss[k] for k in list(ss.keys()) if str(k).startswith("ans_")})
    store = saved_sessions()
    store[ss.rid] = (time.time(), snap)
    if len(store) > 50 and random.random() < 0.05:          # drop old entries now and then
        cutoff = time.time() - KEEP_HOURS * 3600
        for rid in [r for r, (t, _) in list(store.items()) if t < cutoff]:
            store.pop(rid, None)


def init_state():
    ss = st.session_state
    if "rid" not in ss:
        r = st.query_params.get("r")
        saved = saved_sessions().get(r) if r else None
        if saved:
            import copy
            for k, v in copy.deepcopy(saved[1]).items():
                ss[k] = v
            log("resumed", value=f"after a refresh or the phone sleeping, on {ss.step}")
            return
        ss.rid = "R-" + uuid.uuid4().hex[:8].upper()
        ss.step = "start"
        ss.alloc = {}
        ss.events = []
        ss.view = None
        ss.tab = "explore"
        ss.mode = None
        ss.started_at = now()


ACTIONS = {  # event code -> readable action name in the database and dashboard
    "start": "Tapped Continue (started)", "resumed": "Came back (page refreshed or phone slept)", "start_invalid": "Tapped Continue (name or app missing)",
    "list_order": "Fund list order shown", "tab": "Tapped tab", "search": "Searched",
    "sort": "Changed sort", "filter_reset": "Tapped Reset filters", "open_fund": "Opened fund",
    "close_fund": "Tapped Back", "chart_period": "Changed chart period", "choose_mode": "Tapped Buy/SIP button",
    "buy": "Bought (lump sum)", "sip": "Started/added SIP", "buy_failed": "Buy/SIP not accepted",
    "manager_click": "Tapped fund manager name", "return_period": "Changed CAGR period", "use_max": "Tapped Use maximum amount", "complete_open": "Tapped Completed",
    "complete_cancel": "Tapped Keep investing", "complete": "Submitted portfolio",
"survey_incomplete": "Tapped Done (questions unanswered)",
    "survey_done": "Tapped Done (survey completed)",
}


def screen():
    ss = st.session_state
    if ss.get("step") == "start":
        return "start"
    if ss.get("step") == "done":
        return "thank you"
    if ss.get("step") == "survey":
        return "questions"
    return "fund page" if ss.get("view") else ("portfolio" if ss.get("tab") == "portfolio" else "fund list")


def log(event, fund="", value=""):
    ss = st.session_state
    label = ACTIONS.get(event, event.replace("filter_", "Changed filter: "))
    secs = int((datetime.fromisoformat(now()) - datetime.fromisoformat(ss.started_at)).total_seconds())
    ss.events.append({"rid": ss.rid, "name": ss.get("name", ""), "time": now(), "seconds_from_start": secs,
                      "screen": screen(), "action": label, "fund": fund, "detail": str(value), "code": event})
    if len(ss.events) - ss.get("events_saved", 0) >= FLUSH_EVERY and ss.get("step") == "invest":
        flush()


def committed(h):
    return 0 if not h else h.get("buy", 0) + h.get("sip", 0) * SIP_MONTHS


def total_committed(exclude=None):
    return sum(committed(h) for k, h in st.session_state.alloc.items() if k != exclude)


def open_fund(fid, name, position):
    ss = st.session_state
    log("open_fund", name, f"{ss.tab}:{position}")
    ss.view = fid
    ss.mode = None
    ss.nav_count = ss.get("nav_count", 0) + 1


def set_tab(tab):
    log("tab", value=tab)
    st.session_state.tab = tab
    st.session_state.view = None
    st.session_state.nav_count = st.session_state.get("nav_count", 0) + 1


def set_mode(mode, name):
    log("choose_mode", name, mode)
    st.session_state.mode = mode
    st.session_state.mode_count = st.session_state.get("mode_count", 0) + 1


def close_fund(name, tapped=True):
    if tapped:
        log("close_fund", name)
    st.session_state.view = None
    st.session_state.mode = None
    st.session_state.nav_count = st.session_state.get("nav_count", 0) + 1


# ---------------------------------------------------------------- pages
def page_start():
    base_css(APPS["Other / I don't invest yet"])
    _, mid, _ = st.columns([1, 1.6, 1])
    with mid:
        st.markdown("### Invest ₹1,00,000")
        st.write("A short investing exercise for a GIM research study. It takes about 5 minutes. "
                 "No real money is involved and nothing is bought.")
        name = st.text_input("Your name", key="in_name", placeholder="First name is fine")
        age = st.selectbox("Your age", AGES, index=None, placeholder="Select your age range", key="in_age")
        city = st.selectbox("Your location", CITIES, index=None, placeholder="Select your city", key="in_city")
        other = ""
        if city == "Other":
            other = st.text_input("Your city", key="in_city_other", placeholder="Type your city")
        app = st.selectbox("Which app do you mainly use to invest in mutual funds?", list(APPS.keys()),
                           index=None, placeholder="Select an app", key="in_app")
        if st.button("Continue", type="primary", width="stretch"):
            location = (other or "").strip() if city == "Other" else city
            if not (name or "").strip() or not age or not location or not app:
                log("start_invalid", value=f"name={'yes' if (name or '').strip() else 'no'}, age={age or 'none'}, "
                                           f"location={location or 'none'}, app={app or 'none'}")
                st.error("Please fill in your name, age, location and app.")
                return
            ss = st.session_state
            ss.name, ss.app, ss.age, ss.location = name.strip(), app, age, location
            for e in ss.events:          # earlier taps on this page were logged before the name was known
                e["name"] = ss.name
            log("start", value=app)
            ss.step = "invest"
            st.query_params["r"] = ss.rid      # lets the app resume this person after a refresh
            flush()
            st.rerun()


def topbar():
    ss = st.session_state
    st.markdown('<div class="topbar"><div class="dot">₹</div><b>Mutual funds</b></div>', unsafe_allow_html=True)


def wallet():
    ss = st.session_state
    invested = total_committed()
    n = sum(1 for h in ss.alloc.values() if h)
    st.markdown(f"""<div class="wallet"><div class="l">Hi {esc(ss.name)}, cash available to invest</div>
<div class="big">₹{inr(TOTAL - invested)}</div><div class="row">
<div><span>Starting amount</span><b>₹{inr(TOTAL)}</b></div>
<div><span>Invested amount</span><b>₹{inr(invested)}</b></div>
<div><span>Funds chosen</span><b>{n}</b></div></div></div>""", unsafe_allow_html=True)
    st.button("Completed", key="complete_btn", type="primary", width="stretch", icon=":material/check_circle:",
              on_click=lambda: (log("complete_open"), st.session_state.update(confirm=True)))


def page_survey():
    ss = st.session_state
    base_css(APPS[ss.app])
    scroll_top()
    st.markdown("""<div class="qhead"><h3>A few quick questions</h3>
<p>Your portfolio is saved. Please answer every question, then tap Done. There are no right or wrong answers
except where you're asked to pick the correct meaning.</p></div>""", unsafe_allow_html=True)
    missing = ss.get("survey_missing", set())
    for i, q in enumerate(QUESTIONS, 1):
        with st.container(border=True, key=f"q_{q['id']}"):
            flag = ' <span class="qmiss">Please answer</span>' if q["id"] in missing else ""
            st.markdown(f'<div class="qt"><span class="qn">{i}</span>{esc(q["text"])}{flag}</div>',
                        unsafe_allow_html=True)
            k = f"ans_{q['id']}"
            if q["type"] == "single":
                st.radio(q["text"], q["options"], index=None, key=k, label_visibility="collapsed")
                if q.get("other") and ss.get(k) == q["other"]:
                    st.text_input("Please specify", key=k + "_other", placeholder="Type your answer",
                                  label_visibility="collapsed")
            elif q["type"] == "multi":
                if q.get("max"):
                    st.markdown(f'<div class="qsub">{len(ss.get(k) or [])}/{q["max"]} picked</div>',
                                unsafe_allow_html=True)
                st.pills(q["text"], q["options"], selection_mode="multi", key=k, label_visibility="collapsed",
                         **({"on_change": limit_picks, "args": (k, q["max"])} if q.get("max") else {}))
                if q.get("only") and q["only"] in (ss.get(k) or []) and len(ss.get(k)) > 1:
                    st.caption(f"“{q['only']}” can't be combined with other answers.")
            else:
                st.markdown(f'<div class="qsub">Pick the <b>{q["top"]} that matter most</b> '
                            f'({len(ss.get(k + "_top") or [])}/{q["top"]} picked)</div>',
                            unsafe_allow_html=True)
                top = st.pills("Most", q["options"], selection_mode="multi", key=k + "_top",
                               label_visibility="collapsed", on_change=limit_picks, args=(k + "_top", q["top"]))
                st.markdown(f'<div class="qsub">Pick the <b>{q["bottom"]} that matter least</b> '
                            f'({len(ss.get(k + "_bottom") or [])}/{q["bottom"]} picked)</div>',
                            unsafe_allow_html=True)
                if ss.get(k + "_bottom"):          # drop anything that has since been picked as "most"
                    ss[k + "_bottom"] = [o for o in ss[k + "_bottom"] if o not in (top or [])]
                st.pills("Least", [o for o in q["options"] if o not in (top or [])], selection_mode="multi",
                         key=k + "_bottom", label_visibility="collapsed",
                         on_change=limit_picks, args=(k + "_bottom", q["bottom"]))
    if st.button("Done", type="primary", width="stretch", key="survey_done_btn", icon=":material/task_alt:"):
        problems, row = check_answers()
        if problems:
            ss.survey_missing = set(problems)
            log("survey_incomplete", value=", ".join(problems))
            st.rerun()
        log("survey_done")
        ss.status, ss.completed_at, ss.step = "completed", now(), "done"
        flush(answers=row)
        st.rerun()
    if missing:
        st.error(f"Please answer {len(missing)} more question{'s' if len(missing) > 1 else ''} (marked in red).")


def limit_picks(key, n):
    """Keep a chip question to at most n picks: an extra tap is undone."""
    ss = st.session_state
    v = ss.get(key) or []
    if len(v) > n:
        ss[key] = ss.get(key + "_prev", v[:n])
        st.toast(f"You can pick only {n}. Unselect one first to change your choice.")
    ss[key + "_prev"] = list(ss.get(key) or [])


def check_answers():
    """Returns (ids of unanswered or invalid questions, answer rows for the database: one row per question)."""
    ss = st.session_state
    problems, rows, t = [], [], now()
    base = {"rid": ss.rid, "name": ss.name, "app": ss.app, "submitted_at": t}
    for i, q in enumerate(QUESTIONS, 1):
        k = f"ans_{q['id']}"
        if q["type"] == "single":
            v = ss.get(k)
            if q.get("other") and v == q["other"]:
                typed = (ss.get(k + "_other") or "").strip()
                v = f"{q['other']}: {typed}" if typed else None      # "Others" needs the typed answer
            if not v:
                problems.append(q["id"])
            rows.append({**base, "question_no": i, "question": q["text"], "answer": v or "",
                         "correct": ("yes" if v == q["correct"] else "no") if q.get("correct") else ""})
        elif q["type"] == "multi":
            v = ss.get(k) or []
            if not v or (q.get("only") in v and len(v) > 1) or (q.get("max") and len(v) > q["max"]):
                problems.append(q["id"])
            rows.append({**base, "question_no": i, "question": q["text"], "answer": " | ".join(v), "correct": ""})
        else:
            top = ss.get(k + "_top") or []
            bottom = [b for b in (ss.get(k + "_bottom") or []) if b not in top]
            if len(top) != q["top"] or len(bottom) != q["bottom"]:
                problems.append(q["id"])
            rows.append({**base, "question_no": f"{i}a", "question": f"{q['text']} ({q['top']} that matter most)",
                         "answer": " | ".join(top), "correct": ""})
            rows.append({**base, "question_no": f"{i}b", "question": f"{q['text']} ({q['bottom']} that matter least)",
                         "answer": " | ".join(bottom), "correct": ""})
    return problems, rows


@st.dialog("Submit your portfolio?")
def confirm_dialog(df):
    ss = st.session_state
    held = [h for h in ss.alloc.values() if h]
    invested = total_committed()
    if held:
        st.write(f"You've invested **₹{inr(invested)}** across **{len(held)} fund{'s' if len(held) > 1 else ''}**, "
                 f"with **₹{inr(TOTAL - invested)}** left as cash.")
    else:
        st.write("You haven't invested in any fund yet. Your whole ₹1,00,000 will stay as cash.")
    st.write("You won't be able to change it after this. A few quick questions come next.")
    c1, c2 = st.columns(2)
    if c1.button("Keep investing", width="stretch"):
        log("complete_cancel")
        ss.confirm = False
        st.rerun()
    if c2.button("Submit", type="primary", width="stretch"):
        log("complete", value=f"invested ₹{invested}")
        ss.status, ss.step, ss.confirm = "portfolio submitted", "survey", False
        ss.survey_started = now()
        ss.nav_count = ss.get("nav_count", 0) + 1
        flush(portfolio=final_portfolio(df))
        st.rerun()


def page_thanks():
    base_css(APPS[st.session_state.get("app") or "Other / I don't invest yet"])
    st.markdown("""<div class="thanks"><div class="tick">✓</div>
<h2>Thanks for being part of our survey</h2>
<p>Your portfolio and answers have been recorded. You can close this page now.</p></div>""", unsafe_allow_html=True)
    form_url, id_field = secret("form_url"), secret("form_id_field")
    if form_url:
        rid = st.session_state.rid
        link = f"{form_url}?usp=pp_url&{id_field}={rid}" if id_field else form_url
        st.link_button("One last step: a 3-minute questionnaire", link, type="primary", width="stretch")


def tabs():
    n = sum(1 for h in st.session_state.alloc.values() if h)
    with st.container(key="tabs"):
        c1, c2 = st.columns(2, gap="small")
        c1.button("Mutual funds", key="nav_explore", icon=":material/work:", width="stretch",
                  on_click=set_tab, args=("explore",))
        c2.button(f"Portfolio ({n})" if n else "Portfolio", key="nav_portfolio", icon=":material/pie_chart:",
                  width="stretch", on_click=set_tab, args=("portfolio",))


def holding_text(h):
    if not h:
        return ""
    parts = []
    if h.get("buy"):
        parts.append(f"Bought ₹{inr(h['buy'])}")
    if h.get("sip"):
        parts.append(f"SIP ₹{inr(h['sip'])}/month")
    return " · ".join(parts)


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
    st.button("Reset filters", on_click=reset_filters, width="stretch")


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
    st.markdown('<p class="intro">Treat this as your own money. Tap any fund to see its details and invest. '
                "Anything you don't invest stays as cash.</p>", unsafe_allow_html=True)
    from st_keyup import st_keyup
    q_now = st_keyup("Search", key="q", placeholder="Search funds or fund houses", label_visibility="collapsed",
                     debounce=250)
    if (q_now or "") != ss.get("q_logged", ""):
        ss.q_logged = q_now or ""
        if q_now:
            log("search", value=q_now)
    view = df.set_index("fid").loc[ss.order].reset_index()
    view["position"] = range(1, len(view) + 1)
    view, active = apply_filters(view)
    with st.container(key="tools"):
        c1, c2 = st.columns([1, 1.3], gap="small")
        with c1.popover(f"Filters ({len(active)})" if active else "Filters", icon=":material/tune:",
                        width="stretch"):
            filter_panel(df)
        c2.selectbox("Sort", list(SORTS), key="srt", label_visibility="collapsed",
                     on_change=lambda: log("sort", value=st.session_state.srt))
    if q_now:
        q = q_now.lower().strip()
        view = view[view["fund_name"].str.lower().str.contains(q, regex=False)
                    | view["amc"].str.lower().str.contains(q, regex=False)]
    col, asc = SORTS[ss.srt]
    if col:
        view = view.sort_values(col, ascending=asc, na_position="last")

    chips = "".join(f'<span class="pill">{esc(x)}</span>' for x in active)
    st.markdown(f'<div class="chips">{chips}<span class="count">{len(view)} funds</span></div>',
                unsafe_allow_html=True)
    st.session_state.setdefault("srt", "Default")
    if view.empty:
        st.info("No funds match. Try another filter or search.")
    for _, f in view.iterrows():
        h = ss.alloc.get(f["fid"])
        with st.container(border=True, key=f"row_{f['fid']}"):
            c0, c1 = st.columns([1, 6], vertical_alignment="center", gap="small")
            c0.markdown(logo(f["amc"]), unsafe_allow_html=True)
            with c1:
                st.button(f["fund_name"], key=f"open_{f['fid']}", type="tertiary",
                          on_click=open_fund, args=(f["fid"], f["fund_name"], int(f["position"])))
                aum = "–" if pd.isna(f["aum_cr"]) else f"₹{inr(f['aum_cr'], 2)} Cr."
                badge = f'<span class="inv">{holding_text(h)}</span>' if h else ""
                st.markdown(f'<div class="fmeta" style="margin-top:-6px"><span>{esc(f["category"])} · '
                            f'{esc(f["sub_category"])}</span><span>AUM <b>{aum}</b></span>{badge}</div>',
                            unsafe_allow_html=True)


def page_portfolio(df):
    ss = st.session_state
    held = [(fid, h) for fid, h in ss.alloc.items() if h]
    if not held:
        st.markdown("#### Your portfolio")
        st.info("You haven't invested in any fund yet. Everything is currently cash.")
        st.button("Explore mutual funds", type="primary", on_click=set_tab, args=("explore",))
        return
    lump = sum(h.get("buy", 0) for _, h in held)
    sip = sum(h.get("sip", 0) * SIP_MONTHS for _, h in held)
    st.markdown(f"""<div class="stats four" style="margin-bottom:18px">
<div><span>Lump sum</span>₹{inr(lump)}</div><div><span>SIPs ({SIP_MONTHS} mo.)</span>₹{inr(sip)}</div>
<div><span>Cash</span>₹{inr(TOTAL - lump - sip)}</div><div><span>Funds</span>{len(held)}</div></div>""",
                unsafe_allow_html=True)
    fx = df.set_index("fid")
    for fid, h in sorted(held, key=lambda x: -committed(x[1])):
        f = fx.loc[fid]
        with st.container(border=True, key=f"prow_{fid}"):
            c0, c1 = st.columns([1, 6], vertical_alignment="center", gap="small")
            c0.markdown(logo(f["amc"], 40), unsafe_allow_html=True)
            with c1:
                st.button(f["fund_name"], key=f"pf_{fid}", type="tertiary", on_click=open_fund,
                          args=(fid, f["fund_name"], ss.order.index(fid) + 1))
                st.markdown(f'<div class="fmeta" style="margin-top:-6px"><span>{esc(f["category"])} · '
                            f'{esc(f["sub_category"])}</span></div>', unsafe_allow_html=True)
            lines = []
            if h.get("buy"):
                lines.append(f'<div><span>Lump sum · {h["buy"] / f["nav"]:,.3f} units @ ₹{inr(f["nav"], 2)}</span>'
                             f'<b>₹{inr(h["buy"])}</b></div>')
            if h.get("sip"):
                lines.append(f'<div><span>SIP · ₹{inr(h["sip"])}/month × {SIP_MONTHS}</span>'
                             f'<b>₹{inr(h["sip"] * SIP_MONTHS)}</b></div>')
            share = committed(h) / TOTAL
            st.markdown(f'<div class="hold">{"".join(lines)}<div class="tot"><span>Allocation</span>'
                        f'<b>{share:.0%}</b></div><div class="bar"><div style="width:{share * 100:.1f}%"></div>'
                        f'</div></div>', unsafe_allow_html=True)
    st.caption("Your portfolio is saved automatically after every change.")


def page_fund(df):
    ss = st.session_state
    fid = ss.view
    f = df.set_index("fid").loc[fid]
    name = f["fund_name"]
    h = ss.alloc.get(fid)
    room = TOTAL - total_committed()
    sip_ok = not pd.isna(f["min_sip"])

    with st.container(key="backbar"):      # pinned to the top of the screen, so Back is always reachable
        st.button("← Back to " + ("portfolio" if ss.tab == "portfolio" else "all funds"), key="back_btn",
                  on_click=close_fund, args=(name,))
    aum = "–" if pd.isna(f["aum_cr"]) else f"₹{inr(f['aum_cr'], 2)} Cr."
    st.markdown(f'<div class="fhead" style="margin:8px 0 6px">{logo(f["amc"], 60)}<div>'
                f'<div class="fname" style="font-size:22px">{esc(name)}</div>'
                f'<div class="fmeta"><span>{esc(f["amc"])}</span><span>{esc(f["plan"])}</span>'
                f'<span>{esc(f["category"])} · {esc(f["sub_category"])}</span></div></div></div>',
                unsafe_allow_html=True)
    nav_chart(f)
    periods = {"1 Year": "return_1y", "3 Years": "return_3y", "5 Years": "return_5y"}
    rk = f"ret_{fid}"
    with st.container(key="retbox"):
        c1, c2 = st.columns([1, 1.3], vertical_alignment="center", gap="small")
        with c1:
            st.markdown('<div class="retl">CAGR</div>', unsafe_allow_html=True)
            per = st.selectbox("CAGR period", list(periods), index=1, key=rk, label_visibility="collapsed",
                               on_change=lambda: log("return_period", name, st.session_state.get(rk)))
        v = f[periods[per]]
        c2.markdown(f'<div class="retv">{pct(v)}</div>' + ("" if not pd.isna(v) else
                    '<div class="retn">Fund is too new for this period</div>'), unsafe_allow_html=True)
    st.markdown(risk_gauge(f["risk"]), unsafe_allow_html=True)

    # ---- Buy / SIP, placed right under the risk meter so it is on the first screen of a phone
    with st.container(key="actions"):
        if h:
            st.markdown(f'<div class="inv" style="font-size:14px;margin-bottom:8px">You have: {holding_text(h)}</div>',
                        unsafe_allow_html=True)
        c1, c2 = st.columns(2, gap="small")
        c1.button("Buy", key="mode_buy", width="stretch", icon=":material/shopping_bag:",
                  type="primary" if ss.mode != "sip" else "secondary", on_click=set_mode, args=("buy", name))
        c2.button("SIP", key="mode_sip", width="stretch", disabled=not sip_ok, icon=":material/event_repeat:",
                  type="primary" if ss.mode == "sip" else "secondary", on_click=set_mode, args=("sip", name))
        if not sip_ok:
            st.caption("SIP is not available for this fund.")
        if ss.mode:
            sip = ss.mode == "sip"
            with st.container(border=True, key="buybox"):
                has = h.get(ss.mode, 0) if h else 0
                if sip:
                    title = "Add to your monthly SIP" if has else "Start a monthly SIP"
                    st.markdown(f"**{title}**  \nRuns for {SIP_MONTHS} months. "
                                f"You can commit up to ₹{inr(room)} in total, so at most ₹{inr(room // SIP_MONTHS)}/month.")
                else:
                    title = "Buy more" if has else "One-time purchase"
                    st.markdown(f"**{title}**  \nYou can invest up to ₹{inr(room)} here.")
                akey = f"amt_{fid}_{ss.mode}"
                cap = room // SIP_MONTHS if sip else room
                if ss.get("fill_max") == akey:          # "Use maximum" was tapped
                    ss[akey] = int(cap)
                    ss.fill_max = None
                amt = st.number_input("Monthly SIP amount (₹)" if sip else "Amount (₹)", min_value=0,
                                      max_value=TOTAL, step=100 if sip else STEP, key=akey)
                if sip and amt:
                    over = amt * SIP_MONTHS > room
                    st.markdown(f'<div class="sipcalc{" over" if over else ""}">₹{inr(amt)}/month × {SIP_MONTHS} months = '
                                f'<b>₹{inr(amt * SIP_MONTHS)}</b> total</div>', unsafe_allow_html=True)
                go = st.button(("Add to SIP" if has else "Start SIP") if sip else "Buy now", key="go_btn",
                               type="primary", width="stretch")
                if go:
                    amt = int(amt)
                    minimum = f["min_sip"] if sip else f["min_lumpsum"]
                    total = amt * (SIP_MONTHS if sip else 1)
                    problem = None
                    if amt <= 0:
                        problem = "Enter an amount."
                    elif total > room and sip:
                        problem = (f"A ₹{inr(amt)} monthly SIP adds up to ₹{inr(total)} over {SIP_MONTHS} months, but you "
                                   f"have ₹{inr(room)} left. The most you can put in a monthly SIP is ₹{inr(cap)}.")
                    elif total > room:
                        problem = f"You have ₹{inr(room)} left, so you can't invest ₹{inr(amt)}."
                    elif minimum and amt < minimum and not has:
                        problem = f"Minimum {'SIP' if sip else 'investment'} for this fund is ₹{inr(minimum)}."
                    if problem:
                        log("buy_failed", name, f"{ss.mode} ₹{amt}: {problem}")
                        st.error(problem)
                        if total > room and cap > 0 and (has or not minimum or cap >= minimum):
                            st.button(f"Use ₹{inr(cap)}{'/month' if sip else ''} instead", key="use_max",
                                      on_click=lambda k=akey: (log("use_max", name, k),
                                                               st.session_state.update(fill_max=k)))
                    else:
                        new = dict(h or {})
                        new[ss.mode] = new.get(ss.mode, 0) + amt
                        ss.alloc[fid] = new
                        log("sip" if sip else "buy", name, amt)
                        row = f.copy()
                        row["fid"] = fid
                        flush(purchases=[purchase_row(row, ss.mode, amt)])
                        close_fund(name, tapped=False)
                        st.rerun()
        if ss.mode and ss.get("mode_count"):
            scroll_to("buybox", ss.mode_count)

    since = "–" if pd.isna(f["launch_year"]) else esc(int(f["launch_year"]))
    min_sip = "–" if pd.isna(f["min_sip"]) or not f["min_sip"] else f"₹{inr(f['min_sip'])}"
    min_lump = "–" if not f["min_lumpsum"] else f"₹{inr(f['min_lumpsum'])}"
    st.markdown(f"""<div class="stats">
<div><span>NAV</span>₹{inr(f['nav'], 2)}</div>
<div><span>Fund size (AUM)</span>{aum}</div>
<div><span>Expense ratio</span>{pct(f['expense_ratio'], 2)}</div>
<div><span>Min. lump sum</span>{min_lump}</div>
<div><span>Min. SIP</span>{min_sip}</div>
<div><span>Launched</span>{since}</div>
<div style="grid-column:1 / -1"><span>Exit load</span>{esc(f['exit_load'])}</div>
<div style="grid-column:1 / -1;border-bottom:none"><span>Fund manager</span>{manager_links(f['fund_manager'], name)}</div>
</div>""", unsafe_allow_html=True)
    holdings_section(f)
    asof = data_as_of()
    if asof:
        st.markdown(f'<div class="note">Data as of {asof}</div>', unsafe_allow_html=True)


# ---------------------------------------------------------------- database
TABLES = {
    "investors": ["rid", "name", "app", "status", "started_at", "last_active", "completed_at", "minutes_taken",
                  "invested", "cash", "n_funds", "n_sips", "capitalmind_amount", "funds_opened", "taps",
                  "age", "location"],
    "purchases": ["rid", "name", "time", "fund_name", "amc", "category", "sub_category", "type", "amount",
                  "sip_monthly", "committed", "list_position"],
    "portfolio": ["rid", "name", "submitted_at", "fund_name", "amc", "category", "sub_category", "risk",
                  "lump_sum", "sip_monthly", "committed", "share", "list_position"],
    "clicks": ["rid", "name", "time", "seconds_from_start", "screen", "action", "fund", "detail", "code"],
    "answers": ["rid", "name", "app", "submitted_at", "question_no", "question", "answer", "correct"],
}


def use_sheets():
    return bool(secret("gcp_service_account") and secret("sheet_id"))


def _sheet():
    import gspread
    from google.oauth2.service_account import Credentials
    creds = Credentials.from_service_account_info(
        dict(secret("gcp_service_account")), scopes=["https://www.googleapis.com/auth/spreadsheets"])
    return gspread.authorize(creds, http_client=gspread.http_client.BackOffHTTPClient).open_by_key(secret("sheet_id"))


def _write(batch, mem):
    """Runs in the background writer. batch = {"investors": {rid: row}, "purchases": [...], "portfolio": [...],
    "clicks": [...], "answers": [...]}, already merged from many people's saves. Each table is removed from the
    batch once it has been written, so a retry after an error never writes the same rows twice."""
    if use_sheets():
        import gspread
        if "sh" not in mem:
            mem["sh"] = _sheet()
        sh = mem["sh"]

        def ws(name):
            key = "ws_" + name
            if key not in mem:
                try:
                    w = sh.worksheet(name)
                except gspread.WorksheetNotFound:
                    w = sh.add_worksheet(name, rows=1000, cols=len(TABLES[name]))
                head = w.row_values(1)
                if head and head != TABLES[name][:len(head)]:       # old layout: keep it, start a new tab
                    w.update_title(f"{name}_old_{datetime.now(IST):%d%b_%H%M}")
                    w = sh.add_worksheet(name, rows=1000, cols=len(TABLES[name]))
                    head = []
                if head != TABLES[name]:
                    if w.col_count < len(TABLES[name]):
                        w.add_cols(len(TABLES[name]) - w.col_count)
                    w.update(values=[TABLES[name]], range_name="A1")
                mem[key] = w
            return mem[key]

        for name in ("purchases", "portfolio", "clicks", "answers"):
            if batch.get(name):
                ws(name).append_rows([[str(r.get(c, "")) for c in TABLES[name]] for r in batch[name]],
                                     value_input_option="RAW")
            batch.pop(name, None)
        invs = batch.get("investors") or {}
        if not invs:
            return
        w = ws("investors")
        if mem.get("rows") is None:                    # learn where each person's row is (one read)
            col = w.col_values(1)
            mem["rows"] = {rid: i + 1 for i, rid in enumerate(col) if i and rid}
        rows = mem["rows"]
        vals = {rid: [str(r.get(c, "")) for c in TABLES["investors"]] for rid, r in invs.items()}
        old = [rid for rid in vals if rid in rows]
        new = [rid for rid in vals if rid not in rows]
        if old:                                        # everyone already in the sheet: one request
            w.batch_update([{"range": f"A{rows[rid]}", "values": [vals[rid]]} for rid in old],
                           value_input_option="RAW")
            for rid in old:
                invs.pop(rid)
        if new:                                        # everyone new: one request
            res = w.append_rows([vals[rid] for rid in new], value_input_option="RAW")
            start = int(res["updates"]["updatedRange"].split("!")[1].split(":")[0].lstrip("ABCDEFGHIJKLMNOPQRSTUVWXYZ"))
            for i, rid in enumerate(new):
                rows[rid] = start + i
                invs.pop(rid)
    else:
        LOCAL_DIR.mkdir(exist_ok=True)
        for name in ("purchases", "portfolio", "clicks", "answers"):
            if batch.get(name):
                p = LOCAL_DIR / f"{name}.csv"
                new = pd.DataFrame(batch[name], columns=TABLES[name])
                if p.exists() and list(pd.read_csv(p, nrows=0).columns) != TABLES[name]:
                    new = pd.concat([pd.read_csv(p, dtype=str), new]).reindex(columns=TABLES[name])
                    new.to_csv(p, index=False)
                else:
                    new.to_csv(p, mode="a", header=not p.exists(), index=False)
            batch.pop(name, None)
        invs = batch.get("investors") or {}
        if not invs:
            return
        p = LOCAL_DIR / "investors.csv"
        df = pd.read_csv(p, dtype=str) if p.exists() else pd.DataFrame(columns=TABLES["investors"])
        df = df[~df["rid"].isin(list(invs))]
        df = pd.concat([df, pd.DataFrame(list(invs.values()), columns=TABLES["investors"])])
        df.to_csv(p, index=False)
        invs.clear()


def _delete(rids, mem):
    """Remove every row belonging to these respondents from all tables."""
    rids = set(rids)
    if use_sheets():
        import gspread
        if "sh" not in mem:
            mem["sh"] = _sheet()
        for name, cols in TABLES.items():
            try:
                w = mem["sh"].worksheet(name)
            except gspread.WorksheetNotFound:
                continue
            rows = w.get_all_values()
            keep = [r for r in rows[1:] if r and r[0] not in rids]
            if len(keep) == len(rows) - 1:
                continue
            w.clear()
            w.update(values=[rows[0]] + keep, range_name="A1", value_input_option="RAW")
        mem["rows"] = None          # row numbers have shifted; re-read them on the next write
    else:
        for name in TABLES:
            p = LOCAL_DIR / f"{name}.csv"
            if p.exists():
                df = pd.read_csv(p, dtype=str)
                df[~df["rid"].isin(rids)].to_csv(p, index=False)


WRITE_EVERY = 5      # seconds between batches. Google allows ~60 writes a minute; a batch is at most 6 writes.


def _merge(jobs):
    batch = {"investors": {}, "purchases": [], "portfolio": [], "clicks": [], "answers": []}
    for j in jobs:
        if j.get("investor"):
            batch["investors"][j["investor"]["rid"]] = j["investor"]      # only the latest row per person
        for name in ("purchases", "portfolio", "clicks", "answers"):
            batch[name] += j.get(name) or []
    return batch


@st.cache_resource
def writer():
    """One background thread for the whole app. Every few seconds it takes everything people have saved since
    the last round and writes it in a handful of requests, however many people are using the app."""
    q, mem = queue.Queue(), {"rows": None, "written": 0}

    def write(jobs):
        batch = _merge(jobs)
        for attempt in range(10):
            try:
                _write(batch, mem)
                mem["written"] += len(jobs)
                return
            except Exception as e:
                mem.pop("sh", None)
                for k in [k for k in mem if k.startswith("ws_")]:
                    mem.pop(k)
                mem["last_error"] = f"{datetime.now(IST):%d %b %H:%M:%S} {type(e).__name__}: {str(e)[:200]}"
                time.sleep(min(60, 3 * 2 ** attempt))
        # Gave up after ~10 minutes: keep the rows in a local file rather than lose them silently.
        try:
            LOCAL_DIR.mkdir(exist_ok=True)
            with open(LOCAL_DIR / "unsaved.jsonl", "a") as fh:
                fh.write(json.dumps(batch, default=str) + "\n")
            mem["unsaved"] = mem.get("unsaved", 0) + len(jobs)
        except Exception:
            pass

    def work():
        while True:
            jobs = [q.get()]
            time.sleep(WRITE_EVERY)                   # let other people's saves pile up into the same batch
            while True:
                try:
                    jobs.append(q.get_nowait())
                except queue.Empty:
                    break
            try:
                pending = []
                for j in jobs:                        # deletes run in order, between the saves around them
                    if j.get("delete"):
                        if pending:
                            write(pending)
                            pending = []
                        try:
                            _delete(j["delete"], mem)
                        except Exception as e:
                            mem["last_error"] = f"{datetime.now(IST):%d %b %H:%M:%S} delete failed: {e}"
                    else:
                        pending.append(j)
                if pending:
                    write(pending)
            finally:
                for _ in jobs:
                    q.task_done()

    threading.Thread(target=work, daemon=True).start()
    return q, mem


def investor_row(status=None):
    ss = st.session_state
    held = {fid: h for fid, h in ss.alloc.items() if h}
    invested = sum(committed(h) for h in held.values())
    cm = sum(committed(h) for fid, h in held.items() if "capitalmind" in ss.get("amc_of", {}).get(fid, "").lower())
    t = now()
    return {"rid": ss.rid, "name": ss.get("name", ""), "app": ss.get("app", ""),
            "status": status or ss.get("status", "in progress"), "started_at": ss.started_at, "last_active": t,
            "completed_at": ss.get("completed_at", ""),
            "minutes_taken": round((datetime.fromisoformat(t) - datetime.fromisoformat(ss.started_at)).total_seconds() / 60, 1),
            "invested": invested, "cash": TOTAL - invested, "n_funds": len(held),
            "n_sips": sum(1 for h in held.values() if h.get("sip")), "capitalmind_amount": cm,
            "funds_opened": len({e["fund"] for e in ss.events if e["code"] == "open_fund"}),
            "taps": sum(1 for e in ss.events if e["code"] != "list_order"),
            "age": ss.get("age", ""), "location": ss.get("location", "")}


def flush(purchases=None, portfolio=None, answers=None):
    """Queue the investor row, new clicks and any purchase / final portfolio rows for the background writer."""
    ss = st.session_state
    new = ss.events[ss.get("events_saved", 0):]
    ss.events_saved = len(ss.events)
    writer()[0].put({"investor": investor_row(), "clicks": new, "purchases": purchases or [],
                     "portfolio": portfolio or [], "answers": answers or []})


def purchase_row(f, kind, amt):
    ss = st.session_state
    return {"rid": ss.rid, "name": ss.name, "time": now(), "fund_name": f["fund_name"], "amc": f["amc"],
            "category": f["category"], "sub_category": f["sub_category"],
            "type": "SIP" if kind == "sip" else "Lump sum", "amount": amt * (SIP_MONTHS if kind == "sip" else 1),
            "sip_monthly": amt if kind == "sip" else 0, "committed": amt * (SIP_MONTHS if kind == "sip" else 1),
            "list_position": ss.order.index(f["fid"]) + 1}


def final_portfolio(df):
    ss = st.session_state
    fx = df.set_index("fid")
    rows = []
    for fid, h in ss.alloc.items():
        if not h:
            continue
        f = fx.loc[fid]
        rows.append({"rid": ss.rid, "name": ss.name, "submitted_at": now(), "fund_name": f["fund_name"],
                     "amc": f["amc"], "category": f["category"], "sub_category": f["sub_category"], "risk": f["risk"],
                     "lump_sum": h.get("buy", 0), "sip_monthly": h.get("sip", 0), "committed": committed(h),
                     "share": round(committed(h) / TOTAL, 4), "list_position": ss.order.index(fid) + 1})
    return rows


# ---------------------------------------------------------------- researcher dashboard
@st.cache_data(ttl=60, show_spinner="Loading responses…")
def read_tables(_stamp):
    out = {}
    if use_sheets():
        sh = _sheet()
        for name, cols in TABLES.items():
            try:
                rows = sh.worksheet(name).get_all_values()
                out[name] = pd.DataFrame(rows[1:], columns=rows[0]) if rows else pd.DataFrame(columns=cols)
            except Exception:
                out[name] = pd.DataFrame(columns=cols)
    else:
        for name, cols in TABLES.items():
            p = LOCAL_DIR / f"{name}.csv"
            out[name] = pd.read_csv(p, dtype=str) if p.exists() else pd.DataFrame(columns=cols)
    for name, df in out.items():
        for c in ["invested", "cash", "n_funds", "n_sips", "capitalmind_amount", "funds_opened", "taps",
                  "minutes_taken", "amount", "sip_monthly", "committed", "share", "list_position",
                  "lump_sum", "seconds_from_start"]:
            if c in df.columns:
                df[c] = pd.to_numeric(df[c], errors="coerce")
    return out


def page_admin(funds):
    st.markdown("""<style>#MainMenu, footer, header[data-testid="stHeader"] {display:none !important;}
.block-container {max-width: 1200px; padding-top: 2rem;}
[data-testid="stMetric"] {background:#f7f7f8; border-radius:14px; padding:12px 16px;}</style>""",
                unsafe_allow_html=True)
    st.markdown("## Study dashboard")
    pw = secret("admin_password")
    if pw and not st.session_state.get("admin_ok"):
        p = st.text_input("Password", type="password")
        if p and p == pw:
            st.session_state.admin_ok = True
            st.rerun()
        elif p:
            st.error("Wrong password.")
        st.stop()
    if not pw and use_sheets():
        st.warning("Add admin_password to Secrets before sharing the app link, or anyone can open this page.")

    c_a, c_b = st.columns([1, 5], vertical_alignment="center")
    if c_a.button("Refresh", icon=":material/refresh:"):
        st.session_state.stamp = time.time()
    c_b.caption(("Reading from Google Sheets" if use_sheets() else "Reading from local CSV files (Google Sheets not set up)")
                + ". Updates every minute.")
    t = read_tables(st.session_state.get("stamp", 0))
    inv, pur, port, clk, ans = t["investors"], t["purchases"], t["portfolio"], t["clicks"], t["answers"]
    if inv.empty:
        st.info("No responses yet.")
        return
    done = inv[inv["status"] == "completed"]
    cm_names = set(funds.loc[funds["amc"].str.contains("Capitalmind", case=False), "fund_name"])

    m = st.columns(5)
    m[0].metric("Started", len(inv))
    m[1].metric("Completed (incl. questions)", f"{len(done)}  ({len(done) / len(inv):.0%})")
    m[2].metric("Avg. invested", f"₹{inr(done['invested'].mean() if len(done) else 0)}")
    m[3].metric("Avg. funds", f"{done['n_funds'].mean():.1f}" if len(done) else "–")
    m[4].metric("Put money in Capitalmind", f"{(done['capitalmind_amount'] > 0).mean():.0%}" if len(done) else "–")

    tab1, tab2, tab3, tab5, tab4 = st.tabs(["Investors", "One investor", "Funds", "Answers", "Download"])
    with tab1:
        for c in ("age", "location"):
            if c not in inv:
                inv[c] = ""
        show = inv[["name", "age", "location", "app", "status", "invested", "cash", "n_funds", "n_sips", "capitalmind_amount",
                    "funds_opened", "taps", "minutes_taken", "started_at", "rid"]].sort_values("started_at", ascending=False)
        st.dataframe(show, hide_index=True, width="stretch", column_config={
            "invested": st.column_config.NumberColumn("Invested ₹", format="%d"),
            "cash": st.column_config.NumberColumn("Cash ₹", format="%d"),
            "capitalmind_amount": st.column_config.NumberColumn("Capitalmind ₹", format="%d"),
            "n_funds": "Funds", "n_sips": "SIPs", "funds_opened": "Funds opened", "taps": "Taps",
            "minutes_taken": st.column_config.NumberColumn("Minutes", format="%.1f")})
        st.caption("A row appears as soon as someone taps Continue. Status: in progress → portfolio submitted "
                   "(tapped Completed) → completed (answered the questions and tapped Done).")
        with st.expander("Delete test or dummy entries"):
            labels = {r.rid: f"{r['name']} · {r['app']} · {r['status']} · started {r['started_at']}"
                      for _, r in inv.sort_values("started_at", ascending=False).iterrows()}
            pick = st.multiselect("Choose entries to delete", list(labels), format_func=labels.get)
            if pick:
                st.warning(f"This removes {len(pick)} entr{'y' if len(pick) == 1 else 'ies'} and all their purchases "
                           "and taps from every table. It can't be undone.")
                if st.button(f"Delete {len(pick)} entr{'y' if len(pick) == 1 else 'ies'}", type="primary"):
                    q = writer()[0]
                    q.put({"delete": pick})
                    q.join()                                  # wait until it's done before reloading
                    st.session_state.stamp = time.time()
                    st.toast("Deleted.")
                    st.rerun()
    with tab2:
        labels = {r.rid: f"{r['name']} · {r['status']} · started {r['started_at']}"
                  for _, r in inv.sort_values("started_at", ascending=False).iterrows()}
        rid = st.selectbox("Investor", list(labels), format_func=labels.get)
        r = inv[inv["rid"] == rid].iloc[0]
        k = st.columns(4)
        k[0].metric("Invested", f"₹{inr(r['invested'])}")
        k[1].metric("Funds", int(r["n_funds"]))
        k[2].metric("Minutes", r["minutes_taken"])
        k[3].metric("Taps", int(r["taps"]))
        st.markdown("**Portfolio**" + (" (submitted)" if r["status"] == "completed" else " (from purchases so far)"))
        p = port[port["rid"] == rid]
        if p.empty:
            p = pur[pur["rid"] == rid]
            if not p.empty:
                p = p.assign(lump_sum=p["amount"].where(p["type"] == "Lump sum", 0),
                             sip=p["sip_monthly"]).groupby(["fund_name", "category", "sub_category"], as_index=False)[
                    ["lump_sum", "sip", "committed"]].sum().rename(columns={"sip": "sip_monthly"})
                p["share"] = p["committed"] / TOTAL
        if p.empty:
            st.caption("No purchases.")
        else:
            st.dataframe(p[["fund_name", "category", "sub_category", "lump_sum", "sip_monthly", "committed", "share"]]
                         .sort_values("committed", ascending=False), hide_index=True, width="stretch",
                         column_config={"share": st.column_config.ProgressColumn("Share", format="percent",
                                                                                min_value=0, max_value=1)})
        a = ans[ans["rid"] == rid]
        st.markdown("**Answers**")
        if a.empty:
            st.caption("No answers yet.")
        else:
            a = a[a["submitted_at"] == a["submitted_at"].max()]
            st.dataframe(a[["question_no", "question", "answer", "correct"]], hide_index=True, width="stretch",
                         column_config={"question_no": "No."})
        st.markdown("**Every tap, in order**")
        c = clk[clk["rid"] == rid].copy()
        c = c.sort_values("time")
        st.dataframe(c[["seconds_from_start", "screen", "action", "fund", "detail"]], hide_index=True,
                     width="stretch", column_config={"seconds_from_start": "Second"})
    with tab3:
        opened = clk[clk["code"] == "open_fund"].groupby("fund")["rid"].nunique().rename("people_opened")
        src = port if not port.empty else pur
        bought = src.groupby("fund_name").agg(people_bought=("rid", "nunique"), money=("committed", "sum"))
        fx = funds[["fund_name", "amc", "sub_category", "aum_cr"]].set_index("fund_name")
        tbl = fx.join(opened).join(bought).fillna({"people_opened": 0, "people_bought": 0, "money": 0})
        tbl["opened_to_bought"] = (tbl["people_bought"] / tbl["people_opened"].where(tbl["people_opened"] > 0)).astype(float)
        tbl["capitalmind"] = tbl.index.isin(cm_names)
        st.dataframe(tbl.sort_values("money", ascending=False).reset_index(), hide_index=True, width="stretch",
                     column_config={"money": st.column_config.NumberColumn("Money ₹", format="%d"),
                                    "aum_cr": st.column_config.NumberColumn("AUM ₹ Cr", format="%d"),
                                    "opened_to_bought": st.column_config.NumberColumn("Opened → bought",
                                                                                      format="percent"),
                                    "people_opened": "Opened by", "people_bought": "Bought by"})
        st.caption(f"Based on {'submitted portfolios' if not port.empty else 'purchases so far'}. "
                   "Opened → bought shows how often a look turned into money.")
    with tab5:
        if ans.empty:
            st.info("No one has finished the questions yet.")
        else:
            st.caption(f"{ans['rid'].nunique()} people answered.")
            who = inv[inv["rid"].isin(ans["rid"])]
            for col, title, order in (("age", "Age", AGES), ("location", "Location", None)):
                if col in who:
                    v = who[col].fillna("").astype(str)
                    v = v[v != ""].value_counts()
                    if order:
                        v = v.reindex(order, fill_value=0)
                    if len(v):
                        st.markdown(f"**{title}** (from the start page)")
                        st.bar_chart(v, horizontal=True, height=40 + 36 * len(v))
            opts, others = {}, {}
            for q in QUESTIONS:
                opts[q["text"]] = q["options"]
                if q.get("other"):
                    others[q["text"]] = q["other"]
                if q["type"] == "top_bottom":
                    opts[f"{q['text']} ({q['top']} that matter most)"] = q["options"]
                    opts[f"{q['text']} ({q['bottom']} that matter least)"] = q["options"]
            for (no, qt), g in ans.groupby(["question_no", "question"], sort=False):
                st.markdown(f"**{no}. {qt}**")
                v = g["answer"].fillna("").str.split(" \\| ").explode()
                typed = []
                if qt in others:                      # "Others: xyz" counts under Others; typed answers listed below
                    pre = others[qt] + ":"
                    typed = [a[len(pre):].strip() for a in v if a.startswith(pre)]
                    v = v.where(~v.str.startswith(pre), others[qt])
                v = v[v != ""].value_counts()
                if qt in opts:
                    v = v.reindex(opts[qt], fill_value=0)
                st.bar_chart(v, horizontal=True, height=40 + 36 * max(len(v), 1))
                if typed:
                    st.caption("Typed under " + others[qt] + ": " + ", ".join(typed))
                if g["correct"].isin(["yes", "no"]).any():
                    st.caption(f"Answered correctly: {(g['correct'] == 'yes').mean():.0%}")
            with st.expander("All answers as a table"):
                st.dataframe(ans, hide_index=True, width="stretch")
    with tab4:
        for name, df in t.items():
            st.download_button(f"Download {name}.csv", df.to_csv(index=False).encode(), f"{name}.csv",
                               "text/csv", width="stretch")
        q, mem = writer()
        st.caption(f"Saves waiting to be written: {q.qsize()} · written since the app started: {mem['written']}")
        if mem.get("last_error"):
            st.caption(f"Last write error (retried automatically): {mem['last_error']}")
        if mem.get("unsaved"):
            st.error(f"{mem['unsaved']} saves could not be written to Google Sheets after 10 tries.")


# ---------------------------------------------------------------- main
qp = st.query_params
if qp.get("go") == "manager":
    page_redirect(qp)
    st.stop()
if "admin" in qp:
    page_admin(load_funds(str(DATA_FILE))[0])
    st.stop()

init_state()
ss = st.session_state
remember()
if ss.step == "start":
    page_start()
elif ss.step == "survey":
    page_survey()
elif ss.step == "done":
    page_thanks()
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
    ss.amc_of = dict(zip(funds["fid"], funds["amc"]))
    if ss.get("confirm"):
        confirm_dialog(funds)
    if True:
        scroll_top()
        if ss.view:
            page_fund(funds)
        else:
            topbar()
            tabs()
            wallet()
            if ss.tab == "portfolio":
                page_portfolio(funds)
            else:
                page_explore(funds)

remember()        # save again after this run, so the latest taps survive a refresh
