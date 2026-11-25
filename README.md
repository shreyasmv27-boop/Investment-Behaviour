# Mutual fund allocation simulator (draft)

A Streamlit app for the Capitalmind "Crossing the Chasm" study. Each respondent gets ₹1,00,000 of pretend money and builds a portfolio from a list of funds. Every action is logged so we can see what they looked at and how they allocated.

## Flow

1. **Start**: name, and the app they mainly use (Groww, Zerodha Coin, Angel One, PhonePe, Paytm Money, Other).
2. **Mutual funds tab**: fund list showing only logo, name, AUM, plan, category and sub category. Filters on the left: AMC, Category (with sub category), Dividend type, Min. purchase amount, Expense ratio, AUM, Exit load and CAGR. Search and sort at the top. Click a fund name to open it.
3. **Fund page**: NAV chart (3M to Max, from the `nav_history` sheet), 1Y/3Y/5Y CAGR returns, NAV, AUM, expense ratio, minimums, exit load, fund manager and a risk gauge. Two options: **Buy** (one-time amount) or **SIP** (monthly amount for `SIP_MONTHS` months, default 12; the full commitment counts against the ₹1,00,000). Funds with no `min_sip` value don't offer SIP.
4. **Portfolio tab**: holdings with type, NAV, amount and allocation. As in a real app, money can't be pulled back out: Buy adds to the lump sum and SIP adds to the monthly amount. The portfolio is saved automatically after every Buy or SIP.
5. **Questionnaire**: a link in the side menu (shown once `form_url` is set in secrets) opens the Google Form with the respondent ID pre-filled.

## Files

| File | What it is |
|---|---|
| `app.py` | The app |
| `funds.xlsx` | Fund data: 37 real funds. Columns `tag`, `source_url`, `is_capitalmind` are for your analysis only and are never shown to participants |
| `requirements.txt` | Packages Streamlit Cloud installs |
| `.streamlit/config.toml` | Base theme |
| `.streamlit/secrets.toml.example` | Template for the Google Form link and Google Sheets saving |

## funds.xlsx columns

One row per fund, first sheet, these headers in row 1:

| Column | Example | Notes |
|---|---|---|
| fund_name | HDFC Flexi Cap Fund – Direct Growth | |
| amc | HDFC Mutual Fund | |
| category | Equity | Used for the filter tabs: Equity, Debt, Hybrid, Others |
| sub_category | Flexi Cap | Used for the sub category filter |
| plan | Growth | Optional, defaults to Growth |
| nav | 1875.42 | |
| return_1y, return_3y, return_5y | 14.6 | Percent. Leave blank if the fund is too new |
| risk | Very High | One of: Low, Low to Moderate, Moderate, Moderately High, High, Very High |
| min_lumpsum | 1000 | ₹ |
| exit_load | 1% if redeemed within 1 year | Text |
| expense_ratio | 0.68 | Percent |
| fund_manager | R. Mehta (since 2017) | Text |
| min_sip | 100 | Optional |
| aum_cr | 146784.53 | ₹ crore. Required, shown on the fund list |
| launch_year | 2013 | Optional |
| scheme_code | 122639 | AMFI scheme code. Links the fund to its NAV history, and lets `fetch_nav.py` fill NAV, CAGR and launch year |

## NAV history (the chart)

A second sheet called `nav_history` with columns `scheme_code`, `date`, `nav`. The easiest way to build it:

```
pip install pandas openpyxl requests
python fetch_nav.py
```

Run this on your own computer. It finds each fund's AMFI scheme code by name (or uses the one you typed), downloads the NAV history from mfapi.in and writes the `nav_history` sheet. Your own figures are left untouched unless you set `UPDATE_FIGURES = True`. Check the `matched_name` column afterwards. Until this has been run, fund pages simply show no chart.

## Where responses go

Saved after every portfolio change, as a numbered **snapshot**:
- **responses**: one row per snapshot (app, time taken so far, amount invested, cash, number of funds and SIPs, largest share, funds opened)
- **allocations**: one row per fund held in each snapshot (lump sum, monthly SIP, amount committed, share, list position)
- **events**: every action with a timestamp (tab changes, filters, sorts, searches, funds opened, Buy/SIP chosen, amounts set)

For analysis, use each respondent's **highest snapshot number** as their final portfolio. Earlier snapshots show how the portfolio changed along the way.

If Google Sheets is set up in secrets, rows go to tabs with those names. If not, they go to CSV files in `responses/`. **On Streamlit Cloud local files are wiped when the app restarts, so set up Google Sheets before sharing the link.**

## Run locally

```
pip install -r requirements.txt
streamlit run app.py
```

## Deploy on Streamlit Community Cloud

1. Put these files in a GitHub repo (public or private).
2. On share.streamlit.io, choose "Create app", pick the repo, and set the main file to `app.py`.
3. In the app's Settings > Secrets, paste the contents of `secrets.toml.example` with real values.
4. Share the app URL.

### Google Sheets saving

1. In Google Cloud Console, create a project, enable the Google Sheets API, and create a service account. Download its JSON key.
2. Create a Google Sheet and share it (Editor) with the service account's `client_email`.
3. Put the sheet ID and the JSON key fields into Secrets under `sheet_id` and `[gcp_service_account]`.

### Pre-filling the respondent ID in the Google Form

1. In the form, add a short-answer question "Respondent ID".
2. Form menu (three dots) > "Get pre-filled link", type anything in that field, and copy the link.
3. The link contains `entry.<number>=...`. Put `entry.<number>` in `form_id_field` and the form's `viewform` URL in `form_url`.

## Settings in app.py

- `TOTAL`: investable amount (100000)
- Fund `exit_load` text starting with "Nil" counts as no exit load in the filter.
- `SIP_MONTHS`: how many months an SIP runs for when counting it against the total (12)
- `STEP`: step on the amount box (500)
- `SHUFFLE_ORDER`: shows funds in a random order per respondent and logs it, so you can test whether position affects choice. Set to `False` to use the sheet's order.
- `APPS`: the app list and each one's accent colour and layout.

## What gets saved (Google Sheet tabs)

| Tab | One row per | Main columns |
|---|---|---|
| `investors` | investor (updated as they go) | name, app, status (in progress / completed), invested, cash, funds, SIPs, Capitalmind amount, funds opened, taps, minutes |
| `purchases` | every Buy or SIP | fund, type, amount, monthly SIP, list position |
| `portfolio` | fund in the final, submitted portfolio | lump sum, SIP, committed, share |
| `clicks` | every tap | time, seconds from start, screen, action, fund, detail |

Tabs are created automatically the first time someone uses the app. Writing happens in the background, so taps never wait for Google.
Taps on the filter button itself and on the dialog's close (X) can't be recorded; every filter change, the fund manager links and all other buttons are.

## Dashboard

Open `<your app url>/?admin=1` and enter `admin_password` from Secrets. Tabs: Investors (everyone, with status), One investor (their portfolio and every tap in order), Funds (opened vs bought, money per fund), Download (CSV of each table).

## NAV charts

`nav_history.csv` (columns `fund_name`, `date`, `nav`) feeds the fund page charts. `fund_name` must match `funds.xlsx` exactly.
