# Mutual fund allocation simulator (draft)

A Streamlit app for the Capitalmind "Crossing the Chasm" study. Each respondent gets ₹1,00,000 of pretend money and builds a portfolio from a list of funds. Every action is logged so we can see what they looked at and how they allocated.

## Flow

1. **Start**: name, and the app they mainly use (Groww, Zerodha Coin, Angel One, PhonePe, Paytm Money, Other).
2. **Invest**: the fund list in a look that matches their app (accent colour and card or row layout). They can filter by category, search, sort, open fund details, and type an amount for any fund. Anything not invested stays as cash.
3. **Done**: shows their respondent ID and a button to the Google Form, with the ID pre-filled.

## Files

| File | What it is |
|---|---|
| `app.py` | The app |
| `funds.xlsx` | Fund data. **Currently placeholder funds**, replace with the real sheet |
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
| sub_category | Flexi Cap | |
| nav | 1875.42 | |
| return_1y, return_3y, return_5y | 14.6 | Percent. Leave blank if the fund is too new |
| risk | Very High | One of: Low, Low to Moderate, Moderate, Moderately High, High, Very High |
| min_lumpsum | 1000 | ₹ |
| exit_load | 1% if redeemed within 1 year | Text |
| expense_ratio | 0.68 | Percent |
| fund_manager | R. Mehta (since 2017) | Text |
| min_sip | 100 | Optional |
| aum_cr | 65400 | Optional, ₹ crore |
| launch_year | 2013 | Optional |

## Where responses go

Three tables per respondent:
- **responses**: one row per person (app, time taken, amount invested, cash, number of funds, largest share, details opened)
- **allocations**: one row per fund chosen (amount, share, category, risk, position in their list)
- **events**: every action with a timestamp (open/close details, amount changes, filter, sort, search, and the order the list was shown in)

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
- `STEP`: step on the amount box (500)
- `SHUFFLE_ORDER`: shows funds in a random order per respondent and logs it, so you can test whether position affects choice. Set to `False` to use the sheet's order.
- `APPS`: the app list and each one's accent colour and layout.
