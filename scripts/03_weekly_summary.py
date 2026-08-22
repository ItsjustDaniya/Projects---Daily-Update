"""
Weekly Summary Rollup
=======================
Reads the 'daily_summary' tab, rolls it up into weekly buckets (grouped by
Module + Week Start/End), and writes the result to a 'weekly_summary' tab in
the same workbook.

Auth: this file no longer imports common.sheets_auth.get_client() — it's now
self-contained via SERVICE_ACCOUNT_JSON (same pattern as the other pipeline
scripts), so it has no cross-module dependency. This script never touched
Metabase, so no METABASE_API_KEY is needed here.
"""

import json
import os

import pandas as pd
import gspread
from gspread_dataframe import get_as_dataframe, set_with_dataframe
from google.oauth2.service_account import Credentials

# ── I/O sheet config ─────────────────────────────────────────────────────────
INPUT_SHEET_KEY   = "1FhjCMl4pQI-yiYNdo64IRZLGkUaNSAqbraENZdT3CAE"   # source workbook
INPUT_SHEET_TAB   = "daily_summary"                                  # tab to read
OUTPUT_SHEET_KEY  = "1FhjCMl4pQI-yiYNdo64IRZLGkUaNSAqbraENZdT3CAE"   # dest workbook
OUTPUT_SHEET_TAB  = "weekly_summary"                                 # tab to write
WEEK_START = "MON"   # "MON" or "SUN"

# ── Auth (service account, works headlessly in CI) ───────────────────────────
service_account_json = os.getenv("SERVICE_ACCOUNT_JSON")
if not service_account_json:
    raise ValueError("❌ Missing environment variable: SERVICE_ACCOUNT_JSON")

service_info = json.loads(service_account_json)
creds = Credentials.from_service_account_info(
    service_info,
    scopes=[
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ],
)
gc = gspread.authorize(creds)

print("🔎 ENV CHECK")
print(f"   SA client_email    : {service_info.get('client_email')}")


def safe_open_by_key(key):
    """gc.open_by_key() wrapped to fail with an actionable message (the exact
    service-account email to share the sheet with) instead of a bare
    SpreadsheetNotFound traceback."""
    try:
        return gc.open_by_key(key)
    except gspread.exceptions.SpreadsheetNotFound:
        raise RuntimeError(
            f"❌ Could not open Google Sheet with key '{key}'. Share it with "
            f"this service account as Editor: {service_info.get('client_email')}"
        )


def to_weekly(df):
    df.columns = df.columns.str.strip()
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    df = df.dropna(subset=["Date"])
    if WEEK_START.upper() == "SUN":
        offset = (df["Date"].dt.weekday + 1) % 7
    else:
        offset = df["Date"].dt.weekday
    df["Week Start"] = (df["Date"] - pd.to_timedelta(offset, unit="D")).dt.normalize()
    df["Week End"]   = df["Week Start"] + pd.Timedelta(days=6)
    value_cols = [c for c in df.columns
                  if c not in ("Module", "Date", "Week Start", "Week End")]
    # numeric coercion in case the sheet read them as strings
    df[value_cols] = df[value_cols].apply(pd.to_numeric, errors="coerce").fillna(0)
    weekly = (
        df.groupby(["Module", "Week Start", "Week End"], as_index=False)[value_cols]
          .sum()
          .sort_values(["Module", "Week Start"])
          .reset_index(drop=True)
    )
    iso = weekly["Week Start"].dt.isocalendar()
    weekly.insert(
        1, "Week",
        iso["year"].astype(str) + "-W" + iso["week"].astype(str).str.zfill(2)
        + " (" + weekly["Week Start"].dt.strftime("%b %d")
        + " - " + weekly["Week End"].dt.strftime("%b %d") + ")"
    )
    weekly["Week Start"] = weekly["Week Start"].dt.strftime("%Y-%m-%d")
    weekly["Week End"]   = weekly["Week End"].dt.strftime("%Y-%m-%d")
    return weekly


# ── Read ─────────────────────────────────────────────────────────────────────
ws_in = safe_open_by_key(INPUT_SHEET_KEY).worksheet(INPUT_SHEET_TAB)
daily = get_as_dataframe(ws_in, evaluate_formulas=True).dropna(how="all")

# ── Transform ────────────────────────────────────────────────────────────────
weekly = to_weekly(daily)

# ── Write (creates the tab if missing, clears it first) ──────────────────────
sh_out = safe_open_by_key(OUTPUT_SHEET_KEY)
try:
    ws_out = sh_out.worksheet(OUTPUT_SHEET_TAB)
    ws_out.clear()
except gspread.WorksheetNotFound:
    ws_out = sh_out.add_worksheet(title=OUTPUT_SHEET_TAB,
                                  rows=len(weekly) + 10, cols=len(weekly.columns) + 2)
set_with_dataframe(ws_out, weekly)
print(f"Wrote {len(weekly)} weekly rows ({weekly['Module'].nunique()} modules) "
      f"to tab '{OUTPUT_SHEET_TAB}'")
weekly.head()
