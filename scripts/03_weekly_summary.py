import pandas as pd
import gspread
from gspread_dataframe import get_as_dataframe, set_with_dataframe

# ── I/O sheet config ─────────────────────────────────────────────────────────
INPUT_SHEET_KEY   = "1FhjCMl4pQI-yiYNdo64IRZLGkUaNSAqbraENZdT3CAE"   # source workbook
INPUT_SHEET_TAB   = "daily_summary"                                  # tab to read

OUTPUT_SHEET_KEY  = "1FhjCMl4pQI-yiYNdo64IRZLGkUaNSAqbraENZdT3CAE"   # dest workbook
OUTPUT_SHEET_TAB  = "weekly_summary"                                 # tab to write

WEEK_START = "MON"   # "MON" or "SUN"


# ── Auth (service account, works headlessly in CI) ───────────────────────────
import os
import sys
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
from common.sheets_auth import get_client

gc = get_client()


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
ws_in = gc.open_by_key(INPUT_SHEET_KEY).worksheet(INPUT_SHEET_TAB)
daily = get_as_dataframe(ws_in, evaluate_formulas=True).dropna(how="all")

# ── Transform ────────────────────────────────────────────────────────────────
weekly = to_weekly(daily)

# ── Write (creates the tab if missing, clears it first) ──────────────────────
sh_out = gc.open_by_key(OUTPUT_SHEET_KEY)
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