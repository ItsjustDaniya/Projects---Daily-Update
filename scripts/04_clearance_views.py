"""
Clearance views (Batch × days-since-release buckets)
====================================================
SELF-CONTAINED. Runs on its own — no need for eval_turnaround_report.py.

Day-windows count things BY HOW SOON THEY ARRIVED AFTER PROJECT RELEASE:
    <15 days = within 15 days of project_release_date … cumulative … <75 days
    Total    = everything (so >75-day arrivals & missing release dates → Total only)

TWO tables are produced, each in a LONG (Looker-ready) and WIDE (screenshot-style)
form:

  A) SUBMISSION view  — counts submission rows
       days_x_clearance / days_x_clearance_view
       measures:  # submissions | Evaluations % | Clearance %
         # submissions  Total → N ; <Nd → submissions arriving ≤ d days after release
         Evaluations %  evaluated / submissions  (within the window)
         Clearance %    cleared   / submissions  (cleared = evaluated AND
                                                  marks_submission_level ≥ CLEAR_PASS_LEVEL)

  B) UNIQUE-USER view — counts DISTINCT users (user_id), by each user's LATEST submission
       users_x_clearance / users_x_clearance_view
       measures:  Users submitted | Users evaluated | Users cleared
         Each user is first reduced to their single most recent submission
         (by submission_dt); the day window and all three measures are then
         derived from that one row:
         Users submitted  users whose LATEST submission arrived ≤ d days after release
         Users evaluated  …whose LATEST submission is evaluated
         Users cleared    …whose LATEST submission cleared (score ≥ 8)
       (cumulative across windows; Total = all)

⚠️ Unique-user counts are NOT additive — do not SUM them across batches/dates in
   Looker (a user active in two groups would be double-counted). They're correct
   at the grain of the row group (default: one row per Batch).

Data source: set INPUT_CSV (CSV mode, no auth) or leave None (Google Sheets, needs `gc`).
"""

from datetime import timezone, timedelta

import pandas as pd
import gspread
from gspread_dataframe import set_with_dataframe

import os
import sys
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
from common.sheets_auth import get_client

gc = get_client()


# ─────────────────────────────────────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────────────────────────────────────
SPREADSHEET_NAME_2 = "calender - 2"
MENTOR_SHEET       = "Project_evaluations"
OUTPUT_SHEET_KEY   = "19ecxPlN_MsnO8aVXs4FjJcxu0ShZ1ZtXAXJpZZKttj4"

REPORT_MODULES = [
    "DS 02 Spreadsheets", "DS 04 SQL", "DS 03 Power BI",
    "DS 06 EDA 1", "DS 07 EDA 2",
]
EXCLUDED_BATCHES = ["Spreadsheets (T)"]
IST = timezone(timedelta(hours=5, minutes=30))

CLEARANCE_DAY_BUCKETS = [15, 30, 45, 60, 75]   # days since release; cumulative
CLEAR_PASS_LEVEL      = 8                       # "cleared" = evaluated AND
                                               # marks_submission_level >= this
DXC_MODULES           = None                   # None ⇒ all REPORT_MODULES

# Row grouping. Default = one row per Batch (matches the screenshot).
#   add "submission_date" → per-day breakout   |   "submission_month" → per-month
GROUP_DIMS = ["Module_name", "Batch"]
_DIM_LABELS = {
    "Module_name": "Module", "Batch": "Batch",
    "submission_date": "Time", "submission_month": "Month",
}

# Output tabs
OUTPUT_SHEET_DXC_LONG = "days_x_clearance"
OUTPUT_SHEET_DXC_WIDE = "days_x_clearance_view"
OUTPUT_SHEET_USR_LONG = "users_x_clearance"
OUTPUT_SHEET_USR_WIDE = "users_x_clearance_view"

_BUCKET_LABELS       = ["Total"] + [f"<{d} days" for d in CLEARANCE_DAY_BUCKETS]
_METRIC_LABELS_SUB   = ["# submissions", "Evaluations %", "Clearance %"]
_METRIC_LABELS_USERS = ["Users submitted", "Users evaluated", "Users cleared"]

# Resubmission-attempt buckets — a student's cumulative Nth submission per module
# (rank ≤ N). CUMULATIVE, like the day windows. Edit thresholds freely.
RESUB_BUCKETS = [1, 2, 3, 5, 10]
def _resub_label(n): return "1" if n == 1 else f"<{n}"
OUTPUT_SHEET_RES_LONG = "resub_x_clearance"


# ─────────────────────────────────────────────────────────────────────────────
# Load (Google Sheets)
# ─────────────────────────────────────────────────────────────────────────────
def load_from_gsheets() -> pd.DataFrame:
    spreadsheet = gc.open(SPREADSHEET_NAME_2)            # noqa: F821  (gc from auth)
    worksheet   = spreadsheet.worksheet(MENTOR_SHEET)
    print(f"✓ Connected to '{SPREADSHEET_NAME_2}' → '{MENTOR_SHEET}'")
    records = worksheet.get_all_records(numericise_ignore=["all"])
    df = pd.DataFrame(records)
    print(f"✓ Fetched {len(df):,} rows  |  {len(df.columns)} columns")
    return df


# ─────────────────────────────────────────────────────────────────────────────
# Parse & clean
# ─────────────────────────────────────────────────────────────────────────────
def to_ist_naive(col: pd.Series) -> pd.Series:
    dt = pd.to_datetime(col, utc=False, errors="coerce")
    if getattr(dt.dt, "tz", None) is not None:
        dt = dt.dt.tz_convert(IST).dt.tz_localize(None)
    return dt


def parse_and_clean(df: pd.DataFrame) -> pd.DataFrame:
    df.columns = df.columns.str.strip()
    if "user_id" not in df.columns and "User ID" in df.columns:
        df = df.rename(columns={"User ID": "user_id"})
    if "project_release_date" not in df.columns:
        raise KeyError("Expected a 'project_release_date' column for release-based buckets.")

    df = df[df["Submission Time"].astype(str).str.strip().ne("")].copy()

    df["submission_dt"]          = to_ist_naive(df["Submission Time"])
    df["feedback_dt"]            = to_ist_naive(df["feedback_given_time"])
    df["release_dt"]             = to_ist_naive(df["project_release_date"])
    df["submission_date"]        = df["submission_dt"].dt.normalize()
    df["submission_month"]       = df["submission_dt"].dt.to_period("M").dt.to_timestamp()
    df["Module_name"]            = df["Module_name"].astype(str).str.strip()
    df["marks_submission_level"] = pd.to_numeric(df["marks_submission_level"], errors="coerce")

    df["is_evaluated"] = df["marks_submission_level"].notna() & df["feedback_dt"].notna()
    df["tat_hours"]    = (df["feedback_dt"] - df["submission_dt"]).dt.total_seconds() / 3600
    df["days_from_release"] = (df["submission_dt"] - df["release_dt"]).dt.total_seconds() / 86400

    df = df[df["submission_dt"].notna()].copy()
    df = df[df["Module_name"].isin(REPORT_MODULES)].copy()
    df = df[~df["Batch"].astype(str).str.strip().isin(EXCLUDED_BATCHES)].copy()

    # Dedup TRUE storage duplicates by submission_id
    before = len(df)
    unevaluated = (df[~df["is_evaluated"]].sort_values("submission_dt")
                   .drop_duplicates(subset=["submission_id"], keep="last"))
    evaluated   = (df[df["is_evaluated"]].sort_values("feedback_dt")
                   .drop_duplicates(subset=["submission_id"], keep="first"))
    df = pd.concat([unevaluated, evaluated], ignore_index=True).sort_values("submission_dt")
    print(f"  ↳ Removed {before - len(df):,} duplicate rows (by submission_id)")

    # Drop superseded un-evaluated submissions per (user_id, Module_name)
    df = df.sort_values(["user_id", "Module_name", "submission_dt"])
    latest_uneval = (df[~df["is_evaluated"]].groupby(["user_id", "Module_name"])["submission_dt"]
                     .max().rename("latest_uneval_dt").reset_index())
    latest_eval   = (df[df["is_evaluated"]].groupby(["user_id", "Module_name"])["submission_dt"]
                     .max().rename("latest_eval_dt").reset_index())
    df = df.merge(latest_uneval, on=["user_id", "Module_name"], how="left")
    df = df.merge(latest_eval,   on=["user_id", "Module_name"], how="left")
    df = df[
        df["is_evaluated"] |
        (~df["is_evaluated"] &
         (df["submission_dt"] == df["latest_uneval_dt"]) &
         (df["latest_eval_dt"].isna() | (df["latest_uneval_dt"] > df["latest_eval_dt"])))
    ].drop(columns=["latest_uneval_dt", "latest_eval_dt"]).reset_index(drop=True)

    print(f"✓ After cleaning: {len(df):,} rows  |  Modules: {df['Module_name'].unique().tolist()}")
    return df


# ─────────────────────────────────────────────────────────────────────────────
# Metric helpers
# ─────────────────────────────────────────────────────────────────────────────
def _pct(num, den):
    return round(100 * num / den, 1) if den else None


def _masks(grp):
    """Common per-row masks used by both blocks."""
    dfr     = grp["days_from_release"]
    ev_mask = grp["is_evaluated"]
    cl_mask = ev_mask & (grp["marks_submission_level"] >= CLEAR_PASS_LEVEL)
    return dfr, ev_mask, cl_mask


def _clearance_block(grp: pd.DataFrame) -> dict:
    """SUBMISSION counts: # submissions, Evaluations %, Clearance %."""
    N = len(grp)
    dfr, ev_mask, cl_mask = _masks(grp)
    E, C = int(ev_mask.sum()), int(cl_mask.sum())

    out = {
        ("Total", "# submissions"): N,
        ("Total", "Evaluations %"): _pct(E, N),
        ("Total", "Clearance %"):   _pct(C, N),
    }
    for d in CLEARANCE_DAY_BUCKETS:
        within    = dfr <= d
        sub_d     = int(within.sum())
        eval_d    = int((within & ev_mask).sum())
        cleared_d = int((within & cl_mask).sum())
        lab = f"<{d} days"
        out[(lab, "# submissions")] = sub_d
        out[(lab, "Evaluations %")] = _pct(eval_d, sub_d)
        out[(lab, "Clearance %")]   = _pct(cleared_d, sub_d)
    return out


def _user_block(grp: pd.DataFrame) -> dict:
    """DISTINCT-USER counts, based on each user's LATEST submission only.

    Each user is first reduced to their single most recent submission
    (by submission_dt) within the group; all three measures and the
    day-since-release windows are then derived from that one row:
        Users submitted  distinct users (everyone has a latest submission)
        Users evaluated  …whose LATEST submission is evaluated
        Users cleared    …whose LATEST submission cleared (score >= CLEAR_PASS_LEVEL)
    Day windows place each user by their LATEST submission's days_from_release,
    so the windows stay cumulative and Total = all users.
    """
    latest = (grp.sort_values("submission_dt")
                 .drop_duplicates(subset=["user_id"], keep="last"))
    dfr, ev_mask, cl_mask = _masks(latest)
    n = lambda mask: int(mask.sum())   # one row per user ⇒ row count == user count

    out = {
        ("Total", "Users submitted"): int(latest["user_id"].nunique()),
        ("Total", "Users evaluated"): n(ev_mask),
        ("Total", "Users cleared"):   n(cl_mask),
    }
    for d in CLEARANCE_DAY_BUCKETS:
        within = dfr <= d
        lab = f"<{d} days"
        out[(lab, "Users submitted")] = n(within)
        out[(lab, "Users evaluated")] = n(within & ev_mask)
        out[(lab, "Users cleared")]   = n(within & cl_mask)
    return out


def _dim_value(v):
    return v.date() if hasattr(v, "date") else v


# ─────────────────────────────────────────────────────────────────────────────
# Generic builders (shared by both tables)
# ─────────────────────────────────────────────────────────────────────────────
def _build_long(df, block_fn, metric_labels, name) -> pd.DataFrame:
    mods = DXC_MODULES or REPORT_MODULES
    d = df[df["Module_name"].isin(mods)].copy()
    pos = {lbl: i for i, lbl in enumerate(_BUCKET_LABELS)}

    rows = []
    for keys, grp in d.groupby(GROUP_DIMS):
        keys = keys if isinstance(keys, tuple) else (keys,)
        base = {_DIM_LABELS.get(k, k): _dim_value(v) for k, v in zip(GROUP_DIMS, keys)}
        block = block_fn(grp)
        for bucket in _BUCKET_LABELS:
            row = {**base, "Bucket": bucket, "bucket_order": pos[bucket]}
            for m in metric_labels:
                row[m] = block[(bucket, m)]
            rows.append(row)

    dim_cols = [_DIM_LABELS.get(k, k) for k in GROUP_DIMS]
    report = (pd.DataFrame(rows)
              .sort_values(dim_cols + ["bucket_order"]).reset_index(drop=True))
    print(f"✓ {name} (long): {len(report)} rows ({report[dim_cols].drop_duplicates().shape[0]} groups)")
    return report


def _build_wide(df, block_fn, metric_labels, name) -> pd.DataFrame:
    mods = DXC_MODULES or REPORT_MODULES
    d = df[df["Module_name"].isin(mods)].copy()
    dim_cols = [_DIM_LABELS.get(k, k) for k in GROUP_DIMS]

    rows = []
    for keys, grp in d.groupby(GROUP_DIMS):
        keys = keys if isinstance(keys, tuple) else (keys,)
        row = {("", _DIM_LABELS.get(k, k)): _dim_value(v) for k, v in zip(GROUP_DIMS, keys)}
        row.update(block_fn(grp))
        rows.append(row)

    wide = pd.DataFrame(rows)
    ordered = [("", c) for c in dim_cols]
    for b in _BUCKET_LABELS:
        for m in metric_labels:
            ordered.append((b, m))
    wide = wide.reindex(columns=pd.MultiIndex.from_tuples(ordered))
    wide = wide.sort_values([("", c) for c in dim_cols]).reset_index(drop=True)
    print(f"✓ {name} (wide): {len(wide)} rows × {len(_BUCKET_LABELS)} windows")
    return wide


# Public builders ─────────────────────────────────────────────────────────────
def build_days_x_clearance_long(df):
    return _build_long(df, _clearance_block, _METRIC_LABELS_SUB, "days_x_clearance")

def build_days_x_clearance_wide(df):
    return _build_wide(df, _clearance_block, _METRIC_LABELS_SUB, "days_x_clearance_view")

def build_users_x_clearance_long(df):
    return _build_long(df, _user_block, _METRIC_LABELS_USERS, "users_x_clearance")

def build_users_x_clearance_wide(df):
    return _build_wide(df, _user_block, _METRIC_LABELS_USERS, "users_x_clearance_view")


# ─────────────────────────────────────────────────────────────────────────────
# View C — Resubmission × Clearance (per day, batch-wise) → Looker Studio
# ─────────────────────────────────────────────────────────────────────────────
def build_resub_x_clearance_long(df) -> pd.DataFrame:
    """
    Tidy LONG table for Looker Studio. One row per
        Module × Batch × Time(day) × Resubmission bucket × Clearance window
    with COUNT measures: # submissions | Evaluations # | Clearance #.

    • Resubmission bucket = student's cumulative attempt number per module
      (rank ≤ N), computed over the CLEANED rows (all evaluated + each student's
      latest pending), matching the superseded-submission rule in the main report.
    • Clearance window = cumulative days since project release.
    • Both bucket axes are CUMULATIVE and each carries a 'Total' value, so do NOT
      sum across bucket values in Looker — filter/pick one (and don't turn on
      auto-subtotals for the two bucket dimensions). Counts DO sum safely across
      Time and Batch.
    • Zero-count cells are dropped to keep the export small.

    In Looker Studio: pivot table → row dims Batch, Time, Resubmission bucket
    (sort by resub_order); column dim Clearance window (sort by window_order);
    metrics # submissions / Evaluations # / Clearance #.
    """
    mods = DXC_MODULES or REPORT_MODULES
    d = (df[df["Module_name"].isin(mods)]
         .sort_values(["user_id", "Module_name", "submission_dt"]).copy())
    d["attempt"] = d.groupby(["user_id", "Module_name"]).cumcount() + 1

    win = [("Total", None)] + [(f"<{x} days", x) for x in CLEARANCE_DAY_BUCKETS]
    res = [("Total", None)] + [(_resub_label(n), n) for n in RESUB_BUCKETS]
    win_pos = {l: i for i, (l, _) in enumerate(win)}
    res_pos = {l: i for i, (l, _) in enumerate(res)}

    rows = []
    for (mod, batch, day), grp in d.groupby(["Module_name", "Batch", "submission_date"]):
        dfr = grp["days_from_release"]
        att = grp["attempt"]
        ev  = grp["is_evaluated"]
        cl  = ev & (grp["marks_submission_level"] >= CLEAR_PASS_LEVEL)
        for rlabel, rthr in res:
            rmask = (att >= 1) if rthr is None else (att <= rthr)
            for wlabel, wthr in win:
                wmask = (att >= 1) if wthr is None else (dfr <= wthr)
                cell = rmask & wmask
                nsub = int(cell.sum())
                if nsub == 0:
                    continue
                rows.append({
                    "Module": mod, "Batch": batch, "Time": day.date(),
                    "Resubmission bucket": rlabel, "resub_order": res_pos[rlabel],
                    "Clearance window": wlabel, "window_order": win_pos[wlabel],
                    "# submissions": nsub,
                    "Evaluations #": int((cell & ev).sum()),
                    "Clearance #":   int((cell & cl).sum()),
                })
    report = (pd.DataFrame(rows)
              .sort_values(["Module", "Batch", "Time", "resub_order", "window_order"])
              .reset_index(drop=True))
    print(f"✓ resub_x_clearance (long): {len(report)} rows")
    return report


# ─────────────────────────────────────────────────────────────────────────────
# Push helpers
# ─────────────────────────────────────────────────────────────────────────────
def _get_or_create_tab(spreadsheet, tab_name: str):
    try:
        return spreadsheet.worksheet(tab_name)
    except gspread.WorksheetNotFound:
        print(f"  Creating new tab '{tab_name}' …")
        return spreadsheet.add_worksheet(title=tab_name, rows=5000, cols=40)


def push_df(spreadsheet_key: str, tab_name: str, data: pd.DataFrame) -> None:
    sheet = gc.open_by_key(spreadsheet_key)                 # noqa: F821
    ws = _get_or_create_tab(sheet, tab_name)
    ws.clear()
    set_with_dataframe(ws, data, include_index=False, include_column_header=True)
    print(f"  ✓ Written {len(data)} rows to tab '{tab_name}'")


# ─────────────────────────────────────────────────────────────────────────────
# DATA SOURCE  +  RUN     ← edit these two lines
# ─────────────────────────────────────────────────────────────────────────────
INPUT_CSV   = None     # e.g. "query_result_2026-06-19.csv"  (None ⇒ Google Sheets)
DXC_DO_PUSH = True     # False ⇒ just build & preview (use this when reading a CSV)

if __name__ == "__main__":
    if INPUT_CSV:
        print(f"Reading CSV: {INPUT_CSV}")
        raw = pd.read_csv(INPUT_CSV, dtype=str)
    else:
        raw = load_from_gsheets()
    df = parse_and_clean(raw)

    # A) submission view
    dxc_long = build_days_x_clearance_long(df)
    dxc_wide = build_days_x_clearance_wide(df)
    # B) unique-user view
    usr_long = build_users_x_clearance_long(df)
    usr_wide = build_users_x_clearance_wide(df)
    # C) resubmission × clearance, per day, batch-wise (Looker Studio)
    res_long = build_resub_x_clearance_long(df)

    with pd.option_context("display.max_columns", None, "display.width", 220):
        print("\n--- Submissions view (wide, head) ---")
        print(dxc_wide.head(6).to_string(index=False))
        print("\n--- Unique-users view (wide, head) ---")
        print(usr_wide.head(6).to_string(index=False))
        print("\n--- Resubmission × clearance (long, head) ---")
        print(res_long.head(14).to_string(index=False))

    if DXC_DO_PUSH:
        print("\nPushing all tabs to Google Sheets …")
        push_df(OUTPUT_SHEET_KEY, OUTPUT_SHEET_DXC_LONG, dxc_long)
        push_df(OUTPUT_SHEET_KEY, OUTPUT_SHEET_DXC_WIDE, dxc_wide)
        push_df(OUTPUT_SHEET_KEY, OUTPUT_SHEET_USR_LONG, usr_long)
        push_df(OUTPUT_SHEET_KEY, OUTPUT_SHEET_USR_WIDE, usr_wide)
        push_df(OUTPUT_SHEET_KEY, OUTPUT_SHEET_RES_LONG, res_long)
        print("✓ All tabs updated.")