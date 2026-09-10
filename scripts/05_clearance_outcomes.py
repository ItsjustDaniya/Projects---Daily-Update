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
Auth: this file no longer imports common.sheets_auth.get_client() — it's now
self-contained via SERVICE_ACCOUNT_JSON (same pattern as the other pipeline
scripts), so it has no cross-module dependency. This script never touched
Metabase, so no METABASE_API_KEY is needed here.
"""
from datetime import timezone, timedelta
import json
import pandas as pd
import gspread
from gspread_dataframe import set_with_dataframe
import os
from google.oauth2.service_account import Credentials
# -------------------- ENV & AUTH --------------------
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
def safe_open_sheet(title):
    """gc.open() wrapped to fail with an actionable message (the exact
    service-account email to share the sheet with) instead of a bare
    SpreadsheetNotFound traceback."""
    try:
        return gc.open(title)
    except gspread.exceptions.SpreadsheetNotFound:
        raise RuntimeError(
            f"❌ Could not open Google Sheet '{title}'. Either the title "
            f"doesn't match exactly, or it hasn't been shared with this "
            f"service account: {service_info.get('client_email')}. "
            "Share it as Editor, then re-run."
        )
def safe_open_by_key(key):
    """Same as safe_open_sheet, but for gc.open_by_key()."""
    try:
        return gc.open_by_key(key)
    except gspread.exceptions.SpreadsheetNotFound:
        raise RuntimeError(
            f"❌ Could not open Google Sheet with key '{key}'. Share it with "
            f"this service account as Editor: {service_info.get('client_email')}"
        )
# ─────────────────────────────────────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────────────────────────────────────
SPREADSHEET_NAME_2 = "calender - 2"
MENTOR_SHEET       = "Project_evaluations"
OUTPUT_SHEET_KEY   = "19ecxPlN_MsnO8aVXs4FjJcxu0ShZ1ZtXAXJpZZKttj4"
"""
Evaluation Turnaround Report
=============================
Produces output sheets in Google Sheets:
  1. eval_turnaround       — Date × Module × TAT buckets (cumulative)
  2. resub_eval_turnaround — Date × Module × Resubmission buckets + TAT buckets
  3. tat_by_evaluator      — Spreadsheets only: Human (pre-AI) vs AI phase
  4. tat_by_batch          — Spreadsheets only: one row per batch (chronological)
  5. tat_by_batch_pivot    — Same as #4, transposed (batches in columns)
  6. tat_by_batch_date     — Same as #4, but one row per (batch, submission date);
                             TAT buckets shown as COUNTS (whole numbers), not %
TAT buckets are CUMULATIVE on the evaluated side: ≤24h ⊆ ≤48h ⊆ ≤72h ⊆ ≤7d, +>7d.
The Not-Evaluated side stays as separate day buckets (incl. a <1 day bucket).
Phase split (views 3–6) is BATCH-based: a batch's cohort month is parsed from its
name and compared to AI_CUTOFF — submission/evaluation dates never reclassify it.
Usage:
    python eval_turnaround_report.py
    python eval_turnaround_report.py --date 2024-03-21
    python eval_turnaround_report.py --export out
    python eval_turnaround_report.py --no-push
Auth: this file no longer imports common.sheets_auth.get_client() — it's now
self-contained via SERVICE_ACCOUNT_JSON (same pattern as the other pipeline
scripts), so it has no cross-module dependency. This script never touched
Metabase, so no METABASE_API_KEY is needed here.
"""
# ─────────────────────────────────────────────────────────────────────────────
# CELL 1 — Imports
# ─────────────────────────────────────────────────────────────────────────────
import argparse
import re
import json
from datetime import timezone, timedelta
import pandas as pd
import gspread
from gspread_dataframe import set_with_dataframe
import os
import sys
from google.oauth2.service_account import Credentials
# -------------------- ENV & AUTH --------------------
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
def safe_open_sheet(title):
    """gc.open() wrapped to fail with an actionable message (the exact
    service-account email to share the sheet with) instead of a bare
    SpreadsheetNotFound traceback."""
    try:
        return gc.open(title)
    except gspread.exceptions.SpreadsheetNotFound:
        raise RuntimeError(
            f"❌ Could not open Google Sheet '{title}'. Either the title "
            f"doesn't match exactly, or it hasn't been shared with this "
            f"service account: {service_info.get('client_email')}. "
            "Share it as Editor, then re-run."
        )
def safe_open_by_key(key):
    """Same as safe_open_sheet, but for gc.open_by_key()."""
    try:
        return gc.open_by_key(key)
    except gspread.exceptions.SpreadsheetNotFound:
        raise RuntimeError(
            f"❌ Could not open Google Sheet with key '{key}'. Share it with "
            f"this service account as Editor: {service_info.get('client_email')}"
        )
# ─────────────────────────────────────────────────────────────────────────────
# CELL 2 — Configuration
# ─────────────────────────────────────────────────────────────────────────────
SPREADSHEET_NAME_2 = "calender - 2"
MENTOR_SHEET       = "Project_evaluations"
REPORT_MODULES = [
    "DS 02 Spreadsheets",
    "DS 04 SQL",
    "DS 03 Power BI",
    "DS 06 EDA 1",
    "DS 07 EDA 2",
    "DS 08 ML 1",
    "DS 09 ML 2",
    "DS 09 MLops",
    "DS Deep Learning",   # id 210 — add this
]
OUTPUT_SHEET_KEY       = "1FhjCMl4pQI-yiYNdo64IRZLGkUaNSAqbraENZdT3CAE"
OUTPUT_SHEET_TAT       = "eval_turnaround"
OUTPUT_SHEET_RESUB_TAT = "resub_eval_turnaround"
# ── Comparison-view config (views 3–6) ───────────────────────────────────────
SPREADSHEET_MODULE         = "DS 02 Spreadsheets"
AI_CUTOFF                  = pd.Timestamp("2026-02-01")   # cohort month ≥ this → AI
SLA_HOURS                  = 48                            # turnaround target for SLA hit-rate
MATURITY_DAYS              = 7                             # min submission age for a fair comparison
OUTPUT_SHEET_BY_EVAL       = "tat_by_evaluator"
OUTPUT_SHEET_BY_BATCH      = "tat_by_batch"
OUTPUT_SHEET_BY_BATCH_T    = "tat_by_batch_pivot"
OUTPUT_SHEET_BY_BATCH_DATE = "tat_by_batch_date"
IST = timezone(timedelta(hours=5, minutes=30))
# Batches excluded from all reports (exact name match)
EXCLUDED_BATCHES = [
    "Spreadsheets (T)",
]
# ─────────────────────────────────────────────────────────────────────────────
# CELL 3 — Load data from Google Sheets
# ─────────────────────────────────────────────────────────────────────────────
def load_from_gsheets() -> pd.DataFrame:
    spreadsheet = safe_open_sheet(SPREADSHEET_NAME_2)
    worksheet   = spreadsheet.worksheet(MENTOR_SHEET)
    print(f"✓ Connected to '{SPREADSHEET_NAME_2}' → '{MENTOR_SHEET}'")
    records = worksheet.get_all_records(numericise_ignore=["all"])
    df = pd.DataFrame(records)
    print(f"✓ Fetched {len(df):,} rows  |  {len(df.columns)} columns")
    return df
# ─────────────────────────────────────────────────────────────────────────────
# CELL 4 — Parse & clean
# ─────────────────────────────────────────────────────────────────────────────
def to_ist_naive(col: pd.Series) -> pd.Series:
    """Parse a mixed-tz timestamp column → IST, tz-naive."""
    dt = pd.to_datetime(col, utc=False, errors="coerce")
    if dt.dt.tz is not None:
        dt = dt.dt.tz_convert(IST).dt.tz_localize(None)
    return dt
def parse_and_clean(df: pd.DataFrame) -> pd.DataFrame:
    df.columns = df.columns.str.strip()
    df = df[df["Submission Time"].astype(str).str.strip().ne("")].copy()
    df["submission_dt"]          = to_ist_naive(df["Submission Time"])
    df["feedback_dt"]            = to_ist_naive(df["feedback_given_time"])
    df["submission_date"]        = df["submission_dt"].dt.normalize()
    df["Module_name"]            = df["Module_name"].astype(str).str.strip()
    df["marks_submission_level"] = pd.to_numeric(df["marks_submission_level"], errors="coerce")
    # Evaluated = BOTH marks_submission_level AND feedback_given_time are non-null
    df["is_evaluated"] = (
        df["marks_submission_level"].notna() &
        df["feedback_dt"].notna()
    )
    # TAT in hours: feedback_given_time − Submission Time
    df["tat_hours"] = (
        (df["feedback_dt"] - df["submission_dt"])
        .dt.total_seconds() / 3600
    )
    df = df[df["submission_dt"].notna()].copy()
    df = df[df["Module_name"].isin(REPORT_MODULES)].copy()
    df = df[~df["Batch"].astype(str).str.strip().isin(EXCLUDED_BATCHES)].copy()
    print(f"  ↳ Excluded batches: {EXCLUDED_BATCHES}")
    # ── Deduplication ─────────────────────────────────────────────────────────
    # Remove only TRUE storage duplicates (the same physical submission stored
    # more than once). submission_id uniquely identifies a physical submission,
    # so we dedup on it — NOT on marks_submission_level.
    #
    # Why not dedup on level: two genuinely different resubmissions weeks apart
    # can receive the same marks_submission_level (e.g. a student fixed and
    # re-submitted but landed on the same level). Deduping by level would wrongly
    # merge those distinct evaluations and violate "keep every evaluated submission".
    #   UN-EVALUATED rows → keep ONE row per submission_id (latest occurrence)
    #   EVALUATED rows    → keep ONE row per submission_id (earliest feedback,
    #                        i.e. the first time it was evaluated)
    before = len(df)
    unevaluated = (
        df[~df["is_evaluated"]]
        .sort_values("submission_dt")
        .drop_duplicates(subset=["submission_id"], keep="last")
    )
    evaluated = (
        df[df["is_evaluated"]]
        .sort_values("feedback_dt")
        .drop_duplicates(subset=["submission_id"], keep="first")
    )
    df = pd.concat([unevaluated, evaluated], ignore_index=True).sort_values("submission_dt")
    print(f"  ↳ Removed {before - len(df):,} duplicate rows (by submission_id)")
    # ── Drop intermediate / superseded un-evaluated submissions ───────────────
    # For each (user_id, Module_name):
    #   • KEEP every evaluated submission   (each is a real evaluation → TAT)
    #   • KEEP a student's latest un-evaluated submission ONLY IF it is their
    #     most recent submission overall — i.e. it is newer than their latest
    #     EVALUATED submission (or they were never evaluated).
    #
    # Why the extra condition: a student may submit (un-evaluated), then later
    # resubmit and have that newer one evaluated. The earlier un-evaluated row
    # is NOT pending — the student already moved on and got an evaluation — so
    # it must be dropped, not counted as "Not Evaluated".
    #
    # Example (Ram Kumar): his latest un-eval (05-31) is after his last eval
    #   (05-28) → genuinely pending → kept. By contrast a student whose pending
    #   row is followed by a later evaluation is dropped as superseded.
    df = df.sort_values(["user_id", "Module_name", "submission_dt"])
    latest_uneval = (
        df[~df["is_evaluated"]]
        .groupby(["user_id", "Module_name"])["submission_dt"]
        .max()
        .rename("latest_uneval_dt")
        .reset_index()
    )
    latest_eval = (
        df[df["is_evaluated"]]
        .groupby(["user_id", "Module_name"])["submission_dt"]
        .max()
        .rename("latest_eval_dt")
        .reset_index()
    )
    df = df.merge(latest_uneval, on=["user_id", "Module_name"], how="left")
    df = df.merge(latest_eval,   on=["user_id", "Module_name"], how="left")
    df = df[
        df["is_evaluated"] |                                        # keep ALL evaluated
        (
            ~df["is_evaluated"] &
            (df["submission_dt"] == df["latest_uneval_dt"]) &       # latest un-eval, AND
            (df["latest_eval_dt"].isna() |                          # never evaluated, OR
             (df["latest_uneval_dt"] > df["latest_eval_dt"]))       # newer than any eval
        )
    ].drop(columns=["latest_uneval_dt", "latest_eval_dt"]).reset_index(drop=True)
    print(f"✓ After cleaning: {len(df):,} rows  |  Modules: {df['Module_name'].unique().tolist()}")
    print(f"  ↳ All evaluated kept; un-evaluated kept only when it is the student's most recent submission")
    return df
# ─────────────────────────────────────────────────────────────────────────────
# CELL 5 — TAT helper: classify a group of rows into buckets
# ─────────────────────────────────────────────────────────────────────────────
# Evaluated side is CUMULATIVE (≤24h ⊆ ≤48h ⊆ ≤72h ⊆ ≤7d, plus >7d tail).
# Not Evaluated side is separate day buckets, incl. a <1 day bucket so the five
# buckets sum to the total unevaluated count.
# Both output tables (eval_turnaround, resub_eval_turnaround) use this helper.
# ─────────────────────────────────────────────────────────────────────────────
def tat_buckets(grp: pd.DataFrame, as_of: "pd.Timestamp | None" = None) -> dict:
    """
    Given a slice of rows (all same date + module), return:
      - Evaluated TAT buckets — CUMULATIVE: ≤24h ⊆ ≤48h ⊆ ≤72h ⊆ ≤7d, plus >7d tail
      - Not Evaluated day buckets — separate (<1, 1–2, 3–4, 5–6, 7+ days)
      - Avg TAT for evaluated rows
    NOTE: the Evaluated columns nest (each includes the ones before it), so do
    NOT add them across. Valid check: Evaluated ≤7d + Evaluated >7d = total evaluated.
    The Not Evaluated buckets ARE separate and sum to the unevaluated count.
    """
    ev  = grp[grp["is_evaluated"]]
    nev = grp[~grp["is_evaluated"]]
    # Days since submission for un-evaluated rows
    if as_of is not None and len(nev) > 0:
        nev = nev.copy()
        nev["days_pending"] = (
            (as_of - nev["submission_dt"]).dt.total_seconds() / 86400
        )
    else:
        nev = nev.copy()
        nev["days_pending"] = float("nan")
    dp = nev["days_pending"]
    th = ev["tat_hours"]
    return {
        # ── Evaluated: CUMULATIVE turnaround (hours) ──────────────────────────
        "Evaluated ≤24h"          : int((th <=  24).sum()),
        "Evaluated ≤48h"          : int((th <=  48).sum()),
        "Evaluated ≤72h"          : int((th <=  72).sum()),
        "Evaluated ≤7d"           : int((th <= 168).sum()),   # 7 × 24h
        "Evaluated >7d"           : int((th >  168).sum()),    # tail
        "Evaluated Total"         : int(len(ev)),
        # ── Not Evaluated: separate day buckets (sum to unevaluated count) ────
        "Not Evaluated <1 day"    : int((dp < 1).sum()),
        "Not Evaluated 1–2 days"  : int(((dp >= 1) & (dp < 3)).sum()),
        "Not Evaluated 3–4 days"  : int(((dp >= 3) & (dp < 5)).sum()),
        "Not Evaluated 5–6 days"  : int(((dp >= 5) & (dp < 7)).sum()),
        "Not Evaluated 7+ days"   : int((dp >= 7).sum()),
        "Avg TAT hrs (eval only)" : round(ev["tat_hours"].mean(), 1) if len(ev) > 0 else None,
    }
# ─────────────────────────────────────────────────────────────────────────────
# CELL 6 — View 1: Evaluation Turnaround (simple)
# ─────────────────────────────────────────────────────────────────────────────
def build_eval_turnaround(
    df: pd.DataFrame,
    target_date: "pd.Timestamp | None" = None,
) -> pd.DataFrame:
    """
    Per submission-date, per module, per batch:
      Total Submissions | cumulative Evaluated buckets | Not Evaluated buckets | Avg TAT
    """
    today = pd.Timestamp.now(tz=IST).normalize().tz_localize(None)
    if target_date is not None:
        all_dates = [pd.Timestamp(target_date).normalize()]
    else:
        earliest  = df["submission_date"].min()
        all_dates = pd.date_range(start=earliest, end=today, freq="D").tolist()
    rows = []
    for mod in REPORT_MODULES:
        mod_df = df[df["Module_name"] == mod]
        for date in all_dates:
            day_grp = mod_df[mod_df["submission_date"] == date]
            if day_grp.empty:
                continue
            # One row per batch within this date + module
            for batch, grp in day_grp.groupby("Batch"):
                row = {
                    "Module"           : mod,
                    "Submission Date"  : date.date(),
                    "Batch"            : batch,
                    "Total Submissions": len(grp),
                }
                row.update(tat_buckets(grp, as_of=today))
                rows.append(row)
    report = (
        pd.DataFrame(rows)
        .sort_values(["Module", "Submission Date"])
        .reset_index(drop=True)
    )
    print(f"✓ Eval turnaround: {len(report)} rows")
    return report
# ─────────────────────────────────────────────────────────────────────────────
# CELL 7 — View 2: Resubmission buckets + TAT
# ─────────────────────────────────────────────────────────────────────────────
def build_resub_eval_turnaround(
    df: pd.DataFrame,
    target_date: "pd.Timestamp | None" = None,
) -> pd.DataFrame:
    """
    Per submission-date, per module:
      Resubmission buckets (student's cumulative submission count for this module):
        1st–2nd | 3rd–4th | 5th–6th | 7th–8th | 9th+
      Plus the same TAT buckets as View 1.
    Cumulative rank = chronological order of submissions per (user_id, Module_name).
    Rank 1 = student's very first submission for this module, rank 2 = second, etc.
    Note: rank is computed over the cleaned rows (all evaluated + latest un-eval).
    """
    # Assign cumulative submission rank per student per module
    work = df.sort_values(["user_id", "Module_name", "submission_dt"]).copy()
    work["student_sub_rank"] = (
        work.groupby(["user_id", "Module_name"]).cumcount() + 1
    )
    today = pd.Timestamp.now(tz=IST).normalize().tz_localize(None)
    if target_date is not None:
        all_dates = [pd.Timestamp(target_date).normalize()]
    else:
        earliest  = work["submission_date"].min()
        all_dates = pd.date_range(start=earliest, end=today, freq="D").tolist()
    # Bucket definitions: (label, min_rank, max_rank or None)
    RESUB_BUCKETS = [
        ("1st–2nd Submission", 1, 2),
        ("3rd–4th Submission", 3, 4),
        ("5th–6th Submission", 5, 6),
        ("7th–8th Submission", 7, 8),
        ("9th+ Submission",    9, None),
    ]
    rows = []
    for mod in REPORT_MODULES:
        mod_df = work[work["Module_name"] == mod]
        for date in all_dates:
            grp = mod_df[mod_df["submission_date"] == date]
            if grp.empty:
                continue
            for batch, batch_grp in grp.groupby("Batch"):
                rb = batch_grp["student_sub_rank"]
                for label, lo, hi in RESUB_BUCKETS:
                    if hi is None:
                        bucket_mask = rb >= lo
                    else:
                        bucket_mask = (rb >= lo) & (rb <= hi)
                    bucket_grp = batch_grp[bucket_mask]
                    count      = int(bucket_mask.sum())
                    row = {
                        "Module"               : mod,
                        "Submission Date"      : date.date(),
                        "Batch"                : batch,
                        "Submission Bucket"    : label,
                        "Submissions in Bucket": count,
                        "Total Submissions"    : len(batch_grp),
                    }
                    row.update(tat_buckets(bucket_grp, as_of=today))
                    rows.append(row)
    report = (
        pd.DataFrame(rows)
        .sort_values(["Module", "Submission Date"])
        .reset_index(drop=True)
    )
    print(f"✓ Resub + eval turnaround: {len(report)} rows")
    return report
# ─────────────────────────────────────────────────────────────────────────────
# CELL 7B — Comparison views: Evaluator phase & Batch-by-batch (Spreadsheets)
# ─────────────────────────────────────────────────────────────────────────────
# Phase split is BATCH-based: cohort month parsed from the batch NAME vs AI_CUTOFF.
# Filters available on every view: batches, date_from, date_to, mature_only.
_MONTH_RE = re.compile(r"([A-Za-z]+)\s+(\d{4})")
def batch_cohort_month(batch: str) -> "pd.Timestamp":
    """
    Parse the cohort month from a batch name.
    'DS Spreadsheets - January 2026' → Timestamp('2026-01-01'); NaT if not found.
    This is the key to sorting batches correctly: sort by this real date, NOT by
    the batch text (text sorting puts 'April' before 'January').
    """
    m = _MONTH_RE.search(str(batch))
    if not m:
        return pd.NaT
    for fmt in ("%B %Y", "%b %Y"):               # full name, then abbreviated
        try:
            return pd.to_datetime(f"{m.group(1)} {m.group(2)}", format=fmt)
        except ValueError:
            continue
    return pd.NaT
def _prep_spreadsheets(
    df: pd.DataFrame,
    batches=None,
    date_from=None,
    date_to=None,
    mature_only: bool = True,
):
    """Filter to Spreadsheets + apply batch/date/maturity filters; tag phase."""
    today = pd.Timestamp.now(tz=IST).normalize().tz_localize(None)
    d = df[df["Module_name"] == SPREADSHEET_MODULE].copy()
    if batches:
        d = d[d["Batch"].isin(batches)]
    if date_from is not None:
        d = d[d["submission_date"] >= pd.Timestamp(date_from).normalize()]
    if date_to is not None:
        d = d[d["submission_date"] <= pd.Timestamp(date_to).normalize()]
    if mature_only:
        age_days = (today - d["submission_dt"]).dt.total_seconds() / 86400
        d = d[age_days >= MATURITY_DAYS]
    d["cohort_month"]    = d["Batch"].map(batch_cohort_month)
    d["evaluator_phase"] = d["cohort_month"].apply(
        lambda m: "AI" if (pd.notna(m) and m >= AI_CUTOFF) else "Human"
    )
    return d, today
def _tat_summary(grp: pd.DataFrame) -> dict:
    """Comparison metrics for one slice. Rates are % of EVALUATED items."""
    ev   = grp[grp["is_evaluated"]]
    n_ev = len(ev)
    th   = ev["tat_hours"]
    def pct(count):  # share of evaluated
        return round(100 * count / n_ev, 1) if n_ev else None
    return {
        "Total Submissions"           : len(grp),
        "Evaluated"                   : n_ev,
        "Not Evaluated"               : int((~grp["is_evaluated"]).sum()),
        "% ≤24h"                      : pct(int((th <=  24).sum())),
        "% ≤48h"                      : pct(int((th <=  48).sum())),
        "% ≤72h"                      : pct(int((th <=  72).sum())),
        "% ≤7d"                       : pct(int((th <= 168).sum())),
        "% >7d"                       : pct(int((th >  168).sum())),
        f"% within SLA ({SLA_HOURS}h)": pct(int((th <= SLA_HOURS).sum())),
        "Median TAT hrs"              : round(th.median(), 1)      if n_ev else None,
        "P90 TAT hrs"                 : round(th.quantile(0.9), 1) if n_ev else None,
        "Avg TAT hrs"                 : round(th.mean(), 1)        if n_ev else None,
    }
def _tat_summary_counts(grp: pd.DataFrame) -> dict:
    """
    Same shape as _tat_summary but TAT buckets are COUNTS (whole numbers),
    not percentages. Evaluated buckets are cumulative (≤24h ⊆ ≤48h ⊆ ≤72h ⊆ ≤7d),
    plus a >7d tail — so for any row: ≤7d + >7d = Evaluated.
    """
    ev   = grp[grp["is_evaluated"]]
    n_ev = len(ev)
    th   = ev["tat_hours"]
    return {
        "Total Submissions"          : len(grp),
        "Evaluated"                  : n_ev,
        "Not Evaluated"              : int((~grp["is_evaluated"]).sum()),
        "≤24h"                       : int((th <=  24).sum()),
        "≤48h"                       : int((th <=  48).sum()),
        "≤72h"                       : int((th <=  72).sum()),
        "≤7d"                        : int((th <= 168).sum()),
        ">7d"                        : int((th >  168).sum()),
        f"within SLA ({SLA_HOURS}h)" : int((th <= SLA_HOURS).sum()),
        "Median TAT hrs"             : round(th.median(), 1)      if n_ev else None,
        "P90 TAT hrs"                : round(th.quantile(0.9), 1) if n_ev else None,
        "Avg TAT hrs"                : round(th.mean(), 1)        if n_ev else None,
    }
def build_tat_by_evaluator(df, batches=None, date_from=None, date_to=None,
                           mature_only=True) -> pd.DataFrame:
    """View 3 — Human (pre-AI) vs AI, two rows."""
    d, _ = _prep_spreadsheets(df, batches, date_from, date_to, mature_only)
    rows = []
    for phase in ["Human", "AI"]:                # fixed order: pre then post
        grp = d[d["evaluator_phase"] == phase]
        if grp.empty:
            continue
        rows.append({"Evaluator Phase": phase, **_tat_summary(grp)})
    print(f"✓ tat_by_evaluator: {len(rows)} phase rows (mature_only={mature_only})")
    return pd.DataFrame(rows)
def build_tat_by_batch(df, batches=None, date_from=None, date_to=None,
                       mature_only=True) -> pd.DataFrame:
    """View 4 — one row per batch, sorted oldest → newest by cohort month."""
    d, _ = _prep_spreadsheets(df, batches, date_from, date_to, mature_only)
    rows = []
    for batch, grp in d.groupby("Batch"):
        rows.append({
            "Batch"          : batch,
            "_sort_month"    : grp["cohort_month"].iloc[0],     # hidden sort key
            "Cohort Month"   : grp["cohort_month"].iloc[0],
            "Evaluator Phase": grp["evaluator_phase"].iloc[0],
            **_tat_summary(grp),
        })
    report = (
        pd.DataFrame(rows)
        .sort_values("_sort_month")            # ← chronological, not alphabetical
        .drop(columns="_sort_month")
        .reset_index(drop=True)
    )
    if not report.empty:
        report["Cohort Month"] = report["Cohort Month"].dt.strftime("%b %Y")
    print(f"✓ tat_by_batch: {len(report)} batch rows (mature_only={mature_only})")
    return report
# ─────────────────────────────────────────────────────────────────────────────
# CELL 7C — View 6: Batch × Submission Date (TAT buckets as COUNTS)
# ─────────────────────────────────────────────────────────────────────────────
def build_tat_by_batch_date(df, batches=None, date_from=None, date_to=None,
                            mature_only=False) -> pd.DataFrame:
    """
    One row per (Batch, Submission Date). Retains Cohort Month + Evaluator Phase.
    TAT buckets shown as COUNTS (whole numbers), not percentages.
    Sorted oldest → newest by cohort month, then by submission date.
    mature_only defaults to False so EVERY submission is included regardless of
    age — submissions without an evaluation simply count as Not Evaluated.
    """
    d, _ = _prep_spreadsheets(df, batches, date_from, date_to, mature_only)
    rows = []
    for (batch, sub_date), grp in d.groupby(["Batch", "submission_date"]):
        rows.append({
            "Batch"          : batch,
            "Submission Date": sub_date.date(),
            "_sort_month"    : grp["cohort_month"].iloc[0],     # hidden sort key
            "Cohort Month"   : grp["cohort_month"].iloc[0],
            "Evaluator Phase": grp["evaluator_phase"].iloc[0],
            **_tat_summary_counts(grp),          # ← counts instead of %
        })
    report = (
        pd.DataFrame(rows)
        .sort_values(["_sort_month", "Submission Date"])    # chronological
        .drop(columns="_sort_month")
        .reset_index(drop=True)
    )
    if not report.empty:
        report["Cohort Month"] = report["Cohort Month"].dt.strftime("%b %Y")
    print(f"✓ tat_by_batch_date: {len(report)} batch×date rows (mature_only={mature_only})")
    return report
# ─────────────────────────────────────────────────────────────────────────────
# CELL 7D — View 5: batch comparison TRANSPOSED (batches in COLUMNS)
# ─────────────────────────────────────────────────────────────────────────────
def build_tat_by_batch_pivoted(df, header="Cohort Month", **kw) -> pd.DataFrame:
    """
    Transpose View 4: metrics down the rows, batches across the columns
    (oldest batch left → newest right).
    header : which field labels each batch column —
             "Cohort Month" → short labels like 'Sep 2025' (recommended)
             "Batch"        → full names like 'DS Spreadsheets - September 2025'
    Accepts the same filters as build_tat_by_batch (batches, date_from,
    date_to, mature_only) passed via **kw.
    """
    by_batch = build_tat_by_batch(df, **kw)
    if by_batch.empty:
        print("✓ tat_by_batch_pivot: no rows")
        return by_batch
    metric_cols = [c for c in by_batch.columns if c not in ("Batch", "Cohort Month")]
    pivoted = (
        by_batch.set_index(header)[metric_cols]   # index = batch label
        .T                                         # flip: metrics → rows
        .rename_axis("Metric")                     # name the row-label column
        .reset_index()
    )
    print(f"✓ tat_by_batch_pivot: {len(pivoted)} metric rows × "
          f"{len(pivoted.columns) - 1} batch columns")
    return pivoted
# ─────────────────────────────────────────────────────────────────────────────
# CELL 8 — Push to Google Sheets
# ─────────────────────────────────────────────────────────────────────────────
def _get_or_create_tab(spreadsheet, tab_name: str):
    try:
        return spreadsheet.worksheet(tab_name)
    except gspread.WorksheetNotFound:
        print(f"  Creating new tab '{tab_name}' …")
        return spreadsheet.add_worksheet(title=tab_name, rows=5000, cols=40)
def push_many(tabs_and_frames) -> None:
    """Write a list of (tab_name, dataframe) pairs to the output spreadsheet."""
    sheet = safe_open_by_key(OUTPUT_SHEET_KEY)
    for tab, data in tabs_and_frames:
        ws = _get_or_create_tab(sheet, tab)
        ws.clear()
        set_with_dataframe(ws, data, include_index=False, include_column_header=True)
        print(f"  ✓ Written {len(data)} rows to tab '{tab}'")
# ─────────────────────────────────────────────────────────────────────────────
# CELL 9 — CLI argument parsing
# ─────────────────────────────────────────────────────────────────────────────
def _parse_args():
    parser = argparse.ArgumentParser(description="Evaluation Turnaround Report Generator")
    parser.add_argument("--date",    type=str, default=None,
                        help="Run for a single date (YYYY-MM-DD). Omit for full history.")
    parser.add_argument("--export",  type=str, default=None,
                        help="CSV export prefix. E.g. --export out → out_eval_tat.csv, out_resub_tat.csv")
    parser.add_argument("--no-push", action="store_true",
                        help="Skip writing back to Google Sheets.")
    args, _ = parser.parse_known_args()   # ignore Jupyter kernel flags
    return args
# ─────────────────────────────────────────────────────────────────────────────
# CELL 10 — Run
# ─────────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    args          = _parse_args()
    target_date   = pd.Timestamp(args.date) if args.date else None
    do_push       = not args.no_push
    export_prefix = args.export
else:
    # Notebook / Colab defaults
    target_date   = None
    do_push       = True
    export_prefix = None
# ── Comparison-view filters (leave as None to ignore) ────────────────────────
FILTER_BATCHES   = None            # e.g. ["DS Spreadsheets - April 2026"]
FILTER_DATE_FROM = None            # e.g. "2025-01-01"
FILTER_DATE_TO   = None            # e.g. "2027-02-28"
MATURE_ONLY      = True            # False = include very recent (not-yet-resolved) submissions
# 1. Load
raw = load_from_gsheets()
# 2. Clean
df = parse_and_clean(raw)
# 3. Build original views
eval_tat_report  = build_eval_turnaround(df,       target_date=target_date)
resub_tat_report = build_resub_eval_turnaround(df, target_date=target_date)
# 4. Build comparison views (Spreadsheets only)
by_eval = build_tat_by_evaluator(
    df, batches=FILTER_BATCHES, date_from=FILTER_DATE_FROM,
    date_to=FILTER_DATE_TO, mature_only=MATURE_ONLY,
)
by_batch = build_tat_by_batch(
    df, batches=FILTER_BATCHES, date_from=FILTER_DATE_FROM,
    date_to=FILTER_DATE_TO, mature_only=MATURE_ONLY,
)
by_batch_pivot = build_tat_by_batch_pivoted(
    df, header="Cohort Month",          # change to "Batch" for full names
    batches=FILTER_BATCHES, date_from=FILTER_DATE_FROM,
    date_to=FILTER_DATE_TO, mature_only=MATURE_ONLY,
)
# 4b. Batch × Submission Date view — include EVERY submission (immature too);
#     unevaluated ones simply count as Not Evaluated → mature_only=False.
#     TAT buckets here are whole-number COUNTS, not percentages.
by_batch_date = build_tat_by_batch_date(
    df, batches=FILTER_BATCHES, date_from=FILTER_DATE_FROM,
    date_to=FILTER_DATE_TO, mature_only=False,
)
# 5. Preview
print("\n--- Eval Turnaround (sample) ---")
print(eval_tat_report.head(10).to_string(index=False))
print("\n--- Resub + Eval Turnaround (sample) ---")
print(resub_tat_report.head(10).to_string(index=False))
print("\n--- TAT by Evaluator (Human vs AI) ---")
print(by_eval.to_string(index=False))
print("\n--- TAT by Batch ---")
print(by_batch.to_string(index=False))
print("\n--- TAT by Batch (batches in columns) ---")
print(by_batch_pivot.to_string(index=False))
print("\n--- TAT by Batch × Submission Date (counts) ---")
print(by_batch_date.head(15).to_string(index=False))
# 6. Export to CSV (optional)
if export_prefix:
    eval_tat_report.to_csv(f"{export_prefix}_eval_tat.csv",            index=False)
    resub_tat_report.to_csv(f"{export_prefix}_resub_tat.csv",          index=False)
    by_eval.to_csv(f"{export_prefix}_tat_by_evaluator.csv",            index=False)
    by_batch.to_csv(f"{export_prefix}_tat_by_batch.csv",               index=False)
    by_batch_pivot.to_csv(f"{export_prefix}_tat_by_batch_pivot.csv",   index=False)
    by_batch_date.to_csv(f"{export_prefix}_tat_by_batch_date.csv",     index=False)
    print(f"✓ Exported CSVs with prefix '{export_prefix}'")
# 7. Push to Google Sheets
if do_push:
    print("\nPushing to Google Sheets …")
    push_many([
        (OUTPUT_SHEET_TAT,        eval_tat_report),
        (OUTPUT_SHEET_RESUB_TAT,  resub_tat_report),
        (OUTPUT_SHEET_BY_EVAL,    by_eval),
        (OUTPUT_SHEET_BY_BATCH,   by_batch),
        (OUTPUT_SHEET_BY_BATCH_T, by_batch_pivot),
        (OUTPUT_SHEET_BY_BATCH_DATE, by_batch_date),
    ])
    print("✓ All tabs updated.")
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
    spreadsheet = safe_open_sheet(SPREADSHEET_NAME_2)
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
    sheet = safe_open_by_key(spreadsheet_key)
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
