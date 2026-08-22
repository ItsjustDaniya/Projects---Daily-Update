"""
Project Evaluation Daily Report Generator — Extended
=====================================================
Produces two output sheets in Google Sheets:

  1. daily_summary     — per-date / per-module summary with:
                          • Total Submissions (12AM-11AM)
                          • New Add-ons (11AM-11:59PM)
                          • Total New Submissions (sum of above two)
                          • Last Day Evaluations
                          • Pending Evaluations (CURRENT snapshot, not per-date)
                          • Per-mentor breakdown of last-day evals

  2. mentor_deep_dive  — per-mentor / per-date breakdown (unchanged)

Usage:
    python project_evaluation_report_extended.py
    python project_evaluation_report_extended.py --date 2024-03-21
    python project_evaluation_report_extended.py --export out
    python project_evaluation_report_extended.py --no-push

Auth: this file no longer imports common.sheets_auth.get_client() — it's now
self-contained via SERVICE_ACCOUNT_JSON (same pattern as the other pipeline
scripts), so it has no cross-module dependency. This script never touched
Metabase, so no METABASE_API_KEY is needed here.
"""

# ─────────────────────────────────────────────────────────────────────────────
# CELL 1 — Imports
# ─────────────────────────────────────────────────────────────────────────────
import argparse
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
]

OUTPUT_SHEET_KEY    = "1FhjCMl4pQI-yiYNdo64IRZLGkUaNSAqbraENZdT3CAE"
OUTPUT_SHEET_DAILY  = "daily_summary"
OUTPUT_SHEET_MENTOR = "mentor_deep_dive"

SCORE_PASS_THRESHOLD = 8   # marks_obtained >= this → "Passed"

EXCLUDED_BATCHES = [
    "Spreadsheets (T)",
    "DA SQL - Advantage August 2025",
    "DA Spreadsheets - Advantage August 2025",
    "DS Power BI - July & Aug-Adv 2025",
]

NIGHT_START_H = 0
NIGHT_END_H   = 11
DAY_START_H   = 11

IST = timezone(timedelta(hours=5, minutes=30))

RESUB_BUCKETS = [
    ("Re-Sub Attempt 1",    lambda lvl: lvl == 2),
    ("Re-Sub Attempt 2",    lambda lvl: lvl == 3),
    ("Re-Sub Attempt 3",    lambda lvl: lvl == 4),
    ("Re-Sub Attempts 4-6", lambda lvl: (lvl >= 5) & (lvl <= 7)),
    ("Re-Sub Attempts 7-10",lambda lvl: (lvl >= 8) & (lvl <= 11)),
    ("Re-Sub Attempts 10+", lambda lvl: lvl >= 12),
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
    dt = pd.to_datetime(col, utc=False, errors="coerce")
    if dt.dt.tz is not None:
        dt = dt.dt.tz_convert(IST).dt.tz_localize(None)
    return dt


def parse_and_clean(df: pd.DataFrame) -> pd.DataFrame:
    df.columns = df.columns.str.strip()
    df = df[df["Submission Time"].astype(str).str.strip().ne("")].copy()

    df["submission_dt"]   = to_ist_naive(df["Submission Time"])
    df["feedback_dt"]     = to_ist_naive(df["feedback_given_time"])
    df["submission_date"] = df["submission_dt"].dt.normalize()
    df["feedback_date"]   = df["feedback_dt"].dt.normalize()

    df["mentor_name"]       = df["mentor_name"].astype(str).str.strip()
    df["Evaluation Status"] = df["Evaluation Status"].astype(str).str.strip().str.lower()
    df["Module_name"]       = df["Module_name"].astype(str).str.strip()

    df["marks_obtained"]         = pd.to_numeric(df["marks_obtained"],         errors="coerce")
    df["marks_submission_level"] = pd.to_numeric(df["marks_submission_level"], errors="coerce")

    df["is_new_submission"] = df["marks_submission_level"] == 1
    df["is_resubmission"]   = df["marks_submission_level"] > 1
    df["score_passed"]      = df["marks_obtained"] >= SCORE_PASS_THRESHOLD

    df = df[df["Module_name"].isin(REPORT_MODULES)].copy()
    print(f"✓ After cleaning: {len(df):,} rows  |  Modules: {df['Module_name'].unique().tolist()}")
    return df


# ─────────────────────────────────────────────────────────────────────────────
# CELL 5 — Pending evaluations helper (current-snapshot logic)
# ─────────────────────────────────────────────────────────────────────────────

def _build_submitted_base(df: pd.DataFrame) -> pd.DataFrame:
    base = df[
        df["submission_dt"].notna() &
        df["Submission Status"].astype(str).str.strip().str.lower().eq("submitted")
    ].copy()
    batch_col    = base["Batch"].astype(str).str.strip()
    exclude_mask = batch_col.isin(EXCLUDED_BATCHES)
    base         = base[~exclude_mask].copy()
    print(f"  ↳ Excluded {exclude_mask.sum():,} rows from test/advantage batches")
    return base


PENDING_CUTOFF_DATE  = pd.Timestamp("2025-01-01")
OUTPUT_SHEET_PENDING = "pending_raw"


def pending_as_of(submitted: pd.DataFrame, as_of_date: pd.Timestamp, module: str) -> int:
    """
    Count students with a genuinely pending evaluation for `module` as of `as_of_date`.
    (Latest un-evaluated submission per student, not superseded by a later evaluation,
     and the student hasn't already passed the module.)
    """
    mod_sub = submitted[
        (submitted["Module_name"] == module) &
        (submitted["submission_date"] >= PENDING_CUTOFF_DATE) &
        (submitted["submission_date"] < as_of_date)
    ].copy()

    if mod_sub.empty:
        return 0

    mod_sub["is_evaluated"] = (
        mod_sub["marks_submission_level"].notna() &
        mod_sub["feedback_date"].notna() &
        (mod_sub["feedback_date"] <= as_of_date)
    )

    latest_eval_date = (
        mod_sub[mod_sub["is_evaluated"]]
        .groupby("user_id")["submission_date"].max()
        .rename("latest_eval_date")
    )
    mod_sub = mod_sub.join(latest_eval_date, on="user_id")
    mod_sub["latest_eval_date"] = mod_sub["latest_eval_date"].fillna(pd.NaT)

    unevaluated = mod_sub[~mod_sub["is_evaluated"]].copy()
    unevaluated = unevaluated[
        unevaluated["latest_eval_date"].isna() |
        (unevaluated["submission_date"] > unevaluated["latest_eval_date"])
    ]

    passed = set(
        mod_sub[
            mod_sub["marks_obtained"].notna() &
            (mod_sub["marks_obtained"] >= SCORE_PASS_THRESHOLD)
        ]["user_id"].unique()
    )
    unevaluated = unevaluated[~unevaluated["user_id"].isin(passed)]

    if unevaluated.empty:
        return 0

    latest_unevaluated = (
        unevaluated.sort_values("submission_date").groupby("user_id").tail(1)
    )
    return latest_unevaluated["user_id"].nunique()


def _in_window(series: pd.Series, start_h: int, end_h: int) -> pd.Series:
    h = series.dt.hour
    return (h >= start_h) & (h < end_h)


# ─────────────────────────────────────────────────────────────────────────────
# CELL 5b — Build pending raw (student-level list, today's snapshot)
# ─────────────────────────────────────────────────────────────────────────────

def build_pending_raw(df: pd.DataFrame) -> pd.DataFrame:
    submitted = _build_submitted_base(df)
    as_of     = pd.Timestamp.now(tz=IST).normalize().tz_localize(None)
    all_rows  = []

    for mod in REPORT_MODULES:
        mod_sub = submitted[
            (submitted["Module_name"] == mod) &
            (submitted["submission_date"] >= PENDING_CUTOFF_DATE)
        ].copy()
        if mod_sub.empty:
            continue

        mod_sub["is_evaluated"] = (
            mod_sub["marks_submission_level"].notna() &
            mod_sub["feedback_date"].notna() &
            (mod_sub["feedback_date"] <= as_of)
        )
        passed = set(
            mod_sub[
                mod_sub["marks_obtained"].notna() &
                (mod_sub["marks_obtained"] >= SCORE_PASS_THRESHOLD)
            ]["user_id"].unique()
        )

        before_today = mod_sub[mod_sub["submission_date"] < as_of].copy()
        latest_eval_date = (
            before_today[before_today["is_evaluated"]]
            .groupby("user_id")["submission_date"].max()
            .rename("latest_eval_date")
        )
        before_today = before_today.join(latest_eval_date, on="user_id")
        before_today["latest_eval_date"] = before_today["latest_eval_date"].fillna(pd.NaT)

        unevaluated = before_today[~before_today["is_evaluated"]].copy()
        unevaluated = unevaluated[
            unevaluated["latest_eval_date"].isna() |
            (unevaluated["submission_date"] > unevaluated["latest_eval_date"])
        ]
        unevaluated = unevaluated[~unevaluated["user_id"].isin(passed)]

        if not unevaluated.empty:
            pending_old = (
                unevaluated.sort_values("submission_date").groupby("user_id").tail(1)
            ).copy()
            pending_old["pending_type"] = "Pending"
            all_rows.append(pending_old)

        today_subs = mod_sub[
            (mod_sub["submission_date"] == as_of) &
            (~mod_sub["is_evaluated"]) &
            (~mod_sub["user_id"].isin(passed))
        ].copy()
        if not today_subs.empty:
            new_subs = (
                today_subs.sort_values("submission_dt").groupby("user_id").tail(1)
            ).copy()
            new_subs["pending_type"]     = "New Submission"
            new_subs["latest_eval_date"] = pd.NaT
            all_rows.append(new_subs)

    if not all_rows:
        return pd.DataFrame()

    result = pd.concat(all_rows, ignore_index=True)
    total_subs = (
        submitted.groupby(["user_id", "Module_name"]).size()
        .rename("total_submissions").reset_index()
    )
    result = result.merge(total_subs, on=["user_id", "Module_name"], how="left")

    result["Submission Time"]  = pd.to_datetime(result["Submission Time"], errors="coerce").dt.strftime("%Y-%m-%d %H:%M:%S")
    result["submission_date"]  = result["submission_date"].astype(str)
    result["latest_eval_date"] = result["latest_eval_date"].astype(str).replace("NaT", "Never evaluated")

    output = result[[
        "Module_name", "pending_type", "user_id", "student_name", "Batch",
        "submission_date", "Submission Time",
        "total_submissions", "marks_obtained", "latest_eval_date",
    ]].sort_values(["Module_name", "pending_type", "submission_date"]).reset_index(drop=True)

    print(f"✓ Pending raw: {len(output)} students")
    print(output.groupby(["Module_name", "pending_type"]).size().to_string())
    return output


# ─────────────────────────────────────────────────────────────────────────────
# CELL 6 — Daily summary
# ─────────────────────────────────────────────────────────────────────────────

def build_daily_summary(
    df: pd.DataFrame,
    target_date: "pd.Timestamp | None" = None,
) -> pd.DataFrame:

    sub        = df[df["submission_dt"].notna()].copy()
    night      = sub[_in_window(sub["submission_dt"], NIGHT_START_H, NIGHT_END_H)]
    day        = sub[_in_window(sub["submission_dt"], DAY_START_H, 24)]

    evaluated  = df[
        df["marks_submission_level"].notna() &
        df["feedback_dt"].notna()
    ].copy()

    submitted_base = _build_submitted_base(df)
    today = pd.Timestamp.now(tz=IST).normalize().tz_localize(None)

    if target_date is not None:
        all_dates = [pd.Timestamp(target_date)]
    else:
        earliest  = min(df["submission_dt"].dropna().min(), df["feedback_dt"].dropna().min())
        all_dates = pd.date_range(start=pd.Timestamp(earliest).normalize(), end=today, freq="D").tolist()

    # ── CURRENT pending snapshot (NOT per-date) ──────────────────────────────
    # Pending evaluations is a point-in-time state, not a daily event count. If we
    # wrote a value on every date, any date-range SUM in the dashboard would stack
    # the same un-evaluated submission once per day → inflated "cumulative" totals.
    # Instead we compute the pending count ONCE, as of the latest date in the
    # report (the "anchor" = today for a full run), and write it on that single
    # row only; every other date gets 0. So the dashboard's SUM over a range that
    # includes the anchor returns the true current pending, and any submission
    # already evaluated by now simply isn't counted.
    anchor_date = max(all_dates)
    current_pending = {
        m: pending_as_of(submitted_base, anchor_date + pd.Timedelta(days=1), m)
        for m in REPORT_MODULES
    }

    all_mentors = sorted(
        m for m in evaluated["mentor_name"].dropna().unique()
        if m and m not in ("nan", "")
    )

    rows = []
    for mod in REPORT_MODULES:
        for date in all_dates:
            prev_date = date - pd.Timedelta(days=1)

            n_night = int(((night["Module_name"] == mod) & (night["submission_date"] == date)).sum())
            n_day   = int(((day["Module_name"] == mod)   & (day["submission_date"]   == date)).sum())
            n_total = n_night + n_day

            last_eval   = evaluated[
                (evaluated["Module_name"] == mod) &
                (evaluated["feedback_date"] == prev_date)
            ]
            n_last_eval = len(last_eval)

            # Snapshot pending only on the anchor (latest) date; 0 everywhere else.
            n_pending = current_pending[mod] if date == anchor_date else 0

            row = {
                "Module"                        : mod,
                "Date"                          : date.date(),
                "Total Submissions (12AM-11AM)" : n_night,
                "New Add-ons (11AM-11:59PM)"    : n_day,
                "Total New Submissions"         : n_total,
                "Last Day Evaluations"          : n_last_eval,
                "Pending Evaluations"           : n_pending,
            }

            mentor_counts = last_eval.groupby("mentor_name").size()
            for mentor in all_mentors:
                row[f"By {mentor}"] = int(mentor_counts.get(mentor, 0))

            rows.append(row)

    report = pd.DataFrame(rows)
    metric_cols = [
        "Total Submissions (12AM-11AM)",
        "New Add-ons (11AM-11:59PM)",
        "Last Day Evaluations",
        "Pending Evaluations",
    ]
    report = report[report[metric_cols].sum(axis=1) > 0].reset_index(drop=True)

    print(f"✓ Daily summary: {len(report)} rows across {report['Module'].nunique()} module(s)")
    return report


# ─────────────────────────────────────────────────────────────────────────────
# CELL 7 — Mentor deep-dive
# ─────────────────────────────────────────────────────────────────────────────

def build_mentor_deep_dive(
    df: pd.DataFrame,
    target_date: "pd.Timestamp | None" = None,
) -> pd.DataFrame:

    evaluated = df[
        df["marks_submission_level"].notna() &
        df["feedback_dt"].notna() &
        df["mentor_name"].notna() &
        df["mentor_name"].ne("") &
        df["mentor_name"].ne("nan")
    ].copy()

    if target_date is not None:
        target_ts = pd.Timestamp(target_date).normalize()
        evaluated = evaluated[evaluated["feedback_date"] == target_ts]

    rows = []
    for (mod, mentor, fdate), grp in evaluated.groupby(
        ["Module_name", "mentor_name", "feedback_date"]
    ):
        g_lvl      = grp["marks_submission_level"]
        scores     = grp["marks_obtained"].dropna()
        new_subs   = int((g_lvl == 1).sum())
        total_resub= int((g_lvl > 1).sum())
        total      = len(grp)
        count_pass = int(grp["score_passed"].sum())
        count_fail = total - count_pass
        pass_rate  = round(count_pass / total * 100, 1) if total > 0 else 0.0
        min_score  = float(scores.min())            if len(scores) > 0 else None
        max_score  = float(scores.max())            if len(scores) > 0 else None
        avg_score  = round(float(scores.mean()), 2) if len(scores) > 0 else None

        row = {
            "Module"                    : mod,
            "Mentor"                    : mentor,
            "Evaluation Date"           : (fdate + pd.Timedelta(days=1)).date(),
            "New Submissions Evaluated" : new_subs,
        }
        for bucket_name, bucket_fn in RESUB_BUCKETS:
            row[bucket_name] = int(bucket_fn(g_lvl).sum())

        row["Total Re-Submissions Evaluated"] = total_resub
        row["Total Evaluations"]              = total
        row["Min Score Given"]                = min_score
        row["Max Score Given"]                = max_score
        row["Avg Score Given"]                = avg_score
        row[f"Count Score >= {SCORE_PASS_THRESHOLD} (Passed)"] = count_pass
        row[f"Count Score <  {SCORE_PASS_THRESHOLD} (Below)"]  = count_fail
        row["Pass Rate %"]                    = pass_rate
        rows.append(row)

    report = pd.DataFrame(rows).sort_values(
        ["Module", "Mentor", "Evaluation Date"]
    ).reset_index(drop=True)

    print(
        f"✓ Mentor deep-dive: {len(report)} rows  |  "
        f"{report['Mentor'].nunique()} mentor(s)  |  "
        f"{report['Module'].nunique()} module(s)"
    )
    return report


# ─────────────────────────────────────────────────────────────────────────────
# CELL 8 — Display helpers
# ─────────────────────────────────────────────────────────────────────────────

def display_daily_summary(report: pd.DataFrame) -> None:
    if report.empty:
        print("\n  No data found.\n")
        return
    pd.set_option("display.max_columns", None)
    pd.set_option("display.width", 220)
    mentor_cols = [c for c in report.columns if c.startswith("By ")]
    print("\n" + "=" * 85)
    print("  PROJECT EVALUATION — DAILY SUMMARY")
    print("=" * 85)
    for (mod, date), grp in report.groupby(["Module", "Date"]):
        r = grp.iloc[0]
        print(f"\n  Module : {mod}   |   Date : {date}")
        print(f"  {'-'*60}")
        print(f"  {'Total Submissions (12AM-11AM)':<42}: {int(r['Total Submissions (12AM-11AM)'])}")
        print(f"  {'New Add-ons (11AM-11:59PM)':<42}: {int(r['New Add-ons (11AM-11:59PM)'])}")
        print(f"  {'Total New Submissions':<42}: {int(r['Total New Submissions'])}")
        print(f"  {'Last Day Evaluations':<42}: {int(r['Last Day Evaluations'])}")
        print(f"  {'Pending Evaluations (current)':<42}: {int(r['Pending Evaluations'])}")
        active = [(c, int(r[c])) for c in mentor_cols if int(r[c]) > 0]
        if active:
            print(f"\n  Mentor breakdown (last day):")
            for col, cnt in active:
                print(f"    {col:<46}: {cnt}")
    print("\n" + "=" * 85 + "\n")


def display_mentor_deep_dive(report: pd.DataFrame) -> None:
    if report.empty:
        print("\n  No mentor data found.\n")
        return
    pd.set_option("display.max_columns", None)
    pd.set_option("display.width", 260)
    bucket_cols = [b[0] for b in RESUB_BUCKETS]
    print("\n" + "=" * 100)
    print("  MENTOR DEEP-DIVE REPORT")
    print("=" * 100)
    for (mod, mentor), grp in report.groupby(["Module", "Mentor"]):
        print(f"\n  Module : {mod}   |   Mentor : {mentor}")
        print(f"  {'-'*80}")
        for _, r in grp.iterrows():
            resub_parts = "  ".join(f"{b}={int(r[b])}" for b in bucket_cols if int(r[b]) > 0)
            resub_str   = f"  [{resub_parts}]" if resub_parts else ""
            print(
                f"  {str(r['Evaluation Date']):<14} | "
                f"New: {int(r['New Submissions Evaluated']):>3}  "
                f"Re-sub: {int(r['Total Re-Submissions Evaluated']):>3}{resub_str}  "
                f"Total: {int(r['Total Evaluations']):>3}  |  "
                f"min={r['Min Score Given']}  max={r['Max Score Given']}  avg={r['Avg Score Given']}  |  "
                f"Pass(≥{SCORE_PASS_THRESHOLD})={int(r[f'Count Score >= {SCORE_PASS_THRESHOLD} (Passed)']):>3}  "
                f"Below={int(r[f'Count Score <  {SCORE_PASS_THRESHOLD} (Below)']):>3}  "
                f"Pass%={r['Pass Rate %']}%"
            )
    print("\n" + "=" * 100 + "\n")


# ─────────────────────────────────────────────────────────────────────────────
# CELL 9 — Push to Google Sheets
# ─────────────────────────────────────────────────────────────────────────────

def _get_or_create_tab(spreadsheet, tab_name: str):
    try:
        return spreadsheet.worksheet(tab_name)
    except gspread.WorksheetNotFound:
        print(f"  Creating new tab '{tab_name}' …")
        return spreadsheet.add_worksheet(title=tab_name, rows=5000, cols=60)


def push_to_sheets(daily_report, mentor_report, pending_report) -> None:
    sheet = safe_open_by_key(OUTPUT_SHEET_KEY)
    for tab, data in [
        (OUTPUT_SHEET_DAILY,   daily_report),
        (OUTPUT_SHEET_MENTOR,  mentor_report),
        (OUTPUT_SHEET_PENDING, pending_report),
    ]:
        ws = _get_or_create_tab(sheet, tab)
        ws.clear()
        set_with_dataframe(ws, data, include_index=False, include_column_header=True)
        print(f"  ✓ Written {len(data)} rows to tab '{tab}'")


# ─────────────────────────────────────────────────────────────────────────────
# CELL 10 — CLI argument parsing
# ─────────────────────────────────────────────────────────────────────────────

def _parse_args():
    parser = argparse.ArgumentParser(description="Project Evaluation Report Generator")
    parser.add_argument("--date",    type=str, default=None,
                        help="Run for a single date (YYYY-MM-DD). Omit for full history.")
    parser.add_argument("--export",  type=str, default=None,
                        help="CSV export prefix. E.g. --export out → out_daily.csv, out_mentor.csv")
    parser.add_argument("--no-push", action="store_true",
                        help="Skip writing back to Google Sheets.")
    args, _ = parser.parse_known_args()
    return args


# ─────────────────────────────────────────────────────────────────────────────
# CELL 11 — Run
# ─────────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    args          = _parse_args()
    target_date   = pd.Timestamp(args.date) if args.date else None
    do_push       = not args.no_push
    export_prefix = args.export
else:
    target_date   = None
    do_push       = True
    export_prefix = None

raw = load_from_gsheets()
df = parse_and_clean(raw)

daily_report   = build_daily_summary(df,    target_date=target_date)
mentor_report  = build_mentor_deep_dive(df, target_date=target_date)
pending_report = build_pending_raw(df)

display_daily_summary(daily_report)
display_mentor_deep_dive(mentor_report)

if export_prefix:
    daily_report.to_csv(f"{export_prefix}_daily.csv",    index=False)
    mentor_report.to_csv(f"{export_prefix}_mentor.csv",  index=False)
    pending_report.to_csv(f"{export_prefix}_pending.csv", index=False)
    print(f"✓ Exported CSVs with prefix '{export_prefix}'")

if do_push:
    print("\nPushing to Google Sheets …")
    push_to_sheets(daily_report, mentor_report, pending_report)
    print("✓ All tabs updated.")
