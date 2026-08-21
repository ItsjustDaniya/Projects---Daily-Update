"""
Clearance-outcome views (student-level)
=======================================
SELF-CONTAINED companion to days_x_clearance_view.py — same data, same
definitions (Evaluated = marks_submission_level present & feedback present;
Cleared = evaluated AND marks_submission_level >= CLEAR_PASS_LEVEL).

Eight tidy LONG tables, each Looker-Studio ready (one tab each):

  1. attempts_to_clear        — distribution of how many submissions a student
                                needed before first clearing (per Batch)
  2. days_to_clear            — distribution of days from release to a student's
                                first clearing submission (per Batch)
  3. score_distribution       — histogram of marks_submission_level (1–10) with
                                # submissions, # users, % (per Batch)
  4. score_lift_by_attempt    — avg score & clearance rate by attempt number
                                (the learning curve, per Batch)
  5. score_dist_by_attempt    — percentile spread of scores (Min/P10/P25/Median/
                                P75/P90/Max/Avg) at each submission attempt number,
                                showing how the distribution shifts across retries
  6. score_lift_by_day_bucket — score lift across attempts, split by how soon
                                the student first submitted (per Batch)
  7. score_dist_mom           — MoM × Module: percentile spread of evaluated
                                submission scores, bucketed by submission month
  8. wow_submissions          — WoW × Module: submissions, unique users, and
                                % of users clearing, bucketed by submission week

Grain note: views 1–2 are at the unique-USER level (one student counted once
per module). Unique-user counts are NOT additive — don't SUM "Users" across
batches in Looker; the % columns are valid at the Batch row grain.

Views 7–8 are grouped by Module × time period (aggregated across batches, as
requested). Add "Batch" to the relevant groupby if you want them split out.

Data source: reads from Google Sheets using `gc` already defined in your notebook.
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

CLEAR_PASS_LEVEL      = 8                       # cleared = evaluated AND score >= this
CLEARANCE_DAY_BUCKETS = [15, 30, 45, 60, 75]    # days-to-clear windows (cumulative)
ATTEMPT_BUCKETS       = [1, 2, 3, 5, 10]        # attempts-to-clear (cumulative)
SCORE_LEVELS          = list(range(1, 11))       # marks_submission_level 1..10
SCORE_BASIS           = "first"                  # which per-student score to distribute:
                                                 # "first"  = earliest evaluated attempt
                                                 # "best"   = highest score reached
                                                 # "latest" = most recent attempt
SCORE_LIFT_MAX_ATTEMPT     = 10                  # attempts beyond this → "10+"
SCORE_DIST_MAX_ATTEMPT     = 5                   # attempts beyond this → "5+" in view 5
DXC_MODULES                = None               # None ⇒ all REPORT_MODULES

def _attempt_label(n): return "1" if n == 1 else f"<{n}"

OUTPUT_SHEET_ATC          = "attempts_to_clear"
OUTPUT_SHEET_DTC          = "days_to_clear"
OUTPUT_SHEET_SCORE        = "score_distribution"
OUTPUT_SHEET_LIFT         = "score_lift_by_attempt"
OUTPUT_SHEET_SCORE_BY_ATT = "score_dist_by_attempt"
OUTPUT_SHEET_LIFT_BY_DAY  = "score_lift_by_day_bucket"
OUTPUT_SHEET_SCORE_MOM    = "score_dist_mom"        # view 7
OUTPUT_SHEET_WOW          = "wow_submissions"       # view 8


# ─────────────────────────────────────────────────────────────────────────────
# Load + clean
# ─────────────────────────────────────────────────────────────────────────────
def load_from_gsheets() -> pd.DataFrame:
    spreadsheet = gc.open(SPREADSHEET_NAME_2)            # noqa: F821
    worksheet   = spreadsheet.worksheet(MENTOR_SHEET)
    print(f"✓ Connected to '{SPREADSHEET_NAME_2}' → '{MENTOR_SHEET}'")
    records = worksheet.get_all_records(numericise_ignore=["all"])
    df = pd.DataFrame(records)
    print(f"✓ Fetched {len(df):,} rows  |  {len(df.columns)} columns")
    return df


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
        raise KeyError("Expected a 'project_release_date' column.")

    df = df[df["Submission Time"].astype(str).str.strip().ne("")].copy()
    df["submission_dt"]          = to_ist_naive(df["Submission Time"])
    df["feedback_dt"]            = to_ist_naive(df["feedback_given_time"])
    df["release_dt"]             = to_ist_naive(df["project_release_date"])
    df["Module_name"]            = df["Module_name"].astype(str).str.strip()
    df["marks_submission_level"] = pd.to_numeric(df["marks_submission_level"], errors="coerce")

    df["is_evaluated"] = df["marks_submission_level"].notna() & df["feedback_dt"].notna()
    df["days_from_release"] = (df["submission_dt"] - df["release_dt"]).dt.total_seconds() / 86400

    df = df[df["submission_dt"].notna()].copy()
    df = df[df["Module_name"].isin(REPORT_MODULES)].copy()
    df = df[~df["Batch"].astype(str).str.strip().isin(EXCLUDED_BATCHES)].copy()

    before = len(df)
    unevaluated = (df[~df["is_evaluated"]].sort_values("submission_dt")
                   .drop_duplicates(subset=["submission_id"], keep="last"))
    evaluated   = (df[df["is_evaluated"]].sort_values("feedback_dt")
                   .drop_duplicates(subset=["submission_id"], keep="first"))
    df = pd.concat([unevaluated, evaluated], ignore_index=True).sort_values("submission_dt")
    print(f"  ↳ Removed {before - len(df):,} duplicate rows (by submission_id)")

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

    print(f"✓ After cleaning: {len(df):,} rows")
    return df


# ─────────────────────────────────────────────────────────────────────────────
# Shared prep
# ─────────────────────────────────────────────────────────────────────────────
def _pct(num, den):
    return round(100 * num / den, 1) if den else None


def _prep(df) -> pd.DataFrame:
    """Add attempt rank (per user, module) and the cleared flag."""
    mods = DXC_MODULES or REPORT_MODULES
    d = (df[df["Module_name"].isin(mods)]
         .sort_values(["user_id", "Module_name", "submission_dt"]).copy())
    d["attempt"]  = d.groupby(["user_id", "Module_name"]).cumcount() + 1
    d["_cleared"] = d["is_evaluated"] & (d["marks_submission_level"] >= CLEAR_PASS_LEVEL)
    return d


def _per_user_module(d) -> pd.DataFrame:
    """One row per (user, module): batch, total subs, attempts/days to FIRST clear."""
    g = d.groupby(["user_id", "Module_name"], sort=False)
    base = g.agg(Batch=("Batch", "first"), total_subs=("submission_id", "size")).reset_index()
    first_clear = (d[d["_cleared"]].sort_values(["user_id", "Module_name", "attempt"])
                   .groupby(["user_id", "Module_name"], sort=False).first().reset_index())
    fc = (first_clear[["user_id", "Module_name", "attempt", "days_from_release"]]
          .rename(columns={"attempt": "attempts_to_clear", "days_from_release": "days_to_clear"}))
    base = base.merge(fc, on=["user_id", "Module_name"], how="left")
    base["cleared"] = base["attempts_to_clear"].notna()
    return base


# ─────────────────────────────────────────────────────────────────────────────
# 1) Attempts-to-clear distribution
# ─────────────────────────────────────────────────────────────────────────────
def build_attempts_to_clear(df) -> pd.DataFrame:
    pum = _per_user_module(_prep(df))
    spec = ([("Total", None, "total"), ("Cleared (any)", None, "cleared")]
            + [(_attempt_label(n), n, "cum") for n in ATTEMPT_BUCKETS]
            + [("Never cleared", None, "never")])
    rows = []
    for (mod, batch), g in pum.groupby(["Module_name", "Batch"]):
        total = len(g)
        cleared = int(g["cleared"].sum())
        for i, (label, thr, kind) in enumerate(spec):
            if   kind == "total":   cnt = total
            elif kind == "cleared": cnt = cleared
            elif kind == "never":   cnt = total - cleared
            else:                   cnt = int((g["attempts_to_clear"] <= thr).sum())
            rows.append({"Module": mod, "Batch": batch,
                         "Attempts-to-clear bucket": label, "bucket_order": i,
                         "Users": cnt, "% users": _pct(cnt, total)})
    report = (pd.DataFrame(rows)
              .sort_values(["Module", "Batch", "bucket_order"]).reset_index(drop=True))
    print(f"✓ attempts_to_clear: {len(report)} rows")
    return report


# ─────────────────────────────────────────────────────────────────────────────
# 2) Days-to-clear distribution
# ─────────────────────────────────────────────────────────────────────────────
def build_days_to_clear(df) -> pd.DataFrame:
    pum = _per_user_module(_prep(df))
    spec = ([("Total", None, "total"), ("Cleared (any)", None, "cleared")]
            + [(f"<{x} days", x, "cum") for x in CLEARANCE_DAY_BUCKETS]
            + [("Never cleared", None, "never")])
    rows = []
    for (mod, batch), g in pum.groupby(["Module_name", "Batch"]):
        total = len(g)
        cleared = int(g["cleared"].sum())
        med = round(g["days_to_clear"].median(), 1) if cleared else None
        for i, (label, thr, kind) in enumerate(spec):
            if   kind == "total":   cnt = total
            elif kind == "cleared": cnt = cleared
            elif kind == "never":   cnt = total - cleared
            else:                   cnt = int((g["days_to_clear"] <= thr).sum())
            rows.append({"Module": mod, "Batch": batch,
                         "Days-to-clear bucket": label, "bucket_order": i,
                         "Users": cnt, "% users": _pct(cnt, total),
                         "Median days-to-clear": med})
    report = (pd.DataFrame(rows)
              .sort_values(["Module", "Batch", "bucket_order"]).reset_index(drop=True))
    print(f"✓ days_to_clear: {len(report)} rows")
    return report


# ─────────────────────────────────────────────────────────────────────────────
# 3) Score distribution — per STUDENT, by their FIRST/BEST/LATEST score
# ─────────────────────────────────────────────────────────────────────────────
def build_score_distribution(df) -> pd.DataFrame:
    d = _prep(df)
    ev = d[d["is_evaluated"]].sort_values(["user_id", "Module_name", "attempt"])
    g = ev.groupby(["user_id", "Module_name"], sort=False)["marks_submission_level"]
    rep = {"best": g.max(), "latest": g.last()}.get(SCORE_BASIS, g.first())
    score_col = {"first": "First-attempt score", "best": "Best score",
                 "latest": "Latest score"}[SCORE_BASIS]
    rep = rep.rename("rep_score").reset_index()
    students = (d.groupby(["user_id", "Module_name"])
                .agg(Batch=("Batch", "first")).reset_index()
                .merge(rep, on=["user_id", "Module_name"], how="left"))

    rows = []
    for (mod, batch), grp in students.groupby(["Module_name", "Batch"]):
        total = len(grp)
        scores = grp["rep_score"].dropna()
        pcts = {
            "Avg score":    round(scores.mean(), 2) if len(scores) else None,
            "P25 score":    round(scores.quantile(0.25), 1) if len(scores) else None,
            "Median score": round(scores.quantile(0.50), 1) if len(scores) else None,
            "P75 score":    round(scores.quantile(0.75), 1) if len(scores) else None,
            "P90 score":    round(scores.quantile(0.90), 1) if len(scores) else None,
        }
        cum = 0
        def emit(label, order, count, cumulative=None):
            rows.append({"Module": mod, "Batch": batch,
                         score_col: label, "level_order": order,
                         "# students": int(count), "% students": _pct(int(count), total),
                         "Cumulative % students": (_pct(int(cumulative), total)
                                                   if cumulative is not None else None),
                         **pcts})
        emit("Total", 0, total)
        for L in SCORE_LEVELS:
            c = int((grp["rep_score"] == L).sum())
            cum += c
            emit(str(L), L, c, cumulative=cum)
        emit("Not yet evaluated", len(SCORE_LEVELS) + 1, grp["rep_score"].isna().sum())
        emit(f"≥{CLEAR_PASS_LEVEL}", len(SCORE_LEVELS) + 2,
             (grp["rep_score"] >= CLEAR_PASS_LEVEL).sum())
    report = (pd.DataFrame(rows)
              .sort_values(["Module", "Batch", "level_order"]).reset_index(drop=True))
    print(f"✓ score_distribution ({SCORE_BASIS}): {len(report)} rows")
    return report


# ─────────────────────────────────────────────────────────────────────────────
# 4) Score lift by attempt — avg score & clearance rate per attempt number
# ─────────────────────────────────────────────────────────────────────────────
def _lift_summary(ev) -> dict:
    """Per (module, batch): avg 1st-attempt score, avg clearing score, avg lift."""
    e = ev.sort_values(["user_id", "Module_name", "attempt"])
    g = e.groupby(["user_id", "Module_name"], sort=False)
    first = g["marks_submission_level"].first()
    batch = g["Batch"].first()
    clear = (e[e["marks_submission_level"] >= CLEAR_PASS_LEVEL]
             .groupby(["user_id", "Module_name"], sort=False)["marks_submission_level"].first())
    s = pd.DataFrame({"Batch": batch, "first_score": first, "clear_score": clear}).reset_index()
    s["lift"] = s["clear_score"] - s["first_score"]
    out = {}
    for (mod, b), grp in s.groupby(["Module_name", "Batch"]):
        cl = grp[grp["clear_score"].notna()]
        n = len(cl)
        out[(mod, b)] = {
            "Avg 1st-attempt score (cleared)": round(cl["first_score"].mean(), 2) if n else None,
            "Avg clearing score":              round(cl["clear_score"].mean(), 2) if n else None,
            "Avg lift 1st→clear":              round(cl["lift"].mean(), 2) if n else None,
        }
    return out


def build_score_lift(df) -> pd.DataFrame:
    ev = _prep(df)
    ev = ev[ev["is_evaluated"]]
    lift_sum = _lift_summary(ev)
    rows = []
    for (mod, batch), g in ev.groupby(["Module_name", "Batch"]):
        ls = lift_sum.get((mod, batch), {})
        def emit(label, order, sub):
            if len(sub) == 0:
                return
            s = sub["marks_submission_level"]
            rows.append({
                "Module": mod, "Batch": batch,
                "Attempt": label, "attempt_order": order,
                "# submissions": len(sub),
                "# users": int(sub["user_id"].nunique()),
                "Avg score": round(s.mean(), 2),
                "Min": round(s.min(), 1),
                "P25": round(s.quantile(0.25), 1),
                "Median": round(s.median(), 1),
                "P75": round(s.quantile(0.75), 1),
                "Max": round(s.max(), 1),
                "% cleared at attempt": _pct(int((s >= CLEAR_PASS_LEVEL).sum()), len(sub)),
                **ls,
            })
        for a in range(1, SCORE_LIFT_MAX_ATTEMPT + 1):
            emit(str(a), a, g[g["attempt"] == a])
        emit(f"{SCORE_LIFT_MAX_ATTEMPT}+", SCORE_LIFT_MAX_ATTEMPT + 1,
             g[g["attempt"] > SCORE_LIFT_MAX_ATTEMPT])
    report = (pd.DataFrame(rows)
              .sort_values(["Module", "Batch", "attempt_order"]).reset_index(drop=True))
    print(f"✓ score_lift_by_attempt: {len(report)} rows")
    return report


# ─────────────────────────────────────────────────────────────────────────────
# 5) Score distribution by submission attempt
# ─────────────────────────────────────────────────────────────────────────────
# Shows the full score spread (Min/P10/P25/Median/P75/P90/Max/Avg) for ALL
# evaluated submissions at each attempt number — so you can see whether the
# distribution improves after each resubmission.
# Grain: one row per Module × Batch × Attempt bucket.
def build_score_dist_by_attempt(df) -> pd.DataFrame:
    ev = _prep(df)
    ev = ev[ev["is_evaluated"]].copy()

    def _label(a): return str(a) if a <= SCORE_DIST_MAX_ATTEMPT else f"{SCORE_DIST_MAX_ATTEMPT}+"
    def _order(a): return a      if a <= SCORE_DIST_MAX_ATTEMPT else SCORE_DIST_MAX_ATTEMPT + 1

    ev["attempt_label"] = ev["attempt"].apply(_label)
    ev["attempt_order"] = ev["attempt"].apply(_order)

    rows = []
    for (mod, batch, alabel, aorder), grp in ev.groupby(
        ["Module_name", "Batch", "attempt_label", "attempt_order"], sort=False
    ):
        s = grp["marks_submission_level"].dropna()
        if len(s) == 0:
            continue
        rows.append({
            "Module":                    mod,
            "Batch":                     batch,
            "Attempt":                   alabel,
            "attempt_order":             aorder,
            "# submissions":             len(s),
            "# users":                   int(grp["user_id"].nunique()),
            "Min":                       round(s.min(), 1),
            "P10":                       round(s.quantile(0.10), 1),
            "P25":                       round(s.quantile(0.25), 1),
            "Median":                    round(s.quantile(0.50), 1),
            "P75":                       round(s.quantile(0.75), 1),
            "P90":                       round(s.quantile(0.90), 1),
            "Max":                       round(s.max(), 1),
            "Avg score":                 round(s.mean(), 2),
            f"% scored ≥{CLEAR_PASS_LEVEL}": _pct(int((s >= CLEAR_PASS_LEVEL).sum()), len(s)),
            "% scored ≥6":               _pct(int((s >= 6).sum()), len(s)),
        })

    report = (pd.DataFrame(rows)
              .sort_values(["Module", "Batch", "attempt_order"]).reset_index(drop=True))
    print(f"✓ score_dist_by_attempt: {len(report)} rows")
    return report


# ─────────────────────────────────────────────────────────────────────────────
# 6) Score lift by day bucket × attempt
# ─────────────────────────────────────────────────────────────────────────────
# Assigns each student to a "first submission" day bucket based on how many
# days after release they made their FIRST submission (any attempt, not just
# evaluated). Then, within each bucket, shows the score lift across attempt
# numbers — so you can compare whether early starters improve faster.
#
# Grain: Module × Batch × First-sub day bucket × Attempt
# Day buckets are CUMULATIVE (≤15 ⊂ ≤30 ⊂ … ≤75 ⊂ Any). Filter to one
# bucket at a time in Looker Studio, or use as a breakdown dimension.
def build_score_lift_by_day_bucket(df) -> pd.DataFrame:
    d = _prep(df)
    ev = d[d["is_evaluated"]].copy()

    # ── per-user first submission day (any submission, not just evaluated) ──
    first_sub_day = (
        d.groupby(["user_id", "Module_name"])["days_from_release"]
        .min()
        .rename("first_sub_days")
        .reset_index()
    )

    # ── per-user attempt-1 score (first EVALUATED attempt) ──────────────────
    first_score = (
        ev.sort_values(["user_id", "Module_name", "attempt"])
        .groupby(["user_id", "Module_name"], sort=False)["marks_submission_level"]
        .first()
        .rename("first_score")
        .reset_index()
    )

    # ── per-user ever-cleared flag + attempt at which they first cleared ─────
    ever_cleared = (
        ev[ev["_cleared"]]
        .sort_values(["user_id", "Module_name", "attempt"])
        .groupby(["user_id", "Module_name"], sort=False)["attempt"]
        .first()
        .rename("clear_attempt")
        .reset_index()
    )

    # attach to every evaluated submission row
    ev = (ev
          .merge(first_sub_day, on=["user_id", "Module_name"], how="left")
          .merge(first_score,   on=["user_id", "Module_name"], how="left")
          .merge(ever_cleared,  on=["user_id", "Module_name"], how="left"))

    # day-bucket labels (cumulative — student appears in ALL buckets they qualify for)
    day_buckets = [(d_, f"≤{d_} days", i) for i, d_ in enumerate(CLEARANCE_DAY_BUCKETS)]
    day_buckets.append((None, "Any", len(CLEARANCE_DAY_BUCKETS)))   # catch-all

    # attempt bucketing (reuse SCORE_LIFT_MAX_ATTEMPT)
    def _alabel(a): return str(a) if a <= SCORE_LIFT_MAX_ATTEMPT else f"{SCORE_LIFT_MAX_ATTEMPT}+"
    def _aorder(a): return a      if a <= SCORE_LIFT_MAX_ATTEMPT else SCORE_LIFT_MAX_ATTEMPT + 1

    ev["attempt_label"] = ev["attempt"].apply(_alabel)
    ev["attempt_order"] = ev["attempt"].apply(_aorder)

    rows = []
    for (mod, batch), g in ev.groupby(["Module_name", "Batch"], sort=False):

        for (day_thr, day_label, day_order) in day_buckets:
            # filter to students whose first submission falls in this bucket
            if day_thr is not None:
                seg = g[g["first_sub_days"] <= day_thr]
            else:
                seg = g.copy()

            if seg.empty:
                continue

            users_in_seg = seg["user_id"].nunique()

            # cumulative cleared tracker (reset per bucket)
            cleared_so_far = set()

            for (alabel, aorder), ag in (seg.groupby(["attempt_label", "attempt_order"],
                                                      sort=False)):
                s = ag["marks_submission_level"].dropna()
                if len(s) == 0:
                    continue

                # update cumulative cleared set
                newly_cleared = ag.loc[
                    ag["clear_attempt"].notna() & (ag["clear_attempt"] <= ag["attempt"]),
                    "user_id"
                ].unique()
                cleared_so_far.update(newly_cleared)

                avg_lift = (
                    round((s - ag["first_score"]).mean(), 2)
                    if ag["first_score"].notna().any() else None
                )

                rows.append({
                    "Module":                    mod,
                    "Batch":                     batch,
                    "First-sub bucket":          day_label,
                    "bucket_order":              day_order,
                    "Attempt":                   alabel,
                    "attempt_order":             aorder,
                    "# users":                   int(ag["user_id"].nunique()),
                    "Avg first-attempt score":   round(ag["first_score"].mean(), 2)
                                                 if ag["first_score"].notna().any() else None,
                    "Avg score at attempt":      round(s.mean(), 2),
                    "Avg lift vs attempt 1":     avg_lift,
                    "% cleared at attempt":      _pct(int((s >= CLEAR_PASS_LEVEL).sum()), len(s)),
                    "% ever cleared by attempt": _pct(len(cleared_so_far), users_in_seg),
                })

    report = (
        pd.DataFrame(rows)
        .sort_values(["Module", "Batch", "bucket_order", "attempt_order"])
        .reset_index(drop=True)
    )
    print(f"✓ score_lift_by_day_bucket: {len(report)} rows")
    return report


# ─────────────────────────────────────────────────────────────────────────────
# 7) MoM × Module — project score distribution (percentile spread)   ← NEW
# ─────────────────────────────────────────────────────────────────────────────
# Percentile spread of evaluated-submission scores per Module, bucketed by the
# MONTH of submission. One row per Module × Month — a clean MoM trend you can
# plot Median/Avg/percentiles against time. Aggregated across batches; add
# "Batch" to the groupby keys if you want it split out.
def build_score_dist_mom(df) -> pd.DataFrame:
    ev = _prep(df)
    ev = ev[ev["is_evaluated"]].copy()

    dtm = ev["submission_dt"].dt
    ev["Month"]       = dtm.year.astype(str) + "-" + dtm.month.astype(str).str.zfill(2)
    ev["month_order"] = dtm.year * 12 + dtm.month

    rows = []
    for (mod, month, morder), grp in ev.groupby(
        ["Module_name", "Month", "month_order"], sort=False
    ):
        s = grp["marks_submission_level"].dropna()
        if len(s) == 0:
            continue
        rows.append({
            "Module":                       mod,
            "Month":                        month,
            "month_order":                  morder,
            "# submissions":                len(s),
            "# users":                      int(grp["user_id"].nunique()),
            "Min":                          round(s.min(), 1),
            "P10":                          round(s.quantile(0.10), 1),
            "P25":                          round(s.quantile(0.25), 1),
            "Median":                       round(s.quantile(0.50), 1),
            "P75":                          round(s.quantile(0.75), 1),
            "P90":                          round(s.quantile(0.90), 1),
            "Max":                          round(s.max(), 1),
            "Avg score":                    round(s.mean(), 2),
            f"% scored ≥{CLEAR_PASS_LEVEL}": _pct(int((s >= CLEAR_PASS_LEVEL).sum()), len(s)),
            "% scored ≥6":                  _pct(int((s >= 6).sum()), len(s)),
        })

    report = (pd.DataFrame(rows)
              .sort_values(["Module", "month_order"]).reset_index(drop=True))
    print(f"✓ score_dist_mom: {len(report)} rows")
    return report


# ─────────────────────────────────────────────────────────────────────────────
# 8) WoW × Module — submissions, unique users, % of users clearing   ← NEW
# ─────────────────────────────────────────────────────────────────────────────
# Weekly submission activity per Module, bucketed by submission week (Mon–Sun).
#   % users clearing = unique users with a cleared (score ≥ CLEAR_PASS_LEVEL)
#                      submission that week ÷ unique users who submitted that week
# Grain: one row per Module × Week. Aggregated across batches; add "Batch" to
# the groupby keys if you want it split out.
def build_wow_submissions(df) -> pd.DataFrame:
    d = _prep(df)

    ws = (d["submission_dt"].dt.normalize()
          - pd.to_timedelta(d["submission_dt"].dt.weekday, unit="D"))   # Monday of week
    iso = d["submission_dt"].dt.isocalendar()
    d = d.assign(
        week_start = ws.dt.strftime("%Y-%m-%d"),
        week_label = iso["year"].astype(str) + "-W" + iso["week"].astype(str).str.zfill(2),
        week_order = ws.astype("int64"),
    )

    rows = []
    for (mod, wlabel, wstart, worder), grp in d.groupby(
        ["Module_name", "week_label", "week_start", "week_order"], sort=False
    ):
        cleared    = grp[grp["_cleared"]]
        n_users    = int(grp["user_id"].nunique())
        n_cl_users = int(cleared["user_id"].nunique())
        rows.append({
            "Module":           mod,
            "Week":             wlabel,
            "Week start":       wstart,
            "week_order":       worder,
            "# submissions":    len(grp),
            "# unique users":   n_users,
            "# evaluated subs": int(grp["is_evaluated"].sum()),
            "# cleared subs":   len(cleared),
            "# users cleared":  n_cl_users,
            "% users clearing": _pct(n_cl_users, n_users),
        })

    report = (pd.DataFrame(rows)
              .sort_values(["Module", "week_order"]).reset_index(drop=True))
    print(f"✓ wow_submissions: {len(report)} rows")
    return report


# ─────────────────────────────────────────────────────────────────────────────
# Push helpers
# ─────────────────────────────────────────────────────────────────────────────
def _get_or_create_tab(spreadsheet, tab_name):
    try:
        return spreadsheet.worksheet(tab_name)
    except gspread.WorksheetNotFound:
        print(f"  Creating new tab '{tab_name}' …")
        return spreadsheet.add_worksheet(title=tab_name, rows=5000, cols=25)


def push_df(spreadsheet_key, tab_name, data):
    sheet = gc.open_by_key(spreadsheet_key)                 # noqa: F821
    ws = _get_or_create_tab(sheet, tab_name)
    ws.clear()
    set_with_dataframe(ws, data, include_index=False, include_column_header=True)
    print(f"  ✓ Written {len(data)} rows to tab '{tab_name}'")


# ─────────────────────────────────────────────────────────────────────────────
# RUN  — reads from Google Sheets using `gc` already defined in your notebook
# ─────────────────────────────────────────────────────────────────────────────
DXC_DO_PUSH = True     # False ⇒ build & preview only

raw = load_from_gsheets()
df  = parse_and_clean(raw)

atc          = build_attempts_to_clear(df)
dtc          = build_days_to_clear(df)
score        = build_score_distribution(df)
lift         = build_score_lift(df)
score_dist   = build_score_dist_by_attempt(df)
lift_by_day  = build_score_lift_by_day_bucket(df)
score_mom    = build_score_dist_mom(df)            # view 7
wow          = build_wow_submissions(df)           # view 8

with pd.option_context("display.max_columns", None, "display.width", 220):
    b = atc["Batch"].iloc[0]
    print(f"\n--- attempts_to_clear (sample batch: {b}) ---")
    print(atc[atc["Batch"] == b].to_string(index=False))
    print(f"\n--- days_to_clear (sample) ---")
    print(dtc[dtc["Batch"] == b].to_string(index=False))
    print(f"\n--- score_distribution (sample) ---")
    print(score[score["Batch"] == b].to_string(index=False))
    print(f"\n--- score_lift_by_attempt (sample) ---")
    print(lift[lift["Batch"] == b].to_string(index=False))
    print(f"\n--- score_dist_by_attempt (sample) ---")
    m = score_dist["Module"].iloc[0]
    print(score_dist[(score_dist["Batch"] == b) & (score_dist["Module"] == m)]
          .drop(columns=["attempt_order"]).to_string(index=False))
    print(f"\n--- score_lift_by_day_bucket (sample: ≤30 days bucket) ---")
    m = lift_by_day["Module"].iloc[0]
    print(lift_by_day[
        (lift_by_day["Batch"] == b) &
        (lift_by_day["Module"] == m) &
        (lift_by_day["First-sub bucket"] == "≤30 days")
    ].drop(columns=["bucket_order", "attempt_order"]).to_string(index=False))

    print(f"\n--- score_dist_mom (sample module) ---")
    m = score_mom["Module"].iloc[0]
    print(score_mom[score_mom["Module"] == m]
          .drop(columns=["month_order"]).to_string(index=False))
    print(f"\n--- wow_submissions (sample module) ---")
    m = wow["Module"].iloc[0]
    print(wow[wow["Module"] == m]
          .drop(columns=["week_order"]).to_string(index=False))

if DXC_DO_PUSH:
    print("\nPushing eight tabs to Google Sheets …")
    push_df(OUTPUT_SHEET_KEY, OUTPUT_SHEET_ATC,          atc)
    push_df(OUTPUT_SHEET_KEY, OUTPUT_SHEET_DTC,          dtc)
    push_df(OUTPUT_SHEET_KEY, OUTPUT_SHEET_SCORE,        score)
    push_df(OUTPUT_SHEET_KEY, OUTPUT_SHEET_LIFT,         lift)
    push_df(OUTPUT_SHEET_KEY, OUTPUT_SHEET_SCORE_BY_ATT, score_dist)
    push_df(OUTPUT_SHEET_KEY, OUTPUT_SHEET_LIFT_BY_DAY,  lift_by_day)
    push_df(OUTPUT_SHEET_KEY, OUTPUT_SHEET_SCORE_MOM,    score_mom)
    push_df(OUTPUT_SHEET_KEY, OUTPUT_SHEET_WOW,          wow)
    print("✓ All eight tabs updated.")