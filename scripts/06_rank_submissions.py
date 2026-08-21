# %% ─────────────────────────────────────────────────────────────────────────
#  RANK SUBMISSIONS  →  reproduces  Project_RCA_-_submissions_ranked.csv
#
#  Logic (verified to 100% against the sample output):
#    1. Drop unevaluated submissions  (Evaluation Status != "Evaluated").
#       → this removes any un-evaluated attempt sitting *between* evaluated ones,
#         so the rank reflects the evaluated-attempt sequence only.
#    2. Within each (User ID, question_id), rank by submission_id ascending
#       using DENSE ranking → submission_rank.
#         · submission_id is the authoritative creation order of a submission.
#         · dense  ⇒ duplicate/identical submission rows share one rank and the
#                    sequence stays contiguous (1,2,2,3 … not 1,2,2,4).
#         · rows with a blank submission_id get a null rank (kept, unranked).
# ───── ────────────────────────────────────────────────────────────────────────

# %% ── Config ────────────────────────────────────────────────────────────────
SPREADSHEET_NAME_2 = "calender - 2"
MENTOR_SHEET       = "Project_evaluations"          # input worksheet
OUTPUT_SHEET_KEY   = "1CbJL4twtn38a8FbFkkXxZ_HNXFStlgwqLPdPQgwY5zI"
OUTPUT_TAB         = "submissions_ranked"            # output worksheet (created/cleared)

RANK_COL          = "submission_rank"
GROUP_KEYS        = ["user_id", "question_id"]       # rank is independent per student-per-question
ORDER_KEY         = "submission_id"                  # authoritative attempt order
RANK_METHOD       = "dense"                           # ties share a rank, sequence stays contiguous
EVAL_COL          = "Evaluation Status"
EVAL_VALUE        = "Evaluated"
KEEP_UNEVALUATED  = False                             # False ⇒ eliminate unevaluated (the requirement)

WRITE_TO_SHEET    = True                              # push result to OUTPUT_SHEET_KEY / OUTPUT_TAB
WRITE_CSV_PATH    = "submissions_ranked.csv"          # also dump a local CSV (set None to skip)


# %% ── Imports ───────────────────────────────────────────────────────────────
import numpy as np
import pandas as pd


# %% ── Core transform (pure pandas, no I/O) ──────────────────────────────────
def _to_num(series: pd.Series) -> pd.Series:
    """Coerce sheet-formatted numbers like '174,956' or ' 12,865 ' to numeric."""
    return pd.to_numeric(
        series.astype(str).str.replace(",", "", regex=False).str.strip(),
        errors="coerce",
    )


def rank_submissions(df: pd.DataFrame,
                     group_keys=GROUP_KEYS,
                     order_key=ORDER_KEY,
                     rank_method=RANK_METHOD,
                     eval_col=EVAL_COL,
                     eval_value=EVAL_VALUE,
                     keep_unevaluated=KEEP_UNEVALUATED,
                     rank_col=RANK_COL) -> pd.DataFrame:
    """Add `submission_rank` and (by default) drop unevaluated submissions.

    Returns a copy with `rank_col` inserted as the first column; all other
    columns and their original order are preserved.
    """
    work = df.copy()

    # 1) eliminate unevaluated submissions (the in-between ones disappear,
    #    so they never consume a rank position)
    if not keep_unevaluated:
        mask = work[eval_col].astype(str).str.strip().str.casefold() == eval_value.casefold()
        work = work[mask].copy()

    # 2) dense rank by creation order within each (student, question)
    work["_order"] = _to_num(work[order_key])
    work[rank_col] = (
        work.groupby(group_keys, dropna=False)["_order"]
            .rank(method=rank_method)
            .astype("Int64")            # nullable int → blank submission_id stays <NA>
    )
    work = work.drop(columns="_order")

    # 3) move the rank column to the front (matches the reference layout)
    ordered = [rank_col] + [c for c in work.columns if c != rank_col]
    return work[ordered]


# %% ── Google Sheets I/O helpers ─────────────────────────────────────────────
import os
import sys

import gspread
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
from common.sheets_auth import get_client

gc = get_client()

from gspread_dataframe import set_with_dataframe, get_as_dataframe


def read_worksheet(spreadsheet_name: str, worksheet_name: str) -> pd.DataFrame:
    ws = gc.open(spreadsheet_name).worksheet(worksheet_name)
    df = get_as_dataframe(ws, evaluate_formulas=True, header=0)
    return df.dropna(how="all").dropna(axis=1, how="all")  # trim empty padding


def write_worksheet(spreadsheet_key: str, worksheet_name: str, df: pd.DataFrame):
    sh = gc.open_by_key(spreadsheet_key)
    try:
        ws = sh.worksheet(worksheet_name)
        ws.clear()
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=worksheet_name,
                              rows=len(df) + 10, cols=len(df.columns) + 2)
    # <NA> (unranked rows) → blank cells
    out = df.astype(object).where(pd.notna(df), "")
    set_with_dataframe(ws, out, include_index=False, resize=True)


# %% ── Run ───────────────────────────────────────────────────────────────────
df_raw    = read_worksheet(SPREADSHEET_NAME_2, MENTOR_SHEET)
df_ranked = rank_submissions(df_raw)

# quick sanity checks
n_groups  = df_ranked.dropna(subset=[RANK_COL]).groupby(GROUP_KEYS).ngroups
print(f"input rows           : {len(df_raw):,}")
print(f"evaluated (ranked)   : {df_ranked[RANK_COL].notna().sum():,}")
print(f"unranked (no sub id) : {df_ranked[RANK_COL].isna().sum():,}")
print(f"student×question grps : {n_groups:,}")
print(f"max attempts in a grp : {int(df_ranked[RANK_COL].max())}")

if WRITE_CSV_PATH:
    df_ranked.to_csv(WRITE_CSV_PATH, index=False)
    print(f"saved CSV → {WRITE_CSV_PATH}")

if WRITE_TO_SHEET:
    write_worksheet(OUTPUT_SHEET_KEY, OUTPUT_TAB, df_ranked)
    print(f"written → {OUTPUT_TAB} in {OUTPUT_SHEET_KEY}")