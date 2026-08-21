#!/usr/bin/env python3
"""
Runs every report script in scripts/, in an order that respects the one
real dependency in the pipeline (weekly_summary reads the daily_summary tab
that project_evaluation_report writes).

Each script is run in its own subprocess so that one report failing
(e.g. a malformed row, a transient Sheets API error) doesn't stop the
others from running. Exit code is non-zero if any script failed, so the
GitHub Actions run is correctly marked failed/red for follow-up — but
you still get however many reports could be built.
"""

import subprocess
import sys
from pathlib import Path

SCRIPTS_DIR = Path(__file__).parent / "scripts"

# Order matters: 01 must run before 03 (weekly_summary reads its output tab).
# The rest are independent of each other.
PIPELINE = [
    "01_project_evaluation_report.py",
    "03_weekly_summary.py",
    "02_eval_turnaround_report.py",
    "04_clearance_views.py",
    "05_clearance_outcomes.py",
    "06_rank_submissions.py",
]


def main() -> int:
    failures = []

    for script in PIPELINE:
        path = SCRIPTS_DIR / script
        print("\n" + "=" * 80)
        print(f"▶ Running {script}")
        print("=" * 80)

        result = subprocess.run([sys.executable, str(path)])

        if result.returncode != 0:
            print(f"✗ {script} failed (exit code {result.returncode})")
            failures.append(script)
        else:
            print(f"✓ {script} completed")

    print("\n" + "=" * 80)
    if failures:
        print(f"DONE with {len(failures)} failure(s): {', '.join(failures)}")
    else:
        print("DONE — all scripts completed successfully")
    print("=" * 80)

    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
