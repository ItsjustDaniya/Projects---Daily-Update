# Setup

## What changed from the notebook

- The Metabase session/token in the notebook's first two cells was created but
  never actually used anywhere downstream — every report reads from and writes
  to Google Sheets only (via `gspread`). It's been dropped entirely.
- Colab's `google.colab.auth.authenticate_user()` (interactive, needs a
  browser) is replaced everywhere with a Google Cloud **service account**,
  which can authenticate non-interactively in GitHub Actions.
- Cell 7 (older "clearance views" draft) was dropped — cell 8 (the newer
  version, same report + an extra resubmission×clearance table) is what's
  automated as `04_clearance_views.py`.
- Each remaining pipeline became its own script under `scripts/`, matching
  the notebook's original section boundaries. `run_all.py` runs them in
  order once a day.

## 1. Create a Google Cloud service account

1. Go to console.cloud.google.com → select or create a project.
2. **APIs & Services → Library** → enable **Google Sheets API** and
   **Google Drive API**.
3. **APIs & Services → Credentials → Create Credentials → Service account**.
   Give it any name (e.g. `ds-batches-reports`).
4. Open the new service account → **Keys** tab → **Add key → Create new key
   → JSON**. This downloads a `.json` key file — keep it private, don't
   commit it.
5. Note the service account's email address (looks like
   `ds-batches-reports@your-project.iam.gserviceaccount.com`).

## 2. Share your Google Sheets with it

The service account has no access to anything until you share it explicitly,
just like sharing with a colleague's email. Share **Editor** access on:

- The source sheet: **`calender - 2`** (specifically the `Project_evaluations`
  tab the scripts read)
- Every output spreadsheet the scripts write to:
  - `1FhjCMl4pQI-yiYNdo64IRZLGkUaNSAqbraENZdT3CAE`
  - `19ecxPlN_MsnO8aVXs4FjJcxu0ShZ1ZtXAXJpZZKttj4`
  - `1CbJL4twtn38a8FbFkkXxZ_HNXFStlgwqLPdPQgwY5zI`

(Open each by its key: `https://docs.google.com/spreadsheets/d/<KEY>/edit`,
then Share → paste the service account email → Editor.)

## 3. Add the key as a GitHub secret

1. In your GitHub repo: **Settings → Secrets and variables → Actions → New
   repository secret**.
2. Name: `GCP_SERVICE_ACCOUNT_JSON`
3. Value: paste the **entire contents** of the JSON key file you downloaded
   in step 1 (open it in a text editor, select all, copy).

## 4. Push this repo structure

```
.github/workflows/daily-reports.yml
requirements.txt
run_all.py
scripts/
  common/
    __init__.py
    sheets_auth.py
  01_project_evaluation_report.py
  02_eval_turnaround_report.py
  03_weekly_summary.py
  04_clearance_views.py
  05_clearance_outcomes.py
  06_rank_submissions.py
```

Once pushed, the workflow runs daily at 06:30 IST automatically. You can
also trigger it manually any time from the repo's **Actions** tab →
"DS Batches Daily Reports" → **Run workflow**.

## 5. Test it before waiting for the schedule

Locally:

```bash
pip install -r requirements.txt
export GCP_SERVICE_ACCOUNT_JSON="$(cat /path/to/your-key.json)"
python run_all.py
```

Or just push and use **Run workflow** in the Actions tab — check the run's
logs for each script's ✓/✗ status.

## Notes

- `run_all.py` runs each script as a separate process and keeps going even
  if one fails, so a bad row in one report doesn't block the other five.
  The overall job still shows red in Actions if anything failed, so you'll
  see it.
- `01_project_evaluation_report.py` must run before
  `03_weekly_summary.py` — it reads the `daily_summary` tab that 01 writes.
  `run_all.py` already runs them in that order; don't reorder them if you
  edit the pipeline list.
- `06_rank_submissions.py` also writes a local `submissions_ranked.csv` in
  the runner (harmless, discarded when the job ends). Set
  `WRITE_CSV_PATH = None` near the top of that script if you'd rather skip it.
