"""
Shared Google Sheets auth for all report scripts.

Replaces the interactive `google.colab.auth.authenticate_user()` flow used in
the original notebook with a service-account flow that works headlessly in
GitHub Actions (or any CI / cron environment).

Expects the service account's JSON key in the environment variable
GCP_SERVICE_ACCOUNT_JSON (the raw JSON content, not a file path).
See SETUP.md for how to create the service account and populate this secret.
"""

import json
import os

import gspread
from google.oauth2.service_account import Credentials

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

ENV_VAR = "GCP_SERVICE_ACCOUNT_JSON"

_client = None  # cached, so repeated calls within one script don't re-auth


def get_client() -> gspread.Client:
    """Return an authorized gspread client, built from the service account
    JSON stored in the GCP_SERVICE_ACCOUNT_JSON environment variable."""
    global _client
    if _client is not None:
        return _client

    raw = os.environ.get(ENV_VAR)
    if not raw:
        raise RuntimeError(
            f"Environment variable {ENV_VAR} is not set. "
            "In GitHub Actions this should come from a repo secret of the "
            "same name (see SETUP.md)."
        )

    try:
        info = json.loads(raw)
    except json.JSONDecodeError as e:
        raise RuntimeError(
            f"{ENV_VAR} does not contain valid JSON. Make sure the whole "
            "service-account key file's contents were pasted into the secret."
        ) from e

    creds = Credentials.from_service_account_info(info, scopes=SCOPES)
    _client = gspread.authorize(creds)
    return _client
