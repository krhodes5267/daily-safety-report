"""
SETUP_SHEETS.PY -- Create and configure Google Sheets workbooks for Looker Studio
==================================================================================
Run this ONCE after creating a service account to:
  1. Create 2 Google Sheets workbooks
  2. Add all tabs with proper headers
  3. Print the Sheet IDs to add as GitHub secrets

Usage:
  python setup_sheets.py credentials.json your-email@gmail.com

The credentials.json is your Google Cloud service account key file.
Your email is needed to share the workbooks with you (service accounts own them by default).
"""

import sys
import json
import gspread
from google.oauth2.service_account import Credentials

from sheets_pusher import (
    SPEEDING_HEADERS, CAMERA_HEADERS, OBSERVATION_HEADERS,
    INCIDENT_HEADERS, ASSESSMENT_HEADERS,
    TRAINING_HEADERS, CA_HEADERS, DEVICE_HEADERS,
    YTD_HEADERS, DAILY_KPI_HEADERS, EXEC_BRIEFING_HEADERS,
)

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]


def main():
    if len(sys.argv) < 3:
        print("Usage: python setup_sheets.py <credentials.json> <your-email@gmail.com>")
        print("\nThe service account needs Drive API enabled to create spreadsheets.")
        sys.exit(1)

    creds_file = sys.argv[1]
    user_email = sys.argv[2]

    print("Authenticating...")
    creds = Credentials.from_service_account_file(creds_file, scopes=SCOPES)
    client = gspread.authorize(creds)

    # --- Workbook 1: Transactional ---
    print("\nCreating 'BRHAS Safety - Transactional'...")
    wb1 = client.create("BRHAS Safety - Transactional")
    wb1.share(user_email, perm_type="user", role="writer")

    transactional_tabs = [
        ("Speeding", SPEEDING_HEADERS),
        ("Camera", CAMERA_HEADERS),
        ("Observations", OBSERVATION_HEADERS),
        ("Incidents", INCIDENT_HEADERS),
        ("Assessments", ASSESSMENT_HEADERS),
    ]

    # Rename default Sheet1 to first tab
    default_ws = wb1.sheet1
    default_ws.update_title(transactional_tabs[0][0])
    default_ws.update(range_name="A1", values=[transactional_tabs[0][1]])
    default_ws.freeze(rows=1)
    print(f"  Created tab: {transactional_tabs[0][0]} ({len(transactional_tabs[0][1])} columns)")

    for tab_name, headers in transactional_tabs[1:]:
        ws = wb1.add_worksheet(title=tab_name, rows=1000, cols=len(headers))
        ws.update(range_name="A1", values=[headers])
        ws.freeze(rows=1)
        print(f"  Created tab: {tab_name} ({len(headers)} columns)")

    print(f"\n  TRANSACTIONAL_SHEET_ID = {wb1.id}")

    # --- Workbook 2: Snapshots ---
    print("\nCreating 'BRHAS Safety - Snapshots'...")
    wb2 = client.create("BRHAS Safety - Snapshots")
    wb2.share(user_email, perm_type="user", role="writer")

    snapshot_tabs = [
        ("Daily_KPIs", DAILY_KPI_HEADERS),
        ("Training", TRAINING_HEADERS),
        ("Corrective_Actions", CA_HEADERS),
        ("Device_Status", DEVICE_HEADERS),
        ("YTD_Stats", YTD_HEADERS),
        ("Executive_Briefing", EXEC_BRIEFING_HEADERS),
    ]

    default_ws = wb2.sheet1
    default_ws.update_title(snapshot_tabs[0][0])
    default_ws.update(range_name="A1", values=[snapshot_tabs[0][1]])
    default_ws.freeze(rows=1)
    print(f"  Created tab: {snapshot_tabs[0][0]} ({len(snapshot_tabs[0][1])} columns)")

    for tab_name, headers in snapshot_tabs[1:]:
        ws = wb2.add_worksheet(title=tab_name, rows=1000, cols=len(headers))
        ws.update(range_name="A1", values=[headers])
        ws.freeze(rows=1)
        print(f"  Created tab: {tab_name} ({len(headers)} columns)")

    print(f"\n  SNAPSHOTS_SHEET_ID = {wb2.id}")

    # --- Summary ---
    print("\n" + "=" * 60)
    print("SETUP COMPLETE")
    print("=" * 60)
    print(f"\nWorkbooks shared with: {user_email}")
    print(f"\nAdd these as GitHub Secrets:")
    print(f"  TRANSACTIONAL_SHEET_ID = {wb1.id}")
    print(f"  SNAPSHOTS_SHEET_ID     = {wb2.id}")
    print(f"  GOOGLE_SHEETS_CREDS_JSON = <contents of {creds_file}>")
    print(f"\nNext: Run 'python sheets_pusher.py --dry-run' to validate, then push real data.")


if __name__ == "__main__":
    main()
