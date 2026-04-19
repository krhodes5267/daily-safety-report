"""Quick setup: add tabs + headers to existing Google Sheets workbooks."""
import sys, json, gspread
from google.oauth2.service_account import Credentials
from sheets_pusher import (
    SPEEDING_HEADERS, CAMERA_HEADERS, OBSERVATION_HEADERS,
    INCIDENT_HEADERS, ASSESSMENT_HEADERS,
    TRAINING_HEADERS, CA_HEADERS, DEVICE_HEADERS,
    YTD_HEADERS, DAILY_KPI_HEADERS, EXEC_BRIEFING_HEADERS,
)

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

def setup_workbook(client, sheet_id, tabs, label):
    print(f"\nSetting up {label} ({sheet_id})...")
    wb = client.open_by_key(sheet_id)
    existing = [ws.title for ws in wb.worksheets()]
    print(f"  Existing tabs: {existing}")

    for i, (tab_name, headers) in enumerate(tabs):
        if i == 0 and "Sheet1" in existing:
            ws = wb.worksheet("Sheet1")
            ws.update_title(tab_name)
            ws.update(range_name="A1", values=[headers])
            ws.freeze(rows=1)
            print(f"  Renamed Sheet1 -> {tab_name} ({len(headers)} cols)")
        elif tab_name in existing:
            ws = wb.worksheet(tab_name)
            ws.update(range_name="A1", values=[headers])
            ws.freeze(rows=1)
            print(f"  Updated existing tab: {tab_name}")
        else:
            ws = wb.add_worksheet(title=tab_name, rows=1000, cols=len(headers))
            ws.update(range_name="A1", values=[headers])
            ws.freeze(rows=1)
            print(f"  Created tab: {tab_name} ({len(headers)} cols)")

def main():
    trans_id = sys.argv[1]
    snap_id = sys.argv[2]
    creds_file = sys.argv[3] if len(sys.argv) > 3 else "brhas-safety-b5a44478f315.json"

    creds = Credentials.from_service_account_file(creds_file, scopes=SCOPES)
    client = gspread.authorize(creds)

    setup_workbook(client, trans_id, [
        ("Speeding", SPEEDING_HEADERS),
        ("Camera", CAMERA_HEADERS),
        ("Observations", OBSERVATION_HEADERS),
        ("Incidents", INCIDENT_HEADERS),
        ("Assessments", ASSESSMENT_HEADERS),
    ], "Transactional")

    setup_workbook(client, snap_id, [
        ("Daily_KPIs", DAILY_KPI_HEADERS),
        ("Training", TRAINING_HEADERS),
        ("Corrective_Actions", CA_HEADERS),
        ("Device_Status", DEVICE_HEADERS),
        ("YTD_Stats", YTD_HEADERS),
        ("Executive_Briefing", EXEC_BRIEFING_HEADERS),
    ], "Snapshots")

    print("\nDone! All 11 tabs created with headers.")

if __name__ == "__main__":
    main()
