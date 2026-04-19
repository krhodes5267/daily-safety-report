"""
SHEETS_PUSHER.PY -- Push dashboard JSON data to Google Sheets for Looker Studio
================================================================================
Reads output/*.json files produced by the 7 daily data scripts and pushes
them to 2 Google Sheets workbooks:
  1. Transactional (append daily): Speeding, Camera, Observations, Incidents, Assessments
  2. Snapshots (overwrite daily): Training, Corrective_Actions, Device_Status, YTD_Stats,
     Daily_KPIs (append), Executive_Briefing (overwrite)

Auth: Google Cloud service account via GOOGLE_SHEETS_CREDS_JSON env var.
"""

import json
import os
import sys
import datetime

import gspread
from google.oauth2.service_account import Credentials

from api_config import (
    SAFETY_FORMS, KPA_COMPANY_MAP, KPA_SVC_TO_DIVISION,
    GROUP_ID_MAP, Q1_MAN_HOURS, FORM_DIVISION_MAP,
)

# ==============================================================================
# CONFIG
# ==============================================================================
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
TRANSACTIONAL_SHEET_ID = os.environ.get("TRANSACTIONAL_SHEET_ID", "")
SNAPSHOTS_SHEET_ID = os.environ.get("SNAPSHOTS_SHEET_ID", "")
OUTPUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "output")

# Column headers for each tab
SPEEDING_HEADERS = [
    "date", "driver", "vehicle", "yard", "division", "company",
    "max_speed", "posted_speed", "over_by", "duration", "severity", "tier",
    "location", "maps_link",
]
CAMERA_HEADERS = [
    "date", "driver", "vehicle", "yard", "division", "company",
    "event_type", "display_name", "tier", "speed", "duration",
    "coaching_status", "coached_at",
]
OBSERVATION_HEADERS = [
    "date", "observer", "type", "description", "location",
    "division", "company", "kpa_id",
]
INCIDENT_HEADERS = [
    "date", "employee", "type", "description", "location",
    "division", "company", "kpa_id",
]
ASSESSMENT_HEADERS = [
    "date", "assessor", "form_name", "division", "company",
    "location", "kpa_id", "link",
]
TRAINING_HEADERS = [
    "employee", "company", "division", "location",
    "total_assigned", "complete_count", "incomplete_count",
    "compliance_pct", "status", "incomplete_list",
]
CA_HEADERS = [
    "ca_id", "report_number", "origin_form", "division", "company",
    "description", "assigned_to", "created_by",
    "created_date", "due_date", "days_open",
    "is_overdue", "is_safety", "kpa_link",
]
DEVICE_HEADERS = [
    "vehicle", "yard", "division", "availability",
    "issue_type", "devices_affected", "recommended_action",
    "days_inactive", "is_new", "is_escalation",
]
YTD_HEADERS = [
    "company", "division", "man_hours", "recordables",
    "trir", "dart", "days_since_recordable", "last_recordable_date",
]
DAILY_KPI_HEADERS = [
    "date", "company", "division",
    "obs_count", "incident_count", "near_miss_count", "assessment_count",
    "speeding_count", "camera_count", "red_speeding", "red_camera",
    "trir", "dart", "man_hours",
    "obs_per_1k_hrs", "proactive_ratio",
    "safety_ca_open", "safety_ca_overdue", "training_pct",
]
EXEC_BRIEFING_HEADERS = [
    "date", "company", "headline", "status_color",
    "trir", "dart", "days_since_incident",
    "obs_today", "speeding_today", "camera_today",
    "safety_ca_overdue", "training_pct",
]


# ==============================================================================
# AUTH
# ==============================================================================
def get_client():
    """Authenticate with Google Sheets via service account."""
    creds_json = os.environ.get("GOOGLE_SHEETS_CREDS_JSON", "")
    if creds_json:
        info = json.loads(creds_json)
        creds = Credentials.from_service_account_info(info, scopes=SCOPES)
    else:
        creds = Credentials.from_service_account_file("credentials.json", scopes=SCOPES)
    return gspread.authorize(creds)


def load_json(filename):
    """Load a JSON file from the output directory. Returns None if missing."""
    path = os.path.join(OUTPUT_DIR, filename)
    if not os.path.exists(path):
        print(f"  WARNING: {filename} not found, skipping")
        return None
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


# ==============================================================================
# HELPERS
# ==============================================================================
def get_existing_dates(worksheet, date_col=1):
    """Get set of dates already in a transactional tab (column A = date)."""
    try:
        col_values = worksheet.col_values(date_col)
        return set(col_values[1:])  # skip header
    except Exception:
        return set()


def safe_str(val):
    """Convert value to string safely, handling None."""
    if val is None:
        return ""
    return str(val)


def svc_line_to_division(svc_line):
    """Map a KPA service_line to a dashboard division."""
    if not svc_line:
        return ""
    return KPA_SVC_TO_DIVISION.get(svc_line, svc_line)


def svc_line_to_company(svc_line):
    """Infer company from service line. Most are BRHAS."""
    div = svc_line_to_division(svc_line)
    if div in ("Valor",):
        return "Valor"
    if div in ("BTI",):
        return "BTI"
    if div in ("Transcend",):
        return "Transcend"
    return "BRHAS"


def is_safety_form(form_name):
    """Check if a CA origin form is safety-related. Exact match against whitelist."""
    if not form_name:
        return False
    name_lower = form_name.strip().lower()
    for sf in SAFETY_FORMS:
        if sf.lower() == name_lower:
            return True
    # Also match "Unknown (NNNNNN)" pattern if in whitelist
    if name_lower.startswith("unknown ("):
        return form_name.strip() in SAFETY_FORMS
    return False


# ==============================================================================
# TRANSACTIONAL PUSH FUNCTIONS (append with dedup)
# ==============================================================================
def push_speeding(client, data):
    """Push speeding events to Transactional workbook."""
    if not data:
        return
    events = data.get("events", [])
    if not events:
        print("  Speeding: 0 events, skipping")
        return

    wb = client.open_by_key(TRANSACTIONAL_SHEET_ID)
    ws = wb.worksheet("Speeding")

    report_date = data.get("report_date", "")
    existing = get_existing_dates(ws)
    if report_date in existing:
        print(f"  Speeding: {report_date} already exists, skipping")
        return

    rows = []
    for e in events:
        # Derive company from division/yard using GROUP_ID_MAP patterns
        division = e.get("division", "")
        company = "BRHAS"
        if division in ("Valor",):
            company = "Valor"
        elif division in ("BTI",):
            company = "BTI"
        elif division in ("Transcend",):
            company = "Transcend"

        rows.append([
            report_date,
            safe_str(e.get("driver")),
            safe_str(e.get("vehicle")),
            safe_str(e.get("yard")),
            division,
            company,
            e.get("speed", ""),
            e.get("posted_speed", ""),
            e.get("overspeed", ""),
            safe_str(e.get("duration")),
            safe_str(e.get("severity")),
            safe_str(e.get("tier")),
            safe_str(e.get("location")),
            safe_str(e.get("maps_link")),
        ])

    ws.append_rows(rows, value_input_option="USER_ENTERED")
    print(f"  Speeding: appended {len(rows)} rows for {report_date}")


def push_camera(client, data):
    """Push camera events to Transactional workbook."""
    if not data:
        return
    events = data.get("events", [])
    if not events:
        print("  Camera: 0 events, skipping")
        return

    wb = client.open_by_key(TRANSACTIONAL_SHEET_ID)
    ws = wb.worksheet("Camera")

    report_date = data.get("report_date", "")
    existing = get_existing_dates(ws)
    if report_date in existing:
        print(f"  Camera: {report_date} already exists, skipping")
        return

    rows = []
    for e in events:
        yard = e.get("yard", "")
        # Camera events are Casing-only in current pipeline
        division = "Casing"
        company = "BRHAS"

        rows.append([
            report_date,
            safe_str(e.get("driver")),
            safe_str(e.get("vehicle")),
            yard,
            division,
            company,
            safe_str(e.get("event_type")),
            safe_str(e.get("display_name")),
            safe_str(e.get("tier")),
            e.get("speed", ""),
            safe_str(e.get("duration")),
            safe_str(e.get("coaching_status")),
            safe_str(e.get("coached_at")),
        ])

    ws.append_rows(rows, value_input_option="USER_ENTERED")
    print(f"  Camera: appended {len(rows)} rows for {report_date}")


def push_observations(client, data):
    """Push observation details to Transactional workbook."""
    if not data:
        return
    obs = data.get("observations", {})
    details = obs.get("details", [])
    if not details:
        print("  Observations: 0 details, skipping")
        return

    wb = client.open_by_key(TRANSACTIONAL_SHEET_ID)
    ws = wb.worksheet("Observations")

    report_date = data.get("report_date", "")
    existing = get_existing_dates(ws)
    if report_date in existing:
        print(f"  Observations: {report_date} already exists, skipping")
        return

    rows = []
    for d in details:
        svc = d.get("service_line", "")
        rows.append([
            report_date,
            safe_str(d.get("observer")),
            safe_str(d.get("type")),
            safe_str(d.get("description")),
            safe_str(d.get("location")),
            svc_line_to_division(svc),
            svc_line_to_company(svc),
            safe_str(d.get("report_number")),
        ])

    ws.append_rows(rows, value_input_option="USER_ENTERED")
    print(f"  Observations: appended {len(rows)} rows for {report_date}")


def push_incidents(client, data):
    """Push incident details to Transactional workbook."""
    if not data:
        return
    incidents = data.get("incidents", [])
    if not incidents:
        print("  Incidents: 0 events, skipping")
        return

    wb = client.open_by_key(TRANSACTIONAL_SHEET_ID)
    ws = wb.worksheet("Incidents")

    report_date = data.get("report_date", "")
    existing = get_existing_dates(ws)
    if report_date in existing:
        print(f"  Incidents: {report_date} already exists, skipping")
        return

    rows = []
    for inc in incidents:
        svc = inc.get("service_line", "")
        rows.append([
            report_date,
            safe_str(inc.get("employee", inc.get("observer", ""))),
            safe_str(inc.get("type")),
            safe_str(inc.get("description")),
            safe_str(inc.get("location")),
            svc_line_to_division(svc),
            svc_line_to_company(svc),
            safe_str(inc.get("report_number")),
        ])

    ws.append_rows(rows, value_input_option="USER_ENTERED")
    print(f"  Incidents: appended {len(rows)} rows for {report_date}")


def push_assessments(client, data):
    """Push assessment details to Transactional workbook."""
    if not data:
        return
    assessments = data.get("assessments", {})
    form_details = assessments.get("details", [])
    if not form_details:
        print("  Assessments: 0 forms, skipping")
        return

    wb = client.open_by_key(TRANSACTIONAL_SHEET_ID)
    ws = wb.worksheet("Assessments")

    report_date = data.get("report_date", "")
    existing = get_existing_dates(ws)
    if report_date in existing:
        print(f"  Assessments: {report_date} already exists, skipping")
        return

    rows = []
    for form_group in form_details:
        form_name = form_group.get("form_name", "")
        form_id = form_group.get("form_id", 0)
        division = FORM_DIVISION_MAP.get(form_id, "")

        for row in form_group.get("rows", []):
            svc = row.get("service_line", "")
            company = "BRHAS"
            if division in ("Transcend",):
                company = "Transcend"

            rows.append([
                report_date,
                safe_str(row.get("assessor")),
                form_name,
                division or svc_line_to_division(svc),
                company,
                safe_str(row.get("location", row.get("customer", ""))),
                safe_str(row.get("form_id")),  # response ID
                safe_str(row.get("link")),
            ])

    if rows:
        ws.append_rows(rows, value_input_option="USER_ENTERED")
    print(f"  Assessments: appended {len(rows)} rows for {report_date}")


# ==============================================================================
# SNAPSHOT PUSH FUNCTIONS (clear + overwrite)
# ==============================================================================
def push_training(client, data):
    """Push training compliance to Snapshots workbook (overwrite)."""
    if not data:
        return

    non_compliant = data.get("non_compliant_employees", [])
    wb = client.open_by_key(SNAPSHOTS_SHEET_ID)
    ws = wb.worksheet("Training")

    all_rows = [TRAINING_HEADERS]
    for emp in non_compliant:
        incomplete_names = emp.get("incomplete_training_names", [])
        all_rows.append([
            safe_str(emp.get("employee_name")),
            safe_str(emp.get("company")),
            safe_str(emp.get("division")),
            safe_str(emp.get("yard")),
            emp.get("total_assigned", 0),
            emp.get("complete_count", 0),
            emp.get("incomplete_count", 0),
            emp.get("percent_complete", 0),
            safe_str(emp.get("status")),
            "; ".join(incomplete_names) if incomplete_names else "",
        ])

    ws.clear()
    ws.update(range_name="A1", values=all_rows, value_input_option="USER_ENTERED")
    print(f"  Training: wrote {len(all_rows) - 1} non-compliant employees")


def push_corrective_actions(client, data):
    """Push corrective actions to Snapshots workbook (overwrite)."""
    if not data:
        return

    cas = data.get("corrective_actions", [])
    wb = client.open_by_key(SNAPSHOTS_SHEET_ID)
    ws = wb.worksheet("Corrective_Actions")

    all_rows = [CA_HEADERS]
    for ca in cas:
        origin = ca.get("origin_form", "")
        all_rows.append([
            safe_str(ca.get("id")),
            safe_str(ca.get("report_number")),
            origin,
            safe_str(ca.get("division")),
            "",  # company not in CA data -- derived from division
            safe_str(ca.get("description")),
            safe_str(ca.get("assigned_to")),
            safe_str(ca.get("created_by")),
            safe_str(ca.get("created_date")),
            safe_str(ca.get("due_date")),
            ca.get("days_open", 0),
            "TRUE" if ca.get("overdue") else "FALSE",
            "TRUE" if is_safety_form(origin) else "FALSE",
            safe_str(ca.get("kpa_link")),
        ])

    ws.clear()
    ws.update(range_name="A1", values=all_rows, value_input_option="USER_ENTERED")
    print(f"  Corrective Actions: wrote {len(all_rows) - 1} CAs")


def push_device_status(client, data):
    """Push device status issues to Snapshots workbook (overwrite)."""
    if not data:
        return

    issues = data.get("issues", [])
    wb = client.open_by_key(SNAPSHOTS_SHEET_ID)
    ws = wb.worksheet("Device_Status")

    all_rows = [DEVICE_HEADERS]
    for iss in issues:
        all_rows.append([
            safe_str(iss.get("vehicle_number")),
            safe_str(iss.get("yard")),
            "Casing",  # device status is Casing-only
            safe_str(iss.get("availability")),
            safe_str(iss.get("issue_type")),
            safe_str(iss.get("devices_affected")),
            safe_str(iss.get("recommended_action")),
            iss.get("days_inactive", 0),
            "TRUE" if iss.get("is_new") else "FALSE",
            "TRUE" if iss.get("is_escalation") else "FALSE",
        ])

    ws.clear()
    ws.update(range_name="A1", values=all_rows, value_input_option="USER_ENTERED")
    print(f"  Device Status: wrote {len(all_rows) - 1} issues")


def push_ytd_stats(client, data):
    """Push YTD stats to Snapshots workbook (overwrite)."""
    if not data:
        return

    stats_by_company = data.get("stats_by_company", {})
    wb = client.open_by_key(SNAPSHOTS_SHEET_ID)
    ws = wb.worksheet("YTD_Stats")

    all_rows = [YTD_HEADERS]

    for company, stats in stats_by_company.items():
        # Company-level row
        all_rows.append([
            company,
            "ALL",
            stats.get("man_hours", 0),
            stats.get("recordables", 0),
            round(stats.get("trir", 0), 2),
            round(stats.get("dart", 0), 2),
            stats.get("days_since_recordable", ""),
            safe_str(stats.get("last_recordable_date", "")),
        ])

        # Division-level rows
        by_div = stats.get("by_division", {})
        for div_name, div_hours in by_div.items():
            all_rows.append([
                company,
                div_name,
                div_hours,
                "",  # recordables not broken out by division
                "",
                "",
                "",
                "",
            ])

    ws.clear()
    ws.update(range_name="A1", values=all_rows, value_input_option="USER_ENTERED")
    print(f"  YTD Stats: wrote {len(all_rows) - 1} rows")


# ==============================================================================
# COMPUTED PUSH FUNCTIONS
# ==============================================================================
def push_daily_kpis(client, all_data):
    """
    Compute and append daily KPI row(s) to Snapshots workbook.
    One row per company per day. This tab is append-only for trend charts.
    """
    today = all_data.get("report_date", datetime.date.today().isoformat())

    wb = client.open_by_key(SNAPSHOTS_SHEET_ID)
    ws = wb.worksheet("Daily_KPIs")

    existing = get_existing_dates(ws)
    if today in existing:
        print(f"  Daily KPIs: {today} already exists, skipping")
        return

    speeding = all_data.get("speeding", {})
    camera = all_data.get("camera", {})
    kpa = all_data.get("kpa", {})
    ytd = all_data.get("ytd", {})
    training = all_data.get("training", {})
    cas = all_data.get("corrective_actions", {})

    # Count events
    spd_events = speeding.get("events", [])
    cam_events = camera.get("events", [])
    obs_details = kpa.get("observations", {}).get("details", [])
    near_misses = kpa.get("near_misses", [])
    incidents = kpa.get("incidents", [])

    # Assessment count
    assess_total = kpa.get("assessments", {}).get("total", 0)

    # Speeding severity counts
    red_spd = sum(1 for e in spd_events if e.get("tier") == "RED")
    red_cam = sum(1 for e in cam_events if e.get("tier") == "RED")

    # Observation type counts
    obs_by_type = kpa.get("observations", {}).get("by_type", {})
    recognitions = obs_by_type.get("Recognition", 0) + obs_by_type.get("Good Catch", 0)
    reactive = (
        obs_by_type.get("At-Risk Condition", 0)
        + obs_by_type.get("At-Risk Behavior", 0)
        + obs_by_type.get("At-Risk Procedure", 0)
        + len(near_misses)
        + len(incidents)
    )
    proactive_ratio = round(recognitions / reactive, 2) if reactive > 0 else 0

    # Safety CAs
    ca_list = cas.get("corrective_actions", [])
    safety_cas = [c for c in ca_list if is_safety_form(c.get("origin_form", ""))]
    safety_open = len(safety_cas)
    safety_overdue = sum(1 for c in safety_cas if c.get("overdue"))

    # Training
    training_pct = training.get("overall_pct", 0)

    # YTD stats
    man_hours = ytd.get("ytd_man_hours", 0)
    trir = ytd.get("ytd_trir", 0)
    dart = ytd.get("ytd_dart", 0)

    # Obs per 1K man-hours (annualized)
    obs_per_1k = round((len(obs_details) / man_hours) * 1000, 2) if man_hours > 0 else 0

    # Build company-level rows from stats_by_company
    rows = []
    stats_by_co = ytd.get("stats_by_company", {})

    if stats_by_co:
        for company, co_stats in stats_by_co.items():
            co_man_hours = co_stats.get("man_hours", 0)
            co_trir = co_stats.get("trir", 0)
            co_dart = co_stats.get("dart", 0)

            # Filter events by company (approximate -- use division mapping)
            # For now, push one ALL row with aggregate data
            pass

    # Push single aggregate row (all companies combined)
    rows.append([
        today,
        "ALL",
        "ALL",
        len(obs_details),
        len(incidents),
        len(near_misses),
        assess_total,
        len(spd_events),
        len(cam_events),
        red_spd,
        red_cam,
        round(trir, 2),
        round(dart, 2),
        man_hours,
        obs_per_1k,
        proactive_ratio,
        safety_open,
        safety_overdue,
        training_pct,
    ])

    # Also push per-company rows
    for company, co_stats in stats_by_co.items():
        rows.append([
            today,
            company,
            "ALL",
            "",  # can't split obs by company from daily data easily
            co_stats.get("recordables", 0),
            "",
            "",
            "",
            "",
            "",
            "",
            round(co_stats.get("trir", 0), 2),
            round(co_stats.get("dart", 0), 2),
            co_stats.get("man_hours", 0),
            "",
            "",
            "",
            "",
            "",
        ])

    ws.append_rows(rows, value_input_option="USER_ENTERED")
    print(f"  Daily KPIs: appended {len(rows)} rows for {today}")


def push_executive_briefing(client, all_data):
    """
    Compute and overwrite executive briefing in Snapshots workbook.
    One row per company with headline, status color, key metrics.
    """
    today = all_data.get("report_date", datetime.date.today().isoformat())

    wb = client.open_by_key(SNAPSHOTS_SHEET_ID)
    ws = wb.worksheet("Executive_Briefing")

    speeding = all_data.get("speeding", {})
    camera = all_data.get("camera", {})
    kpa = all_data.get("kpa", {})
    ytd = all_data.get("ytd", {})
    training = all_data.get("training", {})
    cas = all_data.get("corrective_actions", {})

    spd_events = speeding.get("events", [])
    cam_events = camera.get("events", [])
    obs = kpa.get("observations", {})
    incidents = kpa.get("incidents", [])

    # Safety CAs
    ca_list = cas.get("corrective_actions", [])
    safety_overdue = sum(
        1 for c in ca_list
        if c.get("overdue") and is_safety_form(c.get("origin_form", ""))
    )

    trir = ytd.get("ytd_trir", 0)
    dart = ytd.get("ytd_dart", 0)
    days_since = ytd.get("days_since_recordable", 0)
    training_pct = training.get("overall_pct", 0)

    # Determine status color and headline
    red_spd = sum(1 for e in spd_events if e.get("tier") == "RED")
    red_cam = sum(1 for e in cam_events if e.get("tier") == "RED")

    # Status logic: green if no incidents + low red events, yellow if moderate, red if incidents
    if len(incidents) > 0:
        status_color = "RED"
        headline = f"INCIDENT REPORTED -- {len(incidents)} incident(s) logged. Immediate review required."
    elif red_spd + red_cam >= 5:
        status_color = "YELLOW"
        headline = f"Elevated risk -- {red_spd} RED speeding + {red_cam} RED camera events."
    elif red_spd + red_cam >= 1:
        status_color = "YELLOW"
        headline = f"Monitor -- {red_spd + red_cam} RED fleet event(s). {obs.get('total', 0)} observations submitted."
    else:
        status_color = "GREEN"
        headline = f"Clean day -- 0 RED events, {obs.get('total', 0)} observations, {days_since}-day safety streak."

    all_rows = [EXEC_BRIEFING_HEADERS]
    all_rows.append([
        today,
        "ALL",
        headline,
        status_color,
        round(trir, 2),
        round(dart, 2),
        days_since,
        obs.get("total", 0),
        len(spd_events),
        len(cam_events),
        safety_overdue,
        training_pct,
    ])

    # Per-company rows
    stats_by_co = ytd.get("stats_by_company", {})
    for company, co_stats in stats_by_co.items():
        co_trir = co_stats.get("trir", 0)
        co_days = co_stats.get("days_since_recordable", "")

        if co_stats.get("recordables", 0) > 0 and co_trir > 1.5:
            co_color = "RED"
            co_headline = f"TRIR {co_trir:.2f} -- above 1.5 threshold."
        elif co_trir > 1.0:
            co_color = "YELLOW"
            co_headline = f"TRIR {co_trir:.2f} -- monitoring threshold."
        else:
            co_color = "GREEN"
            co_headline = f"TRIR {co_trir:.2f} -- within target."

        all_rows.append([
            today,
            company,
            co_headline,
            co_color,
            round(co_trir, 2),
            round(co_stats.get("dart", 0), 2),
            co_days if co_days is not None else "",
            "",  # obs not split by company
            "",
            "",
            "",
            "",
        ])

    ws.clear()
    ws.update(range_name="A1", values=all_rows, value_input_option="USER_ENTERED")
    print(f"  Executive Briefing: wrote {len(all_rows) - 1} rows")


# ==============================================================================
# MAIN
# ==============================================================================
def main():
    print("=" * 60)
    print("SHEETS PUSHER -- Loading JSON data and pushing to Google Sheets")
    print("=" * 60)

    if not TRANSACTIONAL_SHEET_ID or not SNAPSHOTS_SHEET_ID:
        print("ERROR: TRANSACTIONAL_SHEET_ID and SNAPSHOTS_SHEET_ID env vars required")
        sys.exit(1)

    # Load all JSON files
    print("\nLoading JSON files from output/...")
    speeding = load_json("speeding_events.json")
    camera = load_json("camera_events.json")
    kpa = load_json("kpa_data.json")
    ytd = load_json("ytd_stats.json")
    training = load_json("training_compliance.json")
    cas = load_json("corrective_actions.json")
    devices = load_json("device_status.json")

    # Determine report date (use speeding as primary, fallback to today)
    report_date = datetime.date.today().isoformat()
    for src in [speeding, camera, kpa, ytd]:
        if src and src.get("report_date"):
            report_date = src["report_date"]
            break
    print(f"  Report date: {report_date}")

    # Authenticate
    print("\nAuthenticating with Google Sheets...")
    client = get_client()
    print("  Authenticated successfully")

    # Push transactional data (append)
    print("\n--- TRANSACTIONAL (append) ---")
    try:
        push_speeding(client, speeding)
    except Exception as e:
        print(f"  ERROR pushing speeding: {e}")

    try:
        push_camera(client, camera)
    except Exception as e:
        print(f"  ERROR pushing camera: {e}")

    try:
        push_observations(client, kpa)
    except Exception as e:
        print(f"  ERROR pushing observations: {e}")

    try:
        push_incidents(client, kpa)
    except Exception as e:
        print(f"  ERROR pushing incidents: {e}")

    try:
        push_assessments(client, kpa)
    except Exception as e:
        print(f"  ERROR pushing assessments: {e}")

    # Push snapshot data (overwrite)
    print("\n--- SNAPSHOTS (overwrite) ---")
    try:
        push_training(client, training)
    except Exception as e:
        print(f"  ERROR pushing training: {e}")

    try:
        push_corrective_actions(client, cas)
    except Exception as e:
        print(f"  ERROR pushing corrective actions: {e}")

    try:
        push_device_status(client, devices)
    except Exception as e:
        print(f"  ERROR pushing device status: {e}")

    try:
        push_ytd_stats(client, ytd)
    except Exception as e:
        print(f"  ERROR pushing YTD stats: {e}")

    # Push computed data
    print("\n--- COMPUTED ---")
    all_data = {
        "report_date": report_date,
        "speeding": speeding or {},
        "camera": camera or {},
        "kpa": kpa or {},
        "ytd": ytd or {},
        "training": training or {},
        "corrective_actions": cas or {},
        "devices": devices or {},
    }

    try:
        push_daily_kpis(client, all_data)
    except Exception as e:
        print(f"  ERROR pushing daily KPIs: {e}")

    try:
        push_executive_briefing(client, all_data)
    except Exception as e:
        print(f"  ERROR pushing executive briefing: {e}")

    print("\n" + "=" * 60)
    print("SHEETS PUSHER COMPLETE")
    print("=" * 60)


def dry_run():
    """
    Validate all JSON parsing and print what WOULD be pushed to Sheets.
    No Google credentials needed. Run with: python sheets_pusher.py --dry-run
    """
    print("=" * 60)
    print("SHEETS PUSHER -- DRY RUN (no Google Sheets connection)")
    print("=" * 60)

    print("\nLoading JSON files from output/...")
    speeding = load_json("speeding_events.json")
    camera = load_json("camera_events.json")
    kpa = load_json("kpa_data.json")
    ytd = load_json("ytd_stats.json")
    training = load_json("training_compliance.json")
    cas = load_json("corrective_actions.json")
    devices = load_json("device_status.json")

    report_date = ""
    for src in [speeding, camera, kpa, ytd]:
        if src and src.get("report_date"):
            report_date = src["report_date"]
            break
    print(f"  Report date: {report_date}")

    errors = []

    # --- Transactional ---
    print("\n--- TRANSACTIONAL TABS ---")

    # Speeding
    if speeding:
        events = speeding.get("events", [])
        print(f"\n  Speeding: {len(events)} events")
        if events:
            e = events[0]
            print(f"    Sample: {e.get('driver')} | {e.get('vehicle')} | {e.get('tier')} | {e.get('speed')}mph")
            # Validate all expected keys exist
            for key in ["driver", "vehicle", "speed", "posted_speed", "overspeed", "duration", "severity", "tier", "division", "yard"]:
                if key not in e:
                    errors.append(f"Speeding event missing key: {key}")
    else:
        print("\n  Speeding: NO DATA")

    # Camera
    if camera:
        events = camera.get("events", [])
        print(f"\n  Camera: {len(events)} events")
        if events:
            e = events[0]
            print(f"    Sample: {e.get('driver')} | {e.get('vehicle')} | {e.get('display_name')} | {e.get('tier')}")
            for key in ["driver", "vehicle", "event_type", "tier", "yard"]:
                if key not in e:
                    errors.append(f"Camera event missing key: {key}")
    else:
        print("\n  Camera: NO DATA")

    # Observations
    if kpa:
        obs = kpa.get("observations", {})
        details = obs.get("details", [])
        print(f"\n  Observations: {len(details)} details (total: {obs.get('total', 0)})")
        if details:
            d = details[0]
            print(f"    Sample: {d.get('observer')} | {d.get('type')} | {d.get('service_line')}")
            div = svc_line_to_division(d.get("service_line", ""))
            co = svc_line_to_company(d.get("service_line", ""))
            print(f"    Mapped: division={div}, company={co}")

        # Near misses
        near = kpa.get("near_misses", [])
        print(f"\n  Near Misses: {len(near)}")

        # Incidents
        inc = kpa.get("incidents", [])
        print(f"\n  Incidents: {len(inc)}")
        if inc:
            i = inc[0]
            print(f"    Sample: {i.get('employee', i.get('observer', ''))} | {i.get('type')} | {i.get('service_line')}")

        # Assessments
        assess = kpa.get("assessments", {})
        form_details = assess.get("details", [])
        total_rows = sum(len(fg.get("rows", [])) for fg in form_details)
        print(f"\n  Assessments: {assess.get('total', 0)} total, {len(form_details)} form types, {total_rows} detail rows")
        for fg in form_details:
            row_count = len(fg.get("rows", []))
            if row_count > 0:
                div = FORM_DIVISION_MAP.get(fg.get("form_id", 0), "?")
                print(f"    {fg.get('form_name')}: {row_count} rows -> {div}")
    else:
        print("\n  KPA: NO DATA")

    # --- Snapshots ---
    print("\n--- SNAPSHOT TABS ---")

    # Training
    if training:
        nc = training.get("non_compliant_employees", [])
        print(f"\n  Training: {len(nc)} non-compliant employees")
        print(f"    Overall compliance: {training.get('overall_pct', 0)}%")
        print(f"    Total employees: {training.get('total_employees', 0)}")
        if nc:
            e = nc[0]
            print(f"    Sample: {e.get('employee_name')} | {e.get('company')} | {e.get('division')} | {e.get('percent_complete')}%")
            inc_names = e.get("incomplete_training_names", [])
            print(f"    Incomplete programs: {len(inc_names)}")
    else:
        print("\n  Training: NO DATA")

    # Corrective Actions
    if cas:
        ca_list = cas.get("corrective_actions", [])
        safety_count = sum(1 for c in ca_list if is_safety_form(c.get("origin_form", "")))
        overdue_count = sum(1 for c in ca_list if c.get("overdue"))
        print(f"\n  Corrective Actions: {len(ca_list)} total, {safety_count} safety, {overdue_count} overdue")
        if ca_list:
            c = ca_list[0]
            print(f"    Sample: {c.get('origin_form')} | {c.get('assigned_to')} | {c.get('days_open')}d | overdue={c.get('overdue')}")
            print(f"    is_safety={is_safety_form(c.get('origin_form', ''))}")

        # Show form type breakdown
        form_counts = {}
        for c in ca_list:
            f = c.get("origin_form", "Unknown")
            sf = is_safety_form(f)
            key = f"{f} ({'SAFETY' if sf else 'non-safety'})"
            form_counts[key] = form_counts.get(key, 0) + 1
        print("    By form type:")
        for form, count in sorted(form_counts.items(), key=lambda x: -x[1])[:10]:
            print(f"      {count:>4}  {form}")
    else:
        print("\n  Corrective Actions: NO DATA")

    # Device Status
    if devices:
        issues = devices.get("issues", [])
        print(f"\n  Device Status: {len(issues)} issues")
        if issues:
            i = issues[0]
            print(f"    Sample: {i.get('vehicle_number')} | {i.get('yard')} | {i.get('issue_type')}")
    else:
        print("\n  Device Status: NO DATA")

    # YTD Stats
    if ytd:
        sbc = ytd.get("stats_by_company", {})
        print(f"\n  YTD Stats: {len(sbc)} companies")
        for co, st in sbc.items():
            divs = st.get("by_division", {})
            print(f"    {co}: TRIR={st.get('trir', 0):.2f}, DART={st.get('dart', 0):.2f}, "
                  f"hours={st.get('man_hours', 0):,}, {len(divs)} divisions")
    else:
        print("\n  YTD Stats: NO DATA")

    # --- Computed ---
    print("\n--- COMPUTED TABS ---")

    all_data = {
        "report_date": report_date,
        "speeding": speeding or {},
        "camera": camera or {},
        "kpa": kpa or {},
        "ytd": ytd or {},
        "training": training or {},
        "corrective_actions": cas or {},
        "devices": devices or {},
    }

    # Daily KPIs
    spd_events = (speeding or {}).get("events", [])
    cam_events = (camera or {}).get("events", [])
    obs_details = (kpa or {}).get("observations", {}).get("details", [])
    near_misses = (kpa or {}).get("near_misses", [])
    incidents = (kpa or {}).get("incidents", [])
    obs_by_type = (kpa or {}).get("observations", {}).get("by_type", {})

    red_spd = sum(1 for e in spd_events if e.get("tier") == "RED")
    red_cam = sum(1 for e in cam_events if e.get("tier") == "RED")
    recognitions = obs_by_type.get("Recognition", 0) + obs_by_type.get("Good Catch", 0)
    reactive = (obs_by_type.get("At-Risk Condition", 0) + obs_by_type.get("At-Risk Behavior", 0)
                + obs_by_type.get("At-Risk Procedure", 0) + len(near_misses) + len(incidents))
    proactive_ratio = round(recognitions / reactive, 2) if reactive > 0 else 0
    man_hours = (ytd or {}).get("ytd_man_hours", 0)
    obs_per_1k = round((len(obs_details) / man_hours) * 1000, 2) if man_hours > 0 else 0
    trir = (ytd or {}).get("ytd_trir", 0)
    dart = (ytd or {}).get("ytd_dart", 0)

    ca_list = (cas or {}).get("corrective_actions", [])
    safety_cas = [c for c in ca_list if is_safety_form(c.get("origin_form", ""))]

    print(f"\n  Daily KPIs (aggregate row):")
    print(f"    Observations: {len(obs_details)} | Incidents: {len(incidents)} | Near Misses: {len(near_misses)}")
    print(f"    Assessments: {(kpa or {}).get('assessments', {}).get('total', 0)}")
    print(f"    Speeding: {len(spd_events)} ({red_spd} RED) | Camera: {len(cam_events)} ({red_cam} RED)")
    print(f"    TRIR: {trir:.2f} | DART: {dart:.2f} | Man-Hours: {man_hours:,}")
    print(f"    Obs/1K hrs: {obs_per_1k} | Proactive Ratio: {proactive_ratio}")
    print(f"    Safety CAs: {len(safety_cas)} open, {sum(1 for c in safety_cas if c.get('overdue'))} overdue")
    print(f"    Training: {(training or {}).get('overall_pct', 0)}%")

    # Executive Briefing
    days_since = (ytd or {}).get("days_since_recordable", 0)
    if len(incidents) > 0:
        status = "RED"
        headline = f"INCIDENT REPORTED -- {len(incidents)} incident(s) logged."
    elif red_spd + red_cam >= 5:
        status = "YELLOW"
        headline = f"Elevated risk -- {red_spd} RED speeding + {red_cam} RED camera events."
    elif red_spd + red_cam >= 1:
        status = "YELLOW"
        headline = f"Monitor -- {red_spd + red_cam} RED fleet event(s). {len(obs_details)} observations."
    else:
        status = "GREEN"
        headline = f"Clean day -- 0 RED events, {len(obs_details)} observations, {days_since}-day streak."

    print(f"\n  Executive Briefing:")
    print(f"    Status: {status}")
    print(f"    Headline: {headline}")

    # --- Summary ---
    print("\n" + "=" * 60)
    if errors:
        print(f"VALIDATION ERRORS: {len(errors)}")
        for err in errors:
            print(f"  - {err}")
    else:
        print("VALIDATION PASSED -- all data structures look good")

    # Row count summary
    spd_rows = len((speeding or {}).get("events", []))
    cam_rows = len((camera or {}).get("events", []))
    obs_rows = len((kpa or {}).get("observations", {}).get("details", []))
    inc_rows = len((kpa or {}).get("incidents", []))
    assess_rows = sum(len(fg.get("rows", [])) for fg in (kpa or {}).get("assessments", {}).get("details", []))
    train_rows = len((training or {}).get("non_compliant_employees", []))
    ca_rows = len((cas or {}).get("corrective_actions", []))
    dev_rows = len((devices or {}).get("issues", []))
    ytd_rows = sum(1 + len(st.get("by_division", {})) for st in (ytd or {}).get("stats_by_company", {}).values())

    print(f"\nRows to push:")
    print(f"  Transactional:  Speeding={spd_rows}, Camera={cam_rows}, Obs={obs_rows}, Inc={inc_rows}, Assess={assess_rows}")
    print(f"  Snapshots:      Training={train_rows}, CAs={ca_rows}, Devices={dev_rows}, YTD={ytd_rows}")
    print(f"  Computed:       Daily_KPIs=1+{len((ytd or {}).get('stats_by_company', {}))}, Exec_Briefing=1+{len((ytd or {}).get('stats_by_company', {}))}")
    print(f"  TOTAL:          {spd_rows + cam_rows + obs_rows + inc_rows + assess_rows + train_rows + ca_rows + dev_rows + ytd_rows} data rows")
    print("=" * 60)


if __name__ == "__main__":
    if "--dry-run" in sys.argv:
        dry_run()
    else:
        main()
