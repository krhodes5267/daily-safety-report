"""
BACKFILL_SHEETS.PY -- Load historical data into Google Sheets for Looker Studio
================================================================================
Fetches KPA data (2025-01-01 to today) and Motive data (2026-01-01 to today)
from live APIs and pushes to the Transactional Google Sheets workbook.

Usage:
  python backfill_sheets.py credentials.json

Env vars required:
  KPA_API_TOKEN, MOTIVE_API_KEY, TRANSACTIONAL_SHEET_ID, SNAPSHOTS_SHEET_ID
"""

import csv
import json
import os
import sys
import time
import re
from io import StringIO
from datetime import datetime, timedelta, timezone, date

import requests
import gspread
from google.oauth2.service_account import Credentials

from api_config import (
    KPA_API_BASE, MOTIVE_BASE_V1, MOTIVE_BASE_V2,
    OBSERVATION_FORM_ID, INCIDENT_FORM_ID, ASSESSMENT_FORM_IDS,
    OBS_TYPE_HASH, OBS_DESC_HASH, OBS_LOCATION_HASH,
    INC_TYPE_HASH, INC_EMPLOYEE_HASH, INC_DESC_HASH, INC_LOCATION_HASH,
    SVC_LINE_HASH, COMPANY_HASH,
    SERVICE_LINE_HASHES, FORM_DIVISION_MAP,
    CASING_GROUP_IDS, GROUP_ID_MAP,
    KPA_COMPANY_MAP, KPA_SVC_TO_DIVISION,
)

# ==============================================================================
# CONFIG
# ==============================================================================
KPA_TOKEN = os.environ.get("KPA_API_TOKEN", "")
MOTIVE_KEY = os.environ.get("MOTIVE_API_KEY", "")
TRANSACTIONAL_SHEET_ID = os.environ.get("TRANSACTIONAL_SHEET_ID", "")
SNAPSHOTS_SHEET_ID = os.environ.get("SNAPSHOTS_SHEET_ID", "")
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# Date ranges
KPA_START = "2025-01-01"
MOTIVE_START = "2026-01-01"
TODAY = date.today().isoformat()

# Speeding severity tiers
def speeding_tier(speed, posted):
    over = speed - posted
    if over >= 20 or speed >= 90:
        return "RED"
    elif over >= 15:
        return "ORANGE"
    return "YELLOW"

# Camera RED types
RED_CAMERA_TYPES = {
    "cell_phone", "drowsiness", "distraction", "close_following",
    "forward_collision_warning", "collision", "near_collision",
    "stop_sign", "stop_sign_violation", "unsafe_lane_change", "lane_swerving",
}

CASING_GIDS = set(CASING_GROUP_IDS.keys())

# Vehicle number prefix -> (yard, division)
# Patterns: "MID-CSG-1234", "JOU-RAT-2364", "TOW-RAT-2229", "5016C"
YARD_PREFIXES = {
    "MID": "Midland", "BRY": "Bryan", "KIL": "Kilgore",
    "HOB": "Hobbs", "JOU": "Jourdanton", "LAR": "Laredo",
    "SAN": "San Angelo", "SA": "San Angelo",
    "TOW": "Pennsylvania", "PA": "Pennsylvania",
    "OH": "Ohio", "OK": "Oklahoma", "ND": "North Dakota",
    "LVL": "Levelland", "LL": "Levelland", "BAR": "Barstow",
    "DS": "Dallas", "WIN": "Winters", "PER": "Perryton",
}
DIV_PREFIXES = {
    "CSG": "Casing", "RAT": "Rathole", "ANC": "Anchors",
    "PP": "Poly Pipe", "PL": "Pit Lining", "CON": "Construction",
    "ENV": "Environmental", "FEN": "Fencing", "DT": "Drilling Tools",
    "VAL": "Valor", "BTI": "BTI", "TD": "Transcend",
    "WTC": "Water/Construction", "FAB": "Fabrication",
}
SOLO_DIV_PREFIXES = {
    "ENV": "Environmental", "BTI": "BTI", "VAL": "Valor",
    "POL": "Poly Pipe", "PIT": "Pit Lining", "ANC": "Anchors",
    "FEN": "Fencing", "TD": "Transcend", "CON": "Construction",
}

# Assessment form ID -> human-readable name
FORM_NAMES = {
    381707: "CSG - Safety Casing Field Assessment",
    229645: "CSG - Pre/Post Trip Inspection",
    385365: "TD - Rig Inspection",
    226217: "WS - Poly Pipe Field Assessment",
    386087: "WS - Pit Lining Field Assessment",
    172295: "Construction - Site Safety Review",
    153181: "RH - Rathole Field Assessment",
    152018: "Vehicle Inspection Checklist",
    152034: "HSE - Workplace Inspection Checklist",
}


def parse_vehicle_number(vnum):
    """Parse vehicle number into (division, yard, company).

    Patterns:
      YARD-DIV-NUM:  "JOU-RAT-2364"  -> (Rathole, Jourdanton, BRHAS)
      YARD-DIV-NUM:  "LL-RAT-1821"   -> (Rathole, Levelland, BRHAS)
      DIV-NUM:       "ENV-2093E"      -> (Environmental, "", BRHAS)
      DIV-NUM:       "BTI-63138"      -> (BTI, "", BTI)
      NUMX:          "23111C"         -> (Casing, "", BRHAS)
      Other:         "Sales 2560..."  -> ("", "", BRHAS)
    """
    if not vnum:
        return ("", "", "BRHAS")

    clean = vnum.strip().upper().split(" ")[0]
    parts = clean.split("-")

    yard = ""
    div = ""
    company = "BRHAS"

    if len(parts) >= 3:
        yard = YARD_PREFIXES.get(parts[0], "")
        div = DIV_PREFIXES.get(parts[1], "")
        if not yard and parts[0] in SOLO_DIV_PREFIXES:
            div = SOLO_DIV_PREFIXES[parts[0]]
    elif len(parts) == 2:
        if parts[0] in YARD_PREFIXES and parts[1] in DIV_PREFIXES:
            yard = YARD_PREFIXES[parts[0]]
            div = DIV_PREFIXES[parts[1]]
        elif parts[0] in SOLO_DIV_PREFIXES:
            div = SOLO_DIV_PREFIXES[parts[0]]
        elif parts[0] in YARD_PREFIXES:
            yard = YARD_PREFIXES[parts[0]]
        elif parts[1] in DIV_PREFIXES:
            div = DIV_PREFIXES[parts[1]]
    elif len(parts) == 1:
        if re.match(r"^\d+C$", clean):
            div = "Casing"
        elif re.match(r"^\d+R$", clean):
            div = "Rathole"
        elif re.match(r"^\d+A$", clean):
            div = "Anchors"
        elif re.match(r"^\d+V$", clean):
            div = "Valor"
        elif re.search(r"\d+PP$", clean):
            div = "Poly Pipe"
        elif re.search(r"\d+PL$", clean):
            div = "Pit Lining"
        elif re.search(r"\d+E$", clean):
            div = "Environmental"
        elif re.search(r"\d+F$", clean):
            div = "Fencing"

    if div == "Valor":
        company = "Valor"
    elif div == "BTI":
        company = "BTI"
    elif div == "Transcend":
        company = "Transcend"

    return (div, yard, company)


def svc_to_div(svc):
    if not svc:
        return ""
    return KPA_SVC_TO_DIVISION.get(svc, svc)


def svc_to_company(svc):
    div = svc_to_div(svc)
    if div in ("Valor",):
        return "Valor"
    if div in ("BTI",):
        return "BTI"
    if div in ("Transcend",):
        return "Transcend"
    return "BRHAS"


def group_to_yard_div(group_ids):
    """Map Motive group IDs to (division, yard)."""
    for gid in (group_ids or []):
        if gid in GROUP_ID_MAP:
            return GROUP_ID_MAP[gid]
    return ("", "")


def format_duration(seconds):
    """Convert seconds to human-readable duration."""
    if not seconds:
        return ""
    s = int(seconds)
    if s < 60:
        return f"{s}s"
    m = s // 60
    s = s % 60
    if s > 0:
        return f"{m}m {s}s"
    return f"{m}m"


# ==============================================================================
# KPA FETCH
# ==============================================================================
def kpa_fetch_csv(form_id, updated_after_ms):
    """Fetch KPA responses as CSV rows with pagination."""
    all_rows = []
    page = 1
    while True:
        payload = {
            "token": KPA_TOKEN,
            "form_id": form_id,
            "format": "csv",
            "updated_after": updated_after_ms,
            "page": page,
        }
        for attempt in range(3):
            try:
                resp = requests.post(f"{KPA_API_BASE}/responses.flat", json=payload, timeout=180)
                break
            except (requests.exceptions.ConnectionError, requests.exceptions.ReadTimeout):
                print(f"    KPA timeout/reset on form {form_id} page {page}, retry {attempt+1}/3...")
                time.sleep(15 * (attempt + 1))
        else:
            print(f"    Failed after 3 retries on form {form_id} page {page}, stopping this form")
            break
        if not resp.ok:
            print(f"    KPA API error {resp.status_code} for form {form_id} page {page}")
            break
        text = resp.text.strip()
        if not text:
            break
        reader = csv.DictReader(StringIO(text))
        rows = list(reader)
        # Skip label header row
        data = [r for r in rows if r.get("date", "") != "Date"]
        if not data:
            break
        all_rows.extend(data)
        if len(data) < 100:
            break
        page += 1
        time.sleep(1.5)
        if page > 200:  # safety limit
            break
    return all_rows


def get_svc_line(row):
    """Extract service line from KPA row, trying multiple hash fields."""
    for h in SERVICE_LINE_HASHES:
        val = row.get(h, "")
        if val and val not in ("Service Line", "service_line", "Division", "division", ""):
            return val
    return ""


def fetch_kpa_observations(start_date, end_date):
    """Fetch all observations between start and end date."""
    print(f"  Fetching KPA observations ({start_date} to {end_date})...")
    start_dt = datetime.strptime(start_date, "%Y-%m-%d")
    updated_after = int((start_dt - timedelta(days=1)).timestamp() * 1000)

    rows = kpa_fetch_csv(OBSERVATION_FORM_ID, updated_after)
    print(f"    Raw rows: {len(rows)}")

    # Also fetch TD observations (form 484193)
    td_rows = kpa_fetch_csv(484193, updated_after)
    print(f"    TD obs rows: {len(td_rows)}")
    rows.extend(td_rows)

    results = []
    near_misses = []
    for row in rows:
        row_date = (row.get("date") or "")[:10]
        if not row_date or row_date < start_date or row_date > end_date:
            continue

        obs_type = row.get(OBS_TYPE_HASH, "Other") or "Other"
        svc = get_svc_line(row)
        entry = {
            "date": row_date,
            "observer": row.get("Name", "") or row.get("observer", "") or "",
            "type": obs_type,
            "description": row.get(OBS_DESC_HASH, "") or "",
            "location": row.get(OBS_LOCATION_HASH, "") or "",
            "service_line": svc,
            "report_number": row.get("report number", "") or "",
        }

        if "near miss" in obs_type.lower():
            near_misses.append(entry)
        else:
            results.append(entry)

    print(f"    Filtered: {len(results)} observations, {len(near_misses)} near misses")
    return results, near_misses


def fetch_kpa_incidents(start_date, end_date):
    """Fetch all incidents between start and end date."""
    print(f"  Fetching KPA incidents ({start_date} to {end_date})...")
    start_dt = datetime.strptime(start_date, "%Y-%m-%d")
    updated_after = int((start_dt - timedelta(days=1)).timestamp() * 1000)

    rows = kpa_fetch_csv(INCIDENT_FORM_ID, updated_after)
    print(f"    Raw rows: {len(rows)}")

    results = []
    for row in rows:
        row_date = (row.get("date") or "")[:10]
        if not row_date or row_date < start_date or row_date > end_date:
            continue

        svc = get_svc_line(row)
        results.append({
            "date": row_date,
            "employee": row.get(INC_EMPLOYEE_HASH, "") or row.get("Name", "") or "",
            "type": row.get(INC_TYPE_HASH, "") or "",
            "description": row.get(INC_DESC_HASH, "") or "",
            "location": row.get(INC_LOCATION_HASH, "") or "",
            "service_line": svc,
            "report_number": row.get("report number", "") or "",
        })

    print(f"    Filtered: {len(results)} incidents")
    return results


def fetch_kpa_assessments(start_date, end_date):
    """Fetch all assessments across all form types."""
    print(f"  Fetching KPA assessments ({start_date} to {end_date})...")
    start_dt = datetime.strptime(start_date, "%Y-%m-%d")
    updated_after = int((start_dt - timedelta(days=1)).timestamp() * 1000)

    results = []
    for form_id in ASSESSMENT_FORM_IDS:
        time.sleep(1.5)  # rate limit
        rows = kpa_fetch_csv(form_id, updated_after)
        form_name = ""
        division = FORM_DIVISION_MAP.get(form_id, "")

        form_name = FORM_NAMES.get(form_id, f"Form {form_id}")
        count = 0
        for row in rows:
            row_date = (row.get("date") or "")[:10]
            if not row_date or row_date < start_date or row_date > end_date:
                continue

            svc = get_svc_line(row)
            company = "BRHAS"
            if division in ("Transcend",):
                company = "Transcend"

            results.append({
                "date": row_date,
                "assessor": row.get("Name", "") or row.get("observer", "") or "",
                "form_name": form_name,
                "form_id": form_id,
                "division": division or svc_to_div(svc),
                "company": company,
                "location": row.get(OBS_LOCATION_HASH, "") or "",
                "report_number": row.get("report number", "") or "",
                "link": "",
            })
            count += 1

        if count > 0:
            print(f"    Form {form_id} ({form_name or division}): {count} rows")

    print(f"    Total assessments: {len(results)}")
    return results


# ==============================================================================
# MOTIVE FETCH
# ==============================================================================
def fetch_motive_speeding(start_date, end_date):
    """Fetch speeding events from Motive API with pagination."""
    print(f"  Fetching Motive speeding ({start_date} to {end_date})...")
    headers = {"X-Api-Key": MOTIVE_KEY}
    all_events = []
    page = 1

    while page <= 100:
        params = {
            "per_page": 100,
            "page_no": page,
            "start_date": start_date,
            "end_date": end_date,
        }
        for attempt in range(3):
            try:
                resp = requests.get(
                    f"{MOTIVE_BASE_V1}/speeding_events",
                    headers=headers, params=params, timeout=60
                )
                break
            except requests.exceptions.ConnectionError:
                print(f"    Connection reset on page {page}, retry {attempt+1}/3...")
                time.sleep(10 * (attempt + 1))
        else:
            print(f"    Failed after 3 retries on page {page}, stopping")
            break

        if not resp.ok:
            print(f"    API error {resp.status_code} on page {page}")
            break

        data = resp.json()
        items = data.get("speeding_events", [])
        if not items:
            break

        for item in items:
            ev = item.get("speeding_event", item)

            # Driver (often null in speeding API)
            drv_obj = ev.get("driver") or {}
            if isinstance(drv_obj, dict) and drv_obj:
                first = drv_obj.get("first_name", "") or ""
                last = drv_obj.get("last_name", "") or ""
                driver = f"{first} {last}".strip() or drv_obj.get("name", "Unknown")
            else:
                driver = "Unknown"

            # Vehicle number (e.g. "JOU-RAT-2364")
            veh_obj = ev.get("vehicle") or {}
            vehicle = veh_obj.get("number", "") if isinstance(veh_obj, dict) else ""

            # Speed: convert from km/h to mph
            KPH_TO_MPH = 0.621371
            max_speed_kph = float(ev.get("max_vehicle_speed", 0) or 0)
            posted_kph = float(ev.get("min_posted_speed_limit_in_kph", 0) or 0)
            speed = round(max_speed_kph * KPH_TO_MPH, 1)
            posted = round(posted_kph * KPH_TO_MPH, 1)
            over = round(speed - posted, 1)

            duration = ev.get("duration", 0) or 0
            start_time = ev.get("start_time", "")
            severity = ""
            meta = ev.get("metadata") or {}
            if isinstance(meta, dict):
                severity = meta.get("severity", "")

            # Division/yard from vehicle number prefix (e.g. "JOU-RAT-2364")
            div, yard, company = parse_vehicle_number(vehicle)

            tier = speeding_tier(speed, posted)
            lat = ev.get("start_lat") or ev.get("latitude") or ""
            lon = ev.get("start_lon") or ev.get("longitude") or ""
            location = f"{lat}, {lon}" if lat and lon else ""

            ev_date = start_time[:10] if start_time else ""

            all_events.append({
                "date": ev_date,
                "driver": driver,
                "vehicle": vehicle,
                "yard": yard,
                "division": div,
                "company": company,
                "max_speed": speed,
                "posted_speed": posted,
                "over_by": over,
                "duration": format_duration(duration),
                "severity": severity,
                "tier": tier,
                "location": location,
                "maps_link": f"https://www.google.com/maps?q={lat},{lon}" if lat and lon else "",
            })

        print(f"    Page {page}: {len(items)} events (total so far: {len(all_events)})")
        page += 1
        time.sleep(2)

    print(f"    Total speeding events: {len(all_events)}")
    return all_events


def build_casing_vehicle_lookup():
    """Fetch all vehicles from Motive v1 and build Casing lookup."""
    print("    Building Casing vehicle lookup from /v1/vehicles...")
    headers = {"X-Api-Key": MOTIVE_KEY}
    vehicle_drivers = {}
    vehicle_yards = {}
    casing_vehicles = set()
    page = 1

    while True:
        resp = requests.get(
            f"{MOTIVE_BASE_V1}/vehicles",
            headers=headers, params={"per_page": 100, "page_no": page}, timeout=30
        )
        if not resp.ok:
            break
        data = resp.json()
        vehicles = data.get("vehicles", [])
        if not vehicles:
            break

        for wrapper in vehicles:
            v = wrapper.get("vehicle", wrapper)
            num = v.get("number", "")
            if not num:
                continue
            group_ids = v.get("group_ids", [])
            yard = None
            for gid in group_ids:
                if gid in CASING_GROUP_IDS:
                    yard = CASING_GROUP_IDS[gid]
                    break
            if yard is None:
                continue
            casing_vehicles.add(num)
            vehicle_yards[num] = yard
            for field in ("current_driver", "permanent_driver"):
                d = v.get(field)
                if d and isinstance(d, dict):
                    name = f"{d.get('first_name', '')} {d.get('last_name', '')}".strip()
                    if name:
                        vehicle_drivers[num] = name
                        break

        pag = data.get("pagination", {})
        if page * 100 >= pag.get("total", 0):
            break
        page += 1

    print(f"    Found {len(casing_vehicles)} Casing vehicles")
    return vehicle_drivers, vehicle_yards, casing_vehicles


def fetch_motive_camera(start_date, end_date):
    """Fetch camera events from Motive API. Uses vehicle lookup for Casing filtering."""
    print(f"  Fetching Motive camera events (filtering {start_date} to {end_date})...")

    # Build vehicle lookup first
    vehicle_drivers, vehicle_yards, casing_vehicles = build_casing_vehicle_lookup()

    headers = {"X-Api-Key": MOTIVE_KEY}
    all_events = []
    page = 1
    total_fetched = 0

    while page <= 100:
        params = {"per_page": 100, "page_no": page}
        for attempt in range(3):
            try:
                resp = requests.get(
                    f"{MOTIVE_BASE_V2}/driver_performance_events",
                    headers=headers, params=params, timeout=60
                )
                break
            except requests.exceptions.ConnectionError:
                print(f"    Connection reset on page {page}, retry {attempt+1}/3...")
                time.sleep(10 * (attempt + 1))
        else:
            print(f"    Failed after 3 retries on page {page}, stopping")
            break

        if not resp.ok:
            print(f"    API error {resp.status_code} on page {page}")
            break

        data = resp.json()
        items = data.get("driver_performance_events", [])
        if not items:
            break

        for item in items:
            ev = item.get("driver_performance_event", item)

            # Get vehicle number
            veh_obj = ev.get("vehicle") or {}
            if isinstance(veh_obj, dict):
                vehicle_number = veh_obj.get("number", "")
            else:
                vehicle_number = str(veh_obj) if veh_obj else ""

            # Filter to Casing vehicles
            if vehicle_number not in casing_vehicles:
                continue

            # Skip uncoachable (false positives)
            if ev.get("coaching_status", "") == "uncoachable":
                continue

            # Date filter
            start_time = ev.get("start_time", "") or ev.get("event_time", "") or ""
            if not start_time:
                continue
            ev_date = start_time[:10]
            if ev_date < start_date or ev_date > end_date:
                continue

            event_type = ev.get("type", "") or ev.get("event_type", "") or ev.get("behavior_type", "") or ""
            speed = ev.get("start_speed") or ev.get("max_speed") or 0
            duration = ev.get("duration", 0) or 0

            driver = vehicle_drivers.get(vehicle_number, "Unknown")
            yard = vehicle_yards.get(vehicle_number, "Unknown")
            tier = "RED" if event_type in RED_CAMERA_TYPES else "ORANGE"
            display = event_type.replace("_", " ").title() if event_type else ""

            all_events.append({
                "date": ev_date,
                "driver": driver,
                "vehicle": vehicle_number,
                "yard": yard,
                "division": "Casing",
                "company": "BRHAS",
                "event_type": event_type,
                "display_name": display,
                "tier": tier,
                "speed": speed,
                "duration": format_duration(duration),
                "coaching_status": ev.get("coaching_status", ""),
                "coached_at": ev.get("coached_at", "") or "",
            })

        total_fetched += len(items)
        print(f"    Page {page}: {len(items)} raw, {len(all_events)} Casing in-range (fetched: {total_fetched})")

        pag = data.get("pagination", {})
        next_cursor = pag.get("next_cursor") or pag.get("next_page_cursor")
        if next_cursor:
            page += 1
            time.sleep(2)
            continue
        total_api = pag.get("total", 0)
        if total_api and total_fetched >= total_api:
            break
        if not total_api and len(items) < 100:
            break
        page += 1
        time.sleep(2)

    print(f"    Total camera events (Casing, in range): {len(all_events)}")
    return all_events


# ==============================================================================
# PUSH TO SHEETS
# ==============================================================================
def get_sheets_client(creds_file=None):
    """Auth with Google Sheets."""
    creds_json = os.environ.get("GOOGLE_SHEETS_CREDS_JSON", "")
    if creds_json:
        info = json.loads(creds_json)
        creds = Credentials.from_service_account_info(info, scopes=SCOPES)
    elif creds_file:
        creds = Credentials.from_service_account_file(creds_file, scopes=SCOPES)
    else:
        creds = Credentials.from_service_account_file("brhas-safety-b5a44478f315.json", scopes=SCOPES)
    return gspread.authorize(creds)


def push_rows_batched(worksheet, rows, batch_size=500):
    """Append rows in batches to avoid API limits."""
    for i in range(0, len(rows), batch_size):
        batch = rows[i:i + batch_size]
        worksheet.append_rows(batch, value_input_option="USER_ENTERED")
        print(f"      Pushed rows {i+1}-{i+len(batch)}")
        if i + batch_size < len(rows):
            time.sleep(2)


def push_speeding_history(client, events):
    """Push historical speeding to Sheets."""
    if not events:
        return
    wb = client.open_by_key(TRANSACTIONAL_SHEET_ID)
    ws = wb.worksheet("Speeding")

    # Get existing dates to avoid duplicates
    existing = set(ws.col_values(1)[1:])

    rows = []
    for e in events:
        if e["date"] in existing:
            continue
        rows.append([
            e["date"], e["driver"], e["vehicle"], e["yard"],
            e["division"], e["company"], e["max_speed"], e["posted_speed"],
            e["over_by"], e["duration"], e["severity"], e["tier"],
            e["location"], e["maps_link"],
        ])

    if not rows:
        print("    Speeding: all dates already exist, skipping")
        return

    # Sort by date
    rows.sort(key=lambda r: r[0])
    print(f"    Pushing {len(rows)} speeding rows...")
    push_rows_batched(ws, rows)


def push_camera_history(client, events):
    """Push historical camera events to Sheets."""
    if not events:
        return
    wb = client.open_by_key(TRANSACTIONAL_SHEET_ID)
    ws = wb.worksheet("Camera")

    existing = set(ws.col_values(1)[1:])

    rows = []
    for e in events:
        if e["date"] in existing:
            continue
        rows.append([
            e["date"], e["driver"], e["vehicle"], e["yard"],
            e["division"], e["company"], e["event_type"], e["display_name"],
            e["tier"], e["speed"], e["duration"],
            e["coaching_status"], e["coached_at"],
        ])

    if not rows:
        print("    Camera: all dates already exist, skipping")
        return

    rows.sort(key=lambda r: r[0])
    print(f"    Pushing {len(rows)} camera rows...")
    push_rows_batched(ws, rows)


def push_observations_history(client, observations):
    """Push historical observations to Sheets."""
    if not observations:
        return
    wb = client.open_by_key(TRANSACTIONAL_SHEET_ID)
    ws = wb.worksheet("Observations")

    existing = set(ws.col_values(1)[1:])

    rows = []
    for o in observations:
        if o["date"] in existing:
            continue
        svc = o.get("service_line", "")
        rows.append([
            o["date"], o["observer"], o["type"], o["description"],
            o["location"], svc_to_div(svc), svc_to_company(svc),
            o["report_number"],
        ])

    if not rows:
        print("    Observations: all dates already exist, skipping")
        return

    rows.sort(key=lambda r: r[0])
    print(f"    Pushing {len(rows)} observation rows...")
    push_rows_batched(ws, rows)


def push_incidents_history(client, incidents):
    """Push historical incidents to Sheets."""
    if not incidents:
        return
    wb = client.open_by_key(TRANSACTIONAL_SHEET_ID)
    ws = wb.worksheet("Incidents")

    existing = set(ws.col_values(1)[1:])

    rows = []
    for inc in incidents:
        if inc["date"] in existing:
            continue
        svc = inc.get("service_line", "")
        rows.append([
            inc["date"], inc["employee"], inc["type"], inc["description"],
            inc["location"], svc_to_div(svc), svc_to_company(svc),
            inc["report_number"],
        ])

    if not rows:
        print("    Incidents: all dates already exist, skipping")
        return

    rows.sort(key=lambda r: r[0])
    print(f"    Pushing {len(rows)} incident rows...")
    push_rows_batched(ws, rows)


def push_assessments_history(client, assessments):
    """Push historical assessments to Sheets."""
    if not assessments:
        return
    wb = client.open_by_key(TRANSACTIONAL_SHEET_ID)
    ws = wb.worksheet("Assessments")

    existing = set(ws.col_values(1)[1:])

    rows = []
    for a in assessments:
        if a["date"] in existing:
            continue
        rows.append([
            a["date"], a["assessor"], a["form_name"], a["division"],
            a["company"], a["location"], a["report_number"], a["link"],
        ])

    if not rows:
        print("    Assessments: all dates already exist, skipping")
        return

    rows.sort(key=lambda r: r[0])
    print(f"    Pushing {len(rows)} assessment rows...")
    push_rows_batched(ws, rows)


# ==============================================================================
# MAIN
# ==============================================================================
def main():
    print("=" * 60)
    print("BACKFILL SHEETS -- Historical data load")
    print(f"  KPA range:    {KPA_START} to {TODAY}")
    print(f"  Motive range: {MOTIVE_START} to {TODAY}")
    print("=" * 60)

    if not KPA_TOKEN:
        print("ERROR: KPA_API_TOKEN env var required")
        sys.exit(1)
    if not MOTIVE_KEY:
        print("ERROR: MOTIVE_API_KEY env var required")
        sys.exit(1)
    if not TRANSACTIONAL_SHEET_ID:
        print("ERROR: TRANSACTIONAL_SHEET_ID env var required")
        sys.exit(1)

    creds_file = sys.argv[1] if len(sys.argv) > 1 else None

    # --- Fetch KPA data (2025 + 2026) ---
    print("\n--- KPA DATA (2025-01-01 to today) ---")
    observations, near_misses = fetch_kpa_observations(KPA_START, TODAY)
    time.sleep(2)
    incidents = fetch_kpa_incidents(KPA_START, TODAY)
    time.sleep(2)
    assessments = fetch_kpa_assessments(KPA_START, TODAY)

    # --- Fetch Motive data (2026 only) ---
    print("\n--- MOTIVE DATA (2026-01-01 to today) ---")
    speeding = fetch_motive_speeding(MOTIVE_START, TODAY)
    time.sleep(2)
    camera = fetch_motive_camera(MOTIVE_START, TODAY)

    # --- Clear existing data and push ---
    print("\n--- PUSHING TO GOOGLE SHEETS ---")
    print("  Authenticating...")
    client = get_sheets_client(creds_file)

    # Clear all transactional tabs first (fresh backfill)
    print("  Clearing existing transactional data...")
    wb = client.open_by_key(TRANSACTIONAL_SHEET_ID)
    for tab_name in ["Speeding", "Camera", "Observations", "Incidents", "Assessments"]:
        ws = wb.worksheet(tab_name)
        header = ws.row_values(1)
        ws.clear()
        ws.update(range_name="A1", values=[header])
        print(f"    Cleared {tab_name}")
    time.sleep(2)

    # Push all historical data
    print("\n  Pushing historical data...")
    try:
        push_speeding_history(client, speeding)
    except Exception as e:
        print(f"  ERROR pushing speeding: {e}")

    time.sleep(2)
    try:
        push_camera_history(client, camera)
    except Exception as e:
        print(f"  ERROR pushing camera: {e}")

    time.sleep(2)
    try:
        # Combine observations + near misses
        all_obs = observations + near_misses
        push_observations_history(client, all_obs)
    except Exception as e:
        print(f"  ERROR pushing observations: {e}")

    time.sleep(2)
    try:
        push_incidents_history(client, incidents)
    except Exception as e:
        print(f"  ERROR pushing incidents: {e}")

    time.sleep(2)
    try:
        push_assessments_history(client, assessments)
    except Exception as e:
        print(f"  ERROR pushing assessments: {e}")

    # --- Summary ---
    print("\n" + "=" * 60)
    print("BACKFILL COMPLETE")
    print(f"  Observations: {len(observations)} + {len(near_misses)} near misses")
    print(f"  Incidents:    {len(incidents)}")
    print(f"  Assessments:  {len(assessments)}")
    print(f"  Speeding:     {len(speeding)}")
    print(f"  Camera:       {len(camera)}")
    print(f"  TOTAL:        {len(observations) + len(near_misses) + len(incidents) + len(assessments) + len(speeding) + len(camera)} rows")
    print("=" * 60)


if __name__ == "__main__":
    main()
