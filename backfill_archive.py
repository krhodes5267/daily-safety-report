"""
BACKFILL_ARCHIVE.PY -- One-time historical data backfill for dashboard archive
==============================================================================
Fetches ALL data in bulk from KPA and Motive APIs, then partitions by date
into individual archive/YYYY-MM-DD.json files.

Design: fetch-once-partition (NOT per-day API calls).
  - KPA: responses.flat has no updated_before, so we fetch all, filter client-side
  - Motive camera: v2 API ignores date params, so we paginate ALL events once
  - Motive speeding: v1 supports date ranges, but bulk is faster than 480 calls

Usage:
    KPA_API_TOKEN=... MOTIVE_API_KEY=... python backfill_archive.py --start 2025-01-01 --end 2026-04-20
    KPA_API_TOKEN=... MOTIVE_API_KEY=... python backfill_archive.py --start 2026-04-01 --end 2026-04-20  # quick test
"""
import calendar
import csv
import json
import os
import re
import sys
import time
import argparse
from collections import defaultdict
from datetime import datetime, date, timedelta, timezone
from io import StringIO

import requests

from api_config import (
    KPA_API_BASE, MOTIVE_BASE_V1, MOTIVE_BASE_V2,
    OBSERVATION_FORM_ID, INCIDENT_FORM_ID, ASSESSMENT_FORM_IDS,
    OBS_TYPE_HASH, OBS_DESC_HASH, OBS_LOCATION_HASH, OBS_YARD_HASH,
    INC_TYPE_HASH, INC_EMPLOYEE_HASH, INC_DESC_HASH, INC_LOCATION_HASH,
    SVC_LINE_HASH, COMPANY_HASH,
    SERVICE_LINE_HASHES, KPA_COMPANY_MAP, KPA_SVC_TO_DIVISION,
    CASING_GROUP_IDS, GROUP_ID_MAP, FORM_DIVISION_MAP,
    Q1_MAN_HOURS, SAFETY_FORMS,
)

KPA_TOKEN = os.environ.get("KPA_API_TOKEN", "")
MOTIVE_KEY = os.environ.get("MOTIVE_API_KEY", "")

KPH_TO_MPH = 0.621371

KNOWN_YARDS = ["Bryan", "Hobbs", "Jourdanton", "Kilgore", "Laredo", "Midland",
               "Levelland", "Barstow", "Ohio", "Pennsylvania", "Oklahoma",
               "North Dakota", "Lubbock", "Seminole", "Odessa", "Corporate"]


def normalize_yard(raw_yard):
    """Normalize a raw KPA yard/district field to a known yard name."""
    if not raw_yard:
        return ""
    lower = raw_yard.lower().strip()
    for y in KNOWN_YARDS:
        if y.lower() in lower:
            return y
    return ""

# Camera event tier classification
RED_CAMERA_TYPES = {
    "cell_phone", "drowsiness", "distraction", "close_following",
    "forward_collision_warning", "collision", "near_collision",
    "stop_sign", "stop_sign_violation", "unsafe_lane_change", "lane_swerving",
}

# Vehicle name prefix maps (from backfill_sheets.py)
YARD_PREFIXES = {
    "MID": "Midland", "BRY": "Bryan", "KIL": "Kilgore",
    "HOB": "Hobbs", "JOU": "Jourdanton", "LAR": "Laredo",
    "SAN": "San Angelo", "SA": "San Angelo",
    "TOW": "Pennsylvania", "PA": "Pennsylvania",
    "OH": "Ohio", "OK": "Oklahoma", "ND": "North Dakota",
    "BAR": "Barstow", "LL": "Levelland", "WIN": "Unknown",
    "DS": "North Dakota",
}
DIV_PREFIXES = {
    "CSG": "Casing", "RAT": "Rathole", "ANC": "Anchors",
    "PP": "Poly Pipe", "PL": "Pit Lining", "CON": "Construction",
    "ENV": "Environmental", "FEN": "Fencing", "DT": "Drilling Tools",
    "VAL": "Valor", "BTI": "BTI", "TD": "Transcend",
    "WTC": "Water/Construction", "FAB": "Fabrication",
    "PER": "Permian", "SS": "Permian",
}
SOLO_DIV_PREFIXES = {
    "ENV": "Environmental", "BTI": "BTI", "VAL": "Valor",
    "POL": "Poly Pipe", "PIT": "Pit Lining", "ANC": "Anchors",
    "FEN": "Fencing", "TD": "Transcend", "CON": "Construction",
}


# ==============================================================================
# HELPERS
# ==============================================================================
def format_duration(seconds):
    if not seconds:
        return ""
    s = int(seconds)
    if s < 60:
        return f"{s}s"
    m = s // 60
    s = s % 60
    return f"{m}m {s}s" if s > 0 else f"{m}m"


def parse_vehicle_number(vnum):
    """Parse vehicle number into (division, yard, company)."""
    if not vnum:
        return ("", "", "BRHAS")
    clean = vnum.strip().upper().split(" ")[0]
    parts = clean.split("-")
    yard, div, company = "", "", "BRHAS"

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
    elif div == "Permian":
        company = "Permian"
    return (div, yard, company)


def speeding_tier(speed, posted):
    over = speed - posted
    if over >= 20 or speed >= 90:
        return "RED"
    elif over >= 15:
        return "ORANGE"
    return "YELLOW"


def get_svc_line(row):
    for h in SERVICE_LINE_HASHES:
        val = row.get(h, "")
        if val and val not in ("Service Line", "service_line", "Division", "division", ""):
            return val
    return ""


def svc_to_div(svc):
    if not svc:
        return ""
    return KPA_SVC_TO_DIVISION.get(svc, svc)


# ==============================================================================
# KPA BULK FETCH
# ==============================================================================
def kpa_fetch_csv(form_id, updated_after_ms):
    """Fetch ALL KPA responses for a form as CSV rows."""
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
                resp = requests.post(
                    f"{KPA_API_BASE}/responses.flat",
                    json=payload, timeout=180,
                )
                break
            except (requests.exceptions.ConnectionError, requests.exceptions.ReadTimeout):
                print(f"    KPA timeout on form {form_id} page {page}, retry {attempt+1}/3...")
                time.sleep(15 * (attempt + 1))
        else:
            print(f"    Failed after 3 retries on form {form_id} page {page}")
            break
        if not resp.ok:
            print(f"    KPA API error {resp.status_code} for form {form_id} page {page}")
            break
        text = resp.text.strip()
        if not text:
            break
        reader = csv.DictReader(StringIO(text))
        rows = list(reader)
        data = [r for r in rows if r.get("date", "") != "Date"]
        if not data:
            break
        all_rows.extend(data)
        if len(data) < 100:
            break
        page += 1
        time.sleep(1.5)
        if page > 200:
            break
    return all_rows


def fetch_all_kpa_observations(start_date):
    """Fetch all observations from start_date to now, return as list of dicts."""
    print(f"  KPA observations (from {start_date})...")
    start_dt = datetime.strptime(start_date, "%Y-%m-%d")
    updated_after = int((start_dt - timedelta(days=1)).timestamp() * 1000)

    rows = kpa_fetch_csv(OBSERVATION_FORM_ID, updated_after)
    print(f"    Observation rows: {len(rows)}")

    # Also TD observations
    td_rows = kpa_fetch_csv(484193, updated_after)
    print(f"    TD observation rows: {len(td_rows)}")
    rows.extend(td_rows)

    results = []
    for row in rows:
        row_date = (row.get("date") or "")[:10]
        if not row_date or row_date < start_date:
            continue
        obs_type = row.get(OBS_TYPE_HASH, "Other") or "Other"
        svc = get_svc_line(row)
        results.append({
            "date": row_date,
            "observer": row.get("Name", "") or row.get("observer", "") or "",
            "type": obs_type,
            "description": row.get(OBS_DESC_HASH, "") or "",
            "location": row.get(OBS_LOCATION_HASH, "") or "",
            "service_line": svc,
            "report_number": row.get("report number", "") or "",
            "yard": normalize_yard(row.get(OBS_YARD_HASH, "")),
        })
    print(f"    Total observations: {len(results)}")
    return results


def fetch_all_kpa_incidents(start_date):
    """Fetch all incidents from start_date to now."""
    print(f"  KPA incidents (from {start_date})...")
    start_dt = datetime.strptime(start_date, "%Y-%m-%d")
    updated_after = int((start_dt - timedelta(days=1)).timestamp() * 1000)

    rows = kpa_fetch_csv(INCIDENT_FORM_ID, updated_after)
    print(f"    Incident rows: {len(rows)}")

    results = []
    for row in rows:
        row_date = (row.get("date") or "")[:10]
        if not row_date or row_date < start_date:
            continue
        svc = get_svc_line(row)
        results.append({
            "date": row_date,
            "employee": row.get(INC_EMPLOYEE_HASH, "") or row.get("Name", "") or "",
            "type": row.get(INC_TYPE_HASH, "") or "",
            "description": row.get(INC_DESC_HASH, "") or "",
            "location": row.get(INC_LOCATION_HASH, "") or "",
            "service_line": svc,
            "company": row.get(COMPANY_HASH, "") or "",
            "report_number": row.get("report number", "") or "",
        })
    print(f"    Total incidents: {len(results)}")
    return results


FORM_NAME_MAP = {
    381707: "CSG - Safety Casing Field Assessment",
    229645: "CSG - Pre/Post Trip Inspection",
    385365: "TD - Rig Inspection",
    484193: "TD - Observation Card",
    226217: "WS - Poly Pipe Field Assessment",
    386087: "WS - Pit Lining Field Assessment",
    172295: "Construction - Site Safety Review",
    153181: "RH - Rathole Field Assessment",
    152018: "Vehicle Inspection Checklist",
    152034: "HSE - Workplace Inspection Checklist",
}


def fetch_all_kpa_assessments(start_date):
    """Fetch all assessments across all form types."""
    print(f"  KPA assessments (from {start_date})...")
    start_dt = datetime.strptime(start_date, "%Y-%m-%d")
    updated_after = int((start_dt - timedelta(days=1)).timestamp() * 1000)

    results = []
    for form_id in ASSESSMENT_FORM_IDS:
        time.sleep(1.5)
        rows = kpa_fetch_csv(form_id, updated_after)
        division = FORM_DIVISION_MAP.get(form_id, "")
        form_name = FORM_NAME_MAP.get(form_id, f"Form {form_id}")
        count = 0
        for row in rows:
            row_date = (row.get("date") or "")[:10]
            if not row_date or row_date < start_date:
                continue
            svc = get_svc_line(row)
            results.append({
                "date": row_date,
                "assessor": row.get("Name", "") or row.get("observer", "") or "",
                "form_id": form_id,
                "form_name": form_name,
                "division": division or svc_to_div(svc),
                "location": row.get(OBS_LOCATION_HASH, "") or "",
                "report_number": row.get("report number", "") or "",
            })
            count += 1
        if count > 0:
            print(f"    Form {form_id} ({division}): {count} rows")
    print(f"    Total assessments: {len(results)}")
    return results


# ==============================================================================
# MOTIVE BULK FETCH
# ==============================================================================
def fetch_all_motive_speeding(start_date, end_date):
    """Fetch ALL speeding events in date range."""
    print(f"  Motive speeding ({start_date} to {end_date})...")
    headers = {"X-Api-Key": MOTIVE_KEY}
    all_events = []
    page = 1

    while page <= 200:
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
                    headers=headers, params=params, timeout=60,
                )
                break
            except requests.exceptions.ConnectionError:
                print(f"    Connection reset on page {page}, retry {attempt+1}/3...")
                time.sleep(10 * (attempt + 1))
        else:
            print(f"    Failed after 3 retries on page {page}")
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
            drv_obj = ev.get("driver") or {}
            if isinstance(drv_obj, dict) and drv_obj:
                first = drv_obj.get("first_name", "") or ""
                last = drv_obj.get("last_name", "") or ""
                driver = f"{first} {last}".strip() or "Unknown"
            else:
                driver = "Unknown"

            veh_obj = ev.get("vehicle") or {}
            vehicle = veh_obj.get("number", "") if isinstance(veh_obj, dict) else ""

            max_speed_kph = float(ev.get("max_vehicle_speed", 0) or ev.get("avg_vehicle_speed", 0) or 0)
            posted_kph = float(ev.get("min_posted_speed_limit_in_kph", 0) or 0)
            speed = round(max_speed_kph * KPH_TO_MPH, 1)
            posted = round(posted_kph * KPH_TO_MPH, 1)
            over = round(speed - posted, 1)

            duration = ev.get("duration", 0) or 0
            start_time = ev.get("start_time", "")
            tier = speeding_tier(speed, posted)
            div, yard, company = parse_vehicle_number(vehicle)

            lat = ev.get("start_lat") or ev.get("latitude") or ""
            lon = ev.get("start_lon") or ev.get("longitude") or ""
            location = f"{lat}, {lon}" if lat and lon else ""
            ev_date = start_time[:10] if start_time else ""

            all_events.append({
                "date": ev_date,
                "driver": driver,
                "vehicle": vehicle,
                "speed": speed,
                "posted_speed": posted,
                "overspeed": over,
                "duration": format_duration(duration),
                "severity": (ev.get("metadata") or {}).get("severity", "") if isinstance(ev.get("metadata"), dict) else "",
                "time": start_time,
                "location": location,
                "maps_link": f"https://www.google.com/maps?q={lat},{lon}" if lat and lon else "",
                "tier": tier,
                "division": div or "Unknown",
                "yard": yard or "Unknown",
            })

        if page % 10 == 0:
            print(f"    Page {page}: {len(all_events)} total so far")
        total = data.get("total", 0)
        if total and page * 100 >= total:
            break
        page += 1
        time.sleep(2)

    print(f"    Total speeding events: {len(all_events)}")
    return all_events


def fetch_all_motive_mileage(start_date, end_date):
    """Fetch ALL IFTA trip mileage in date range, month by month."""
    print(f"  Motive mileage (IFTA trips, {start_date} to {end_date})...")
    headers = {"X-Api-Key": MOTIVE_KEY}
    all_trips = []

    # Fetch month by month (API may limit large ranges)
    current = datetime.strptime(start_date, "%Y-%m-%d").date()
    end_dt = datetime.strptime(end_date, "%Y-%m-%d").date()

    while current <= end_dt:
        month_end = current.replace(day=28) + timedelta(days=4)
        month_end = month_end.replace(day=1) - timedelta(days=1)  # last day of month
        if month_end > end_dt:
            month_end = end_dt

        page = 1
        month_count = 0
        while page <= 200:
            params = {
                "per_page": 100,
                "page_no": page,
                "start_date": current.isoformat(),
                "end_date": month_end.isoformat(),
            }
            for attempt in range(3):
                try:
                    resp = requests.get(
                        f"{MOTIVE_BASE_V1}/ifta/trips",
                        headers=headers, params=params, timeout=60,
                    )
                    break
                except requests.exceptions.ConnectionError:
                    time.sleep(10 * (attempt + 1))
            else:
                break

            if not resp.ok:
                break

            data = resp.json()
            items = data.get("ifta_trips", [])
            if not items:
                break

            for item in items:
                trip = item.get("ifta_trip", item)
                veh = trip.get("vehicle") or {}
                vehicle = veh.get("number", "") if isinstance(veh, dict) else ""
                distance = float(trip.get("distance", 0) or 0)
                trip_date = (trip.get("date") or "")[:10]
                if distance > 0 and trip_date:
                    div, yard, company = parse_vehicle_number(vehicle)
                    all_trips.append({
                        "date": trip_date,
                        "vehicle": vehicle,
                        "distance": round(distance, 2),
                        "division": div or "Unknown",
                        "yard": yard or "Unknown",
                    })
                    month_count += 1

            total = data.get("pagination", {}).get("total", 0)
            if total and page * 100 >= total:
                break
            page += 1
            time.sleep(1)

        print(f"    {current.strftime('%Y-%m')}: {month_count} trips")
        current = month_end + timedelta(days=1)

    print(f"    Total mileage trips: {len(all_trips)}")
    return all_trips


def fetch_all_motive_camera():
    """Fetch ALL camera events (no date filter -- API ignores it)."""
    print(f"  Motive camera events (all pages, no date filter)...")

    # Build Casing vehicle lookup first
    print("    Building Casing vehicle lookup...")
    headers = {"X-Api-Key": MOTIVE_KEY}
    vehicle_drivers = {}
    vehicle_yards = {}
    casing_vehicles = set()
    page = 1

    while True:
        resp = requests.get(
            f"{MOTIVE_BASE_V1}/vehicles",
            headers=headers, params={"per_page": 100, "page_no": page}, timeout=30,
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

    # Fetch all camera events
    all_events = []
    page = 1
    total_fetched = 0

    while page <= 200:
        params = {"per_page": 100, "page_no": page}
        for attempt in range(3):
            try:
                resp = requests.get(
                    f"{MOTIVE_BASE_V2}/driver_performance_events",
                    headers=headers, params=params, timeout=60,
                )
                break
            except requests.exceptions.ConnectionError:
                print(f"    Connection reset on page {page}, retry {attempt+1}/3...")
                time.sleep(10 * (attempt + 1))
        else:
            print(f"    Failed after 3 retries on page {page}")
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
            veh_obj = ev.get("vehicle") or {}
            vehicle_number = veh_obj.get("number", "") if isinstance(veh_obj, dict) else ""

            if vehicle_number not in casing_vehicles:
                continue

            # Skip uncoachable (dismissed/false positives)
            if ev.get("coaching_status", "") == "uncoachable":
                continue

            start_time = ev.get("start_time", "") or ev.get("event_time", "") or ""
            if not start_time:
                continue
            ev_date = start_time[:10]

            event_type = ev.get("type", "") or ev.get("event_type", "") or ev.get("behavior_type", "") or ""
            speed_kph = ev.get("start_speed") or ev.get("max_speed") or 0
            speed = round(float(speed_kph) * KPH_TO_MPH, 1) if speed_kph else 0
            duration = ev.get("duration", 0) or 0

            driver = vehicle_drivers.get(vehicle_number, "Unknown")
            yard = vehicle_yards.get(vehicle_number, "Unknown")
            tier = "RED" if event_type in RED_CAMERA_TYPES else "ORANGE"
            display = event_type.replace("_", " ").title() if event_type else ""

            all_events.append({
                "date": ev_date,
                "driver": driver,
                "vehicle": vehicle_number,
                "event_type": event_type,
                "raw_type": event_type,
                "display_name": display,
                "tier": tier,
                "speed": speed,
                "duration": format_duration(duration),
                "time": start_time,
                "yard": yard,
                "coaching_status": ev.get("coaching_status", "") or "unknown",
                "coached_at": ev.get("coached_at", "") or "",
            })

        total_fetched += len(items)
        if page % 10 == 0:
            print(f"    Page {page}: {total_fetched} raw, {len(all_events)} Casing kept")

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

    print(f"    Total camera events (Casing, non-uncoachable): {len(all_events)}")
    return all_events


# ==============================================================================
# YTD STATS COMPUTATION
# ==============================================================================
def compute_ytd_stats(incidents, as_of_date):
    """Compute YTD TRIR/DART stats as of a specific date."""
    # Man-hours (prorated from Q1 data)
    hours_by_company = {}
    for company, data in Q1_MAN_HOURS.items():
        monthly = data["monthly"]
        march_hours = monthly.get("2026-03", 0)
        company_ytd = 0.0
        for m in range(1, as_of_date.month + 1):
            key = f"{as_of_date.year}-{m:02d}"
            if key in monthly:
                if m == as_of_date.month:
                    days_in_month = calendar.monthrange(as_of_date.year, m)[1]
                    company_ytd += round(monthly[key] * as_of_date.day / days_in_month)
                else:
                    company_ytd += monthly[key]
            else:
                if m == as_of_date.month:
                    days_in_month = calendar.monthrange(as_of_date.year, m)[1]
                    company_ytd += round(march_hours * as_of_date.day / days_in_month)
                else:
                    company_ytd += march_hours
        hours_by_company[company] = round(company_ytd)

    ytd_hours = sum(hours_by_company.values())

    # Count recordables up to as_of_date
    ytd_rec = 0
    overall_last = None
    for inc in incidents:
        if "Recordable" not in (inc.get("type") or ""):
            continue
        d_str = inc.get("date", "")[:10]
        try:
            d = datetime.strptime(d_str, "%Y-%m-%d").date()
        except Exception:
            continue
        if d.year == as_of_date.year and d <= as_of_date:
            ytd_rec += 1
        if overall_last is None or d > overall_last:
            if d <= as_of_date:
                overall_last = d

    ytd_trir = round(ytd_rec * 200000 / ytd_hours, 2) if ytd_hours > 0 else 0
    days_rec = (as_of_date - overall_last).days if overall_last else None

    return {
        "report_date": as_of_date.isoformat(),
        "ytd_trir": ytd_trir,
        "ytd_dart": ytd_trir,
        "ytd_recordables": ytd_rec,
        "ytd_man_hours": ytd_hours,
        "days_since_lti": days_rec,
        "days_since_recordable": days_rec,
        "last_recordable_date": overall_last.isoformat() if overall_last else None,
    }


# ==============================================================================
# PARTITION + WRITE
# ==============================================================================
def partition_by_date(items):
    """Group a list of dicts by their 'date' field."""
    buckets = defaultdict(list)
    for item in items:
        d = item.get("date", "")[:10]
        if d:
            buckets[d].append(item)
    return buckets


def build_speeding_day(events):
    """Build speeding section for one day."""
    summary = {"red": 0, "orange": 0, "yellow": 0}
    by_div = {}
    for e in events:
        tier = e.get("tier", "YELLOW").lower()
        summary[tier] = summary.get(tier, 0) + 1
        div = e.get("division", "Unknown")
        by_div[div] = by_div.get(div, 0) + 1
    return {
        "events": events,
        "total_events": len(events),
        "summary": summary,
        "by_division": by_div,
        "repeat_offenders": {},
    }


def build_camera_day(events):
    """Build camera section for one day."""
    summary = {"red": 0, "orange": 0, "yellow": 0}
    by_yard = {}
    driver_counts = defaultdict(int)
    for e in events:
        tier = e.get("tier", "ORANGE").lower()
        summary[tier] = summary.get(tier, 0) + 1
        yard = e.get("yard", "Unknown")
        by_yard[yard] = by_yard.get(yard, 0) + 1
        driver_counts[e.get("driver", "Unknown")] += 1
    repeat = {d: c for d, c in driver_counts.items() if c >= 2 and d != "Unknown"}
    return {
        "events": events,
        "total_events": len(events),
        "summary": summary,
        "by_yard": by_yard,
        "repeat_offenders": repeat,
    }


def build_mileage_day(trips):
    """Build mileage section for one day."""
    total = 0
    by_division = defaultdict(float)
    by_yard = defaultdict(float)
    vehicles = defaultdict(float)
    for t in trips:
        dist = t.get("distance", 0)
        total += dist
        div = t.get("division", "Unknown")
        yard = t.get("yard", "Unknown")
        by_division[div] += dist
        by_yard[yard] += dist
        vehicles[t.get("vehicle", "")] += dist
    return {
        "total_miles": round(total, 1),
        "by_division": {k: round(v, 1) for k, v in sorted(by_division.items(), key=lambda x: -x[1])},
        "by_yard": {k: round(v, 1) for k, v in sorted(by_yard.items(), key=lambda x: -x[1])},
        "vehicle_count": len(vehicles),
    }


def build_kpa_day(obs_list, near_misses, inc_list, assess_list):
    """Build KPA section for one day."""
    by_type = {}
    for o in obs_list:
        t = o.get("type", "Other")
        by_type[t] = by_type.get(t, 0) + 1
    return {
        "observations": {
            "total": len(obs_list),
            "by_type": by_type,
            "details": obs_list,
        },
        "near_misses": near_misses,
        "incidents": inc_list,
        "assessments": {
            "total": len(assess_list),
            "details": assess_list,
        },
    }


# ==============================================================================
# MAIN
# ==============================================================================
def main():
    parser = argparse.ArgumentParser(description="Backfill archive files from APIs")
    parser.add_argument("--start", required=True, help="Start date YYYY-MM-DD")
    parser.add_argument("--end", required=True, help="End date YYYY-MM-DD")
    parser.add_argument("--output-dir", default="archive", help="Output directory")
    parser.add_argument("--skip-motive", action="store_true", help="Skip Motive API calls")
    parser.add_argument("--skip-kpa", action="store_true", help="Skip KPA API calls")
    args = parser.parse_args()

    start_date = args.start
    end_date = args.end
    output_dir = args.output_dir

    if not KPA_TOKEN:
        print("ERROR: KPA_API_TOKEN not set")
        sys.exit(1)
    if not MOTIVE_KEY:
        print("ERROR: MOTIVE_API_KEY not set")
        sys.exit(1)

    print("=" * 60)
    print(f"  BACKFILL ARCHIVE: {start_date} to {end_date}")
    print(f"  Output: {output_dir}/")
    print("=" * 60)

    # ------------------------------------------------------------------
    # PHASE A: Bulk fetch all data
    # ------------------------------------------------------------------
    print("\n== PHASE A: Bulk API Fetches ==\n")

    # KPA data
    all_observations = []
    all_incidents = []
    all_assessments = []
    if not args.skip_kpa:
        all_observations = fetch_all_kpa_observations(start_date)
        time.sleep(2)
        all_incidents = fetch_all_kpa_incidents(start_date)
        time.sleep(2)
        all_assessments = fetch_all_kpa_assessments(start_date)
        time.sleep(2)

    # Motive data (2026+ only)
    motive_start = max(start_date, "2026-01-01")
    all_speeding = []
    all_camera = []
    all_mileage = []
    if not args.skip_motive and motive_start <= end_date:
        all_speeding = fetch_all_motive_speeding(motive_start, end_date)
        time.sleep(5)
        all_mileage = fetch_all_motive_mileage(motive_start, end_date)
        time.sleep(5)
        all_camera = fetch_all_motive_camera()

    # ------------------------------------------------------------------
    # PHASE B: Partition by date + write archive files
    # ------------------------------------------------------------------
    print(f"\n== PHASE B: Partition + Write ==\n")

    obs_by_date = partition_by_date(all_observations)
    inc_by_date = partition_by_date(all_incidents)
    assess_by_date = partition_by_date(all_assessments)
    spd_by_date = partition_by_date(all_speeding)
    cam_by_date = partition_by_date(all_camera)
    mile_by_date = partition_by_date(all_mileage)

    # Separate near misses from observations
    obs_clean = defaultdict(list)
    nm_by_date = defaultdict(list)
    for d, obs_list in obs_by_date.items():
        for o in obs_list:
            if "near miss" in (o.get("type") or "").lower():
                nm_by_date[d].append(o)
            else:
                obs_clean[d].append(o)

    # Generate date range
    current = datetime.strptime(start_date, "%Y-%m-%d").date()
    end_dt = datetime.strptime(end_date, "%Y-%m-%d").date()

    os.makedirs(output_dir, exist_ok=True)
    file_count = 0
    total_bytes = 0

    while current <= end_dt:
        d_str = current.isoformat()

        # Speeding + camera + mileage only for 2026+
        spd_data = None
        cam_data = None
        mileage_data = None
        if d_str >= "2026-01-01":
            spd_events = spd_by_date.get(d_str, [])
            spd_data = build_speeding_day(spd_events)
            cam_events = cam_by_date.get(d_str, [])
            cam_data = build_camera_day(cam_events)
            mile_trips = mile_by_date.get(d_str, [])
            mileage_data = build_mileage_day(mile_trips) if mile_trips else None

        # KPA
        kpa_data = build_kpa_day(
            obs_clean.get(d_str, []),
            nm_by_date.get(d_str, []),
            inc_by_date.get(d_str, []),
            assess_by_date.get(d_str, []),
        )

        # YTD stats (only for 2026)
        ytd_data = None
        if current.year >= 2026:
            ytd_data = compute_ytd_stats(all_incidents, current)

        archive = {
            "date": d_str,
            "generated_at": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
            "speeding": spd_data,
            "camera": cam_data,
            "kpa": kpa_data,
            "ytd": ytd_data,
            "mileage": mileage_data,
            "cas": None,
            "training": None,
            "devices": None,
        }

        path = os.path.join(output_dir, f"{d_str}.json")
        with open(path, "w", encoding="utf-8") as f:
            json.dump(archive, f, separators=(",", ":"))

        size = os.path.getsize(path)
        total_bytes += size
        file_count += 1

        if file_count % 30 == 0:
            print(f"    Written {file_count} files ({total_bytes/1024/1024:.1f} MB)... latest: {d_str}")

        current += timedelta(days=1)

    print(f"\n  Done! {file_count} archive files, {total_bytes/1024/1024:.1f} MB total")
    print(f"  Location: {output_dir}/")


if __name__ == "__main__":
    main()
