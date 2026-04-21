"""
DASHBOARD_API.PY -- Flask API for Safety Dashboard historical data
==================================================================
Provides date-range queries for KPA, Motive, and YTD stats.
Deploy to Render.com free tier with gunicorn.

Usage (local):
    KPA_API_TOKEN=... MOTIVE_API_KEY=... python dashboard_api.py

Endpoints:
    GET /api/health          -> {"status": "ok"}
    GET /api/data?start=...&end=...  -> All data for date range
"""

import calendar
import csv
import json
import os
import re
import time
from datetime import datetime, date, timedelta, timezone
from io import StringIO

from flask import Flask, jsonify, request
from flask_cors import CORS
import requests as http_requests

from api_config import (
    KPA_API_BASE, MOTIVE_BASE_V1, MOTIVE_BASE_V2,
    OBSERVATION_FORM_ID, INCIDENT_FORM_ID, ASSESSMENT_FORM_IDS,
    OBS_TYPE_HASH, OBS_DESC_HASH, OBS_LOCATION_HASH, OBS_YARD_HASH,
    INC_TYPE_HASH, INC_EMPLOYEE_HASH, INC_DESC_HASH, INC_LOCATION_HASH,
    SVC_LINE_HASH, COMPANY_HASH,
    SERVICE_LINE_HASHES,
    KPA_COMPANY_MAP, KPA_SVC_TO_DIVISION, SAFETY_FORMS,
    CASING_GROUP_IDS, GROUP_ID_MAP, FORM_DIVISION_MAP,
    Q1_MAN_HOURS, KPA_RESPONSE_URL,
)

# ==============================================================================
# CONFIG
# ==============================================================================

app = Flask(__name__)
CORS(app, origins=["https://krhodes5267.github.io", "http://localhost:*"])

KPA_TOKEN = os.environ.get("KPA_API_TOKEN", "")
MOTIVE_KEY = os.environ.get("MOTIVE_API_KEY", "")

CENTRAL_TZ = timezone(timedelta(hours=-6))

# In-memory cache: {cache_key: {"data": ..., "expires": timestamp}}
_cache = {}
CACHE_TTL_SECONDS = 900  # 15 min for historical, overridden to 300 for today


# ==============================================================================
# CACHE
# ==============================================================================

def cache_get(key):
    entry = _cache.get(key)
    if entry and time.time() < entry["expires"]:
        return entry["data"]
    return None


def cache_set(key, data, ttl=CACHE_TTL_SECONDS):
    _cache[key] = {"data": data, "expires": time.time() + ttl}


# ==============================================================================
# KPA HELPERS
# ==============================================================================

def kpa_post(endpoint, payload, timeout=60):
    """POST to KPA API with token in JSON body."""
    payload["token"] = KPA_TOKEN
    resp = http_requests.post(
        KPA_API_BASE + "/" + endpoint,
        json=payload,
        timeout=timeout,
    )
    return resp


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


def kpa_csv_rows(form_id, updated_after_ms=None):
    """Fetch all rows from a KPA form as CSV, return list of dicts."""
    payload = {"form_id": form_id, "format": "csv"}
    if updated_after_ms:
        payload["updated_after"] = updated_after_ms
    resp = kpa_post("responses.flat", payload)
    if not resp.text or not resp.text.strip():
        return []
    reader = csv.DictReader(StringIO(resp.text))
    rows = list(reader)
    # Skip label header row
    if rows and rows[0].get("date", "") == "Date":
        rows = rows[1:]
    return rows


def kpa_paginated(endpoint, data_key, extra_params=None, max_pages=50):
    """Fetch paginated KPA JSON endpoint."""
    all_rows = []
    page = 1
    while page <= max_pages:
        payload = {"limit": 500, "page": page}
        if extra_params:
            payload.update(extra_params)
        resp = kpa_post(endpoint, payload, timeout=120)
        data = json.loads(resp.text)
        items = data.get(data_key, [])
        if not items:
            break
        all_rows.extend(items)
        last_page = data.get("paging", {}).get("last_page", 1)
        if page >= last_page:
            break
        page += 1
        time.sleep(1.5)
    return all_rows


# ==============================================================================
# MOTIVE HELPERS
# ==============================================================================

def motive_get(url, params=None, timeout=30):
    """GET from Motive API with API key header."""
    headers = {"X-Api-Key": MOTIVE_KEY}
    return http_requests.get(url, headers=headers, params=params, timeout=timeout)


# Cache vehicle -> (division, yard) mapping (built from vehicle list)
_vehicle_map_cache = {"data": None, "ts": 0}


def get_vehicle_division_map():
    """Fetch all vehicles and build number -> (division, yard) map using group IDs."""
    now = time.time()
    if _vehicle_map_cache["data"] and (now - _vehicle_map_cache["ts"]) < 3600:
        return _vehicle_map_cache["data"]

    veh_map = {}
    page = 1
    while page <= 50:
        try:
            resp = motive_get(MOTIVE_BASE_V1 + "/vehicles",
                              {"per_page": 100, "page_no": page}, timeout=30)
            data = resp.json()
        except Exception:
            break
        vehicles = data.get("vehicles", [])
        if not vehicles:
            break
        for v in vehicles:
            veh = v.get("vehicle", v)
            num = veh.get("number", "")
            gids = veh.get("group_ids", []) or []
            div, yard = resolve_vehicle_division(num, gids)
            if num:
                veh_map[num] = (div, yard)
        page += 1
    _vehicle_map_cache["data"] = veh_map
    _vehicle_map_cache["ts"] = now
    return veh_map


def resolve_vehicle_division(vehicle_name, group_ids=None):
    """Resolve a Motive vehicle to (division, yard) tuple."""
    if group_ids:
        for gid in group_ids:
            if gid in GROUP_ID_MAP:
                return GROUP_ID_MAP[gid]
    # Fallback: parse vehicle name prefix
    name = vehicle_name or ""
    if re.match(r"^\d+C\b", name) or re.match(r"^\d+C\s", name):
        return ("Casing", "Unknown")
    prefixes = [
        ("LL-RAT", "Rathole", "Levelland"),
        ("MID-RAT", "Rathole", "Midland"),
        ("BS-RAT", "Rathole", "Barstow"),
        ("JD-RAT", "Rathole", "Jourdanton"),
        ("DS-RAT", "Rathole", "North Dakota"),
        ("OH-RAT", "Rathole", "Ohio"),
        ("PA-RAT", "Rathole", "Pennsylvania"),
        ("OK-RAT", "Rathole", "Oklahoma"),
        ("TOW-RAT", "Rathole", "Unknown"),
        ("RAT-", "Rathole", "Unknown"),
        ("POL-", "Poly Pipe", "Midland"),
        ("PL-", "Pit Lining", "Midland"),
        ("ANC-", "Anchors", "Midland"),
        ("CON-", "Construction", "Midland"),
        ("ENV-", "Environmental", "Midland"),
        ("FEN-", "Fencing", "Midland"),
        ("BTI-", "BTI", "Midland"),
        ("TD", "Transcend", "Unknown"),
        ("VAL-", "Valor", "Levelland"),
    ]
    for prefix, div, yard in prefixes:
        if name.upper().startswith(prefix.upper()):
            return (div, yard)
    return ("Unknown", "Unknown")


# ==============================================================================
# DATA FETCHERS
# ==============================================================================

def fetch_speeding(start_date, end_date):
    """Fetch speeding events from Motive for date range."""
    if not MOTIVE_KEY:
        return []

    # Pre-fetch vehicle list to resolve division/yard via group IDs
    veh_map = get_vehicle_division_map()

    events = []
    page = 1
    while True:
        params = {
            "per_page": 100,
            "page_no": page,
            "start_date": start_date,
            "end_date": end_date,
        }
        try:
            resp = motive_get(MOTIVE_BASE_V1 + "/speeding_events", params)
            data = resp.json()
        except Exception:
            break
        items = data.get("speeding_events", [])
        if not items:
            break
        for item in items:
            ev = item if "vehicle_name" in item else item.get("speeding_event", item)
            # Motive API returns speeds in km/h -- convert to mph
            speed_kph = ev.get("avg_vehicle_speed", 0) or ev.get("speed", 0) or 0
            posted_kph = ev.get("min_posted_speed_limit_in_kph", 0) or ev.get("speed_limit", 0) or 0
            speed = round(speed_kph * 0.621371, 1)
            posted = round(posted_kph * 0.621371, 0)
            over = round(speed - posted, 1)
            # Tier (mph-based)
            if over >= 20 or speed >= 90:
                tier = "RED"
            elif over >= 15:
                tier = "ORANGE"
            else:
                tier = "YELLOW"
            # Vehicle -> division/yard via pre-fetched map
            veh_obj = ev.get("vehicle", {}) or {}
            vname = veh_obj.get("number", "") or ev.get("vehicle_name", "")
            # Look up from vehicle map first (has group IDs)
            if vname in veh_map:
                div, yard = veh_map[vname]
            else:
                # Fallback to name-based resolution
                div, yard = resolve_vehicle_division(vname)
            lat = ev.get("start_lat") or ev.get("latitude") or ev.get("start_latitude")
            lon = ev.get("start_lon") or ev.get("longitude") or ev.get("start_longitude")
            driver = ev.get("driver_name", "")
            if not driver:
                drv_obj = ev.get("driver") or {}
                if isinstance(drv_obj, dict):
                    driver = drv_obj.get("name", "") or "{} {}".format(
                        drv_obj.get("first_name", ""), drv_obj.get("last_name", "")).strip()
            events.append({
                "driver": driver or "Unknown",
                "vehicle": vname,
                "speed": speed,
                "posted_speed": posted,
                "overspeed": over,
                "duration": str(ev.get("duration", 0) or 0) + "s",
                "time": ev.get("start_time", ""),
                "location": "{}, {}".format(lat, lon) if lat and lon else "",
                "maps_link": "https://www.google.com/maps?q={},{}".format(lat, lon) if lat and lon else "",
                "tier": tier,
                "division": div,
                "yard": yard,
            })
        page += 1
        if page > 50:
            break
    return events


def fetch_camera(start_date, end_date):
    """Fetch camera events from Motive V2 for date range (Casing only).

    NOTE: Motive V2 driver_performance_events may ignore date params.
    We fetch all events and filter client-side by Central Time window.
    """
    if not MOTIVE_KEY:
        return []
    events = []
    # Central time window for filtering
    start_central = datetime.strptime(start_date, "%Y-%m-%d").replace(tzinfo=CENTRAL_TZ)
    end_central = datetime.strptime(end_date, "%Y-%m-%d").replace(
        hour=23, minute=59, second=59, tzinfo=CENTRAL_TZ
    )

    page = 1
    while True:
        params = {
            "per_page": 100,
            "page_no": page,
        }
        try:
            resp = motive_get(MOTIVE_BASE_V2 + "/driver_performance_events", params)
            data = resp.json()
        except Exception:
            break
        items = data.get("driver_performance_events", [])
        if not items:
            break
        for item in items:
            ev = item.get("driver_performance_event", item)
            # Filter to Casing vehicles
            veh = ev.get("vehicle", {}) or {}
            gids = veh.get("group_ids", []) or []
            casing_gids = set(CASING_GROUP_IDS.keys())
            is_casing = bool(set(gids) & casing_gids)
            if not is_casing:
                vname = veh.get("name", "")
                if not re.match(r"^\d+C\b", vname or ""):
                    continue
            # Client-side date filter (API may ignore date params)
            ev_time = ev.get("start_time", "")
            if ev_time:
                try:
                    ev_dt = datetime.fromisoformat(ev_time.replace("Z", "+00:00"))
                    ev_central = ev_dt.astimezone(CENTRAL_TZ)
                    if ev_central < start_central or ev_central > end_central:
                        continue
                except Exception:
                    pass
            # Resolve yard
            yard = "Unknown"
            for gid in gids:
                if gid in CASING_GROUP_IDS:
                    yard = CASING_GROUP_IDS[gid]
                    break
            raw_type = ev.get("type", "") or ev.get("event_type", "")
            display = raw_type.replace("_", " ").title()
            speed = ev.get("start_speed") or ev.get("max_speed") or 0
            # Tier based on type
            red_types = {
                "cell_phone", "drowsiness", "distraction", "close_following",
                "forward_collision_warning", "collision", "near_collision",
                "stop_sign", "unsafe_lane_change", "lane_swerving",
            }
            tier = "RED" if raw_type in red_types else "ORANGE"

            events.append({
                "driver": (ev.get("driver", {}) or {}).get("name", "Unknown"),
                "vehicle": veh.get("name", "Unknown"),
                "event_type": raw_type,
                "display_name": display,
                "tier": tier,
                "speed": round(speed, 1),
                "duration": str(ev.get("duration", 0) or 0) + "s",
                "time": ev.get("start_time", ""),
                "yard": yard,
            })
        # Pagination
        paging = data.get("pagination", {})
        total = paging.get("total", 0)
        if len(events) >= total or page > 50:
            break
        page += 1

    return events


def fetch_observations(start_date, end_date):
    """Fetch KPA observations for date range."""
    if not KPA_TOKEN:
        return {"total": 0, "by_type": {}, "details": []}
    # Use updated_after for efficiency (1 day before start)
    start_dt = datetime.strptime(start_date, "%Y-%m-%d")
    updated_after = int((start_dt - timedelta(days=1)).timestamp() * 1000)
    rows = kpa_csv_rows(OBSERVATION_FORM_ID, updated_after)

    details = []
    by_type = {}
    for row in rows:
        row_date = (row.get("date") or "")[:10]
        if row_date < start_date or row_date > end_date:
            continue
        obs_type = row.get(OBS_TYPE_HASH, "Other") or "Other"
        desc = row.get(OBS_DESC_HASH, "") or ""
        location = row.get(OBS_LOCATION_HASH, "") or ""
        yard = normalize_yard(row.get(OBS_YARD_HASH, ""))
        # Find service line
        svc = ""
        for h in SERVICE_LINE_HASHES:
            val = row.get(h, "")
            if val and val != "Service Line":
                svc = val
                break
        by_type[obs_type] = by_type.get(obs_type, 0) + 1
        details.append({
            "report_number": row.get("report number", ""),
            "date": row.get("date", ""),
            "observer": row.get("Name", "") or row.get("observer", ""),
            "type": obs_type,
            "description": desc,
            "location": location,
            "service_line": svc,
            "yard": yard,
        })

    # Near misses are observations with type "Near Miss"
    near_misses = [d for d in details if (d.get("type") or "").lower() == "near miss"]

    return {"total": len(details), "by_type": by_type, "details": details}, near_misses


def fetch_incidents(start_date, end_date):
    """Fetch KPA incidents for date range."""
    if not KPA_TOKEN:
        return []
    start_dt = datetime.strptime(start_date, "%Y-%m-%d")
    updated_after = int((start_dt - timedelta(days=1)).timestamp() * 1000)
    rows = kpa_csv_rows(INCIDENT_FORM_ID, updated_after)

    incidents = []
    for row in rows:
        row_date = (row.get("date") or "")[:10]
        if row_date < start_date or row_date > end_date:
            continue
        inc_type = row.get(INC_TYPE_HASH, "") or ""
        desc = row.get(INC_DESC_HASH, "") or ""
        employee = row.get(INC_EMPLOYEE_HASH, "") or ""
        location = row.get(INC_LOCATION_HASH, "") or ""
        svc = ""
        for h in SERVICE_LINE_HASHES:
            val = row.get(h, "")
            if val and val not in ("Service Line", "service_line", "Division", "division"):
                svc = val
                break
        incidents.append({
            "report_number": row.get("report number", ""),
            "date": row.get("date", ""),
            "type": inc_type,
            "employee": employee,
            "description": desc,
            "location": location,
            "service_line": svc,
        })
    return incidents


def fetch_assessments(start_date, end_date):
    """Fetch KPA assessments for date range across all assessment forms."""
    if not KPA_TOKEN:
        return {"total": 0, "findings": 0, "details": []}

    # Get form names
    try:
        form_resp = kpa_post("forms.list", {}, timeout=60)
        form_data = json.loads(form_resp.text)
        form_names = {}
        for f in form_data.get("forms", []):
            form_names[f.get("id", 0)] = f.get("name", "Unknown")
    except Exception:
        form_names = {}

    start_dt = datetime.strptime(start_date, "%Y-%m-%d")
    updated_after = int((start_dt - timedelta(days=1)).timestamp() * 1000)

    total = 0
    findings = 0
    details = []

    for form_id in ASSESSMENT_FORM_IDS:
        try:
            rows = kpa_csv_rows(form_id, updated_after)
        except Exception:
            continue

        form_name = form_names.get(form_id, "Unknown ({})".format(form_id))
        form_rows = []

        for row in rows:
            row_date = (row.get("date") or "")[:10]
            if row_date < start_date or row_date > end_date:
                continue
            # Assessor
            assessor = row.get("Name", "") or row.get("name", "") or row.get("observer", "") or "Unknown"
            # Location
            location = row.get("lg5pnj4chjadnv46", "") or row.get("7vj2l992y7fwqhwz", "") or "Unknown"
            # Service line
            svc = ""
            for h in SERVICE_LINE_HASHES:
                val = row.get(h, "")
                if val and val not in ("Service Line", "service_line", "Division", "division"):
                    svc = val
                    break
            report_num = row.get("report number", "")
            form_rows.append({
                "assessor": assessor,
                "location": location,
                "customer": "",
                "form_id": report_num,
                "link": "{}/{}".format(KPA_RESPONSE_URL, report_num),
                "issue": "None noted",
                "service_line": svc,
            })

        if form_rows:
            total += len(form_rows)
            details.append({
                "form_name": form_name,
                "form_id": form_id,
                "count": len(form_rows),
                "rows": form_rows,
            })

    return {"total": total, "findings": findings, "details": details}


def fetch_corrective_actions():
    """Fetch open corrective actions from KPA."""
    if not KPA_TOKEN:
        return {"total_open": 0, "corrective_actions": []}

    # Fetch users for name lookup
    try:
        user_resp = kpa_post("users.list", {}, timeout=120)
        user_data = json.loads(user_resp.text)
        users = {}
        for u in user_data.get("users", []):
            uid = str(u.get("id", ""))
            name = "{} {}".format(u.get("firstname", ""), u.get("lastname", "")).strip()
            if uid and name:
                users[uid] = name
    except Exception:
        users = {}

    # Fetch form names
    try:
        form_resp = kpa_post("forms.list", {}, timeout=60)
        form_data = json.loads(form_resp.text)
        forms = {}
        for f in form_data.get("forms", []):
            forms[str(f.get("id", ""))] = f.get("name", "Unknown")
    except Exception:
        forms = {}

    # Fetch all followups
    followups = kpa_paginated("followups.list", "followups")

    today_int = int(date.today().strftime("%Y%m%d"))
    cas = []
    for fu in followups:
        if not fu.get("open", True):
            continue
        form_id = str(fu.get("form_id", ""))
        origin_form = forms.get(form_id, "Unknown ({})".format(form_id))
        due_int = fu.get("due", 0) or 0
        overdue = due_int > 0 and due_int < today_int
        created_ms = fu.get("created_on", "0")
        try:
            created_dt = datetime.fromtimestamp(int(created_ms) / 1000, tz=timezone.utc)
            created_str = created_dt.strftime("%Y-%m-%d")
            days_open = (date.today() - created_dt.date()).days
        except Exception:
            created_str = ""
            days_open = 0
        due_str = ""
        if due_int > 0:
            try:
                due_str = "{}-{}-{}".format(str(due_int)[:4], str(due_int)[4:6], str(due_int)[6:8])
            except Exception:
                pass
        assignee_id = str(fu.get("m_assignee_id", ""))
        creator_id = str(fu.get("m_observer_id", ""))
        desc = ""
        messages = fu.get("messages", [])
        if messages:
            desc = messages[0].get("note", "") if messages else ""
        # Determine division from origin form
        fid_int = int(form_id) if form_id.isdigit() else 0
        division = FORM_DIVISION_MAP.get(fid_int, "Shared")

        cas.append({
            "id": str(fu.get("id", "")),
            "report_number": str(fu.get("response_id", "")),
            "kpa_link": "{}/{}".format(KPA_RESPONSE_URL, fu.get("response_id", "")),
            "form_id": form_id,
            "origin_form": origin_form,
            "division": division,
            "description": desc[:200],
            "assigned_to": users.get(assignee_id, "Unknown"),
            "created_by": users.get(creator_id, "Unknown"),
            "created_date": created_str,
            "due_date": due_str,
            "days_open": days_open,
            "overdue": overdue,
        })

    return {
        "total_open": len(cas),
        "total_overdue": sum(1 for c in cas if c["overdue"]),
        "corrective_actions": cas,
    }


def fetch_ytd_stats(as_of_date=None):
    """Compute YTD stats: man-hours + TRIR from payroll + KPA incidents."""
    if as_of_date is None:
        as_of_date = date.today()

    # Man-hours
    hours_by_company = {}
    for company, data in Q1_MAN_HOURS.items():
        monthly = data["monthly"]
        march_hours = monthly.get("2026-03", 0)
        company_ytd = 0.0
        for m in range(1, as_of_date.month + 1):
            key = "{}-{:02d}".format(as_of_date.year, m)
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
        q1_total = data["total"]
        scale = company_ytd / q1_total if q1_total > 0 else 1.0
        prorated_divs = {}
        for div_name, div_q1 in data["by_division"].items():
            prorated_divs[div_name] = round(div_q1 * scale)
        hours_by_company[company] = {
            "total": round(company_ytd),
            "by_division": prorated_divs,
        }

    ytd_hours = sum(c["total"] for c in hours_by_company.values())

    # Incidents from KPA (builds by_company and service_lines in one pass)
    ytd_rec = 0
    overall_last = None
    by_company = {}
    svc_lines_by_company = {}
    if KPA_TOKEN:
        try:
            rows = kpa_csv_rows(INCIDENT_FORM_ID)
            for row in rows:
                inc_type = row.get(INC_TYPE_HASH, "")
                if "Recordable" not in inc_type:
                    continue
                d_str = (row.get("date") or "")[:10]
                try:
                    d = datetime.strptime(d_str, "%Y-%m-%d").date()
                except Exception:
                    continue
                raw_co = (row.get(COMPANY_HASH) or "").strip() or "Unknown"
                co = KPA_COMPANY_MAP.get(raw_co, raw_co)
                if "butch" in raw_co.lower() and "truck" in raw_co.lower():
                    co = "BTI"
                elif "transcend" in raw_co.lower():
                    co = "Transcend"
                elif "valor" in raw_co.lower():
                    co = "Valor"
                elif "permian" in raw_co.lower():
                    co = "Permian"
                elif "butch" in raw_co.lower():
                    co = "BRHAS"
                if d.year == as_of_date.year:
                    ytd_rec += 1
                if overall_last is None or d > overall_last:
                    overall_last = d
                if co not in by_company:
                    by_company[co] = {"last_date": d.isoformat(), "ytd_count": 0, "days_since": 0}
                if d > datetime.strptime(by_company[co]["last_date"], "%Y-%m-%d").date():
                    by_company[co]["last_date"] = d.isoformat()
                if d.year == as_of_date.year:
                    by_company[co]["ytd_count"] += 1
                    # Build service_lines
                    svc = ""
                    for h in SERVICE_LINE_HASHES:
                        val = row.get(h, "")
                        if val and val not in ("Service Line", "service_line", "Division", "division"):
                            svc = KPA_SVC_TO_DIVISION.get(val, val)
                            break
                    if svc:
                        if co not in svc_lines_by_company:
                            svc_lines_by_company[co] = {}
                        if svc not in svc_lines_by_company[co]:
                            svc_lines_by_company[co][svc] = {"ytd_count": 0, "last_date": d.isoformat(), "days_since": 0}
                        svc_lines_by_company[co][svc]["ytd_count"] += 1
                        if d > datetime.strptime(svc_lines_by_company[co][svc]["last_date"], "%Y-%m-%d").date():
                            svc_lines_by_company[co][svc]["last_date"] = d.isoformat()
        except Exception:
            pass

    for co_data in by_company.values():
        last = datetime.strptime(co_data["last_date"], "%Y-%m-%d").date()
        co_data["days_since"] = (as_of_date - last).days
    for co_svcs in svc_lines_by_company.values():
        for sl_data in co_svcs.values():
            last = datetime.strptime(sl_data["last_date"], "%Y-%m-%d").date()
            sl_data["days_since"] = (as_of_date - last).days

    ytd_trir = round(ytd_rec * 200000 / ytd_hours, 2) if ytd_hours > 0 else 0
    days_rec = (as_of_date - overall_last).days if overall_last else None

    stats_by_company = {}
    for co, hrs in hours_by_company.items():
        co_inc = by_company.get(co, {})
        co_rec = co_inc.get("ytd_count", 0)
        co_hrs = hrs["total"]
        stats_by_company[co] = {
            "man_hours": co_hrs,
            "recordables": co_rec,
            "trir": round(co_rec * 200000 / co_hrs, 2) if co_hrs > 0 else 0,
            "dart": round(co_rec * 200000 / co_hrs, 2) if co_hrs > 0 else 0,
            "days_since_recordable": co_inc.get("days_since"),
            "last_recordable_date": co_inc.get("last_date"),
            "by_division": hrs["by_division"],
            "service_lines": svc_lines_by_company.get(co, {}),
        }

    return {
        "report_date": as_of_date.isoformat(),
        "generated_at": datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S"),
        "ytd_trir": ytd_trir,
        "ytd_dart": ytd_trir,
        "ytd_recordables": ytd_rec,
        "ytd_man_hours": ytd_hours,
        "days_since_lti": days_rec,
        "days_since_recordable": days_rec,
        "last_recordable_date": overall_last.isoformat() if overall_last else None,
        "by_company": by_company,
        "stats_by_company": stats_by_company,
    }


# ==============================================================================
# ROUTES
# ==============================================================================

@app.route("/api/health")
def health():
    return jsonify({
        "status": "ok",
        "has_kpa": bool(KPA_TOKEN),
        "has_motive": bool(MOTIVE_KEY),
    })


@app.route("/api/data")
def get_data():
    """Fetch all dashboard data for a date range.

    Params:
        start: YYYY-MM-DD (required)
        end: YYYY-MM-DD (required)
    """
    start = request.args.get("start")
    end = request.args.get("end")
    if not start or not end:
        return jsonify({"error": "start and end params required"}), 400

    # Validate dates
    try:
        datetime.strptime(start, "%Y-%m-%d")
        datetime.strptime(end, "%Y-%m-%d")
    except ValueError:
        return jsonify({"error": "Invalid date format, use YYYY-MM-DD"}), 400

    # Check cache
    cache_key = "data_{}_{}".format(start, end)
    cached = cache_get(cache_key)
    if cached:
        return jsonify(cached)

    # Determine if this is today's data (shorter cache)
    is_today = start == date.today().isoformat() and end == date.today().isoformat()
    ttl = 300 if is_today else CACHE_TTL_SECONDS

    # Convert Central dates to UTC for Motive
    start_central = datetime.strptime(start, "%Y-%m-%d").replace(tzinfo=CENTRAL_TZ)
    end_central = datetime.strptime(end, "%Y-%m-%d").replace(hour=23, minute=59, second=59, tzinfo=CENTRAL_TZ)
    start_utc = start_central.astimezone(timezone.utc).strftime("%Y-%m-%d")
    end_utc = end_central.astimezone(timezone.utc).strftime("%Y-%m-%d")

    # Fetch all data
    result = {}
    errors = []

    try:
        spd_events = fetch_speeding(start_utc, end_utc)
        spd_summary = {"red": 0, "orange": 0, "yellow": 0}
        spd_by_div = {}
        for ev in spd_events:
            tier = ev.get("tier", "YELLOW").lower()
            spd_summary[tier] = spd_summary.get(tier, 0) + 1
            div = ev.get("division", "Unknown")
            spd_by_div[div] = spd_by_div.get(div, 0) + 1
        result["speeding"] = {
            "report_date": end,
            "events": spd_events,
            "total_events": len(spd_events),
            "summary": spd_summary,
            "by_division": spd_by_div,
            "repeat_offenders": {},
        }
    except Exception as e:
        errors.append("speeding: {}".format(str(e)))
        result["speeding"] = None

    # Camera events require vehicle cross-reference (V1 + V2) which is slow.
    # For historical queries, camera data is not available via API.
    # Dashboard will use today's static JSON for camera data.
    result["camera"] = None

    try:
        obs_data, near_misses = fetch_observations(start, end)
        inc_data = []
        try:
            inc_data = fetch_incidents(start, end)
        except Exception as e:
            errors.append("incidents: {}".format(str(e)))
        assess_data = {"total": 0, "findings": 0, "details": []}
        try:
            assess_data = fetch_assessments(start, end)
        except Exception as e:
            errors.append("assessments: {}".format(str(e)))
        result["kpa"] = {
            "report_date": end,
            "observations": obs_data,
            "near_misses": near_misses,
            "incidents": inc_data,
            "assessments": assess_data,
        }
    except Exception as e:
        errors.append("kpa: {}".format(str(e)))
        result["kpa"] = None

    try:
        end_dt = datetime.strptime(end, "%Y-%m-%d").date()
        result["ytd"] = fetch_ytd_stats(end_dt)
    except Exception as e:
        errors.append("ytd: {}".format(str(e)))
        result["ytd"] = None

    try:
        result["cas"] = fetch_corrective_actions()
    except Exception as e:
        errors.append("cas: {}".format(str(e)))
        result["cas"] = None

    # Training, devices, and camera are point-in-time or require complex cross-refs.
    # Dashboard uses today's static JSON for these.
    result["training"] = None
    result["devices"] = None
    result["_note"] = "Training, device, and camera data use today's static snapshots. Speeding, KPA, YTD stats, and CAs are fetched live for the requested date range."

    if errors:
        result["_errors"] = errors

    cache_set(cache_key, result, ttl)
    return jsonify(result)


# ==============================================================================
# MAIN
# ==============================================================================

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    print("Starting Safety Dashboard API on port {}".format(port))
    print("  KPA token: {}".format("SET" if KPA_TOKEN else "MISSING"))
    print("  Motive key: {}".format("SET" if MOTIVE_KEY else "MISSING"))
    app.run(host="0.0.0.0", port=port, debug=True)
