"""
DAILY CORRECTIVE ACTIONS TRACKER
=================================
Pulls all open corrective actions (follow-ups) from KPA and writes
output/corrective_actions.json.

Auth, API call pattern, and error handling match daily_kpa_report.py.

Usage:
    KPA_API_TOKEN=... python daily_corrective_actions.py
"""

import json
import os
import sys
import time
from datetime import datetime, timezone

import requests

# ==============================================================================
# SETUP - API keys from environment variables
# ==============================================================================

API_TOKEN = os.environ.get("KPA_API_TOKEN")
if not API_TOKEN:
    print("ERROR: KPA_API_TOKEN environment variable is not set.")
    sys.exit(1)

API_BASE = "https://api.kpaehs.com/v1"
KPA_RESPONSE_URL = "https://brhas-ees.kpaehs.com/forms/responses/view"

# Form ID -> human-readable name (for origin_form enrichment)
FORM_NAMES = {
    151085: "Observation Card",
    151622: "Incident Report",
    180243: "Root Cause Analysis",
    180035: "5 Why",
    170742: "JSA Log",
    367333: "JSA Review Log",
    152018: "Vehicle Inspection Checklist",
    151213: "Non-DOT Pre-Trip Inspection",
    152034: "HSE - Workplace Inspection Checklist",
    180746: "Safety Meeting Sign In Log",
    161692: "Tailgate Talk",
    267566: "Hot Work Permit",
    180935: "Fall Protection Inspection",
    381707: "CSG - Safety Casing Field Assessment",
    187229: "CSG - JSA Review",
    187860: "CSG - Workplace Inspection",
    229645: "Non-DOT Pre/Post Trip",
    385365: "TD - Rig Inspection",
    484193: "TD - Observation Card",
    226217: "WS - Poly Pipe Field Assessment",
    386087: "WS - Pit Lining Field Assessment",
    172295: "Construction - Site Safety Review",
    153181: "RH - Rathole Field Assessment",
    199992: "Equipment Inspection",
}

# Form ID -> division mapping (for dashboard filtering)
FORM_DIVISION_MAP = {
    385365: "Transcend",       # TD - Rig Inspection
    484193: "Transcend",       # TD - Observation Card
    226217: "Poly Pipe",       # WS - Poly Pipe Field Assessment
    386087: "Pit Lining",      # WS - Pit Lining Field Assessment
    172295: "Construction",    # Construction - Site Safety Review
    153181: "Rathole",         # RH - Rathole Field Assessment
    381707: "Casing",          # CSG - Safety Casing Field Assessment
    229645: "Casing",          # CSG - Non-DOT Pre/Post Trip
}
# Shared forms (Observation Card, Vehicle Inspection, etc.) -> "Shared"


# ==============================================================================
# KPA API CALLS
# ==============================================================================

def call_kpa(endpoint, params):
    """Make request to KPA API (POST with JSON body)."""
    url = f"{API_BASE}/{endpoint}"
    payload = {"token": API_TOKEN}
    payload.update(params)

    try:
        response = requests.post(url, json=payload, timeout=30)
        return response.text
    except Exception as e:
        print(f"ERROR: {e}")
        return None


def call_kpa_json_paginated(endpoint, params, data_key, max_pages=50):
    """Paginated KPA call for JSON endpoints.

    JSON endpoints support max limit of 500.
    Matches data_fetcher._call_json_paginated pattern.
    """
    all_rows = []
    page = 1

    while True:
        p = dict(params)
        p["page"] = page
        url = f"{API_BASE}/{endpoint}"
        payload = {"token": API_TOKEN, "limit": 500}
        payload.update(p)

        for attempt in range(3):
            try:
                resp = requests.post(url, json=payload, timeout=60)
                break
            except Exception as e:
                if attempt < 2:
                    print(f"  Retry {attempt + 1} for {endpoint} page {page}: {e}")
                    time.sleep(2)
                else:
                    print(f"  ERROR: {endpoint} page {page} failed after 3 attempts: {e}")
                    return all_rows

        try:
            data = json.loads(resp.text)
        except json.JSONDecodeError:
            print(f"  ERROR: Invalid JSON from {endpoint} page {page}")
            break

        items = data.get(data_key, [])
        if not items:
            break
        all_rows.extend(items)

        paging = data.get("paging", {})
        last_page = paging.get("last_page", 1)
        if page >= last_page:
            break
        if page >= max_pages:
            print(f"  Capped at {max_pages} pages ({len(all_rows)} rows)")
            break
        page += 1
        time.sleep(1.5)  # KPA rate limit delay

    return all_rows


def fetch_users():
    """Fetch all KPA users and build {user_id: 'First Last'} lookup.

    users.list is NOT paginated and does NOT accept limit -- just token.
    Needs 120s timeout (large payload).
    Field names are lowercase: firstname, lastname (not camelCase).
    """
    url = f"{API_BASE}/users.list"
    payload = {"token": API_TOKEN}
    try:
        resp = requests.post(url, json=payload, timeout=120)
        data = json.loads(resp.text)
        users = data.get("users", [])
    except Exception as e:
        print(f"  Warning: Could not fetch users: {e}")
        return {}

    lookup = {}
    for u in users:
        uid = u.get("id", "")
        first = u.get("firstname", "")
        last = u.get("lastname", "")
        if uid and (first or last):
            lookup[uid] = f"{first} {last}".strip()
    print(f"  Users loaded: {len(lookup)}")
    return lookup


def fetch_all_followups():
    """Fetch all follow-ups from KPA followups.list endpoint."""
    rows = call_kpa_json_paginated("followups.list", {}, "followups")
    print(f"  Total follow-ups fetched: {len(rows)}")
    return rows


# ==============================================================================
# PROCESSING
# ==============================================================================

def process_followups(raw_followups, user_lookup):
    """Filter to open follow-ups and enrich with calculated fields."""
    today = datetime.now(timezone.utc).date()
    open_items = []

    for row in raw_followups:
        # open is a boolean (True/False) from KPA
        if not row.get("open"):
            continue

        form_id = row.get("form_id")

        # Due date: KPA returns integer like 20260410
        due_int = row.get("due")
        if due_int and isinstance(due_int, int):
            due_date = f"{str(due_int)[:4]}-{str(due_int)[4:6]}-{str(due_int)[6:8]}"
        else:
            due_date = ""

        # Created date: KPA returns STRING of epoch milliseconds (e.g. '1776090373470')
        created_on = row.get("created_on", "")
        created_date = ""
        days_open = 0
        if created_on:
            try:
                epoch_ms = int(str(created_on))
                created_dt = datetime.fromtimestamp(epoch_ms / 1000, tz=timezone.utc)
                created_date = created_dt.strftime("%Y-%m-%d")
                days_open = (today - created_dt.date()).days
            except (ValueError, OSError):
                pass

        # Overdue check
        overdue = False
        if due_date:
            try:
                due_dt = datetime.strptime(due_date, "%Y-%m-%d").date()
                overdue = today > due_dt
            except ValueError:
                pass

        # Extract finding/description from messages (note can be None)
        msgs = row.get("messages", [])
        description = ""
        if msgs and isinstance(msgs, list):
            note = msgs[0].get("note")
            description = note if note else ""

        # Form name lookup (form_id already extracted above for division filter)
        origin_form = FORM_NAMES.get(form_id, f"Unknown ({form_id})")

        # response_id matches 'report number' from responses.flat
        response_id = row.get("response_id", "")
        kpa_link = f"{KPA_RESPONSE_URL}/{response_id}" if response_id else ""

        # Assigned to (resolve user ID to name)
        assignee_id = row.get("m_assignee_id", "")
        assigned_to = user_lookup.get(assignee_id, "Unassigned")

        # Observer/creator
        observer_id = row.get("m_observer_id", "")
        created_by = user_lookup.get(observer_id, "Unknown")

        # Division from form ID mapping; shared forms tagged as "Shared"
        division = FORM_DIVISION_MAP.get(form_id, "Shared")

        open_items.append({
            "id": row.get("id", ""),
            "report_number": str(response_id),
            "kpa_link": kpa_link,
            "form_id": form_id,
            "origin_form": origin_form,
            "division": division,
            "description": description,
            "assigned_to": assigned_to,
            "created_by": created_by,
            "created_date": created_date,
            "due_date": due_date,
            "days_open": days_open,
            "overdue": overdue,
        })

    # Sort by days_open descending (oldest first)
    open_items.sort(key=lambda x: x["days_open"], reverse=True)
    return open_items


def build_summary(open_items):
    """Build summary counts for the JSON output."""
    total = len(open_items)
    overdue = sum(1 for item in open_items if item["overdue"])

    # By assignee
    by_assignee = {}
    for item in open_items:
        name = item["assigned_to"]
        if name not in by_assignee:
            by_assignee[name] = {"total": 0, "overdue": 0}
        by_assignee[name]["total"] += 1
        if item["overdue"]:
            by_assignee[name]["overdue"] += 1

    # By origin form
    by_form = {}
    for item in open_items:
        form = item["origin_form"]
        by_form[form] = by_form.get(form, 0) + 1

    # By age bucket
    by_age = {
        "0-7 days": 0,
        "8-30 days": 0,
        "31-60 days": 0,
        "61-90 days": 0,
        "90+ days": 0,
    }
    for item in open_items:
        d = item["days_open"]
        if d <= 7:
            by_age["0-7 days"] += 1
        elif d <= 30:
            by_age["8-30 days"] += 1
        elif d <= 60:
            by_age["31-60 days"] += 1
        elif d <= 90:
            by_age["61-90 days"] += 1
        else:
            by_age["90+ days"] += 1

    return {
        "total_open": total,
        "total_overdue": overdue,
        "by_assignee": by_assignee,
        "by_origin_form": by_form,
        "by_age": by_age,
    }


# ==============================================================================
# MAIN
# ==============================================================================

def main():
    print("=" * 60)
    print("  DAILY CORRECTIVE ACTIONS TRACKER")
    print("=" * 60)
    print()

    print("[1/3] Fetching KPA users...")
    user_lookup = fetch_users()

    print("[2/3] Fetching follow-ups from KPA...")
    raw_followups = fetch_all_followups()

    print("[3/3] Processing open corrective actions...")
    open_items = process_followups(raw_followups, user_lookup)
    summary = build_summary(open_items)

    # Build JSON output (matching existing report pattern)
    now = datetime.now(timezone.utc)
    json_data = {
        "report_date": now.strftime("%Y-%m-%d"),
        "generated_at": now.strftime("%Y-%m-%d %H:%M:%S"),
        "summary": summary,
        "corrective_actions": open_items,
    }

    # Write to output/corrective_actions.json
    os.makedirs("output", exist_ok=True)
    json_path = os.path.join("output", "corrective_actions.json")
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(json_data, f, indent=2, default=str)

    # Print confirmation
    print()
    print(f"  Open corrective actions: {summary['total_open']}")
    print(f"  Overdue: {summary['total_overdue']}")
    print(f"  Age breakdown:")
    for bucket, count in summary["by_age"].items():
        print(f"    {bucket}: {count}")
    print(f"  Top assignees:")
    ranked = sorted(
        summary["by_assignee"].items(),
        key=lambda x: x[1]["total"],
        reverse=True,
    )
    for name, counts in ranked[:10]:
        overdue_tag = f" ({counts['overdue']} overdue)" if counts["overdue"] else ""
        print(f"    {name}: {counts['total']}{overdue_tag}")
    print()
    print(f"  JSON written: {json_path}")
    print("  Done.")


if __name__ == "__main__":
    main()
