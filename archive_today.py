"""
ARCHIVE_TODAY.PY -- Combine daily output JSONs into single archive file
======================================================================
Reads from output/ directory (already populated by daily scripts) and
writes a combined archive/YYYY-MM-DD.json file. The archive date comes
from the report_date field in the output JSON, NOT from today's date
(daily scripts fetch yesterday's data).

Also fetches IFTA mileage from Motive API for the archive date (not
available from any existing daily script).

Usage:
    python archive_today.py
    python archive_today.py --output-dir archive/
"""
import json
import os
import sys
import argparse
from collections import defaultdict
from datetime import datetime, timezone, timedelta

import requests

from api_config import MOTIVE_BASE_V1, GROUP_ID_MAP

MOTIVE_KEY = os.environ.get("MOTIVE_API_KEY", "")

# Vehicle prefix -> division (same as backfill_archive.py)
DIV_PREFIXES = {
    "CSG": "Casing", "RAT": "Rathole", "ANC": "Anchors",
    "PP": "Poly Pipe", "PL": "Pit Lining", "CON": "Construction",
    "ENV": "Environmental", "FEN": "Fencing", "DT": "Drilling Tools",
    "VAL": "Valor", "BTI": "BTI", "TD": "Transcend",
    "WTC": "Water/Construction", "FAB": "Fabrication",
    "PER": "Permian", "SS": "Permian",
}
YARD_PREFIXES = {
    "MID": "Midland", "BRY": "Bryan", "KIL": "Kilgore",
    "HOB": "Hobbs", "JOU": "Jourdanton", "LAR": "Laredo",
    "LL": "Levelland", "BAR": "Barstow",
    "TOW": "Pennsylvania", "PA": "Pennsylvania",
    "OH": "Ohio", "OK": "Oklahoma", "ND": "North Dakota",
}


def parse_vehicle_division(vehicle):
    """Parse division from vehicle number."""
    parts = vehicle.replace("-", " ").replace("_", " ").upper().split()
    div = "Unknown"
    for p in parts:
        for prefix, d in DIV_PREFIXES.items():
            if p.startswith(prefix):
                return d
    return div


def fetch_daily_mileage(target_date):
    """Fetch IFTA trip mileage from Motive for a single day."""
    if not MOTIVE_KEY:
        print("  Mileage: MOTIVE_API_KEY not set, skipping")
        return None

    headers = {"X-Api-Key": MOTIVE_KEY}
    by_division = defaultdict(float)
    total_miles = 0
    vehicle_set = set()
    page = 1

    while page <= 50:
        try:
            resp = requests.get(
                f"{MOTIVE_BASE_V1}/ifta/trips",
                headers=headers,
                params={"per_page": 100, "page_no": page,
                        "start_date": target_date, "end_date": target_date},
                timeout=60,
            )
            if not resp.ok:
                break
            data = resp.json()
        except Exception as e:
            print(f"  Mileage: API error on page {page}: {e}")
            break

        items = data.get("ifta_trips", [])
        if not items:
            break

        for item in items:
            trip = item.get("ifta_trip", item)
            veh = trip.get("vehicle") or {}
            vehicle = veh.get("number", "") if isinstance(veh, dict) else ""
            distance = float(trip.get("distance", 0) or 0)
            if distance > 0:
                total_miles += distance
                vehicle_set.add(vehicle)
                div = parse_vehicle_division(vehicle)
                by_division[div] += distance

        pag = data.get("pagination", {})
        if pag.get("total") and page * 100 >= pag["total"]:
            break
        page += 1

    if total_miles > 0:
        print(f"  Mileage: {total_miles:,.1f} miles, {len(vehicle_set)} vehicles")
        return {
            "total_miles": round(total_miles, 1),
            "by_division": {k: round(v, 1) for k, v in
                            sorted(by_division.items(), key=lambda x: -x[1])},
            "vehicle_count": len(vehicle_set),
        }
    return None


def main():
    parser = argparse.ArgumentParser(description="Archive daily output JSON files")
    parser.add_argument("--output-dir", default="archive", help="Archive directory")
    args = parser.parse_args()

    output_dir = "output"

    # Map output files to archive keys
    file_map = {
        "speeding_events.json": "speeding",
        "camera_events.json": "camera",
        "kpa_data.json": "kpa",
        "ytd_stats.json": "ytd",
        "corrective_actions.json": "cas",
        "training_compliance.json": "training",
        "device_status.json": "devices",
    }

    # Load all output files
    loaded = {}
    for filename, key in file_map.items():
        path = os.path.join(output_dir, filename)
        if os.path.exists(path):
            with open(path, encoding="utf-8") as f:
                loaded[key] = json.load(f)
        else:
            loaded[key] = None

    # Determine archive date from report_date fields (data is for yesterday)
    archive_date = None
    for key in ("speeding", "camera", "kpa"):
        data = loaded.get(key)
        if data and "report_date" in data:
            archive_date = data["report_date"]
            break

    if not archive_date:
        # Fallback: use ytd report_date or yesterday
        ytd = loaded.get("ytd")
        if ytd and "report_date" in ytd:
            archive_date = ytd["report_date"]
        else:
            archive_date = (datetime.now(timezone.utc) - timedelta(days=1)).strftime("%Y-%m-%d")
            print(f"  WARNING: No report_date found, using fallback: {archive_date}")

    # Staleness guard: check generated_at is recent (within 4 hours)
    for key in ("speeding", "kpa"):
        data = loaded.get(key)
        if data and "generated_at" in data:
            try:
                gen_str = data["generated_at"]
                # Handle both "2026-04-18 17:29:06" and ISO format
                for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%dT%H:%M:%S", "%Y-%m-%dT%H:%M:%SZ"):
                    try:
                        gen_dt = datetime.strptime(gen_str, fmt)
                        break
                    except ValueError:
                        continue
                else:
                    continue
                age_hours = (datetime.now(timezone.utc).replace(tzinfo=None) - gen_dt).total_seconds() / 3600
                if age_hours > 4:
                    print(f"  WARNING: {key} data is {age_hours:.1f} hours old -- skipping archive")
                    sys.exit(0)
            except Exception:
                pass

    # Filter camera events: drop uncoachable
    camera = loaded.get("camera")
    if camera and isinstance(camera.get("events"), list):
        original_count = len(camera["events"])
        camera["events"] = [
            e for e in camera["events"]
            if e.get("coaching_status", "") != "uncoachable"
        ]
        filtered_count = len(camera["events"])
        camera["total_events"] = filtered_count
        if original_count != filtered_count:
            print(f"  Camera: filtered {original_count - filtered_count} uncoachable events ({original_count} -> {filtered_count})")

    # Fetch mileage for archive date (not in any existing daily script)
    mileage_data = fetch_daily_mileage(archive_date)

    # Build archive object
    archive = {
        "date": archive_date,
        "generated_at": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "speeding": loaded.get("speeding"),
        "camera": loaded.get("camera"),
        "kpa": loaded.get("kpa"),
        "ytd": loaded.get("ytd"),
        "mileage": mileage_data,
        "cas": None,       # Point-in-time, not archived
        "training": None,  # Point-in-time, not archived
        "devices": None,   # Point-in-time, not archived
    }

    # Write archive file
    os.makedirs(args.output_dir, exist_ok=True)
    output_path = os.path.join(args.output_dir, f"{archive_date}.json")
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(archive, f, separators=(",", ":"))

    size_kb = os.path.getsize(output_path) / 1024
    print(f"  Archived {archive_date}: {output_path} ({size_kb:.1f} KB)")


if __name__ == "__main__":
    main()
