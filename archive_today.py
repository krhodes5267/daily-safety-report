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

import time as _time

import requests

from api_config import MOTIVE_BASE_V1, GROUP_ID_MAP, CASING_GROUP_IDS

MOTIVE_KEY = os.environ.get("MOTIVE_API_KEY", "")


def _api_get(url, headers, params, timeout=60, retries=1):
    """GET with simple retry on failure. Returns response or None."""
    for attempt in range(retries + 1):
        try:
            resp = requests.get(url, headers=headers, params=params, timeout=timeout)
            if resp.ok:
                return resp
            if attempt < retries:
                print(f"  Retry: {url.split('/')[-1]} returned {resp.status_code}, retrying in 5s...")
                _time.sleep(5)
        except Exception as e:
            if attempt < retries:
                print(f"  Retry: {url.split('/')[-1]} failed ({e}), retrying in 5s...")
                _time.sleep(5)
            else:
                print(f"  API error: {url.split('/')[-1]} failed after {retries + 1} attempts: {e}")
    return None

# Vehicle prefix -> division (same as backfill_archive.py)
DIV_PREFIXES = {
    "CSG": "Casing", "CAS": "Casing", "RAT": "Rathole", "ANC": "Anchors",
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
    """Parse division from vehicle number using prefix/suffix conventions."""
    parts = vehicle.replace("-", " ").replace("_", " ").upper().split()
    for p in parts:
        for prefix, d in DIV_PREFIXES.items():
            if p.startswith(prefix):
                return d
    # Suffix-based fallback: vehicles like "1733C" (C=Casing), "18125E" (E=Environmental)
    clean = vehicle.strip().rstrip(" ").upper()
    if clean and clean[-1] == "C" and clean[:-1].isdigit():
        return "Casing"
    if clean and clean[-1] == "E" and clean[:-1].isdigit():
        return "Environmental"
    return "Unknown"


def _build_vehicle_lookup():
    """Build vehicle_number -> yard, division, and driver lookups from Motive /v1/vehicles API.
    Uses GROUP_ID_MAP to map ALL divisions (not just Casing) to yards/divisions."""
    headers = {"X-Api-Key": MOTIVE_KEY}
    vehicle_yards = {}
    vehicle_divisions = {}
    vehicle_drivers = {}
    page = 1
    while True:
        resp = _api_get(
            f"{MOTIVE_BASE_V1}/vehicles",
            headers=headers, params={"per_page": 100, "page_no": page}, timeout=30,
        )
        if not resp:
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
            short_num = num.split(" ")[0].strip()
            # Also clean double-space descriptions: "BTI-6172" from "BTI-6172  - DRIVER NAME"
            clean_num = num.split("  ")[0].strip() if "  " in num else short_num
            matched = False
            for gid in v.get("group_ids", []):
                if gid in GROUP_ID_MAP:
                    div, yard = GROUP_ID_MAP[gid]
                    vehicle_yards[num] = yard
                    vehicle_yards[short_num] = yard
                    vehicle_yards[clean_num] = yard
                    vehicle_divisions[num] = div
                    vehicle_divisions[short_num] = div
                    vehicle_divisions[clean_num] = div
                    matched = True
                    break
            # Fallback: no group match -- use prefix/suffix parsing
            if not matched:
                div = parse_vehicle_division(clean_num)
                if div != "Unknown":
                    vehicle_divisions[num] = div
                    vehicle_divisions[short_num] = div
                    vehicle_divisions[clean_num] = div
                    yard = parse_vehicle_yard(clean_num)
                    if yard != "Unknown":
                        vehicle_yards[num] = yard
                        vehicle_yards[short_num] = yard
                        vehicle_yards[clean_num] = yard
            for field in ("current_driver", "permanent_driver"):
                d = v.get(field)
                if d and isinstance(d, dict):
                    name = f"{d.get('first_name', '')} {d.get('last_name', '')}".strip()
                    if name:
                        vehicle_drivers[num] = name
                        vehicle_drivers[short_num] = name
                        vehicle_drivers[clean_num] = name
                        break
        pag = data.get("pagination", {})
        if page * 100 >= pag.get("total", 0):
            break
        page += 1
    return vehicle_yards, vehicle_divisions, vehicle_drivers


def fetch_daily_mileage(target_date):
    """Fetch IFTA trip mileage from Motive for a single day."""
    if not MOTIVE_KEY:
        print("  Mileage: MOTIVE_API_KEY not set, skipping")
        return None

    headers = {"X-Api-Key": MOTIVE_KEY}
    by_division = defaultdict(float)
    by_vehicle = defaultdict(float)
    total_miles = 0
    vehicle_set = set()
    page = 1

    while page <= 50:
        resp = _api_get(
            f"{MOTIVE_BASE_V1}/ifta/trips",
            headers=headers,
            params={"per_page": 100, "page_no": page,
                    "start_date": target_date, "end_date": target_date},
            timeout=60,
        )
        if not resp:
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
            if distance > 0:
                total_miles += distance
                vehicle_set.add(vehicle)
                div = parse_vehicle_division(vehicle)
                by_division[div] += distance
                by_vehicle[vehicle] += distance

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
            "by_vehicle": {k: round(v, 1) for k, v in
                           sorted(by_vehicle.items(), key=lambda x: -x[1])},
            "vehicle_count": len(vehicle_set),
        }
    return None


def fetch_driver_scorecards(start_date, end_date):
    """Fetch per-driver safety scorecards from Motive scorecard_summary API.

    Returns list of dicts with driver name, miles, score, events.
    Uses start_date/end_date to scope the data to the archive period.
    """
    if not MOTIVE_KEY:
        print("  Scorecards: MOTIVE_API_KEY not set, skipping")
        return None

    headers = {"X-Api-Key": MOTIVE_KEY}
    drivers = []
    page = 1

    while page <= 20:
        resp = _api_get(
            f"{MOTIVE_BASE_V1}/scorecard_summary",
            headers=headers,
            params={
                "per_page": 100,
                "page_no": page,
                "start_date": start_date,
                "end_date": end_date,
            },
            timeout=60,
        )
        if not resp:
            break
        data = resp.json()

        rollups = data.get("driver_performance_rollups", [])
        if not rollups:
            break

        for r in rollups:
            dr = r.get("driver_performance_rollup", {})
            driver = dr.get("driver")
            if not driver or not isinstance(driver, dict):
                continue
            status = driver.get("status", "")
            if status != "active":
                continue
            km = dr.get("total_kilometers", 0) or 0
            if km <= 0:
                continue
            name = (
                (driver.get("first_name", "") + " " + driver.get("last_name", ""))
                .strip()
            )
            drivers.append({
                "name": name,
                "driver_id": driver.get("id"),
                "miles": round(km * 0.621371, 1),
                "score": dr.get("score", 0),
                "hard_brakes": dr.get("num_hard_brakes", 0),
                "hard_accels": dr.get("num_hard_accels", 0),
                "hard_corners": dr.get("num_hard_corners", 0),
                "coached": dr.get("num_coached_events", 0),
            })

        pag = data.get("pagination", {})
        total = pag.get("total", 0)
        if total and page * 100 >= total:
            break
        page += 1

    if drivers:
        drivers.sort(key=lambda x: -x["miles"])
        print(f"  Scorecards: {len(drivers)} active drivers with mileage")
        return drivers
    return None


def parse_vehicle_yard(vehicle):
    """Parse yard from vehicle number using prefix conventions."""
    parts = vehicle.replace("-", " ").replace("_", " ").upper().split()
    for p in parts:
        for prefix, yard in YARD_PREFIXES.items():
            if p.startswith(prefix):
                return yard
    return "Unknown"


def fetch_vehicle_odometers():
    """Fetch current odometer readings for ALL vehicles via /v1/vehicle_locations.

    Returns dict: {vehicle_number: odometer_miles} or None on failure.
    This endpoint has NO lag -- returns real-time readings.
    """
    if not MOTIVE_KEY:
        print("  Odometer: MOTIVE_API_KEY not set, skipping")
        return None

    headers = {"X-Api-Key": MOTIVE_KEY}
    odometers = {}
    page = 1

    while page <= 100:
        resp = _api_get(
            f"{MOTIVE_BASE_V1}/vehicle_locations",
            headers=headers,
            params={"per_page": 100, "page_no": page},
            timeout=60,
        )
        if not resp:
            break
        data = resp.json()

        items = data.get("vehicles", [])
        if not items:
            break

        for item in items:
            veh_data = item.get("vehicle", item)
            veh_num = veh_data.get("number", "")
            loc = veh_data.get("current_location") or {}
            odo = loc.get("odometer") if loc else None
            if veh_num and odo is not None:
                try:
                    odo_val = float(odo)
                    if odo_val > 0:
                        # Clean vehicle number (take first token before long descriptions)
                        clean_num = veh_num.split("  ")[0].strip() if "  " in veh_num else veh_num.strip()
                        odometers[clean_num] = round(odo_val, 1)
                except (ValueError, TypeError):
                    pass

        pag = data.get("pagination", {})
        total = pag.get("total", 0)
        if total and page * 100 >= total:
            break
        page += 1

    if odometers:
        print(f"  Odometer: {len(odometers)} vehicles with readings")
    else:
        print("  Odometer: no readings returned")
    return odometers if odometers else None


def compute_odometer_mileage(current_odometers, archive_dir, archive_date,
                             vehicle_divisions=None, vehicle_yards_map=None):
    """Compute daily mileage from odometer deltas vs previous day's readings.

    Args:
        current_odometers: {vehicle_number: odometer_miles} from today
        archive_dir: path to archive directory
        archive_date: date string YYYY-MM-DD for the archive being built
        vehicle_divisions: {vehicle_number: division} from Motive group IDs (optional)
        vehicle_yards_map: {vehicle_number: yard} from Motive group IDs (optional)

    Returns:
        mileage dict with total_miles, by_division, by_yard, vehicle_count, odometers
        or None if no previous baseline available.
    """
    if not current_odometers:
        return None

    # Try to find previous day's odometer readings (check up to 3 days back)
    prev_odometers = None
    from datetime import datetime as dt2, timedelta as td2
    target = dt2.strptime(archive_date, "%Y-%m-%d").date()
    for days_back in range(1, 4):
        prev_date = (target - td2(days=days_back)).isoformat()
        prev_path = os.path.join(archive_dir, f"{prev_date}.json")
        if os.path.exists(prev_path):
            try:
                with open(prev_path, encoding="utf-8") as f:
                    prev_archive = json.load(f)
                prev_odo = (prev_archive.get("mileage") or {}).get("odometers")
                if prev_odo and len(prev_odo) > 0:
                    prev_odometers = prev_odo
                    print(f"  Odometer: using baseline from {prev_date} ({len(prev_odo)} vehicles)")
                    break
            except Exception:
                pass

    if not prev_odometers:
        print("  Odometer: no previous baseline found -- storing baseline only")
        return {
            "total_miles": 0,
            "by_division": {},
            "by_yard": {},
            "vehicle_count": len(current_odometers),
            "odometers": current_odometers,
            "source": "odometer_baseline",
        }

    # Compute deltas
    total_miles = 0
    by_division = defaultdict(float)
    by_yard = defaultdict(float)
    vehicles_with_miles = set()

    for veh, odo in current_odometers.items():
        prev_odo = prev_odometers.get(veh)
        if prev_odo is None:
            continue  # New vehicle, no baseline
        delta = odo - prev_odo
        # Sanity check: skip negative deltas (odometer reset) or impossibly high (>2000 mi/day)
        if delta <= 0 or delta > 2000:
            continue
        total_miles += delta
        # Use group-based lookup first (accurate), fall back to prefix parsing
        div = (vehicle_divisions or {}).get(veh) or parse_vehicle_division(veh)
        yard = (vehicle_yards_map or {}).get(veh) or parse_vehicle_yard(veh)
        by_division[div] += delta
        by_yard[yard] += delta
        vehicles_with_miles.add(veh)

    print(f"  Odometer: {total_miles:,.1f} daily miles from {len(vehicles_with_miles)} vehicles")

    return {
        "total_miles": round(total_miles, 1),
        "by_division": {k: round(v, 1) for k, v in sorted(by_division.items(), key=lambda x: -x[1])},
        "by_yard": {k: round(v, 1) for k, v in sorted(by_yard.items(), key=lambda x: -x[1])},
        "vehicle_count": len(vehicles_with_miles),
        "odometers": current_odometers,
        "source": "odometer_delta",
    }


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

    # Skip if already archived recently (prevents duplicate work on re-runs)
    existing_path = os.path.join(args.output_dir, f"{archive_date}.json")
    if os.path.exists(existing_path):
        try:
            age_sec = (_time.time() - os.path.getmtime(existing_path))
            if age_sec < 7200:  # 2 hours
                print(f"  SKIP: {archive_date} already archived {age_sec / 60:.0f}min ago")
                sys.exit(0)
        except Exception:
            pass

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
                age_hours = (datetime.now() - gen_dt).total_seconds() / 3600
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

    # Build vehicle lookup once (used for speeding enrichment + mileage division mapping)
    vehicle_yards = {}
    vehicle_divisions = {}
    vehicle_drivers = {}
    if MOTIVE_KEY:
        vehicle_yards, vehicle_divisions, vehicle_drivers = _build_vehicle_lookup()

    # Enrich speeding yard + driver data
    speeding = loaded.get("speeding")
    if speeding and isinstance(speeding.get("events"), list) and vehicle_yards:
        yard_fixed = 0
        driver_fixed = 0
        for e in speeding["events"]:
            veh_full = e.get("vehicle", "")
            veh_short = veh_full.split(" ")[0].strip()
            if e.get("yard", "Unknown") == "Unknown":
                yard = vehicle_yards.get(veh_full) or vehicle_yards.get(veh_short)
                if yard:
                    e["yard"] = yard
                    yard_fixed += 1
            if e.get("driver", "Unknown") == "Unknown":
                drv = vehicle_drivers.get(veh_full) or vehicle_drivers.get(veh_short)
                if drv:
                    e["driver"] = drv
                    driver_fixed += 1
        if yard_fixed or driver_fixed:
            print(f"  Speeding: enriched {yard_fixed} yards, {driver_fixed} drivers from vehicle lookup")

    # Fetch mileage: prefer odometer deltas (no lag), fall back to IFTA
    mileage_data = None
    odometers = fetch_vehicle_odometers()
    if odometers:
        mileage_data = compute_odometer_mileage(
            odometers, args.output_dir, archive_date,
            vehicle_divisions=vehicle_divisions, vehicle_yards_map=vehicle_yards,
        )
    if not mileage_data or mileage_data.get("total_miles", 0) == 0:
        # Fall back to IFTA (may have lag but covers the gap)
        ifta_data = fetch_daily_mileage(archive_date)
        if ifta_data and ifta_data.get("total_miles", 0) > 0:
            # Preserve odometer readings even if using IFTA for miles
            if mileage_data and mileage_data.get("odometers"):
                ifta_data["odometers"] = mileage_data["odometers"]
                ifta_data["source"] = "ifta_with_odometer_baseline"
            else:
                ifta_data["source"] = "ifta"
            mileage_data = ifta_data
        elif mileage_data and mileage_data.get("odometers"):
            # No IFTA either, but at least save the odometer baseline
            pass
        else:
            mileage_data = None

    # Fetch driver scorecards (rolling 30-day window ending on archive date)
    scorecard_start = (datetime.strptime(archive_date, "%Y-%m-%d") - timedelta(days=29)).strftime("%Y-%m-%d")
    driver_scorecards = fetch_driver_scorecards(scorecard_start, archive_date)

    # Build archive object
    archive = {
        "date": archive_date,
        "generated_at": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "speeding": loaded.get("speeding"),
        "camera": loaded.get("camera"),
        "kpa": loaded.get("kpa"),
        "ytd": loaded.get("ytd"),
        "mileage": mileage_data,
        "driver_scorecards": driver_scorecards,
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
    # Summary log for GitHub Action monitoring
    spd_ct = len((loaded.get("speeding") or {}).get("events", []))
    cam_ct = len((loaded.get("camera") or {}).get("events", []))
    kpa_ok = "Y" if loaded.get("kpa") else "N"
    mi_total = round((mileage_data or {}).get("total_miles", 0))
    sc_ct = len(driver_scorecards) if driver_scorecards else 0
    print(f"  Archived {archive_date}: {output_path} ({size_kb:.1f} KB)")
    print(f"  Summary: speeding={spd_ct} camera={cam_ct} kpa={kpa_ok} mileage={mi_total}mi scorecards={sc_ct}")


if __name__ == "__main__":
    main()
