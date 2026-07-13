"""
CAMERA_CONTROL.PY -- Motive Dashcam Camera Toggle
==================================================
Turn dashcam cameras ON/OFF for personal-use (non-DOT) vehicles.
Uses the Motive Camera Control API (PUT /v1/cameras/{eld_device_id}).

Usage:
    py -3 camera_control.py list                       # List all vehicles with ELD device IDs
    py -3 camera_control.py list --division Casing     # Filter by division
    py -3 camera_control.py status <eld_device_id>     # Check camera status for a device
    py -3 camera_control.py off <eld_device_id>        # Turn cameras OFF
    py -3 camera_control.py on <eld_device_id>         # Turn cameras ON
    py -3 camera_control.py off <id1> <id2> <id3>      # Bulk toggle multiple devices

Requires:
    MOTIVE_API_KEY environment variable (or hardcoded below for local use)
"""

import os
import sys
import time
import json
import requests
from api_config import MOTIVE_BASE_V1, GROUP_ID_MAP

API_KEY = os.environ.get("MOTIVE_API_KEY", "8d3dd502-36c0-47c4-ade3-a1fbbef0c05c")

HEADERS = {
    "X-Api-Key": API_KEY,
    "Accept": "application/json",
    "Content-Type": "application/json",
}


def get_vehicle_gateways(division_filter=None):
    """Fetch all vehicle gateways (ELD devices) from Motive."""
    url = f"{MOTIVE_BASE_V1}/vehicle_gateways"
    all_devices = []
    page = 1

    while True:
        resp = requests.get(url, headers=HEADERS, params={"page_no": page, "per_page": 100})
        resp.raise_for_status()
        data = resp.json()

        gateways = data.get("vehicle_gateways", [])
        if not gateways:
            break

        for gw in gateways:
            device = gw.get("vehicle_gateway", gw)
            all_devices.append(device)

        # Check for next page
        pagination = data.get("pagination", {})
        total = pagination.get("total", 0)
        per_page = pagination.get("per_page", 100)
        total_pages = (total + per_page - 1) // per_page if total else 1
        if page >= total_pages:
            break
        page += 1

    # Enrich with division/yard from GROUP_ID_MAP
    results = []
    for d in all_devices:
        vehicle = d.get("vehicle", {}) or {}
        group_ids = []
        # group_ids can be a flat list of ints or a list of dicts
        raw_groups = vehicle.get("group_ids", []) or vehicle.get("groups", []) or []
        for g in raw_groups:
            if isinstance(g, dict):
                gid = g.get("id")
            else:
                gid = g
            if gid:
                group_ids.append(gid)

        division, yard = "Unknown", "Unknown"
        for gid in group_ids:
            if gid in GROUP_ID_MAP:
                division, yard = GROUP_ID_MAP[gid]
                break

        entry = {
            "eld_device_id": d.get("id"),
            "serial": d.get("serial_number", ""),
            "model": d.get("model", ""),
            "vehicle_id": vehicle.get("id", ""),
            "vehicle_number": vehicle.get("number", ""),
            "vehicle_name": vehicle.get("name", vehicle.get("number", "")),
            "division": division,
            "yard": yard,
            "group_ids": group_ids,
        }
        results.append(entry)

    if division_filter:
        results = [r for r in results if r["division"].lower() == division_filter.lower()]

    results.sort(key=lambda x: (x["division"], x["yard"], x["vehicle_number"]))
    return results


def get_vehicles_list(division_filter=None):
    """Fetch vehicles from /v1/vehicles endpoint as fallback."""
    url = f"{MOTIVE_BASE_V1}/vehicles"
    all_vehicles = []
    page = 1

    while True:
        resp = requests.get(url, headers=HEADERS, params={"page_no": page, "per_page": 100})
        resp.raise_for_status()
        data = resp.json()

        vehicles = data.get("vehicles", [])
        if not vehicles:
            break

        for v in vehicles:
            vehicle = v.get("vehicle", v)
            all_vehicles.append(vehicle)

        pagination = data.get("pagination", {})
        total = pagination.get("total", 0)
        per_page = pagination.get("per_page", 100)
        total_pages = (total + per_page - 1) // per_page if total else 1
        if page >= total_pages:
            break
        page += 1
        print(f"  Page {page}/{total_pages}...")

    results = []
    for v in all_vehicles:
        eld = v.get("eld_device", {}) or {}
        eld_id = eld.get("id")
        if not eld_id:
            continue

        group_ids = []
        raw_groups = v.get("group_ids", []) or v.get("groups", []) or []
        for g in raw_groups:
            if isinstance(g, dict):
                gid = g.get("id")
            else:
                gid = g
            if gid:
                group_ids.append(gid)

        division, yard = "Unknown", "Unknown"
        for gid in group_ids:
            if gid in GROUP_ID_MAP:
                division, yard = GROUP_ID_MAP[gid]
                break

        entry = {
            "eld_device_id": eld_id,
            "serial": eld.get("serial_number", ""),
            "model": eld.get("model", ""),
            "vehicle_id": v.get("id", ""),
            "vehicle_number": v.get("number", ""),
            "vehicle_name": v.get("name", v.get("number", "")),
            "division": division,
            "yard": yard,
            "group_ids": group_ids,
        }
        results.append(entry)

    if division_filter:
        results = [r for r in results if r["division"].lower() == division_filter.lower()]

    results.sort(key=lambda x: (x["division"], x["yard"], x["vehicle_number"]))
    return results


def camera_toggle(eld_device_id, state):
    """Toggle camera ON or OFF. Returns (req_id, req_status)."""
    state = state.upper()
    if state not in ("ON", "OFF"):
        raise ValueError("State must be ON or OFF")

    url = f"{MOTIVE_BASE_V1}/cameras/{eld_device_id}"
    resp = requests.put(url, headers=HEADERS, json={"camera_state": state})
    resp.raise_for_status()
    data = resp.json()
    return data.get("req_id"), data.get("req_status")


def camera_poll(eld_device_id, req_id, max_attempts=10, interval=8):
    """Poll camera control job status until completion or timeout."""
    url = f"{MOTIVE_BASE_V1}/cameras/{eld_device_id}/{req_id}"

    for attempt in range(max_attempts):
        time.sleep(interval)
        try:
            resp = requests.get(url, headers=HEADERS)
            if resp.status_code == 429:
                print(f"  Attempt {attempt + 1}/{max_attempts}: rate limited, backing off...")
                time.sleep(15)
                continue
            resp.raise_for_status()
            data = resp.json()
            status = data.get("req_status", "Unknown")

            if status in ("Succeeded", "Failed", "Error"):
                return status
            print(f"  Attempt {attempt + 1}/{max_attempts}: {status}...")
        except requests.HTTPError as e:
            print(f"  Attempt {attempt + 1}/{max_attempts}: {e}")

    return "Timeout (command was submitted and may still complete)"


def cmd_list(args):
    """List vehicles with their ELD device IDs."""
    division = None
    if "--division" in args:
        idx = args.index("--division")
        if idx + 1 < len(args):
            division = args[idx + 1]

    print("Fetching vehicles from Motive...")
    vehicles = get_vehicles_list(division)

    if not vehicles:
        # Fallback to vehicle_gateways endpoint
        print("Trying vehicle_gateways endpoint...")
        vehicles = get_vehicle_gateways(division)

    if not vehicles:
        print("No vehicles found.")
        return

    print(f"\n{'ELD Device ID':<16} {'Vehicle #':<18} {'Division':<16} {'Yard':<16} {'Model'}")
    print("-" * 85)
    for v in vehicles:
        print(f"{v['eld_device_id']:<16} {v['vehicle_number']:<18} {v['division']:<16} {v['yard']:<16} {v['model']}")

    print(f"\nTotal: {len(vehicles)} vehicles")


def cmd_toggle(state, device_ids):
    """Toggle cameras ON or OFF for one or more devices."""
    if not device_ids:
        print(f"Usage: camera_control.py {state.lower()} <eld_device_id> [<id2> ...]")
        sys.exit(1)

    for eld_id in device_ids:
        print(f"\n[{eld_id}] Sending camera {state.upper()} request...")
        try:
            req_id, req_status = camera_toggle(eld_id, state)
            print(f"  Request submitted: req_id={req_id}, status={req_status}")

            if req_status == "Submitted":
                print(f"  Polling for completion...")
                final = camera_poll(eld_id, req_id)
                if final == "Succeeded":
                    print(f"  Camera {state.upper()} -- CONFIRMED")
                else:
                    print(f"  Final status: {final}")
        except requests.HTTPError as e:
            print(f"  ERROR: {e}")
            if e.response is not None:
                print(f"  Response: {e.response.text}")
        except Exception as e:
            print(f"  ERROR: {e}")


def cmd_status(device_ids):
    """Check current camera status for device(s)."""
    if not device_ids:
        print("Usage: camera_control.py status <eld_device_id> [<id2> ...]")
        sys.exit(1)

    for eld_id in device_ids:
        print(f"\n[{eld_id}] Checking camera status...")
        # The API doesn't have a direct status endpoint, but we can check
        # device info through the vehicle gateway
        url = f"{MOTIVE_BASE_V1}/vehicle_gateways/{eld_id}"
        try:
            resp = requests.get(url, headers=HEADERS)
            resp.raise_for_status()
            data = resp.json()
            gw = data.get("vehicle_gateway", data)
            print(f"  Device: {gw.get('serial_number', 'N/A')}")
            print(f"  Model: {gw.get('model', 'N/A')}")
            vehicle = gw.get("vehicle", {}) or {}
            print(f"  Vehicle: {vehicle.get('number', 'N/A')}")
            cam_status = gw.get("camera_status") or gw.get("status", "Check Motive Dashboard")
            print(f"  Camera Status: {cam_status}")
        except requests.HTTPError as e:
            print(f"  ERROR: {e}")
        except Exception as e:
            print(f"  ERROR: {e}")


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)

    command = sys.argv[1].lower()

    if command == "list":
        cmd_list(sys.argv[2:])
    elif command == "off":
        cmd_toggle("OFF", sys.argv[2:])
    elif command == "on":
        cmd_toggle("ON", sys.argv[2:])
    elif command == "status":
        cmd_status(sys.argv[2:])
    elif command in ("help", "--help", "-h"):
        print(__doc__)
    else:
        print(f"Unknown command: {command}")
        print(__doc__)
        sys.exit(1)


if __name__ == "__main__":
    main()
