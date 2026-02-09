#!/usr/bin/env python3
"""
INVESTIGATION: Find where Motive stores video URLs for driver performance events.
Temporary diagnostic script — run via GitHub Actions then delete.
"""

import requests
import json
import os
import sys

MOTIVE_API_KEY = os.environ.get("MOTIVE_API_KEY")
if not MOTIVE_API_KEY:
    print("ERROR: MOTIVE_API_KEY not set")
    sys.exit(1)

HEADERS = {"X-Api-Key": MOTIVE_API_KEY}
BASE_V1 = "https://api.gomotive.com/v1"
BASE_V2 = "https://api.gomotive.com/v2"


def safe_get(url, params=None, label=""):
    """Make a GET request, print full response, return data."""
    print(f"\n{'='*80}")
    print(f"  {label}")
    print(f"  GET {url}")
    if params:
        print(f"  Params: {params}")
    print(f"{'='*80}")
    try:
        resp = requests.get(url, headers=HEADERS, params=params or {}, timeout=30)
        print(f"  Status: {resp.status_code}")
        try:
            data = resp.json()
            # Truncate very large responses but print structure
            text = json.dumps(data, indent=2, default=str)
            if len(text) > 8000:
                print(text[:8000])
                print(f"\n  ... [TRUNCATED — full response is {len(text)} chars]")
            else:
                print(text)
            return resp.status_code, data
        except Exception:
            print(f"  Body (not JSON): {resp.text[:2000]}")
            return resp.status_code, None
    except Exception as e:
        print(f"  ERROR: {e}")
        return 0, None


def find_fields_recursive(obj, prefix="", target_keywords=None):
    """Recursively find all field names, flagging those matching keywords."""
    if target_keywords is None:
        target_keywords = ["url", "video", "media", "link", "footage",
                          "recording", "clip", "s3", "camera", "dash"]
    results = []
    if isinstance(obj, dict):
        for key, val in obj.items():
            full_path = f"{prefix}.{key}" if prefix else key
            # Check if key matches any target keyword
            key_lower = key.lower()
            matched = [kw for kw in target_keywords if kw in key_lower]
            val_preview = repr(val)[:120] if not isinstance(val, (dict, list)) else type(val).__name__
            results.append((full_path, val_preview, matched))
            # Recurse into dicts and lists
            if isinstance(val, dict):
                results.extend(find_fields_recursive(val, full_path, target_keywords))
            elif isinstance(val, list) and val and isinstance(val[0], dict):
                results.extend(find_fields_recursive(val[0], f"{full_path}[0]", target_keywords))
    return results


def main():
    print("\n" + "#" * 80)
    print("# MOTIVE VIDEO URL INVESTIGATION")
    print("#" * 80)

    # =========================================================================
    # STEP 1: Fetch recent events, dump ALL field names
    # =========================================================================
    print("\n\n" + "=" * 80)
    print("STEP 1: Fetch events and analyze ALL field names")
    print("=" * 80)

    status, data = safe_get(
        f"{BASE_V2}/driver_performance_events",
        {"per_page": 5, "page_no": 1},
        "Fetch 5 recent events from v2/driver_performance_events"
    )

    event_id = None
    if data and "driver_performance_events" in data:
        events = data["driver_performance_events"]
        print(f"\n  Got {len(events)} events")

        for i, wrapper in enumerate(events):
            evt = wrapper.get("driver_performance_event", wrapper)
            print(f"\n  --- Event {i+1}: type={evt.get('type')}, id={evt.get('id')} ---")

            if i == 0:
                event_id = evt.get("id")

                # Print EVERY field name
                print(f"\n  ALL FIELD NAMES IN EVENT:")
                fields = find_fields_recursive(evt)
                for path, val_preview, matched in fields:
                    flag = " <<<< MATCH" if matched else ""
                    print(f"    {path}: {val_preview}{flag}")

                # Print full JSON of first event
                print(f"\n  FULL JSON OF EVENT 1:")
                debug = {}
                for k, v in evt.items():
                    if isinstance(v, list) and len(v) > 5:
                        debug[k] = f"[array of {len(v)} items, first={v[0]}]"
                    else:
                        debug[k] = v
                print(json.dumps(debug, indent=2, default=str))

            # Check camera_media for all events
            cm = evt.get("camera_media")
            print(f"    camera_media = {cm}")

        # Find an event with camera_media not null
        print(f"\n  Scanning all 5 events for non-null camera_media...")
        for i, wrapper in enumerate(events):
            evt = wrapper.get("driver_performance_event", wrapper)
            cm = evt.get("camera_media")
            if cm:
                print(f"    EVENT {i+1} HAS camera_media! type={evt.get('type')}")
                print(f"    camera_media = {json.dumps(cm, indent=2, default=str)}")
                break
        else:
            print(f"    All 5 events have camera_media=null")

    # =========================================================================
    # STEP 2: Try to fetch MORE events to find one with camera_media
    # =========================================================================
    print("\n\n" + "=" * 80)
    print("STEP 2: Search for an event WITH camera_media (scan up to 500 events)")
    print("=" * 80)

    found_camera_event = None
    for page in range(1, 6):  # 5 pages x 100 = up to 500 events
        status, data = safe_get(
            f"{BASE_V2}/driver_performance_events",
            {"per_page": 100, "page_no": page},
            f"Page {page} — scanning for non-null camera_media"
        )
        if not data:
            break
        events = data.get("driver_performance_events", [])
        if not events:
            break

        # Only print summary, not full response
        types_on_page = {}
        camera_count = 0
        for wrapper in events:
            evt = wrapper.get("driver_performance_event", wrapper)
            t = evt.get("type", "unknown")
            types_on_page[t] = types_on_page.get(t, 0) + 1
            cm = evt.get("camera_media")
            if cm:
                camera_count += 1
                if not found_camera_event:
                    found_camera_event = evt
                    print(f"\n  FOUND EVENT WITH camera_media!")
                    print(f"  Event ID: {evt.get('id')}")
                    print(f"  Type: {evt.get('type')}")
                    print(f"  camera_media value:")
                    print(json.dumps(cm, indent=2, default=str))
                    # Also dump ALL fields of this event
                    print(f"\n  ALL FIELDS of this camera event:")
                    fields = find_fields_recursive(evt)
                    for path, val_preview, matched in fields:
                        flag = " <<<< MATCH" if matched else ""
                        print(f"    {path}: {val_preview}{flag}")

        print(f"\n  Page {page}: {len(events)} events, {camera_count} with camera_media")
        print(f"  Types: {types_on_page}")

        if found_camera_event:
            break

    if not found_camera_event:
        print("\n  NO events found with camera_media across 500 events!")

    # =========================================================================
    # STEP 3: Try individual event detail endpoints
    # =========================================================================
    if event_id:
        print("\n\n" + "=" * 80)
        print(f"STEP 3: Try detail endpoints for event_id={event_id}")
        print("=" * 80)

        endpoints = [
            (f"{BASE_V2}/driver_performance_events/{event_id}",
             "v2/driver_performance_events/{id}"),
            (f"{BASE_V1}/driver_performance_events/{event_id}",
             "v1/driver_performance_events/{id}"),
            (f"{BASE_V2}/driver_performance_events/{event_id}/media",
             "v2/driver_performance_events/{id}/media"),
            (f"{BASE_V1}/driver_performance_events/{event_id}/media",
             "v1/driver_performance_events/{id}/media"),
            (f"{BASE_V1}/safety/events/{event_id}",
             "v1/safety/events/{id}"),
            (f"{BASE_V2}/safety_events/{event_id}",
             "v2/safety_events/{id}"),
            (f"{BASE_V1}/safety_events/{event_id}",
             "v1/safety_events/{id}"),
            (f"{BASE_V2}/driver_performance_events/{event_id}/video",
             "v2/driver_performance_events/{id}/video"),
        ]

        for url, label in endpoints:
            status, data = safe_get(url, {}, label)
            if data and status == 200:
                # Scan for video/media fields
                fields = find_fields_recursive(data)
                matches = [(p, v) for p, v, m in fields if m]
                if matches:
                    print(f"\n  VIDEO/MEDIA FIELDS FOUND:")
                    for path, val in matches:
                        print(f"    {path}: {val}")

    # =========================================================================
    # STEP 4: Try standalone media/video endpoints
    # =========================================================================
    print("\n\n" + "=" * 80)
    print("STEP 4: Try standalone media/video endpoints")
    print("=" * 80)

    media_endpoints = [
        (f"{BASE_V1}/media", {"per_page": 3}, "v1/media"),
        (f"{BASE_V2}/media", {"per_page": 3}, "v2/media"),
        (f"{BASE_V1}/videos", {"per_page": 3}, "v1/videos"),
        (f"{BASE_V2}/videos", {"per_page": 3}, "v2/videos"),
        (f"{BASE_V1}/safety/media", {"per_page": 3}, "v1/safety/media"),
        (f"{BASE_V2}/safety/media", {"per_page": 3}, "v2/safety/media"),
        (f"{BASE_V1}/camera_media", {"per_page": 3}, "v1/camera_media"),
        (f"{BASE_V2}/camera_media", {"per_page": 3}, "v2/camera_media"),
        (f"{BASE_V1}/driver_performance_events/media", {"per_page": 3},
         "v1/driver_performance_events/media"),
    ]

    for url, params, label in media_endpoints:
        status, data = safe_get(url, params, label)
        if data and status == 200:
            fields = find_fields_recursive(data)
            matches = [(p, v) for p, v, m in fields if m]
            if matches:
                print(f"\n  VIDEO/MEDIA FIELDS FOUND:")
                for path, val in matches:
                    print(f"    {path}: {val}")

    # =========================================================================
    # STEP 5: Check for dashboard/coaching URLs in events
    # =========================================================================
    print("\n\n" + "=" * 80)
    print("STEP 5: Check for dashboard/coaching URL fields")
    print("=" * 80)

    # Re-fetch first event and check all string fields for URL patterns
    status, data = safe_get(
        f"{BASE_V2}/driver_performance_events",
        {"per_page": 1, "page_no": 1},
        "Re-check first event for any URL-like string values"
    )
    if data and "driver_performance_events" in data:
        evt = data["driver_performance_events"][0]
        evt = evt.get("driver_performance_event", evt)

        print(f"\n  Checking all string values for URL patterns (http/https):")
        def find_urls(obj, prefix=""):
            if isinstance(obj, dict):
                for k, v in obj.items():
                    fp = f"{prefix}.{k}" if prefix else k
                    if isinstance(v, str) and ("http" in v.lower() or "://" in v):
                        print(f"    {fp} = {v}")
                    elif isinstance(v, (dict, list)):
                        find_urls(v, fp)
            elif isinstance(obj, list):
                for i, item in enumerate(obj[:3]):
                    find_urls(item, f"{prefix}[{i}]")
        find_urls(evt)

        # Also check specific fields
        for field in ["coaching_url", "dashboard_url", "event_url", "web_url",
                      "detail_url", "review_url", "permalink", "share_url",
                      "coaching_status_url", "fleet_url"]:
            val = evt.get(field)
            if val:
                print(f"    {field} = {val}")

    # =========================================================================
    # STEP 6: Try /v1/vehicle_media and /v1/driver_performance_events with
    #         specific camera-event types to find one with video
    # =========================================================================
    print("\n\n" + "=" * 80)
    print("STEP 6: Search specifically for camera-triggered event types")
    print("=" * 80)

    camera_types = ["distraction", "cell_phone", "drowsiness", "camera_obstruction",
                    "close_following", "smoking"]
    for etype in camera_types:
        # Try type filter
        status, data = safe_get(
            f"{BASE_V2}/driver_performance_events",
            {"per_page": 3, "page_no": 1, "event_types": etype},
            f"Filter by event_types={etype}"
        )
        if data:
            events = data.get("driver_performance_events", [])
            total = data.get("total", data.get("pagination", {}).get("total", len(events)))
            print(f"  Total for type '{etype}': {total}")
            for wrapper in events:
                evt = wrapper.get("driver_performance_event", wrapper)
                cm = evt.get("camera_media")
                if cm:
                    print(f"  FOUND camera_media for {etype} event!")
                    print(f"  camera_media = {json.dumps(cm, indent=2, default=str)}")
                    # Print ALL fields
                    print(f"\n  ALL FIELDS:")
                    fields = find_fields_recursive(evt)
                    for path, val_preview, matched in fields:
                        flag = " <<<< MATCH" if matched else ""
                        print(f"    {path}: {val_preview}{flag}")
                    break

        # Also try type_filter param
        status2, data2 = safe_get(
            f"{BASE_V2}/driver_performance_events",
            {"per_page": 3, "page_no": 1, "type": etype},
            f"Filter by type={etype}"
        )
        if data2:
            events2 = data2.get("driver_performance_events", [])
            total2 = data2.get("total", data2.get("pagination", {}).get("total", len(events2)))
            if str(total2) != str(total):
                print(f"  Total for type param '{etype}': {total2}")
                for wrapper in events2:
                    evt = wrapper.get("driver_performance_event", wrapper)
                    cm = evt.get("camera_media")
                    if cm:
                        print(f"  FOUND camera_media for {etype} event (type param)!")
                        print(f"  camera_media = {json.dumps(cm, indent=2, default=str)}")
                        break

    # =========================================================================
    # STEP 7: Construct fallback dashboard URLs
    # =========================================================================
    print("\n\n" + "=" * 80)
    print("STEP 7: Fallback — Construct Motive dashboard URLs")
    print("=" * 80)

    if event_id:
        urls = [
            f"https://app.gomotive.com/safety/events/{event_id}",
            f"https://app.gomotive.com/en/safety/events/{event_id}",
            f"https://fleet.keeptruckin.com/safety-analytics/events/{event_id}",
            f"https://app.gomotive.com/en/safety-analytics/events/{event_id}",
            f"https://app.gomotive.com/en/fleet/safety/events/{event_id}",
        ]
        print(f"\n  Event ID: {event_id}")
        print(f"  Candidate dashboard URLs to try manually:")
        for u in urls:
            print(f"    {u}")

    print("\n\n" + "#" * 80)
    print("# INVESTIGATION COMPLETE")
    print("#" * 80 + "\n")


if __name__ == "__main__":
    main()
