"""
ARCHIVE_TODAY.PY -- Combine daily output JSONs into single archive file
======================================================================
Reads from output/ directory (already populated by daily scripts) and
writes a combined archive/YYYY-MM-DD.json file. The archive date comes
from the report_date field in the output JSON, NOT from today's date
(daily scripts fetch yesterday's data).

Usage:
    python archive_today.py
    python archive_today.py --output-dir archive/
"""
import json
import os
import sys
import argparse
from datetime import datetime, timezone, timedelta


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

    # Build archive object
    archive = {
        "date": archive_date,
        "generated_at": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "speeding": loaded.get("speeding"),
        "camera": loaded.get("camera"),
        "kpa": loaded.get("kpa"),
        "ytd": loaded.get("ytd"),
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
