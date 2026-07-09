#!/usr/bin/env python3
"""
CASING DAILY SAFETY DIRECTOR BRIEFING EMAIL
=============================================
Consolidates Casing-only speeding, camera, drowsiness, and distracted driving
data into a single director-level summary email sent to krhodes@brhas.com.

Reads pre-generated JSON files (no duplicate API calls):
  - output/speeding_events.json  (from daily_speeding_report.py)
  - output/camera_events.json    (from daily_casing_camera_report.py)

Makes ONE API call when drowsiness events exist:
  - Motive /v1/hours_of_service   (driving duration for fatigue context)

Usage:
    python daily_casing_director_email.py             # Send email
    python daily_casing_director_email.py --no-email   # Console summary only
"""

import json
import os
import sys
import smtplib
import base64
import requests
from datetime import datetime, timedelta, timezone
from html import escape as html_escape
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from collections import Counter

try:
    from zoneinfo import ZoneInfo
    CENTRAL_TZ = ZoneInfo("America/Chicago")
except Exception:
    CENTRAL_TZ = timezone(timedelta(hours=-6))

# ==============================================================================
# CONFIGURATION
# ==============================================================================

# Branding colors (same as daily_speeding_report.py)
C_RED = "#C00000"
C_DARK = "#800000"
C_AMBER = "#FF8C00"
C_YELLOW_DARK = "#CC9900"
C_GREEN = "#008000"

# Recipient
DIRECTOR_RECIPIENT = "krhodes@brhas.com"

# Motive API (for HOS lookup only)
MOTIVE_API_KEY = os.environ.get("MOTIVE_API_KEY", "")
MOTIVE_BASE = "https://api.gomotive.com/v1"

# Gmail
GMAIL_ADDRESS = os.environ.get("GMAIL_ADDRESS", "")
GMAIL_APP_PASSWORD = os.environ.get("GMAIL_APP_PASSWORD", "")

# Paths
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(SCRIPT_DIR, "output")
LOGOS_DIR = os.path.join(SCRIPT_DIR, "logos")

# Event type classification (matches normalized event_type field in camera_events.json)
DROWSINESS_TYPES = {"drowsiness"}
DISTRACTION_TYPES = {"distraction", "cell_phone"}

# Dashboard URL
DASHBOARD_URL = "https://krhodes5267.github.io/daily-safety-report/"

# Logo cache
_LOGO_CACHE = {}


# ==============================================================================
# UTILITY FUNCTIONS
# ==============================================================================

def _h(text):
    """HTML-escape text safely."""
    return html_escape(str(text)) if text else ""


def _get_logo_b64(filename):
    """Load a logo file and return base64 data URI. Cached."""
    if filename in _LOGO_CACHE:
        return _LOGO_CACHE[filename]
    path = os.path.join(LOGOS_DIR, filename)
    if not os.path.exists(path):
        _LOGO_CACHE[filename] = ""
        return ""
    with open(path, "rb") as f:
        data = base64.b64encode(f.read()).decode("ascii")
    ext = filename.rsplit(".", 1)[-1].lower()
    mime = "image/png" if ext == "png" else "image/jpeg"
    uri = f"data:{mime};base64,{data}"
    _LOGO_CACHE[filename] = uri
    return uri


def _build_logo_html(max_height="50px"):
    """Build inline logo HTML for the Butch's logo."""
    main_logo = _get_logo_b64("Butchs.jpg")
    if not main_logo:
        return ""
    return (
        f'<div style="text-align:center;padding:10px 0;">'
        f'<img src="{main_logo}" alt="BRHAS" style="height:{max_height};margin:0 10px;">'
        f'</div>'
    )


# ==============================================================================
# DATA LOADING
# ==============================================================================

def load_speeding_events():
    """Load speeding_events.json and filter to Casing division only."""
    path = os.path.join(OUTPUT_DIR, "speeding_events.json")
    if not os.path.exists(path):
        print(f"  WARNING: {path} not found. Speeding data will be empty.")
        return [], ""
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)

    report_date = data.get("report_date", "")
    all_events = data.get("events", [])

    # Filter to Casing division only
    casing_events = [e for e in all_events if e.get("division", "") == "Casing"]
    print(f"  Speeding: {len(casing_events)} Casing events (of {len(all_events)} total)")
    return casing_events, report_date


def load_camera_events():
    """Load camera_events.json (already Casing-only)."""
    path = os.path.join(OUTPUT_DIR, "camera_events.json")
    if not os.path.exists(path):
        print(f"  WARNING: {path} not found. Camera data will be empty.")
        return [], ""
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)

    report_date = data.get("report_date", "")
    events = data.get("events", [])
    print(f"  Camera: {len(events)} events")
    return events, report_date


# ==============================================================================
# EVENT CATEGORIZATION
# ==============================================================================

def categorize_camera_events(camera_events):
    """Split camera events into drowsiness, distraction, and other."""
    drowsiness_events = []
    distraction_events = []
    other_events = []

    for e in camera_events:
        etype = e.get("event_type", "").lower()
        if etype in DROWSINESS_TYPES:
            drowsiness_events.append(e)
        elif etype in DISTRACTION_TYPES:
            distraction_events.append(e)
        else:
            other_events.append(e)

    return drowsiness_events, distraction_events, other_events


# ==============================================================================
# HOS API - DRIVING HOURS FOR DROWSINESS CONTEXT
# ==============================================================================

def fetch_hos_driving_hours(driver_names, report_date_str):
    """Fetch driving_duration from Motive HOS API for a date.

    Returns dict mapping driver_name -> driving_hours (float).
    Only matches drivers whose names are in driver_names set.
    """
    if not MOTIVE_API_KEY:
        print("  HOS: MOTIVE_API_KEY not set, skipping HOS lookup")
        return {}

    if not driver_names:
        return {}

    headers = {"X-Api-Key": MOTIVE_API_KEY}
    all_records = []
    page = 1

    try:
        while True:
            resp = requests.get(
                f"{MOTIVE_BASE}/hours_of_service",
                headers=headers,
                params={
                    "start_date": report_date_str,
                    "end_date": report_date_str,
                    "per_page": 100,
                    "page_no": page,
                },
                timeout=60,
            )
            resp.raise_for_status()
            data = resp.json()
            records = data.get("hours_of_services", [])
            if not records:
                break
            all_records.extend(records)
            pag = data.get("pagination", {})
            total = pag.get("total", 0)
            if page * 100 >= total:
                break
            page += 1

        print(f"  HOS: Fetched {len(all_records)} records across {page} page(s)")
    except Exception as e:
        print(f"  HOS: API call failed: {e}")
        return {}

    # Build driver_name -> driving_hours map (case-insensitive match)
    target_lower = {name.lower(): name for name in driver_names}
    driver_hours = {}

    for wrapper in all_records:
        rec = wrapper.get("hours_of_service", wrapper)
        driver = rec.get("driver", {})
        name = f"{driver.get('first_name', '')} {driver.get('last_name', '')}".strip()
        name_lower = name.lower()

        if name_lower in target_lower:
            original_name = target_lower[name_lower]
            driving_sec = rec.get("driving_duration", 0) or 0
            driver_hours[original_name] = driver_hours.get(original_name, 0) + driving_sec / 3600.0

    return driver_hours


# ==============================================================================
# HTML EMAIL BUILDER
# ==============================================================================

def build_director_email(speeding_events, camera_events, drowsiness_events,
                         distraction_events, driver_hours, report_date_str):
    """Build the full Casing director briefing HTML email."""
    # Parse report date for display
    try:
        report_date = datetime.strptime(report_date_str, "%Y-%m-%d")
        date_display = report_date.strftime("%A, %B %d, %Y")
    except Exception:
        date_display = report_date_str

    now_central = datetime.now(timezone.utc).astimezone(CENTRAL_TZ)

    # Pre-calculate tier counts
    spd_red = [e for e in speeding_events if e.get("tier") == "RED"]
    spd_orange = [e for e in speeding_events if e.get("tier") == "ORANGE"]
    spd_yellow = [e for e in speeding_events if e.get("tier") == "YELLOW"]

    cam_red = [e for e in camera_events if e.get("tier") == "RED"]
    cam_orange = [e for e in camera_events if e.get("tier") == "ORANGE"]
    cam_yellow = [e for e in camera_events if e.get("tier") == "YELLOW"]

    parts = []

    # ---- HEADER ----
    parts.append(f"""<html><head><meta charset="utf-8"></head>
<body style="margin:0;padding:0;background:#f4f4f4;">
<table width="100%" cellpadding="0" cellspacing="0" style="background:#f4f4f4;">
<tr><td align="center">
<table width="600" cellpadding="0" cellspacing="0" style="background:#ffffff;border:1px solid #ddd;margin:20px auto;font-family:Calibri,Arial,Helvetica,sans-serif;font-size:14px;color:#333;">

<tr><td style="background:#ffffff;padding:15px 20px;text-align:center;">
  {_build_logo_html(max_height="50px")}
</td></tr>
<tr><td style="background:{C_RED};padding:20px 30px;text-align:center;">
  <div style="font-size:14px;font-weight:bold;color:#ffffff;letter-spacing:1px;">BRHAS CASING DIVISION</div>
  <div style="font-size:22px;font-weight:bold;color:#ffffff;margin:6px 0;">DAILY SAFETY DIRECTOR BRIEFING</div>
  <div style="font-size:12px;color:#ffffff;">{date_display}</div>
</td></tr>""")

    # ---- DASHBOARD TILES ----
    tile = (
        "display:inline-block;width:120px;text-align:center;"
        "padding:15px 10px;margin:5px;border-radius:6px;"
    )
    parts.append(f"""
<tr><td style="padding:20px 30px;text-align:center;">
  <div style="{tile}background:#f8f0f0;border:2px solid {C_RED};">
    <div style="font-size:28px;font-weight:bold;color:{C_RED};">{len(speeding_events)}</div>
    <div style="font-size:11px;color:#666;">SPEEDING</div>
    <div style="font-size:10px;color:#FF0000;">{len(spd_red)} RED</div>
  </div>
  <div style="{tile}background:#f8f0f0;border:2px solid {C_AMBER};">
    <div style="font-size:28px;font-weight:bold;color:{C_AMBER};">{len(camera_events)}</div>
    <div style="font-size:11px;color:#666;">CAMERA</div>
    <div style="font-size:10px;color:#FF0000;">{len(cam_red)} RED</div>
  </div>
  <div style="{tile}background:#FFE0E0;border:2px solid #FF0000;">
    <div style="font-size:28px;font-weight:bold;color:#FF0000;">{len(drowsiness_events)}</div>
    <div style="font-size:11px;color:#666;">DROWSINESS</div>
  </div>
  <div style="{tile}background:#FFF0E0;border:2px solid {C_AMBER};">
    <div style="font-size:28px;font-weight:bold;color:{C_AMBER};">{len(distraction_events)}</div>
    <div style="font-size:11px;color:#666;">DISTRACTION</div>
  </div>
</td></tr>""")

    # ---- SECTION 1: CRITICAL ALERTS ----
    critical_items = []

    # Drowsiness alerts (most critical -- life safety)
    for e in drowsiness_events:
        hours = driver_hours.get(e.get("driver", ""), None)
        hours_str = f"{hours:.1f} hrs driven that day" if hours is not None else "Hours driven: N/A"
        critical_items.append(
            f'<div style="background:#FFE0E0;border-left:4px solid #FF0000;'
            f'padding:8px 12px;margin:6px 0;font-size:12px;">'
            f'<b style="color:#FF0000;">DROWSINESS / FATIGUE</b> | '
            f'{_h(e.get("driver", "Unknown"))} | {_h(e.get("vehicle", ""))}<br>'
            f'Speed: {e.get("speed", "")} mph | {_h(e.get("yard", ""))} | '
            f'{_h(e.get("time", ""))} | <b>{hours_str}</b> | '
            f'Coaching: {_h(e.get("coaching_status", ""))}'
            f'</div>'
        )

    # Distraction / cell phone alerts
    for e in distraction_events:
        critical_items.append(
            f'<div style="background:#FFE0E0;border-left:4px solid #FF0000;'
            f'padding:8px 12px;margin:6px 0;font-size:12px;">'
            f'<b style="color:#FF0000;">'
            f'{_h(e.get("display_name", "DISTRACTION").upper())}</b> | '
            f'{_h(e.get("driver", "Unknown"))} | {_h(e.get("vehicle", ""))}<br>'
            f'Speed: {e.get("speed", "")} mph | {_h(e.get("yard", ""))} | '
            f'{_h(e.get("time", ""))}'
            f'</div>'
        )

    # RED speeding events
    for e in sorted(spd_red, key=lambda x: x.get("overspeed", 0), reverse=True):
        map_link = ""
        if e.get("maps_link"):
            map_link = f' | <a href="{_h(e["maps_link"])}">Map</a>'
        critical_items.append(
            f'<div style="background:#FFE0E0;border-left:4px solid #FF0000;'
            f'padding:8px 12px;margin:6px 0;font-size:12px;">'
            f'<b style="color:#FF0000;">SPEEDING +{e.get("overspeed", "")} over</b> '
            f'({e.get("speed", "")} in a {e.get("posted_speed", "")} zone)<br>'
            f'{_h(e.get("driver", "Unknown"))} | {_h(e.get("vehicle", ""))} | '
            f'{_h(e.get("yard", ""))}{map_link}'
            f'</div>'
        )

    if critical_items:
        parts.append(f"""
<tr><td style="padding:15px 30px;">
  <h3 style="color:#FF0000;margin:0 0 10px 0;font-size:16px;border-bottom:2px solid #FF0000;padding-bottom:4px;">
    CRITICAL ALERTS ({len(critical_items)})
  </h3>
  {"".join(critical_items)}
</td></tr>""")

    # ---- SECTION 2: SPEEDING SUMMARY ----
    if speeding_events:
        # Per-yard breakdown
        yard_counts = Counter()
        yard_red_counts = Counter()
        for e in speeding_events:
            yard = e.get("yard", "") or "Unknown"
            yard_counts[yard] += 1
            if e.get("tier") == "RED":
                yard_red_counts[yard] += 1

        yard_rows = ""
        for yard in sorted(yard_counts.keys()):
            red_ct = yard_red_counts.get(yard, 0)
            red_badge = ""
            if red_ct:
                red_badge = (
                    f' <span style="background:#FF0000;color:#fff;padding:1px 6px;'
                    f'border-radius:3px;font-size:10px;">{red_ct} RED</span>'
                )
            yard_rows += (
                f'<tr>'
                f'<td style="padding:4px 10px;border-bottom:1px solid #eee;">{_h(yard)}</td>'
                f'<td style="padding:4px 10px;border-bottom:1px solid #eee;text-align:center;">{yard_counts[yard]}</td>'
                f'<td style="padding:4px 10px;border-bottom:1px solid #eee;">{red_badge}</td>'
                f'</tr>'
            )

        # Repeat offenders (3+ speeding events)
        driver_counts = Counter(
            e.get("driver", "Unknown") for e in speeding_events
            if e.get("driver", "Unknown") != "Unknown"
        )
        repeats = {n: c for n, c in driver_counts.items() if c >= 3}
        repeat_html = ""
        if repeats:
            for name, count in sorted(repeats.items(), key=lambda x: -x[1]):
                driver_evts = [e for e in speeding_events if e.get("driver") == name]
                worst = max(driver_evts, key=lambda x: x.get("overspeed", 0))
                repeat_html += (
                    f'<div style="background:#fff5f5;border-left:4px solid {C_RED};'
                    f'padding:6px 12px;margin:4px 0;font-size:12px;">'
                    f'<b>{_h(name)}: {count} events</b> '
                    f'(worst: +{worst.get("overspeed", "")} over) | '
                    f'{_h(worst.get("yard", ""))}'
                    f'</div>'
                )

        parts.append(f"""
<tr><td style="padding:15px 30px;">
  <h3 style="color:{C_RED};margin:0 0 10px 0;font-size:16px;border-bottom:2px solid {C_RED};padding-bottom:4px;">
    SPEEDING SUMMARY - RED: {len(spd_red)} | ORANGE: {len(spd_orange)} | YELLOW: {len(spd_yellow)}
  </h3>
  <table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;font-size:13px;">
    <tr style="background:#f0f0f0;">
      <th style="padding:4px 10px;text-align:left;">Yard</th>
      <th style="padding:4px 10px;text-align:center;">Events</th>
      <th style="padding:4px 10px;text-align:left;"></th>
    </tr>
    {yard_rows}
  </table>
  {repeat_html}
</td></tr>""")
    else:
        parts.append(f"""
<tr><td style="padding:15px 30px;">
  <h3 style="color:{C_GREEN};margin:0 0 10px 0;font-size:16px;">
    SPEEDING - No Casing speeding events yesterday
  </h3>
</td></tr>""")

    # ---- SECTION 3: CAMERA / DRIVECAM SUMMARY ----
    if camera_events:
        # Per-yard breakdown
        cam_yard_counts = Counter()
        cam_yard_red = Counter()
        for e in camera_events:
            yard = e.get("yard", "") or "Unknown"
            cam_yard_counts[yard] += 1
            if e.get("tier") == "RED":
                cam_yard_red[yard] += 1

        cam_yard_rows = ""
        for yard in sorted(cam_yard_counts.keys()):
            red_ct = cam_yard_red.get(yard, 0)
            red_badge = ""
            if red_ct:
                red_badge = (
                    f' <span style="background:#FF0000;color:#fff;padding:1px 6px;'
                    f'border-radius:3px;font-size:10px;">{red_ct} RED</span>'
                )
            cam_yard_rows += (
                f'<tr>'
                f'<td style="padding:4px 10px;border-bottom:1px solid #eee;">{_h(yard)}</td>'
                f'<td style="padding:4px 10px;border-bottom:1px solid #eee;text-align:center;">{cam_yard_counts[yard]}</td>'
                f'<td style="padding:4px 10px;border-bottom:1px solid #eee;">{red_badge}</td>'
                f'</tr>'
            )

        # Event type distribution
        type_counts = Counter(
            e.get("display_name", e.get("event_type", "Unknown")) for e in camera_events
        )
        type_str = " | ".join(f"{name}: {ct}" for name, ct in type_counts.most_common())

        # Repeat offenders (2+ camera events)
        cam_driver_counts = Counter(
            e.get("driver", "Unknown") for e in camera_events
            if e.get("driver", "Unknown") != "Unknown"
        )
        cam_repeats = {n: c for n, c in cam_driver_counts.items() if c >= 2}
        cam_repeat_html = ""
        if cam_repeats:
            for name, count in sorted(cam_repeats.items(), key=lambda x: -x[1]):
                driver_evts = [e for e in camera_events if e.get("driver") == name]
                types = ", ".join(sorted(set(
                    e.get("display_name", "") for e in driver_evts
                )))
                cam_repeat_html += (
                    f'<div style="background:#fff5f5;border-left:4px solid {C_AMBER};'
                    f'padding:6px 12px;margin:4px 0;font-size:12px;">'
                    f'<b>{_h(name)}: {count} events</b> ({_h(types)})'
                    f'</div>'
                )

        parts.append(f"""
<tr><td style="padding:15px 30px;">
  <h3 style="color:{C_RED};margin:0 0 10px 0;font-size:16px;border-bottom:2px solid {C_RED};padding-bottom:4px;">
    CAMERA EVENTS - RED: {len(cam_red)} | ORANGE: {len(cam_orange)} | YELLOW: {len(cam_yellow)}
  </h3>
  <table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;font-size:13px;">
    <tr style="background:#f0f0f0;">
      <th style="padding:4px 10px;text-align:left;">Yard</th>
      <th style="padding:4px 10px;text-align:center;">Events</th>
      <th style="padding:4px 10px;text-align:left;"></th>
    </tr>
    {cam_yard_rows}
  </table>
  <div style="font-size:12px;color:#666;margin:8px 0;"><b>Types:</b> {_h(type_str)}</div>
  {cam_repeat_html}
</td></tr>""")
    else:
        parts.append(f"""
<tr><td style="padding:15px 30px;">
  <h3 style="color:{C_GREEN};margin:0 0 10px 0;font-size:16px;">
    CAMERA EVENTS - No Casing camera events yesterday
  </h3>
</td></tr>""")

    # ---- SECTION 4: FATIGUE & DROWSINESS DETAIL ----
    if drowsiness_events:
        drowsy_details = ""
        for e in drowsiness_events:
            hours = driver_hours.get(e.get("driver", ""), None)
            hours_display = f"{hours:.1f} hours" if hours is not None else "N/A"
            coached = e.get("coaching_status", "")
            coached_badge = ""
            if coached == "coached":
                coached_badge = (
                    '<span style="background:#008000;color:#fff;padding:1px 6px;'
                    'border-radius:3px;font-size:10px;">COACHED</span>'
                )
            elif coached == "coachable":
                coached_badge = (
                    '<span style="background:#FF8C00;color:#fff;padding:1px 6px;'
                    'border-radius:3px;font-size:10px;">PENDING</span>'
                )
            else:
                coached_badge = _h(coached)

            drowsy_details += (
                f'<div style="background:#fff5f5;border-left:4px solid #FF0000;'
                f'padding:10px 12px;margin:6px 0;font-size:12px;">'
                f'<b style="color:#FF0000;">{_h(e.get("driver", "Unknown"))}</b> | '
                f'{_h(e.get("vehicle", ""))} | {_h(e.get("yard", ""))}<br>'
                f'Time: {_h(e.get("time", ""))} | Speed: {e.get("speed", "")} mph | '
                f'<b>Hours driven that day: {hours_display}</b><br>'
                f'Coaching: {coached_badge}'
                f'</div>'
            )

        parts.append(f"""
<tr><td style="padding:15px 30px;">
  <h3 style="color:#FF0000;margin:0 0 10px 0;font-size:16px;border-bottom:2px solid #FF0000;padding-bottom:4px;">
    FATIGUE & DROWSINESS DETAIL ({len(drowsiness_events)})
  </h3>
  <div style="font-size:11px;color:#888;margin-bottom:8px;">
    Hours driven = cumulative driving time for the calendar day (from Motive HOS data).
  </div>
  {drowsy_details}
</td></tr>""")

    # ---- NO EVENTS AT ALL ----
    if not speeding_events and not camera_events:
        parts.append(f"""
<tr><td style="padding:30px;text-align:center;">
  <div style="font-size:18px;font-weight:bold;color:{C_GREEN};">
    No Casing safety events yesterday!
  </div>
</td></tr>""")

    # ---- FOOTER ----
    parts.append(f"""
<tr><td style="background:{C_DARK};padding:15px 30px;text-align:center;">
  <div style="color:#ffcccc;font-size:10px;">Casing Division Safety Director Briefing</div>
  <div style="color:#ffffff;font-size:10px;margin-top:4px;">
    Generated {now_central.strftime('%I:%M %p CT')} | BRHAS HSE Department
  </div>
  <div style="color:#ffcccc;font-size:10px;margin-top:4px;">
    <a href="{DASHBOARD_URL}" style="color:#ffcccc;">View Dashboard</a>
  </div>
</td></tr>

</table>
</td></tr></table>
</body></html>""")

    return "\n".join(parts)


# ==============================================================================
# SEND EMAIL
# ==============================================================================

def send_director_email(html_body, report_date_str):
    """Send the director briefing email via Gmail SMTP."""
    if not GMAIL_ADDRESS or not GMAIL_APP_PASSWORD:
        print("  Email skipped -- GMAIL_ADDRESS or GMAIL_APP_PASSWORD not set.")
        return

    try:
        report_date = datetime.strptime(report_date_str, "%Y-%m-%d")
        subject = f"Casing Safety Director Briefing - {report_date.strftime('%B %d, %Y')}"
    except Exception:
        subject = f"Casing Safety Director Briefing - {report_date_str}"

    try:
        msg = MIMEMultipart("mixed")
        msg["From"] = GMAIL_ADDRESS
        msg["To"] = DIRECTOR_RECIPIENT
        msg["Subject"] = subject
        msg.attach(MIMEText(html_body, "html"))

        with smtplib.SMTP("smtp.gmail.com", 587, timeout=30) as server:
            server.starttls()
            server.login(GMAIL_ADDRESS, GMAIL_APP_PASSWORD)
            server.sendmail(GMAIL_ADDRESS, DIRECTOR_RECIPIENT, msg.as_string())

        print(f"  Email sent to {DIRECTOR_RECIPIENT}")
    except Exception as e:
        print(f"  Email failed: {e}")


# ==============================================================================
# MAIN
# ==============================================================================

def main():
    today = datetime.now(timezone.utc).astimezone(CENTRAL_TZ)
    yesterday = today - timedelta(days=1)

    print("\n" + "=" * 80)
    print("CASING DAILY SAFETY DIRECTOR BRIEFING")
    print(f"Report for: {yesterday.strftime('%A, %B %d, %Y')}")
    print("=" * 80)

    # --- Load data from JSON files ---
    print("\n[1] Loading speeding events...")
    speeding_events, spd_date = load_speeding_events()

    print("[2] Loading camera events...")
    camera_events, cam_date = load_camera_events()

    report_date = cam_date or spd_date or yesterday.strftime("%Y-%m-%d")

    # --- Categorize camera events ---
    print("[3] Categorizing camera events...")
    drowsiness_events, distraction_events, other_cam = categorize_camera_events(camera_events)
    print(f"    Drowsiness: {len(drowsiness_events)} | Distraction: {len(distraction_events)} | Other: {len(other_cam)}")

    # --- Fetch HOS data for drowsiness drivers ---
    driver_hours = {}
    if drowsiness_events:
        print("[4] Fetching HOS driving hours for drowsiness drivers...")
        driver_names = set(
            e.get("driver", "") for e in drowsiness_events
            if e.get("driver") and e.get("driver") != "Unknown"
        )
        driver_hours = fetch_hos_driving_hours(driver_names, report_date)
        for name in driver_names:
            hrs = driver_hours.get(name)
            if hrs is not None:
                print(f"    {name}: {hrs:.1f} hrs")
            else:
                print(f"    {name}: N/A")
    else:
        print("[4] No drowsiness events -- skipping HOS lookup")

    # --- Console summary ---
    print(f"\n--- DIRECTOR SUMMARY ---")
    print(f"  Speeding:    {len(speeding_events)} (RED: {len([e for e in speeding_events if e.get('tier')=='RED'])})")
    print(f"  Camera:      {len(camera_events)} (RED: {len([e for e in camera_events if e.get('tier')=='RED'])})")
    print(f"  Drowsiness:  {len(drowsiness_events)}")
    print(f"  Distraction: {len(distraction_events)}")

    # --- Build email ---
    print("\n[5] Building HTML email...")
    html_body = build_director_email(
        speeding_events, camera_events, drowsiness_events,
        distraction_events, driver_hours, report_date
    )

    # --- Save HTML preview ---
    preview_path = os.path.join(OUTPUT_DIR, "director_briefing_preview.html")
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    with open(preview_path, "w", encoding="utf-8") as f:
        f.write(html_body)
    print(f"    Preview: {preview_path}")

    # --- Send email ---
    if "--no-email" not in sys.argv:
        print("[6] Sending email...")
        send_director_email(html_body, report_date)
    else:
        print("[6] Email skipped (--no-email flag)")

    print("\n" + "=" * 80)
    print("COMPLETE")
    print("=" * 80 + "\n")


if __name__ == "__main__":
    main()
