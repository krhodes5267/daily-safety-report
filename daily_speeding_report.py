#!/usr/bin/env python3
"""
DAILY SPEEDING REPORT - AUTOMATED (GitHub Actions)
===================================================
Runs daily at 5:05 AM Central via GitHub Actions.

Pulls speeding events from Motive API (last 24 hours) and generates:
- Word document (.docx) with full event table
- HTML email with the same data, sent via Gmail SMTP

Thresholds (whichever is worse wins):
- RED: 15+ mph over posted limit OR 90+ mph absolute (termination-level)
- ORANGE: 10-14 mph over posted limit (formal coaching)
- YELLOW: 6-9 mph over posted limit (monitoring)
- Repeat offenders: drivers with 2+ events flagged separately

Uses the /v1/speeding_events endpoint which returns actual speeding-over-posted-limit
events with posted speed limits, duration, severity, and GPS coordinates.
"""

import requests
import smtplib
import os
import sys
from datetime import datetime, timedelta
from html import escape as html_escape
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from collections import Counter
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

# ==============================================================================
# SETUP
# ==============================================================================

MOTIVE_API_KEY = os.environ.get("MOTIVE_API_KEY")
if not MOTIVE_API_KEY:
    print("ERROR: MOTIVE_API_KEY environment variable is not set.")
    sys.exit(1)

MOTIVE_BASE_URL = "https://api.gomotive.com/v1"

# Logos are optional - only exist on local machines
LOGOS_PATH = os.path.expanduser("~/Downloads")
LOGOS = {
    'butchs': 'Butchs.jpg',
    'trucking': 'ButchTrucking-01.jpg',
    'permian': 'Permian-01.jpg',
    'hutchs': 'rsz_hutchs_logo-1.png',
    'transcend': 'Transcend.jpg',
    'valor': 'Valor Logo.jpg',
}

MOTIVE_GROUP_MAP = {
    'group_midland_rathole': 'Midland Rathole',
    'group_jourdanton_rathole': 'Jourdanton Rathole',
    'group_levelland_rathole': 'Levelland Rathole',
    'group_ohio_rathole': 'Ohio Rathole',
    'group_pennsylvania_rathole': 'Pennsylvania Rathole',
    'group_oklahoma_rathole': 'Oklahoma Rathole',
    'group_barstow_rathole': 'Barstow Rathole',
    'group_north_dakota_rathole': 'North Dakota Rathole',
    'group_midland_casing': 'Midland Casing',
    'group_bryan_casing': 'Bryan Casing',
    'group_kilgore_casing': 'Kilgore Casing',
    'group_hobbs_casing': 'Hobbs Casing',
    'group_jourdanton_casing': 'Jourdanton Casing',
    'group_laredo_casing': 'Laredo Casing',
    'group_san_angelo_casing': 'San Angelo Casing',
    'group_anchors': 'Anchors',
    'group_environmental': 'Environmental',
    'group_fencing': 'Fencing',
    'group_construction': 'Construction',
    'group_poly_pipe': 'Poly Pipe',
    'group_pit_lining': 'Pit Lining',
    'group_downhole_tools': 'Downhole Tools',
    'group_trucking': "Butch's Trucking",
    'group_transcend': 'Transcend Drilling',
    'group_valor': 'Valor Energy Services',
}


# ==============================================================================
# MOTIVE API - PULL SPEEDING EVENTS
# ==============================================================================

KMH_TO_MPH = 0.621371


def _get_driver_name(evt):
    """Get driver name from the event, falling back to vehicle number.

    The 'driver' field is often None. When it is, the driver name
    is embedded in the vehicle number (e.g., 'POL-2324PP - Yem Bobey').
    """
    driver = evt.get("driver")
    if driver and isinstance(driver, dict):
        first = driver.get("first_name", "")
        last = driver.get("last_name", "")
        name = f"{first} {last}".strip()
        if name:
            return name

    # Fallback: parse from vehicle number ("POL-2324PP - Yem Bobey")
    vehicle = evt.get("vehicle", {})
    if isinstance(vehicle, dict):
        veh_num = vehicle.get("number", "")
        if " - " in veh_num:
            return veh_num.split(" - ", 1)[1].strip()

    return "Unknown"


def _format_duration(seconds):
    """Format duration in seconds to a readable string."""
    if not seconds or not isinstance(seconds, (int, float)):
        return "N/A"
    seconds = int(seconds)
    if seconds < 60:
        return f"{seconds}s"
    minutes = seconds // 60
    secs = seconds % 60
    if secs:
        return f"{minutes}m {secs}s"
    return f"{minutes}m"


def get_24h_speeding_events():
    """Pull all speeding events from the last 24 hours using /v1/speeding_events.

    This endpoint returns actual speeding-over-posted-limit events with
    posted speed limits, duration, severity, and GPS coordinates.
    """
    end_time = datetime.utcnow()
    start_time = end_time - timedelta(hours=24)

    start_iso = start_time.strftime("%Y-%m-%dT%H:%M:%SZ")
    end_iso = end_time.strftime("%Y-%m-%dT%H:%M:%SZ")

    headers = {"X-Api-Key": MOTIVE_API_KEY}

    all_events = []
    page = 1

    while True:
        params = {
            "per_page": 100,
            "page_no": page,
            "start_time": start_iso,
            "end_time": end_iso,
        }

        try:
            response = requests.get(
                f"{MOTIVE_BASE_URL}/speeding_events",
                headers=headers,
                params=params,
                timeout=30,
            )
            response.raise_for_status()

            data = response.json()
            events = data.get("speeding_events", [])

            if not events:
                break

            for wrapper in events:
                # Each event is nested: {"speeding_event": {actual data}}
                evt = wrapper.get("speeding_event", wrapper)
                enriched = enrich_event(evt)
                all_events.append(enriched)

            # Pagination: API returns total/per_page/page_no at top level
            total = data.get("total", 0)
            if page * 100 < total:
                page += 1
            else:
                break

        except Exception as e:
            print(f"    Error fetching page {page}: {e}")
            break

    return sorted(all_events, key=lambda x: x["speed"], reverse=True)


def enrich_event(event):
    """Classify and enrich a speeding event from /v1/speeding_events."""
    # Convert speeds from km/h to mph
    max_speed_kmh = event.get("max_vehicle_speed") or event.get("avg_vehicle_speed") or 0
    max_speed = round(max_speed_kmh * KMH_TO_MPH, 1)

    posted_speed_kmh = event.get("min_posted_speed_limit_in_kph") or 0
    posted_speed = round(posted_speed_kmh * KMH_TO_MPH, 1)

    over_speed_kmh = event.get("max_over_speed_in_kph") or event.get("avg_over_speed_in_kph") or 0
    overspeed = round(over_speed_kmh * KMH_TO_MPH, 1)

    # Tier classification: check BOTH absolute speed AND over-limit, worst wins
    if overspeed >= 15 or max_speed >= 90:
        tier = "RED"
    elif overspeed >= 10:
        tier = "ORANGE"
    elif overspeed >= 6:
        tier = "YELLOW"
    else:
        tier = "YELLOW"  # All API events are already 6+ over

    driver_name = _get_driver_name(event)

    vehicle = event.get("vehicle", {})
    vehicle_number = (
        vehicle.get("number", "Unknown") if isinstance(vehicle, dict) else str(vehicle)
    )

    # Duration
    duration_secs = event.get("duration", 0)
    duration_str = _format_duration(duration_secs)

    # Severity from metadata
    metadata = event.get("metadata", {})
    severity = metadata.get("severity", "unknown") if isinstance(metadata, dict) else "unknown"

    # Timestamp
    timestamp = event.get("start_time") or event.get("end_time", "")
    try:
        event_time = datetime.fromisoformat(timestamp.replace("Z", "+00:00"))
        formatted_time = event_time.strftime("%m/%d/%Y %H:%M:%S")
    except Exception:
        formatted_time = str(timestamp)

    # GPS coordinates are top-level
    latitude = event.get("start_lat")
    longitude = event.get("start_lon")
    maps_link = (
        f"https://www.google.com/maps?q={latitude},{longitude}"
        if latitude and longitude
        else "N/A"
    )

    return {
        "driver": driver_name,
        "vehicle": vehicle_number,
        "speed": max_speed,
        "posted_speed": posted_speed,
        "overspeed": overspeed,
        "duration": duration_str,
        "severity": severity,
        "time": formatted_time,
        "location": f"{latitude:.4f}, {longitude:.4f}" if latitude and longitude else "Unknown",
        "maps_link": maps_link,
        "tier": tier,
    }


def get_repeat_offenders(events):
    """Find drivers with 2+ speeding events."""
    driver_counts = Counter(e["driver"] for e in events)
    repeats = {name: count for name, count in driver_counts.items() if count >= 2}
    return repeats


# ==============================================================================
# BUILD WORD DOCUMENT
# ==============================================================================

def logo_exists(name):
    return os.path.exists(os.path.join(LOGOS_PATH, LOGOS.get(name, "")))


def get_logo_path(name):
    return os.path.join(LOGOS_PATH, LOGOS[name])


def create_word_document(events, yesterday_date):
    """Build the speeding report Word document."""
    doc = Document()

    for section in doc.sections:
        section.top_margin = Inches(0.75)
        section.bottom_margin = Inches(0.75)
        section.left_margin = Inches(0.75)
        section.right_margin = Inches(0.75)

    # --- Logos ---
    logo_para = doc.add_paragraph()
    logo_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    logos_added = 0

    if logo_exists("butchs"):
        try:
            logo_para.add_run().add_picture(get_logo_path("butchs"), width=Inches(2.0))
            logos_added += 1
        except Exception:
            pass

    if logos_added == 0:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run("BRHAS Safety Companies")
        run.font.size = Pt(16)
        run.font.bold = True
        run.font.color.rgb = RGBColor(192, 0, 0)

    doc.add_paragraph()
    sister_para = doc.add_paragraph()
    sister_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for name in ["trucking", "permian", "hutchs", "transcend", "valor"]:
        if logo_exists(name):
            try:
                sister_para.add_run().add_picture(get_logo_path(name), width=Inches(1.2))
                sister_para.add_run("  ")
            except Exception:
                pass

    doc.add_paragraph()

    # --- Title ---
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_run = title.add_run("Daily Speeding Report")
    title_run.font.size = Pt(22)
    title_run.font.bold = True
    title_run.font.color.rgb = RGBColor(192, 0, 0)

    date_para = doc.add_paragraph()
    date_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    date_run = date_para.add_run(yesterday_date.strftime("%A, %B %d, %Y"))
    date_run.font.size = Pt(12)
    date_run.italic = True

    gen_para = doc.add_paragraph()
    gen_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    gen_run = gen_para.add_run(
        f"Generated: {datetime.now().strftime('%B %d, %Y at %H:%M:%S')}"
    )
    gen_run.font.size = Pt(9)
    gen_run.font.color.rgb = RGBColor(128, 0, 0)

    doc.add_paragraph()

    # --- Summary ---
    red_events = [e for e in events if e["tier"] == "RED"]
    yellow_events = [e for e in events if e["tier"] == "YELLOW"]
    orange_events = [e for e in events if e["tier"] == "ORANGE"]
    repeats = get_repeat_offenders(events)

    p = doc.add_paragraph()
    p.add_run(f"Total Speeding Events: {len(events)}").font.bold = True

    if red_events:
        p = doc.add_paragraph()
        run = p.add_run(f"  RED - TERMINATION (15+ over or 90+ mph): {len(red_events)}")
        run.font.color.rgb = RGBColor(255, 0, 0)
        run.font.bold = True

    if orange_events:
        p = doc.add_paragraph()
        run = p.add_run(
            f"  ORANGE - FORMAL COACHING (10-14 over): {len(orange_events)}"
        )
        run.font.color.rgb = RGBColor(255, 153, 0)
        run.font.bold = True

    if yellow_events:
        p = doc.add_paragraph()
        run = p.add_run(f"  YELLOW - MONITORING (6-9 over): {len(yellow_events)}")
        run.font.color.rgb = RGBColor(204, 102, 0)
        run.font.bold = True

    doc.add_paragraph()

    # --- Repeat Offenders ---
    if repeats:
        p = doc.add_paragraph()
        run = p.add_run("REPEAT OFFENDERS (2+ events)")
        run.font.size = Pt(14)
        run.font.bold = True
        run.font.color.rgb = RGBColor(192, 0, 0)

        for name, count in sorted(repeats.items(), key=lambda x: x[1], reverse=True):
            driver_events = [e for e in events if e["driver"] == name]
            worst = max(driver_events, key=lambda x: x["speed"])
            p = doc.add_paragraph()
            run = p.add_run(f"  {name}: {count} events")
            run.font.bold = True
            run.font.color.rgb = RGBColor(192, 0, 0)
            p.add_run(f" (worst: {worst['speed']} mph)")

        doc.add_paragraph()

    # --- Event Table ---
    if events:
        p = doc.add_paragraph()
        run = p.add_run("ALL SPEEDING EVENTS")
        run.font.size = Pt(14)
        run.font.bold = True
        run.font.color.rgb = RGBColor(192, 0, 0)

        doc.add_paragraph()

        table = doc.add_table(rows=1, cols=9)
        table.style = "Light Grid Accent 1"

        headers = ["Tier", "Driver", "Vehicle", "Max Speed", "Limit", "Over", "Duration", "Location", "Time"]
        for i, h in enumerate(headers):
            table.rows[0].cells[i].text = h
            run = table.rows[0].cells[i].paragraphs[0].runs[0]
            run.bold = True
            run.font.size = Pt(8)

        for event in events:
            cells = table.add_row().cells
            cells[0].text = event["tier"]
            tier_run = cells[0].paragraphs[0].runs[0]
            if event["tier"] == "RED":
                tier_run.font.color.rgb = RGBColor(255, 0, 0)
            elif event["tier"] == "YELLOW":
                tier_run.font.color.rgb = RGBColor(204, 102, 0)
            elif event["tier"] == "ORANGE":
                tier_run.font.color.rgb = RGBColor(255, 153, 0)
            tier_run.bold = True

            cells[1].text = event["driver"]
            cells[2].text = event["vehicle"]
            cells[3].text = f"{event['speed']} mph"
            cells[4].text = f"{event['posted_speed']} mph"
            cells[5].text = f"+{event['overspeed']} mph"
            cells[6].text = event["duration"]
            cells[7].text = event["location"]
            cells[8].text = event["time"]

            for cell in cells[1:]:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.font.size = Pt(8)
    else:
        p = doc.add_paragraph()
        run = p.add_run("No speeding events in the last 24 hours")
        run.font.color.rgb = RGBColor(0, 128, 0)
        run.font.bold = True

    # --- Footer ---
    doc.add_paragraph()
    doc.add_paragraph()
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("END OF REPORT")
    run.font.size = Pt(10)
    run.font.italic = True
    run.font.color.rgb = RGBColor(192, 0, 0)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(
        "Butch's Rat Hole & Anchor Service Inc. | HSE Department"
    )
    run.font.size = Pt(9)
    run.font.color.rgb = RGBColor(128, 0, 0)

    return doc


# ==============================================================================
# BUILD HTML EMAIL
# ==============================================================================

C_RED = "#C00000"
C_DARK = "#800000"
C_ORANGE = "#CC6600"
C_AMBER = "#FF9900"
C_GREEN = "#008000"


def _h(text):
    """HTML-escape text safely."""
    return html_escape(str(text)) if text else ""


def build_html_report(events, yesterday_date):
    """Build HTML email body mirroring the Word doc."""
    red_events = [e for e in events if e["tier"] == "RED"]
    yellow_events = [e for e in events if e["tier"] == "YELLOW"]
    orange_events = [e for e in events if e["tier"] == "ORANGE"]
    repeats = get_repeat_offenders(events)

    parts = []

    # --- Wrapper + Header ---
    parts.append(f"""<html><head><meta charset="utf-8"></head>
<body style="margin:0;padding:0;background:#f4f4f4;">
<table width="100%" cellpadding="0" cellspacing="0" style="background:#f4f4f4;">
<tr><td align="center">
<table width="700" cellpadding="0" cellspacing="0" style="background:#ffffff;border:1px solid #ddd;margin:20px auto;font-family:Arial,Helvetica,sans-serif;font-size:14px;color:#333;">

<tr><td style="background:{C_RED};padding:30px 40px;text-align:center;">
  <div style="font-size:16px;font-weight:bold;color:#ffffff;letter-spacing:1px;">BRHAS Safety Companies</div>
  <div style="font-size:28px;font-weight:bold;color:#ffffff;margin:10px 0;">DAILY SPEEDING REPORT</div>
  <div style="font-size:13px;font-style:italic;color:#ffcccc;">HSE Management Summary</div>
  <div style="font-size:12px;color:#ffffff;margin-top:8px;">Report Date: {yesterday_date.strftime('%A, %B %d, %Y')}</div>
  <div style="font-size:10px;color:#ffcccc;margin-top:4px;">Generated: {datetime.now().strftime('%B %d, %Y at %H:%M:%S')}</div>
</td></tr>""")

    # --- Summary ---
    summary = f"<b>Total Speeding Events: {len(events)}</b><br><br>"
    if red_events:
        summary += f'<div style="color:#FF0000;font-weight:bold;margin:4px 0 4px 20px;">&#128308; RED - TERMINATION (15+ over or 90+ mph): {len(red_events)}</div>'
    if orange_events:
        summary += f'<div style="color:{C_AMBER};font-weight:bold;margin:4px 0 4px 20px;">&#128993; ORANGE - FORMAL COACHING (10-14 over): {len(orange_events)}</div>'
    if yellow_events:
        summary += f'<div style="color:{C_ORANGE};font-weight:bold;margin:4px 0 4px 20px;">&#128992; YELLOW - MONITORING (6-9 over): {len(yellow_events)}</div>'
    if not events:
        summary += f'<b style="color:{C_GREEN};">&#9989; No speeding events in the last 24 hours!</b>'

    parts.append(f"""
<tr><td style="padding:25px 40px;">
  <h2 style="color:{C_RED};margin:0 0 15px 0;font-size:18px;border-bottom:2px solid {C_RED};padding-bottom:5px;">EXECUTIVE SUMMARY</h2>
  {summary}
</td></tr>""")

    # --- Repeat Offenders ---
    if repeats:
        repeat_html = ""
        for name, count in sorted(repeats.items(), key=lambda x: x[1], reverse=True):
            driver_events = [e for e in events if e["driver"] == name]
            worst = max(driver_events, key=lambda x: x["speed"])
            repeat_html += f'<div style="background:#fff5f5;border-left:4px solid {C_RED};padding:12px 15px;margin:10px 0;">'
            repeat_html += f'<b style="color:{C_RED};">{_h(name)}: {count} events</b><br>'
            repeat_html += f'Worst: {worst["speed"]} mph<br>'
            repeat_html += "</div>"

        parts.append(f"""
<tr><td style="padding:25px 40px;">
  <h2 style="color:{C_RED};margin:0 0 15px 0;font-size:18px;border-bottom:2px solid {C_RED};padding-bottom:5px;">REPEAT OFFENDERS (2+ events)</h2>
  {repeat_html}
</td></tr>""")

    # --- RED Events Detail ---
    if red_events:
        red_html = ""
        for e in red_events:
            red_html += f'<div style="background:#fff5f5;border-left:4px solid #FF0000;padding:12px 15px;margin:10px 0;">'
            red_html += f'<b style="color:#FF0000;">TERMINATION LEVEL</b><br>'
            red_html += f'<b>Driver:</b> {_h(e["driver"])} | <b>Vehicle:</b> {_h(e["vehicle"])}<br>'
            red_html += f'<b>Max Speed:</b> {e["speed"]} mph | <b>Limit:</b> {e["posted_speed"]} mph | <b>Over:</b> +{e["overspeed"]} mph<br>'
            red_html += f'<b>Duration:</b> {_h(e["duration"])} | <b>Severity:</b> {_h(e["severity"])}<br>'
            red_html += f'<b>Time:</b> {_h(e["time"])} | <b>Location:</b> {_h(e["location"])}<br>'
            if e["maps_link"] != "N/A":
                red_html += f'<a href="{_h(e["maps_link"])}">View on Google Maps</a><br>'
            red_html += "</div>"

        parts.append(f"""
<tr><td style="padding:25px 40px;border-top:3px solid #FF0000;">
  <h2 style="color:#FF0000;margin:0 0 15px 0;font-size:18px;">RED ALERTS - TERMINATION (15+ over or 90+ mph) ({len(red_events)})</h2>
  {red_html}
</td></tr>""")

    # --- ORANGE Events Detail ---
    if orange_events:
        orange_html = ""
        for e in orange_events:
            orange_html += f'<div style="background:#fffbf0;border-left:4px solid {C_AMBER};padding:12px 15px;margin:10px 0;">'
            orange_html += f'<b style="color:{C_AMBER};">FORMAL COACHING REQUIRED</b><br>'
            orange_html += f'<b>Driver:</b> {_h(e["driver"])} | <b>Vehicle:</b> {_h(e["vehicle"])}<br>'
            orange_html += f'<b>Max Speed:</b> {e["speed"]} mph | <b>Limit:</b> {e["posted_speed"]} mph | <b>Over:</b> +{e["overspeed"]} mph<br>'
            orange_html += f'<b>Duration:</b> {_h(e["duration"])} | <b>Severity:</b> {_h(e["severity"])}<br>'
            orange_html += f'<b>Time:</b> {_h(e["time"])} | <b>Location:</b> {_h(e["location"])}<br>'
            if e["maps_link"] != "N/A":
                orange_html += f'<a href="{_h(e["maps_link"])}">View on Google Maps</a><br>'
            orange_html += "</div>"

        parts.append(f"""
<tr><td style="padding:25px 40px;border-top:3px solid {C_AMBER};">
  <h2 style="color:{C_AMBER};margin:0 0 15px 0;font-size:18px;">ORANGE ALERTS - FORMAL COACHING (10-14 over) ({len(orange_events)})</h2>
  {orange_html}
</td></tr>""")

    # --- Full Event Table ---
    if events:
        table_rows = ""
        for e in events:
            if e["tier"] == "RED":
                tier_color = "#FF0000"
                bg = "#fff5f5"
            elif e["tier"] == "YELLOW":
                tier_color = C_ORANGE
                bg = "#fffbf0"
            else:
                tier_color = C_AMBER
                bg = "#ffffff"

            link_cell = ""
            if e["maps_link"] != "N/A":
                link_cell = f'<a href="{_h(e["maps_link"])}" style="font-size:11px;">Map</a>'

            table_rows += f"""<tr style="background:{bg};">
  <td style="padding:6px 8px;border:1px solid #ddd;"><b style="color:{tier_color};">{e["tier"]}</b></td>
  <td style="padding:6px 8px;border:1px solid #ddd;">{_h(e["driver"])}</td>
  <td style="padding:6px 8px;border:1px solid #ddd;">{_h(e["vehicle"])}</td>
  <td style="padding:6px 8px;border:1px solid #ddd;text-align:center;font-weight:bold;">{e["speed"]} mph</td>
  <td style="padding:6px 8px;border:1px solid #ddd;text-align:center;">{e["posted_speed"]} mph</td>
  <td style="padding:6px 8px;border:1px solid #ddd;text-align:center;font-weight:bold;">+{e["overspeed"]}</td>
  <td style="padding:6px 8px;border:1px solid #ddd;text-align:center;">{_h(e["duration"])}</td>
  <td style="padding:6px 8px;border:1px solid #ddd;font-size:12px;">{_h(e["time"])}</td>
  <td style="padding:6px 8px;border:1px solid #ddd;">{link_cell}</td>
</tr>"""

        parts.append(f"""
<tr><td style="padding:25px 40px;border-top:2px solid #ddd;">
  <h2 style="color:{C_RED};margin:0 0 15px 0;font-size:18px;border-bottom:2px solid {C_RED};padding-bottom:5px;">ALL SPEEDING EVENTS ({len(events)})</h2>
  <table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;font-size:12px;">
    <tr style="background:{C_RED};">
      <th style="padding:8px;color:#fff;border:1px solid #ddd;">Tier</th>
      <th style="padding:8px;color:#fff;border:1px solid #ddd;">Driver</th>
      <th style="padding:8px;color:#fff;border:1px solid #ddd;">Vehicle</th>
      <th style="padding:8px;color:#fff;border:1px solid #ddd;">Max Speed</th>
      <th style="padding:8px;color:#fff;border:1px solid #ddd;">Limit</th>
      <th style="padding:8px;color:#fff;border:1px solid #ddd;">Over</th>
      <th style="padding:8px;color:#fff;border:1px solid #ddd;">Duration</th>
      <th style="padding:8px;color:#fff;border:1px solid #ddd;">Time</th>
      <th style="padding:8px;color:#fff;border:1px solid #ddd;">Map</th>
    </tr>
    {table_rows}
  </table>
</td></tr>""")

    # --- Footer ---
    parts.append(f"""
<tr><td style="background:{C_DARK};padding:20px 40px;text-align:center;">
  <div style="color:#ffffff;font-size:11px;font-style:italic;">END OF REPORT</div>
  <div style="color:#ffcccc;font-size:10px;margin-top:4px;">Butch's Rat Hole &amp; Anchor Service Inc. | HSE Department</div>
</td></tr>

</table>
</td></tr></table>
</body></html>""")

    return "\n".join(parts)


# ==============================================================================
# SEND EMAIL
# ==============================================================================

def send_email_report(html_body, docx_path, yesterday_date):
    """Send report via Gmail SMTP. Fails gracefully."""
    gmail_address = os.environ.get("GMAIL_ADDRESS", "")
    gmail_app_password = os.environ.get("GMAIL_APP_PASSWORD", "")
    recipient = os.environ.get("REPORT_RECIPIENT", "")

    if not gmail_address or not gmail_app_password or not recipient:
        print(
            "  Email skipped - GMAIL_ADDRESS, GMAIL_APP_PASSWORD, or REPORT_RECIPIENT not set."
        )
        return

    subject = f"Daily Speeding Report - {yesterday_date.strftime('%B %d, %Y')}"

    try:
        msg = MIMEMultipart("mixed")
        msg["From"] = gmail_address
        msg["To"] = recipient
        msg["Subject"] = subject

        msg.attach(MIMEText(html_body, "html"))

        if os.path.exists(docx_path):
            with open(docx_path, "rb") as f:
                part = MIMEBase(
                    "application",
                    "vnd.openxmlformats-officedocument.wordprocessingml.document",
                )
                part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header(
                "Content-Disposition",
                f'attachment; filename="{os.path.basename(docx_path)}"',
            )
            msg.attach(part)

        with smtplib.SMTP("smtp.gmail.com", 587) as server:
            server.starttls()
            server.login(gmail_address, gmail_app_password)
            server.sendmail(gmail_address, recipient, msg.as_string())

        print(f"  Email sent to {recipient}")
    except Exception as e:
        print(f"  Email failed: {e}")


# ==============================================================================
# MAIN
# ==============================================================================

def main():
    today = datetime.now()
    yesterday = today - timedelta(days=1)

    print("\n" + "=" * 80)
    print("DAILY SPEEDING REPORT - AUTOMATED")
    print(f"Report for: {yesterday.strftime('%A, %B %d, %Y')}")
    print("=" * 80)
    print("\n  Thresholds (whichever is worse wins):")
    print("    RED:    15+ over posted limit OR 90+ mph (termination)")
    print("    ORANGE: 10-14 over posted limit (formal coaching)")
    print("    YELLOW: 6-9 over posted limit (monitoring)")
    print("    Repeat: 2+ events flagged\n")

    print("[1] Fetching speeding events from Motive...")
    events = get_24h_speeding_events()
    print(f"    Found {len(events)} events")

    if events:
        red = len([e for e in events if e["tier"] == "RED"])
        yellow = len([e for e in events if e["tier"] == "YELLOW"])
        orange = len([e for e in events if e["tier"] == "ORANGE"])
        repeats = get_repeat_offenders(events)
        print(f"    RED: {red} | YELLOW: {yellow} | ORANGE: {orange}")
        if repeats:
            print(
                f"    Repeat offenders: {', '.join(f'{n} ({c}x)' for n, c in repeats.items())}"
            )

    print("\n[2] Creating Word document...")
    doc = create_word_document(events, yesterday)

    date_str = yesterday.strftime("%Y-%m-%d")
    output_file = f"DailySpeedingReport_{date_str}.docx"
    doc.save(output_file)
    print(f"    Saved: {output_file}")

    print("\n[3] Building HTML email...")
    html_body = build_html_report(events, yesterday)

    print("[4] Sending email...")
    send_email_report(html_body, output_file, yesterday)

    print("\n" + "=" * 80)
    print("COMPLETE")
    print("=" * 80 + "\n")


if __name__ == "__main__":
    main()
