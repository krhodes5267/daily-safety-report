#!/usr/bin/env python3
"""
DAILY CASING DASHBOARD SUMMARY EMAIL
======================================
Reads all JSON data files from output/ and sends a single summary email
with key metrics and a link to the full dashboard on GitHub Pages.

Replaces the 4 individual daily emails (camera, speeding, KPA, device status)
with one consolidated view. Tiered distribution emails (safety reps, dispatchers,
managers) still go out from the individual scripts.
"""

import json
import os
import smtplib
import sys
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from html import escape as html_escape


def load_json(filename):
    """Load a JSON file from output/ directory. Returns empty dict on failure."""
    path = os.path.join("output", filename)
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError) as e:
        print(f"  Warning: Could not load {path}: {e}")
        return {}


def build_summary_email(dashboard_url):
    """Build HTML summary email from all JSON data files."""
    camera = load_json("camera_events.json")
    speeding = load_json("speeding_events.json")
    kpa = load_json("kpa_data.json")
    devices = load_json("device_status.json")
    cas = load_json("corrective_actions.json")
    ytd = load_json("ytd_stats.json")
    training = load_json("training_compliance.json")

    # Extract key metrics
    camera_total = camera.get("total_events", 0)
    camera_red = camera.get("summary", {}).get("red", 0)

    # Speeding: filter to Casing only
    speeding_events = speeding.get("events", [])
    casing_speeding = [e for e in speeding_events if e.get("division") == "Casing"]
    speeding_total = len(casing_speeding)
    speeding_red = sum(1 for e in casing_speeding if e.get("tier") == "RED")

    obs = kpa.get("observations", {})
    obs_total = obs.get("total", 0)
    obs_by_type = obs.get("by_type", {})
    near_misses = obs_by_type.get("Near Miss", 0)
    incidents = kpa.get("incidents", [])
    incident_count = len(incidents)

    device_total = devices.get("total_issues", 0)
    device_escalations = devices.get("summary", {}).get("escalations", 0)

    ca_overdue = cas.get("summary", {}).get("total_overdue", 0)
    ca_open = cas.get("summary", {}).get("total_open", 0)

    ytd_trir = ytd.get("ytd_trir") or "N/A"
    ytd_dart = ytd.get("ytd_dart") or "N/A"
    ytd_hours = ytd.get("ytd_man_hours", 0) or 0
    days_recordable = ytd.get("days_since_recordable") or "N/A"
    days_lti = ytd.get("days_since_lti") or "N/A"

    training_pct_raw = training.get("overall_pct")
    try:
        training_pct = float(training_pct_raw) if training_pct_raw is not None else None
    except (ValueError, TypeError):
        training_pct = None

    report_date = camera.get("report_date") or speeding.get("report_date") or "Unknown"

    # Format date for display
    try:
        date_display = datetime.strptime(report_date, "%Y-%m-%d").strftime("%B %d, %Y")
    except (ValueError, TypeError):
        date_display = "Unknown"

    # Red flags
    red_flags = []
    if incident_count > 0:
        red_flags.append(f"{incident_count} incident(s) reported")
    if camera_red > 0:
        red_flags.append(f"{camera_red} RED camera event(s)")
    if speeding_red > 0:
        red_flags.append(f"{speeding_red} RED speeding event(s)")
    if near_misses > 0:
        red_flags.append(f"{near_misses} near miss(es)")

    # Repeat offenders
    camera_repeats = camera.get("repeat_offenders", {})
    repeats_html = ""
    if camera_repeats:
        items = ", ".join(f"{html_escape(str(name))} ({count}x)" for name, count in camera_repeats.items())
        repeats_html = f'<p style="color:#A73333;font-size:13px;margin:4px 0 0 0;">Repeat offenders: {items}</p>'

    # Build HTML
    html = f"""<!DOCTYPE html>
<html>
<head><meta charset="utf-8"></head>
<body style="margin:0;padding:0;font-family:Calibri,Arial,sans-serif;background:#f4f4f4;">
<table width="100%" cellpadding="0" cellspacing="0" style="max-width:680px;margin:0 auto;background:#ffffff;">

<!-- Header -->
<tr><td style="background:#1a1a1a;padding:20px 24px;">
  <h1 style="color:#ffffff;font-size:22px;margin:0;">Casing Safety Dashboard</h1>
  <p style="color:#aaaaaa;font-size:14px;margin:4px 0 0 0;">{date_display}</p>
</td></tr>

<!-- Red Flags -->
{"" if not red_flags else '''
<tr><td style="background:#FFF3F3;padding:14px 24px;border-left:4px solid #A73333;">
  <p style="color:#A73333;font-weight:bold;font-size:15px;margin:0;">ATTENTION REQUIRED</p>
  <ul style="color:#333;font-size:14px;margin:6px 0 0 0;padding-left:20px;">
''' + "".join(f'<li>{flag}</li>' for flag in red_flags) + '''
  </ul>''' + repeats_html + '''
</td></tr>
'''}

<!-- Quick Stats Grid -->
<tr><td style="padding:20px 24px;">
  <table width="100%" cellpadding="0" cellspacing="0">
    <tr>
      <td width="25%" style="text-align:center;padding:12px 4px;border:1px solid #eee;">
        <div style="font-size:28px;font-weight:bold;color:#1a1a1a;">{camera_total}</div>
        <div style="font-size:12px;color:#666;">Camera Events</div>
      </td>
      <td width="25%" style="text-align:center;padding:12px 4px;border:1px solid #eee;">
        <div style="font-size:28px;font-weight:bold;color:#1a1a1a;">{speeding_total}</div>
        <div style="font-size:12px;color:#666;">Speeding (Casing)</div>
      </td>
      <td width="25%" style="text-align:center;padding:12px 4px;border:1px solid #eee;">
        <div style="font-size:28px;font-weight:bold;color:#1a1a1a;">{obs_total}</div>
        <div style="font-size:12px;color:#666;">Observations</div>
      </td>
      <td width="25%" style="text-align:center;padding:12px 4px;border:1px solid #eee;">
        <div style="font-size:28px;font-weight:bold;color:#666666;">{incident_count}</div>
        <div style="font-size:12px;color:#666;">Incidents</div>
      </td>
    </tr>
    <tr>
      <td width="25%" style="text-align:center;padding:12px 4px;border:1px solid #eee;">
        <div style="font-size:28px;font-weight:bold;color:#A73333;">{ca_overdue}</div>
        <div style="font-size:12px;color:#666;">Overdue CAs</div>
      </td>
      <td width="25%" style="text-align:center;padding:12px 4px;border:1px solid #eee;">
        <div style="font-size:28px;font-weight:bold;color:#1a1a1a;">{device_total}</div>
        <div style="font-size:12px;color:#666;">Device Issues</div>
      </td>
      <td width="25%" style="text-align:center;padding:12px 4px;border:1px solid #eee;">
        <div style="font-size:28px;font-weight:bold;color:{"#A73333" if training_pct is not None and training_pct < 80 else "#1a1a1a"};">{f"{training_pct:.1f}%" if training_pct is not None else "N/A"}</div>
        <div style="font-size:12px;color:#666;">Training</div>
      </td>
      <td width="25%" style="text-align:center;padding:12px 4px;border:1px solid #eee;">
        <div style="font-size:28px;font-weight:bold;color:#1a1a1a;">{near_misses}</div>
        <div style="font-size:12px;color:#666;">Near Misses</div>
      </td>
    </tr>
  </table>
</td></tr>

<!-- YTD Metrics -->
<tr><td style="padding:0 24px 16px;">
  <table width="100%" cellpadding="0" cellspacing="0" style="background:#f8f8f8;border:1px solid #e0e0e0;">
    <tr>
      <td style="padding:10px 16px;text-align:center;border-right:1px solid #e0e0e0;">
        <div style="font-size:11px;color:#666;">YTD TRIR</div>
        <div style="font-size:20px;font-weight:bold;">{ytd_trir}</div>
      </td>
      <td style="padding:10px 16px;text-align:center;border-right:1px solid #e0e0e0;">
        <div style="font-size:11px;color:#666;">YTD DART</div>
        <div style="font-size:20px;font-weight:bold;">{ytd_dart}</div>
      </td>
      <td style="padding:10px 16px;text-align:center;border-right:1px solid #e0e0e0;">
        <div style="font-size:11px;color:#666;">Man-Hours</div>
        <div style="font-size:20px;font-weight:bold;">{ytd_hours:,}</div>
      </td>
      <td style="padding:10px 16px;text-align:center;border-right:1px solid #e0e0e0;">
        <div style="font-size:11px;color:#666;">Days Since Recordable</div>
        <div style="font-size:20px;font-weight:bold;">{days_recordable}</div>
      </td>
      <td style="padding:10px 16px;text-align:center;">
        <div style="font-size:11px;color:#666;">Days Since LTI</div>
        <div style="font-size:20px;font-weight:bold;">{days_lti}</div>
      </td>
    </tr>
  </table>
</td></tr>

<!-- By Yard Summary -->
<tr><td style="padding:0 24px 16px;">
  <p style="font-size:14px;font-weight:bold;color:#333;margin:0 0 8px 0;">By Yard</p>
  <table width="100%" cellpadding="6" cellspacing="0" style="font-size:13px;border-collapse:collapse;">
    <tr style="background:#1a1a1a;color:#fff;">
      <th style="text-align:left;padding:6px 8px;">Yard</th>
      <th style="text-align:center;padding:6px 8px;">Camera</th>
      <th style="text-align:center;padding:6px 8px;">Speeding</th>
      <th style="text-align:center;padding:6px 8px;">Devices</th>
    </tr>"""

    yards = ["Midland", "Bryan", "Kilgore", "Hobbs", "Jourdanton", "Laredo", "San Angelo"]
    camera_by_yard = camera.get("by_yard", {})
    casing_speed_by_yard = {}
    for e in casing_speeding:
        y = e.get("yard", "Unknown")
        casing_speed_by_yard[y] = casing_speed_by_yard.get(y, 0) + 1
    device_by_yard = devices.get("by_yard", {})

    for i, yard in enumerate(yards):
        bg = "#f9f9f9" if i % 2 == 0 else "#ffffff"
        cam = camera_by_yard.get(yard, 0)
        spd = casing_speed_by_yard.get(yard, 0)
        dev = device_by_yard.get(yard, 0)
        html += f"""
    <tr style="background:{bg};">
      <td style="padding:6px 8px;">{yard}</td>
      <td style="text-align:center;padding:6px 8px;">{cam}</td>
      <td style="text-align:center;padding:6px 8px;">{spd}</td>
      <td style="text-align:center;padding:6px 8px;">{dev}</td>
    </tr>"""

    html += f"""
  </table>
</td></tr>

<!-- Dashboard Button -->
<tr><td style="padding:20px 24px;text-align:center;">
  <a href="{html_escape(dashboard_url, quote=True)}" style="display:inline-block;background:#A73333;color:#ffffff;font-size:16px;font-weight:bold;padding:14px 40px;text-decoration:none;border-radius:4px;">
    OPEN FULL DASHBOARD
  </a>
  <p style="font-size:12px;color:#999;margin:10px 0 0 0;">
    Full interactive dashboard with trends, driver analysis, training details, and more.
  </p>
</td></tr>

<!-- Footer -->
<tr><td style="background:#f4f4f4;padding:16px 24px;text-align:center;">
  <p style="font-size:11px;color:#999;margin:0;">
    Butch's Rat Hole and Anchor Service, Inc. -- Casing Division Safety Dashboard<br>
    Auto-generated {datetime.now().strftime("%B %d, %Y at %I:%M %p")}
  </p>
</td></tr>

</table>
</body>
</html>"""

    return html, date_display


def send_summary_email(html, date_display):
    """Send summary email via Gmail SMTP."""
    gmail_address = os.environ.get("GMAIL_ADDRESS", "")
    gmail_app_password = os.environ.get("GMAIL_APP_PASSWORD", "")
    recipient = os.environ.get("REPORT_RECIPIENT", "")

    if not gmail_address or not gmail_app_password or not recipient:
        print("  Email skipped -- GMAIL_ADDRESS, GMAIL_APP_PASSWORD, or REPORT_RECIPIENT not set.")
        return False

    subject = f"Casing Safety Dashboard -- {date_display}"

    try:
        msg = MIMEMultipart("alternative")
        msg["From"] = gmail_address
        msg["To"] = recipient
        msg["Subject"] = subject
        msg.attach(MIMEText(html, "html"))

        with smtplib.SMTP("smtp.gmail.com", 587, timeout=30) as server:
            server.starttls()
            server.login(gmail_address, gmail_app_password)
            server.sendmail(gmail_address, recipient, msg.as_string())

        print(f"  Summary email sent to {recipient}")
        return True
    except Exception as e:
        print(f"  Email failed: {e}")
        return False


def main():
    dashboard_url = os.environ.get(
        "DASHBOARD_URL",
        "https://krhodes5267.github.io/daily-safety-report/"
    )

    print("=" * 60)
    print("CASING SAFETY DASHBOARD -- SUMMARY EMAIL")
    print("=" * 60)

    html, date_display = build_summary_email(dashboard_url)
    print(f"\n  Report date: {date_display}")
    print(f"  Dashboard URL: {dashboard_url}")

    if "--dry-run" in sys.argv:
        print("\n  Dry run -- email not sent.")
        # Write HTML to file for preview
        with open("output/dashboard_email_preview.html", "w", encoding="utf-8") as f:
            f.write(html)
        print("  Preview saved to output/dashboard_email_preview.html")
    else:
        send_summary_email(html, date_display)

    print("\nDONE")


if __name__ == "__main__":
    main()
