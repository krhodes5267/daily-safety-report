#!/usr/bin/env python3
"""
DISTRIBUTE SAFETY REP EMAILS
=============================
Reads the events JSON exported by daily_speeding_report.py and sends
individual emails to each safety rep with ONLY their division/yard tables.

Runs as a workflow step right after the main speeding report.
"""

import json
import os
import sys
import smtplib
from datetime import datetime, timedelta, timezone
from html import escape as html_escape
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

try:
    from zoneinfo import ZoneInfo
    CENTRAL_TZ = ZoneInfo("America/Chicago")
except Exception:
    CENTRAL_TZ = timezone(timedelta(hours=-6))

# ==============================================================================
# SAFETY REP -> EMAIL & DIVISION MAPPING
# Each rep gets ONE email with all their division/yard tables combined.
# ==============================================================================

SAFETY_REP_MAP = [
    {
        "name": "John Snodgrass",
        "emails": ["jsnodgrass@texasrigmove.com"],
        "divisions": [
            ("Rathole", "Midland"),
            ("Rathole", "Levelland"),
            ("Rathole", "Barstow"),
            ("Transcend Drilling", ""),
            ("Fencing", ""),
            ("Anchors", ""),
            ("Environmental", ""),
            ("Valor Energy Services", ""),
        ],
    },
    {
        "name": "Hancock & Salazar",
        "emails": ["mhancock@brhas.com", "msalazar@brhas.com"],
        "divisions": [
            ("Casing", "Midland"),
            ("Casing", "San Angelo"),
        ],
    },
    {
        "name": "Justin Conrad",
        "emails": ["jconrad@brhas.com"],
        "divisions": [
            ("Casing", "Bryan"),
        ],
    },
    {
        "name": "James Barnett",
        "emails": ["jbarnett@brhas.com"],
        "divisions": [
            ("Casing", "Kilgore"),
        ],
    },
    {
        "name": "Allen Batts",
        "emails": ["abatts@brhas.com"],
        "divisions": [
            ("Casing", "Hobbs"),
        ],
    },
    {
        "name": "Joey Speyrer",
        "emails": ["jspeyrer@brhas.com"],
        "divisions": [
            ("Casing", "Jourdanton"),
            ("Casing", "Laredo"),
        ],
    },
    {
        "name": "Jose Romero",
        "emails": ["jose.romero@brhas.com"],
        "divisions": [
            ("Poly Pipe", ""),
            ("Pit Lining", ""),
            ("Construction", ""),
        ],
    },
    {
        "name": "Sean Fry",
        "emails": ["sean.fry@brhas.com"],
        "divisions": [
            ("Rathole", "Ohio"),
            ("Rathole", "Pennsylvania"),
        ],
    },
    {
        "name": "Wes Franklin",
        "emails": ["wes@texasrigmove.com"],
        "divisions": [
            ("Rathole", "Midland"),
            ("Rathole", "Levelland"),
            ("Rathole", "Barstow"),
            ("Rathole", "Oklahoma"),
            ("Rathole", "North Dakota"),
            ("Fencing", ""),
            ("Anchors", ""),
            ("Environmental", ""),
            ("Valor Energy Services", ""),
        ],
    },
    {
        "name": "Leean Benevides",
        "emails": ["leean.benavides@brhas.com"],
        "divisions": [
            ("Rathole", "Jourdanton"),
        ],
    },
    {
        "name": "Bernard Bradley",
        "emails": ["bbradley@brhas.com"],
        "divisions": [
            ("Butch's Trucking", ""),
        ],
    },
    {
        "name": "Charley Langwell",
        "emails": ["clangwell@brhas.com"],
        "divisions": [
            ("Transcend Drilling", ""),
        ],
    },
    {
        "name": "Kelly Rhodes",
        "emails": ["krhodes@brhas.com"],
        "divisions": [
            ("Downhole Tools", ""),
        ],
    },
]

# Divisions that use yard breakdown in headers
YARD_DIVISIONS = {"Rathole", "Casing"}


# ==============================================================================
# HTML HELPERS
# ==============================================================================

C_RED = "#C00000"
C_AMBER = "#FF8C00"
C_YELLOW_DARK = "#CC9900"


def _h(text):
    """HTML-escape text safely."""
    return html_escape(str(text)) if text else ""


def _tier_colors(tier):
    """Return (text_color, bg_color) for a tier."""
    if tier == "RED":
        return "#FF0000", "#FFE0E0"
    elif tier == "ORANGE":
        return C_AMBER, "#FFF0E0"
    else:
        return C_YELLOW_DARK, "#FFFFF0"


def _build_event_table_html(events):
    """Build an HTML table for a list of events."""
    rows = ""
    for e in events:
        tc, bg = _tier_colors(e["tier"])
        map_cell = f'<a href="{_h(e["maps_link"])}" style="font-size:11px;">Map</a>' if e.get("maps_link") else ""
        rows += f"""<tr style="background:{bg};">
  <td style="padding:5px 6px;border:1px solid #ddd;"><b style="color:{tc};">{e["tier"]}</b></td>
  <td style="padding:5px 6px;border:1px solid #ddd;">{_h(e["driver"])}</td>
  <td style="padding:5px 6px;border:1px solid #ddd;">{_h(e["vehicle"])}</td>
  <td style="padding:5px 6px;border:1px solid #ddd;text-align:center;font-weight:bold;">{e["speed"]}</td>
  <td style="padding:5px 6px;border:1px solid #ddd;text-align:center;">{e["posted_speed"]}</td>
  <td style="padding:5px 6px;border:1px solid #ddd;text-align:center;font-weight:bold;color:{tc};">+{e["overspeed"]}</td>
  <td style="padding:5px 6px;border:1px solid #ddd;text-align:center;">{_h(e["duration"])}</td>
  <td style="padding:5px 6px;border:1px solid #ddd;font-size:11px;">{_h(e["time"])}</td>
  <td style="padding:5px 6px;border:1px solid #ddd;">{map_cell}</td>
</tr>"""

    return f"""<table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;font-size:12px;font-family:Calibri,Arial,Helvetica,sans-serif;">
  <tr style="background:{C_RED};">
    <th style="padding:6px;color:#fff;border:1px solid #ddd;">Tier</th>
    <th style="padding:6px;color:#fff;border:1px solid #ddd;">Driver</th>
    <th style="padding:6px;color:#fff;border:1px solid #ddd;">Vehicle</th>
    <th style="padding:6px;color:#fff;border:1px solid #ddd;">Speed</th>
    <th style="padding:6px;color:#fff;border:1px solid #ddd;">Limit</th>
    <th style="padding:6px;color:#fff;border:1px solid #ddd;">Over</th>
    <th style="padding:6px;color:#fff;border:1px solid #ddd;">Dur.</th>
    <th style="padding:6px;color:#fff;border:1px solid #ddd;">Time</th>
    <th style="padding:6px;color:#fff;border:1px solid #ddd;">Map</th>
  </tr>
  {rows}
</table>"""


def _build_section_label(division, yard):
    """Build the section header text for a division/yard."""
    if division in YARD_DIVISIONS and yard:
        return f"{division} - {yard} Yard"
    return division


def _build_rep_email_html(sections):
    """Build the full HTML email body for a safety rep.

    sections: list of (label, events_list) tuples
    Just the tables, no fluff.
    """
    parts = []

    for label, events in sections:
        events_sorted = sorted(events, key=lambda x: x["overspeed"], reverse=True)
        count = len(events_sorted)
        parts.append(
            f'<h3 style="color:{C_RED};font-family:Calibri,Arial,Helvetica,sans-serif;'
            f'margin:20px 0 8px 0;font-size:16px;">'
            f'{_h(label)} &mdash; {count} event{"s" if count != 1 else ""}</h3>'
        )
        parts.append(_build_event_table_html(events_sorted))

    body = "\n".join(parts)

    return f"""<html><head><meta charset="utf-8"></head>
<body style="margin:0;padding:20px;background:#ffffff;font-family:Calibri,Arial,Helvetica,sans-serif;">
{body}
</body></html>"""


# ==============================================================================
# ERROR EMAIL
# ==============================================================================

def _send_error_email(error_msg, report_date):
    """Send error notification to Kelly."""
    gmail_address = os.environ.get("GMAIL_ADDRESS", "")
    gmail_app_password = os.environ.get("GMAIL_APP_PASSWORD", "")

    if not gmail_address or not gmail_app_password:
        print(f"  Cannot send error email (no Gmail creds): {error_msg}")
        return

    try:
        msg = MIMEMultipart("alternative")
        msg["From"] = gmail_address
        msg["To"] = "krhodes@brhas.com"
        msg["Subject"] = f"Speeding Report Distribution ERROR - {report_date.strftime('%B %d, %Y')}"
        msg.attach(MIMEText(
            f"<html><body><p style='color:red;font-weight:bold;'>{_h(error_msg)}</p></body></html>",
            "html",
        ))

        with smtplib.SMTP("smtp.gmail.com", 587) as server:
            server.starttls()
            server.login(gmail_address, gmail_app_password)
            server.sendmail(gmail_address, ["krhodes@brhas.com"], msg.as_string())

        print(f"  Error email sent to krhodes@brhas.com")
    except Exception as e2:
        print(f"  Failed to send error email: {e2}")


# ==============================================================================
# MAIN
# ==============================================================================

def main():
    today = datetime.now(timezone.utc).astimezone(CENTRAL_TZ)
    yesterday = today - timedelta(days=1)
    date_str = yesterday.strftime("%Y-%m-%d")

    json_file = f"speeding_events_{date_str}.json"
    if not os.path.exists(json_file):
        print(f"ERROR: Events file not found: {json_file}")
        _send_error_email(f"Distribution failed: {json_file} not found", yesterday)
        sys.exit(1)

    print(f"\n{'='*70}")
    print("SAFETY REP EMAIL DISTRIBUTION")
    print(f"Report date: {yesterday.strftime('%A, %B %d, %Y')}")
    print(f"{'='*70}\n")

    # Load events
    with open(json_file, "r") as f:
        events = json.load(f)

    print(f"Loaded {len(events)} events from {json_file}")

    if not events:
        print("No events — nothing to distribute. Done.")
        return

    # Gmail config
    gmail_address = os.environ.get("GMAIL_ADDRESS", "")
    gmail_app_password = os.environ.get("GMAIL_APP_PASSWORD", "")

    if not gmail_address or not gmail_app_password:
        print("ERROR: GMAIL_ADDRESS or GMAIL_APP_PASSWORD not set.")
        sys.exit(1)

    report_date_str = yesterday.strftime("%B %d, %Y")
    subject = f"Daily Speeding Report - {report_date_str}"

    sent_count = 0
    skipped_count = 0
    fail_count = 0

    for rep in SAFETY_REP_MAP:
        rep_name = rep["name"]
        rep_emails = rep["emails"]
        rep_divisions = rep["divisions"]

        # Collect events for this rep's divisions
        sections = []
        for div, yard in rep_divisions:
            matching = [e for e in events if e["division"] == div and e["yard"] == yard]
            if matching:
                label = _build_section_label(div, yard)
                sections.append((label, matching))

        if not sections:
            print(f"  SKIP: {rep_name} — no violations in {len(rep_divisions)} division(s)")
            skipped_count += 1
            continue

        # Build email — just the tables
        total_events = sum(len(evts) for _, evts in sections)
        html_body = _build_rep_email_html(sections)

        # Send
        try:
            msg = MIMEMultipart("alternative")
            msg["From"] = gmail_address
            msg["To"] = ", ".join(rep_emails)
            msg["Subject"] = subject
            msg.attach(MIMEText(html_body, "html"))

            with smtplib.SMTP("smtp.gmail.com", 587) as server:
                server.starttls()
                server.login(gmail_address, gmail_app_password)
                server.sendmail(gmail_address, rep_emails, msg.as_string())

            section_labels = [label for label, _ in sections]
            print(f"  SENT: {rep_name} ({', '.join(rep_emails)}) — {total_events} events across {', '.join(section_labels)}")
            sent_count += 1

        except Exception as e:
            print(f"  FAIL: {rep_name} ({', '.join(rep_emails)}) — {e}")
            _send_error_email(f"Failed to send to {rep_name}: {e}", yesterday)
            fail_count += 1

    print(f"\n{'='*70}")
    print(f"DISTRIBUTION COMPLETE: {sent_count} sent, {skipped_count} skipped (no data), {fail_count} failed")
    print(f"{'='*70}\n")


if __name__ == "__main__":
    main()
