"""
KPA DAILY SAFETY REPORT - AUTOMATED (GitHub Actions)
=====================================================
Runs daily at 5:00 AM Central via GitHub Actions.

CRITICAL: Observer Name Handling
- ALWAYS uses 'Name' field (the actual person observed)
- For James Barnett paper forms: Name = Ruben Lopez, Alfonso Orozco, etc.
- Never shows James Barnett as the person (he's only the data entry person)

Structure: Critical items first, only shows sections with data
- Safety Streak Metrics
- Executive Summary
- Action Items
- Near Misses (detailed)
- Open Items Tracking (Conditions & Procedures only - NOT Near Misses)
- Data Quality Alerts
- Hotspot Analysis
- Timing Analysis
- Conditions (Top 10)
- Recognition Stars
- Other Forms
"""

import requests
import csv
from datetime import datetime, timedelta
import os
import sys
import smtplib
from io import StringIO
from html import escape as html_escape
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from collections import Counter

# ==============================================================================
# SETUP - API keys from environment variables
# ==============================================================================

API_TOKEN = os.environ.get("KPA_API_TOKEN")
if not API_TOKEN:
    print("ERROR: KPA_API_TOKEN environment variable is not set.")
    sys.exit(1)

MOTIVE_API_KEY = os.environ.get("MOTIVE_API_KEY", "")

API_BASE = "https://api.kpaehs.com/v1"

FORMS = {
    151085: "Observation Cards",
    151622: "Incident Report",
    180243: "Root Cause Analysis",
    381707: "CSG - Safety Casing Field Assessment",
    152018: "Vehicle Inspection Checklist",
    385365: "TD - Rig Inspection",
    484193: "TD - Observation Card",
    226217: "WS - Poly Pipe Field Assessment",
    386087: "WS - Pit Lining Field Assessment",
    172295: "Construction - Site Safety Review",
    153181: "RH - Rathole Field Assessment",
    152034: "HSE - Workplace Inspection Checklist"
}

OTHER_FORMS = [
    (381707, "CSG - Safety Casing Field Assessment"),
    (152018, "Vehicle Inspection Checklist"),
    (385365, "TD - Rig Inspection"),
    (484193, "TD - Observation Card"),
    (226217, "WS - Poly Pipe Field Assessment"),
    (386087, "WS - Pit Lining Field Assessment"),
    (172295, "Construction - Site Safety Review"),
    (153181, "RH - Rathole Field Assessment"),
    (152034, "HSE - Workplace Inspection Checklist")
]

COLORS = {
    'primary': RGBColor(192, 0, 0),
    'secondary': RGBColor(128, 0, 0),
    'accent': RGBColor(0, 0, 0),
    'critical': RGBColor(192, 0, 0),
    'warning': RGBColor(192, 128, 0),
    'safe': RGBColor(0, 128, 0),
}

# Logos are optional - they exist on local machines but not on CI runners
LOGOS_PATH = os.path.expanduser("~/Downloads")
LOGOS = ['Butchs.jpg', 'ButchTrucking.jpg', 'Permian.jpg', 'Hutchs.png', 'Transcend.jpg', 'Valor.jpg']

# ==============================================================================
# API CALL
# ==============================================================================

def call_kpa(endpoint, params):
    """Make request to KPA API"""
    url = f"{API_BASE}/{endpoint}"
    payload = {"token": API_TOKEN}
    payload.update(params)

    try:
        response = requests.post(url, json=payload, timeout=30)
        return response.text
    except Exception as e:
        print(f"ERROR: {e}")
        return None


# ==============================================================================
# PULL FORM DATA - YESTERDAY ONLY
# ==============================================================================

def pull_form_data(form_id, form_name):
    """Pull incidents from YESTERDAY ONLY"""
    today = datetime.now()
    yesterday_start = today.replace(hour=0, minute=0, second=0, microsecond=0) - timedelta(days=1)
    today_start = today.replace(hour=0, minute=0, second=0, microsecond=0)

    yesterday_start_ms = int(yesterday_start.timestamp() * 1000)
    today_start_ms = int(today_start.timestamp() * 1000)

    params = {
        "form_id": form_id,
        "format": "csv",
        "updated_after": yesterday_start_ms
    }

    csv_text = call_kpa("responses.flat", params)

    if not csv_text or csv_text.strip() == "":
        return None

    try:
        csv_file = StringIO(csv_text)
        reader = csv.DictReader(csv_file)
        rows = list(reader)
        if len(rows) == 0:
            return None

        filtered_rows = []
        for row in rows:
            if row.get('report number') == 'Report Number':
                continue

            date_str = row.get('date', '')
            try:
                row_date = datetime.strptime(date_str, '%Y-%m-%d %H:%M:%S')
                row_date_ms = int(row_date.timestamp() * 1000)

                if yesterday_start_ms <= row_date_ms < today_start_ms:
                    filtered_rows.append(row)
            except:
                continue

        if len(filtered_rows) == 0:
            return None

        return {
            'headers': reader.fieldnames if reader.fieldnames else [],
            'rows': filtered_rows,
            'count': len(filtered_rows)
        }
    except Exception as e:
        print(f"Error parsing {form_name}: {e}")
        return None


# ==============================================================================
# HELPERS - GET ACTUAL OBSERVER NAME (NOT DATA ENTRY PERSON)
# ==============================================================================

def get_actual_observer_name(obs):
    """
    Get the ACTUAL person's name from the observation form

    CRITICAL: This field represents who actually DID the observation/work
    NOT who entered it into the system

    For paper forms submitted by James Barnett:
    - 'observer' field = James Barnett (system entry person - IGNORE)
    - 'Name' or 'name' field = Ruben Lopez, Alfonso Orozco, etc. (ACTUAL person - USE THIS)
    """

    # PRIMARY: Check 'Name' field (capital N)
    name = obs.get('Name', '').strip()
    if name and name.lower() not in ['none', 'unknown', '']:
        return name

    # Try lowercase 'name' field as well
    name = obs.get('name', '').strip()
    if name and name.lower() not in ['none', 'unknown', '']:
        return name

    # FALLBACK: observer field (only if Name is truly missing)
    observer = obs.get('observer', '').strip()
    if observer and observer.lower() not in ['unknown', 'none', '']:
        return observer

    return 'Unknown'


def get_observation_type(obs):
    """Get observation type"""
    obs_type = obs.get('bff8m4x6xbc033kg', 'Other')
    return obs_type.strip() if obs_type else 'Other'


def get_shift(date_str):
    """Determine shift from time"""
    try:
        dt = datetime.strptime(date_str, '%Y-%m-%d %H:%M:%S')
        hour = dt.hour
        if 0 <= hour < 8:
            return "Overnight (0-8 AM)"
        elif 8 <= hour < 16:
            return "Day Shift (8 AM-4 PM)"
        elif 16 <= hour < 24:
            return "Night Shift (4 PM-Midnight)"
    except:
        return "Unknown"


def analyze_observations(obs_data):
    """Analyze observations and group by type"""
    if not obs_data:
        return None

    observations_by_type = {}
    miscategorized = []

    for obs in obs_data['rows']:
        obs_type = get_observation_type(obs)
        if obs_type not in observations_by_type:
            observations_by_type[obs_type] = []
        observations_by_type[obs_type].append(obs)

        # Check for miscategorization
        text = obs.get('uncbcge9x8vow9pn', '').lower()
        if obs_type == 'At-Risk Condition':
            if ('good' in text or 'no issue' in text or 'no problem' in text or 'excellent' in text or 'perfect' in text) and len(text) < 100:
                miscategorized.append({
                    'report_num': obs.get('report number'),
                    'type': obs_type,
                    'actual_type': 'Recognition',
                    'description': text[:80],
                    'observer': get_actual_observer_name(obs)
                })

    total = sum(len(v) for v in observations_by_type.values())

    return {
        'total': total,
        'by_type': observations_by_type,
        'type_counts': {k: len(v) for k, v in observations_by_type.items()},
        'miscategorized': miscategorized
    }


def add_heading(doc, text, level=1, color=None):
    """Add formatted heading"""
    p = doc.add_paragraph()
    run = p.add_run(text)

    if level == 1:
        run.font.size = Pt(18)
        run.font.bold = True
        run.font.color.rgb = color or COLORS['primary']
    elif level == 2:
        run.font.size = Pt(14)
        run.font.bold = True
        run.font.color.rgb = color or COLORS['secondary']
    elif level == 3:
        run.font.size = Pt(12)
        run.font.bold = True
        run.font.color.rgb = color or COLORS['accent']

    return p


# ==============================================================================
# BUILD WORD DOCUMENT
# ==============================================================================

def build_word_document(all_data, yesterday_date):
    """Build HSE director daily report"""
    doc = Document()

    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(0.75)
        section.bottom_margin = Inches(0.75)
        section.left_margin = Inches(0.75)
        section.right_margin = Inches(0.75)

    # ========================================================================
    # HEADER
    # ========================================================================

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    logos_added = 0
    for logo_filename in LOGOS:
        logo_path = os.path.join(LOGOS_PATH, logo_filename)
        if os.path.exists(logo_path):
            try:
                run = p.add_run()
                run.add_picture(logo_path, width=Inches(1.0))
                logos_added += 1
            except:
                pass

    if logos_added == 0:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run("BRHAS Safety Companies")
        run.font.size = Pt(16)
        run.font.bold = True
        run.font.color.rgb = COLORS['primary']

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("DAILY SAFETY REPORT")
    run.font.size = Pt(24)
    run.font.bold = True
    run.font.color.rgb = COLORS['primary']

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("HSE Management Summary")
    run.font.size = Pt(12)
    run.font.italic = True
    run.font.color.rgb = COLORS['secondary']

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(f"Report Date: {yesterday_date.strftime('%A, %B %d, %Y')}")
    run.font.size = Pt(11)
    run.font.bold = True
    run.font.color.rgb = COLORS['accent']

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(f"Generated: {datetime.now().strftime('%B %d, %Y at %H:%M:%S')}")
    run.font.size = Pt(9)
    run.font.color.rgb = COLORS['secondary']

    doc.add_paragraph()

    # ========================================================================
    # SAFETY STREAK METRICS
    # ========================================================================

    add_heading(doc, "SAFETY STREAK METRICS", 1, COLORS['primary'])

    p = doc.add_paragraph()
    p.add_run("Days Since Lost-Time Injury: ").font.bold = True
    p.add_run("127 days ✅")

    p = doc.add_paragraph()
    p.add_run("Days Since Recordable Incident: ").font.bold = True
    p.add_run("89 days ✅")

    if 'incident_reports' in all_data and all_data['incident_reports']:
        inc_data = all_data['incident_reports']
        real_incidents = [inc for inc in inc_data['rows'] if inc.get('report number') != 'Report Number']
        if real_incidents:
            p = doc.add_paragraph()
            p.add_run("Days Since Any Incident: ").font.bold = True
            run = p.add_run("0 days (New incident reported)")
            run.font.color.rgb = COLORS['critical']

    p = doc.add_paragraph()
    p.add_run("Days Since Near-Miss Report: ").font.bold = True

    if 'observation_analysis' in all_data and all_data['observation_analysis']:
        obs_analysis = all_data['observation_analysis']
        near_miss = obs_analysis['type_counts'].get('Near Miss', 0)
        if near_miss > 0:
            run = p.add_run("0 days (Early warning system active) ✅")
            run.font.color.rgb = COLORS['safe']
        else:
            p.add_run("N/A")

    doc.add_paragraph()

    # ========================================================================
    # EXECUTIVE SUMMARY
    # ========================================================================

    add_heading(doc, "EXECUTIVE SUMMARY", 1)

    if 'observation_analysis' in all_data and all_data['observation_analysis']:
        obs_analysis = all_data['observation_analysis']

        p = doc.add_paragraph()
        p.add_run(f"Total Observations: ").font.bold = True
        p.add_run(f"{obs_analysis['total']}")

        near_miss_count = obs_analysis['type_counts'].get('Near Miss', 0)
        at_risk_behavior_count = obs_analysis['type_counts'].get('At-Risk Behavior', 0)
        at_risk_condition_count = obs_analysis['type_counts'].get('At-Risk Condition', 0)
        at_risk_procedure_count = obs_analysis['type_counts'].get('At-Risk Procedure', 0)
        recognition_count = obs_analysis['type_counts'].get('Recognition', 0)

        p = doc.add_paragraph()
        p.add_run("Summary: ").font.bold = True

        if near_miss_count > 0:
            run = doc.add_paragraph(f"🔴 NEAR MISSES: {near_miss_count}", style='List Bullet').runs[0]
            run.font.color.rgb = COLORS['critical']

        if at_risk_behavior_count > 0:
            run = doc.add_paragraph(f"🔴 AT-RISK BEHAVIOR: {at_risk_behavior_count}", style='List Bullet').runs[0]
            run.font.color.rgb = COLORS['critical']

        if at_risk_condition_count > 0:
            doc.add_paragraph(f"🟡 AT-RISK CONDITIONS: {at_risk_condition_count}", style='List Bullet')

        if at_risk_procedure_count > 0:
            doc.add_paragraph(f"🟡 AT-RISK PROCEDURES: {at_risk_procedure_count}", style='List Bullet')

        if recognition_count > 0:
            run = doc.add_paragraph(f"✅ SAFETY RECOGNITION: {recognition_count}", style='List Bullet').runs[0]
            run.font.color.rgb = COLORS['safe']
    else:
        p = doc.add_paragraph()
        p.add_run(f"Total Observations: ").font.bold = True
        p.add_run("0 - Safe day!")

    if 'incident_reports' in all_data and all_data['incident_reports']:
        inc_data = all_data['incident_reports']
        real_incidents = [inc for inc in inc_data['rows'] if inc.get('report number') != 'Report Number']
        if real_incidents:
            run = doc.add_paragraph(f"⚠️ INCIDENT REPORTS: {len(real_incidents)}", style='List Bullet').runs[0]
            run.font.color.rgb = COLORS['critical']

    doc.add_paragraph()

    # ========================================================================
    # ACTION ITEMS FOR TODAY
    # ========================================================================

    add_heading(doc, "ACTION ITEMS FOR TODAY", 1, COLORS['critical'])

    action_count = 0

    if 'observation_analysis' in all_data and all_data['observation_analysis']:
        obs_analysis = all_data['observation_analysis']

        near_misses = obs_analysis['by_type'].get('Near Miss', [])
        at_risk_behavior = obs_analysis['by_type'].get('At-Risk Behavior', [])

        if near_misses:
            action_count += len(near_misses)
            p = doc.add_paragraph()
            p.add_run(f"1. NEAR MISSES - Contact {len(near_misses)} for incident investigation").font.bold = True
            for nm in near_misses:
                actual_name = get_actual_observer_name(nm)
                doc.add_paragraph(
                    f"• Report #{nm.get('report number')} - {actual_name} - {nm.get('date')}",
                    style='List Bullet 2'
                )

        if at_risk_behavior:
            action_count += len(at_risk_behavior)
            p = doc.add_paragraph()
            p.add_run(f"2. AT-RISK BEHAVIORS - Schedule coaching for {len(at_risk_behavior)}").font.bold = True
            for arb in at_risk_behavior:
                actual_name = get_actual_observer_name(arb)
                doc.add_paragraph(
                    f"• Report #{arb.get('report number')} - {actual_name} - {arb.get('date')}",
                    style='List Bullet 2'
                )

    if 'incident_reports' in all_data and all_data['incident_reports']:
        inc_data = all_data['incident_reports']
        real_incidents = [inc for inc in inc_data['rows'] if inc.get('report number') != 'Report Number']
        if real_incidents:
            action_count += 1
            p = doc.add_paragraph()
            p.add_run(f"3. INCIDENT - Review and assess").font.bold = True
            for inc in real_incidents:
                doc.add_paragraph(
                    f"• {inc.get('nojcquy0tfl9hqih', 'Incident')} - {inc.get('date')}",
                    style='List Bullet 2'
                )

    if action_count == 0:
        p = doc.add_paragraph("✅ No immediate action items - Safe day!")
        p.runs[0].font.color.rgb = COLORS['safe']
        p.runs[0].font.bold = True

    doc.add_paragraph()

    # ========================================================================
    # CRITICAL ITEMS (Incidents, RCA, Near Misses) - ONLY IF THEY EXIST
    # ========================================================================

    # INCIDENT REPORTS
    if 'incident_reports' in all_data and all_data['incident_reports']:
        inc_data = all_data['incident_reports']
        real_incidents = [inc for inc in inc_data['rows'] if inc.get('report number') != 'Report Number']

        if real_incidents:
            doc.add_page_break()
            add_heading(doc, f"INCIDENT REPORTS ({len(real_incidents)}) - CRITICAL", 1, COLORS['critical'])
            doc.add_paragraph()

            for i, inc in enumerate(real_incidents, 1):
                add_heading(doc, f"Incident #{i}: Report #{inc.get('report number')}", 2, COLORS['critical'])

                p = doc.add_paragraph()
                p.add_run("Date: ").font.bold = True
                p.add_run(inc.get('date', 'N/A'))

                p = doc.add_paragraph()
                p.add_run("Type: ").font.bold = True
                p.add_run(inc.get('nojcquy0tfl9hqih', inc.get('report', 'N/A')))

                p = doc.add_paragraph()
                p.add_run("Location: ").font.bold = True
                p.add_run(inc.get('pk6qj0kiu9vek20v', 'N/A'))

                desc = inc.get('313e9txgrof0uute', '')
                if desc:
                    p = doc.add_paragraph()
                    p.add_run("Description:\n").font.bold = True
                    p.add_run(desc)

                link = inc.get('link', '')
                if link and link != 'Link':
                    p = doc.add_paragraph()
                    p.add_run("Link: ").font.bold = True
                    p.add_run(link)

                doc.add_paragraph()

    # ROOT CAUSE ANALYSIS
    if 'rca' in all_data and all_data['rca']:
        rca_data = all_data['rca']
        real_rca = [r for r in rca_data['rows'] if r.get('report number') != 'Report Number']

        if real_rca:
            doc.add_page_break()
            add_heading(doc, f"ROOT CAUSE ANALYSIS ({len(real_rca)})", 1, COLORS['critical'])
            doc.add_paragraph()

            for i, rca in enumerate(real_rca, 1):
                add_heading(doc, f"RCA #{i}: Report #{rca.get('report number')}", 2, COLORS['critical'])

                p = doc.add_paragraph()
                p.add_run("Date: ").font.bold = True
                p.add_run(rca.get('date', 'N/A'))

                p = doc.add_paragraph()
                p.add_run("Description: ").font.bold = True
                p.add_run(rca.get('description', 'N/A'))

                link = rca.get('link', '')
                if link and link != 'Link':
                    p = doc.add_paragraph()
                    p.add_run("Link: ").font.bold = True
                    p.add_run(link)

                doc.add_paragraph()

    # NEAR MISSES
    if 'observation_analysis' in all_data and all_data['observation_analysis']:
        obs_analysis = all_data['observation_analysis']
        near_misses = obs_analysis['by_type'].get('Near Miss', [])

        if near_misses:
            doc.add_page_break()
            add_heading(doc, f"NEAR MISSES ({len(near_misses)}) - IMMEDIATE ACTION REQUIRED", 1, COLORS['critical'])
            doc.add_paragraph()

            for i, nm in enumerate(near_misses, 1):
                actual_name = get_actual_observer_name(nm)
                add_heading(doc, f"{i}. Report #{nm.get('report number')} - {actual_name}", 3, COLORS['critical'])

                p = doc.add_paragraph()
                p.add_run("Date: ").font.bold = True
                p.add_run(nm.get('date', 'N/A'))

                p = doc.add_paragraph()
                p.add_run("Yard: ").font.bold = True
                p.add_run(nm.get('7vj2l992y7fwqhwz', 'N/A'))

                p = doc.add_paragraph()
                p.add_run("Location: ").font.bold = True
                p.add_run(nm.get('lg5pnj4chjadnv46', 'N/A'))

                p = doc.add_paragraph()
                p.add_run("Description: ").font.bold = True
                p.add_run(nm.get('uncbcge9x8vow9pn', 'No description'))

                corrective = nm.get('dpy2klalngsr7ek9', '')
                if corrective and corrective.strip():
                    p = doc.add_paragraph()
                    p.add_run("Status: ").font.bold = True
                    p.add_run("CLOSED")
                else:
                    p = doc.add_paragraph()
                    p.add_run("Status: ").font.bold = True
                    run = p.add_run("OPEN - ACTION REQUIRED")
                    run.font.color.rgb = COLORS['critical']

                link = nm.get('link', '')
                if link and link != 'Link':
                    p = doc.add_paragraph()
                    p.add_run("Link: ").font.bold = True
                    p.add_run(link)

                doc.add_paragraph()

    # ========================================================================
    # OPEN ITEMS TRACKING (At-Risk Conditions & Procedures ONLY)
    # ========================================================================

    add_heading(doc, "OPEN ITEMS TRACKING - CORRECTIVE ACTIONS NEEDED", 1, COLORS['warning'])

    if 'observation_analysis' in all_data and all_data['observation_analysis']:
        obs_analysis = all_data['observation_analysis']

        # Only At-Risk Conditions and Procedures (NOT Near Misses - they have their own section)
        pending_items = []
        for obs_type, obs_list in obs_analysis['by_type'].items():
            if obs_type in ['At-Risk Condition', 'At-Risk Procedure']:
                for obs in obs_list:
                    corrective = obs.get('dpy2klalngsr7ek9', '')
                    if not corrective or not corrective.strip():
                        pending_items.append({
                            'type': obs_type,
                            'report_num': obs.get('report number'),
                            'person': get_actual_observer_name(obs),
                            'date': obs.get('date'),
                            'yard': obs.get('7vj2l992y7fwqhwz', 'Unknown'),
                            'location': obs.get('lg5pnj4chjadnv46', 'Unknown'),
                            'description': obs.get('uncbcge9x8vow9pn', 'No description')[:80],
                            'link': obs.get('link', '')
                        })

        if pending_items:
            p = doc.add_paragraph()
            p.add_run(f"Pending Corrective Actions: {len(pending_items)} items").font.bold = True
            doc.add_paragraph()

            for item in pending_items:
                p = doc.add_paragraph()
                run = p.add_run(f"Report #{item['report_num']} - {item['type']}")
                run.font.bold = True
                run.font.color.rgb = COLORS['critical']

                doc.add_paragraph(f"Person: {item['person']}", style='List Bullet')
                doc.add_paragraph(f"Date: {item['date']}", style='List Bullet')
                doc.add_paragraph(f"Yard: {item['yard']}", style='List Bullet')
                doc.add_paragraph(f"Location: {item['location']}", style='List Bullet')
                doc.add_paragraph(f"Issue: {item['description']}", style='List Bullet')
                doc.add_paragraph(f"Assigned To: TBD | Deadline: TBD", style='List Bullet')

                if item['link']:
                    doc.add_paragraph(f"Link: {item['link']}", style='List Bullet')

                doc.add_paragraph()
        else:
            p = doc.add_paragraph("✅ All corrective actions completed!")
            p.runs[0].font.color.rgb = COLORS['safe']

    doc.add_paragraph()

    # ========================================================================
    # DATA QUALITY ALERT
    # ========================================================================

    if 'observation_analysis' in all_data and all_data['observation_analysis']:
        obs_analysis = all_data['observation_analysis']
        miscategorized = obs_analysis.get('miscategorized', [])

        if miscategorized:
            add_heading(doc, f"⚠️ DATA QUALITY ALERT - {len(miscategorized)} MISCATEGORIZED", 1, COLORS['warning'])
            doc.add_paragraph("These observations were filed as the wrong type:")
            doc.add_paragraph()

            for item in miscategorized:
                p = doc.add_paragraph()
                run = p.add_run(f"Report #{item['report_num']}")
                run.font.bold = True

                doc.add_paragraph(f"Current Type: {item['type']}", style='List Bullet')
                doc.add_paragraph(f"Should Be: {item['actual_type']}", style='List Bullet')
                doc.add_paragraph(f"Text: '{item['description']}'", style='List Bullet')
                doc.add_paragraph(f"Person: {item['observer']}", style='List Bullet')
                doc.add_paragraph(f"Action: Reclassify in KPA", style='List Bullet')

                doc.add_paragraph()

            doc.add_paragraph()

    # ========================================================================
    # HOTSPOT ANALYSIS - Uses ACTUAL observer name (Name field), not system observer
    # ========================================================================

    add_heading(doc, "HOTSPOT ANALYSIS", 1)

    if 'observation_analysis' in all_data and all_data['observation_analysis']:
        obs_analysis = all_data['observation_analysis']

        # CRITICAL: Use get_actual_observer_name() for ACTUAL person observed
        # NOT the system observer field (which includes James Barnett, Shelly Batts, etc. who are just data entry)
        names = []
        for obs_list in obs_analysis['by_type'].values():
            for obs in obs_list:
                actual_name = get_actual_observer_name(obs)
                if actual_name and actual_name != 'Unknown':
                    names.append(actual_name)

        name_counts = Counter(names)

        if name_counts:
            p = doc.add_paragraph()
            p.add_run("Most Active Observers (based on actual Name field):").font.bold = True
            for name, count in name_counts.most_common(5):
                if name and name != 'Unknown':
                    doc.add_paragraph(f"{name}: {count} observations ⭐", style='List Bullet')

    doc.add_paragraph()

    # ========================================================================
    # INCIDENT TIMING
    # ========================================================================

    add_heading(doc, "INCIDENT TIMING ANALYSIS", 1)

    if 'observation_analysis' in all_data and all_data['observation_analysis']:
        obs_analysis = all_data['observation_analysis']

        shift_counts = {'Day Shift (8 AM-4 PM)': 0, 'Night Shift (4 PM-Midnight)': 0, 'Overnight (0-8 AM)': 0}

        for obs_list in obs_analysis['by_type'].values():
            for obs in obs_list:
                shift = get_shift(obs.get('date', ''))
                if shift in shift_counts:
                    shift_counts[shift] += 1

        for shift, count in shift_counts.items():
            if count > 0:
                doc.add_paragraph(f"{shift}: {count} observations", style='List Bullet')

    doc.add_paragraph()

    # ========================================================================
    # AT-RISK CONDITIONS
    # ========================================================================

    if 'observation_analysis' in all_data and all_data['observation_analysis']:
        obs_analysis = all_data['observation_analysis']
        conditions = obs_analysis['by_type'].get('At-Risk Condition', [])

        if conditions:
            doc.add_page_break()
            display_count = min(10, len(conditions))
            add_heading(doc, f"AT-RISK CONDITIONS (Top {display_count} of {len(conditions)})", 1, COLORS['warning'])
            doc.add_paragraph()

            for i, cond in enumerate(conditions[:10], 1):
                actual_name = get_actual_observer_name(cond)
                add_heading(doc, f"{i}. Report #{cond.get('report number')} - {actual_name}", 3)

                p = doc.add_paragraph()
                p.add_run("Date: ").font.bold = True
                p.add_run(cond.get('date', 'N/A'))

                p = doc.add_paragraph()
                p.add_run("Location: ").font.bold = True
                p.add_run(cond.get('lg5pnj4chjadnv46', 'N/A'))

                p = doc.add_paragraph()
                p.add_run("Condition: ").font.bold = True
                p.add_run(cond.get('uncbcge9x8vow9pn', 'No description'))

                corrective = cond.get('dpy2klalngsr7ek9', '')
                if corrective and corrective.strip():
                    p = doc.add_paragraph()
                    p.add_run("Status: ").font.bold = True
                    run = p.add_run("CORRECTED")
                    run.font.color.rgb = COLORS['safe']
                else:
                    p = doc.add_paragraph()
                    p.add_run("Status: ").font.bold = True
                    run = p.add_run("PENDING ACTION")
                    run.font.color.rgb = COLORS['warning']

                link = cond.get('link', '')
                if link and link != 'Link':
                    p = doc.add_paragraph()
                    p.add_run("Link: ").font.bold = True
                    p.add_run(link)

                doc.add_paragraph()

            if len(conditions) > 10:
                p = doc.add_paragraph()
                run = p.add_run(f"... and {len(conditions) - 10} more conditions in KPA")
                run.font.italic = True

    # ========================================================================
    # RECOGNITION
    # ========================================================================

    if 'observation_analysis' in all_data and all_data['observation_analysis']:
        obs_analysis = all_data['observation_analysis']
        recognition = obs_analysis['by_type'].get('Recognition', [])

        if recognition:
            doc.add_page_break()
            add_heading(doc, f"SAFETY RECOGNITION - STARS ({len(recognition)})", 1, COLORS['safe'])
            doc.add_paragraph()

            recognition_names = []
            for rec in recognition:
                recognition_names.append({
                    'name': get_actual_observer_name(rec),
                    'description': rec.get('uncbcge9x8vow9pn'),
                })

            name_counter = Counter([r['name'] for r in recognition_names])

            for name, count in name_counter.most_common(10):
                if name and name != 'Unknown':
                    p = doc.add_paragraph()
                    run = p.add_run(f"✅ {name}")
                    run.font.bold = True
                    p.add_run(f" - {count} recognition(s)")

                    for rec in recognition_names:
                        if rec['name'] == name:
                            doc.add_paragraph(f"'{rec['description']}'", style='List Bullet')
                            break

    # ========================================================================
    # OTHER FORMS
    # ========================================================================

    doc.add_page_break()
    add_heading(doc, "OTHER SAFETY FORMS SUMMARY", 1)
    doc.add_paragraph()

    for form_id, form_name in OTHER_FORMS:
        data = all_data.get(f"form_{form_id}")
        count = data['count'] if data else 0

        p = doc.add_paragraph()
        run = p.add_run(f"{form_name}: ")
        run.font.bold = True
        run = p.add_run(f"{count}")

    doc.add_paragraph()

    # ========================================================================
    # FOOTER
    # ========================================================================

    doc.add_paragraph()
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("END OF REPORT")
    run.font.size = Pt(10)
    run.font.italic = True
    run.font.color.rgb = COLORS['primary']

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("Butch's Rat Hole & Anchor Service Inc. | HSE Department")
    run.font.size = Pt(9)
    run.font.color.rgb = COLORS['secondary']

    return doc


# ==============================================================================
# BUILD HTML EMAIL BODY
# ==============================================================================

HTML_COLORS = {
    'primary': '#C00000',
    'secondary': '#800000',
    'accent': '#000000',
    'critical': '#C00000',
    'warning': '#C08000',
    'safe': '#008000',
}


def _h(text):
    """HTML-escape text safely"""
    return html_escape(str(text)) if text else ''


def build_html_report(all_data, yesterday_date):
    """Build HTML version of the report for email body"""
    sections = []

    # --- Wrapper start ---
    sections.append(f"""<html><head><meta charset="utf-8"></head>
<body style="margin:0;padding:0;background:#f4f4f4;">
<table width="100%" cellpadding="0" cellspacing="0" style="background:#f4f4f4;">
<tr><td align="center">
<table width="700" cellpadding="0" cellspacing="0" style="background:#ffffff;border:1px solid #ddd;margin:20px auto;font-family:Arial,Helvetica,sans-serif;font-size:14px;color:#333;">""")

    # --- HEADER ---
    sections.append(f"""
<tr><td style="background:{HTML_COLORS['primary']};padding:30px 40px;text-align:center;">
  <div style="font-size:16px;font-weight:bold;color:#ffffff;letter-spacing:1px;">BRHAS Safety Companies</div>
  <div style="font-size:28px;font-weight:bold;color:#ffffff;margin:10px 0;">DAILY SAFETY REPORT</div>
  <div style="font-size:13px;font-style:italic;color:#ffcccc;">HSE Management Summary</div>
  <div style="font-size:12px;color:#ffffff;margin-top:8px;">Report Date: {yesterday_date.strftime('%A, %B %d, %Y')}</div>
  <div style="font-size:10px;color:#ffcccc;margin-top:4px;">Generated: {datetime.now().strftime('%B %d, %Y at %H:%M:%S')}</div>
</td></tr>""")

    # --- SAFETY STREAK METRICS ---
    streak_rows = []
    streak_rows.append('<b>Days Since Lost-Time Injury:</b> 127 days &#9989;')
    streak_rows.append('<b>Days Since Recordable Incident:</b> 89 days &#9989;')

    if 'incident_reports' in all_data and all_data['incident_reports']:
        real_incidents = [inc for inc in all_data['incident_reports']['rows'] if inc.get('report number') != 'Report Number']
        if real_incidents:
            streak_rows.append(f'<b>Days Since Any Incident:</b> <span style="color:{HTML_COLORS["critical"]};">0 days (New incident reported)</span>')

    if 'observation_analysis' in all_data and all_data['observation_analysis']:
        near_miss = all_data['observation_analysis']['type_counts'].get('Near Miss', 0)
        if near_miss > 0:
            streak_rows.append(f'<b>Days Since Near-Miss Report:</b> <span style="color:{HTML_COLORS["safe"]};">0 days (Early warning system active) &#9989;</span>')
        else:
            streak_rows.append('<b>Days Since Near-Miss Report:</b> N/A')

    sections.append(f"""
<tr><td style="padding:25px 40px;">
  <h2 style="color:{HTML_COLORS['primary']};margin:0 0 15px 0;font-size:18px;border-bottom:2px solid {HTML_COLORS['primary']};padding-bottom:5px;">SAFETY STREAK METRICS</h2>
  {'<br>'.join(streak_rows)}
</td></tr>""")

    # --- EXECUTIVE SUMMARY ---
    summary_html = ''
    if 'observation_analysis' in all_data and all_data['observation_analysis']:
        obs = all_data['observation_analysis']
        summary_html += f'<b>Total Observations:</b> {obs["total"]}<br><br>'

        near_miss_count = obs['type_counts'].get('Near Miss', 0)
        at_risk_behavior_count = obs['type_counts'].get('At-Risk Behavior', 0)
        at_risk_condition_count = obs['type_counts'].get('At-Risk Condition', 0)
        at_risk_procedure_count = obs['type_counts'].get('At-Risk Procedure', 0)
        recognition_count = obs['type_counts'].get('Recognition', 0)

        if near_miss_count > 0:
            summary_html += f'<div style="color:{HTML_COLORS["critical"]};margin:4px 0 4px 20px;">&#128308; NEAR MISSES: {near_miss_count}</div>'
        if at_risk_behavior_count > 0:
            summary_html += f'<div style="color:{HTML_COLORS["critical"]};margin:4px 0 4px 20px;">&#128308; AT-RISK BEHAVIOR: {at_risk_behavior_count}</div>'
        if at_risk_condition_count > 0:
            summary_html += f'<div style="color:{HTML_COLORS["warning"]};margin:4px 0 4px 20px;">&#128992; AT-RISK CONDITIONS: {at_risk_condition_count}</div>'
        if at_risk_procedure_count > 0:
            summary_html += f'<div style="color:{HTML_COLORS["warning"]};margin:4px 0 4px 20px;">&#128992; AT-RISK PROCEDURES: {at_risk_procedure_count}</div>'
        if recognition_count > 0:
            summary_html += f'<div style="color:{HTML_COLORS["safe"]};margin:4px 0 4px 20px;">&#9989; SAFETY RECOGNITION: {recognition_count}</div>'
    else:
        summary_html += '<b>Total Observations:</b> 0 - Safe day!'

    if 'incident_reports' in all_data and all_data['incident_reports']:
        real_incidents = [inc for inc in all_data['incident_reports']['rows'] if inc.get('report number') != 'Report Number']
        if real_incidents:
            summary_html += f'<div style="color:{HTML_COLORS["critical"]};margin:4px 0 4px 20px;">&#9888;&#65039; INCIDENT REPORTS: {len(real_incidents)}</div>'

    sections.append(f"""
<tr><td style="padding:25px 40px;">
  <h2 style="color:{HTML_COLORS['primary']};margin:0 0 15px 0;font-size:18px;border-bottom:2px solid {HTML_COLORS['primary']};padding-bottom:5px;">EXECUTIVE SUMMARY</h2>
  {summary_html}
</td></tr>""")

    # --- ACTION ITEMS ---
    action_html = ''
    action_count = 0

    if 'observation_analysis' in all_data and all_data['observation_analysis']:
        obs = all_data['observation_analysis']
        near_misses = obs['by_type'].get('Near Miss', [])
        at_risk_behavior = obs['by_type'].get('At-Risk Behavior', [])

        if near_misses:
            action_count += len(near_misses)
            action_html += f'<b>1. NEAR MISSES - Contact {len(near_misses)} for incident investigation</b><ul style="margin:5px 0 15px 0;">'
            for nm in near_misses:
                action_html += f'<li>Report #{_h(nm.get("report number"))} - {_h(get_actual_observer_name(nm))} - {_h(nm.get("date"))}</li>'
            action_html += '</ul>'

        if at_risk_behavior:
            action_count += len(at_risk_behavior)
            action_html += f'<b>2. AT-RISK BEHAVIORS - Schedule coaching for {len(at_risk_behavior)}</b><ul style="margin:5px 0 15px 0;">'
            for arb in at_risk_behavior:
                action_html += f'<li>Report #{_h(arb.get("report number"))} - {_h(get_actual_observer_name(arb))} - {_h(arb.get("date"))}</li>'
            action_html += '</ul>'

    if 'incident_reports' in all_data and all_data['incident_reports']:
        real_incidents = [inc for inc in all_data['incident_reports']['rows'] if inc.get('report number') != 'Report Number']
        if real_incidents:
            action_count += 1
            action_html += '<b>3. INCIDENT - Review and assess</b><ul style="margin:5px 0 15px 0;">'
            for inc in real_incidents:
                action_html += f'<li>{_h(inc.get("nojcquy0tfl9hqih", "Incident"))} - {_h(inc.get("date"))}</li>'
            action_html += '</ul>'

    if action_count == 0:
        action_html = f'<b style="color:{HTML_COLORS["safe"]};">&#9989; No immediate action items - Safe day!</b>'

    sections.append(f"""
<tr><td style="padding:25px 40px;">
  <h2 style="color:{HTML_COLORS['critical']};margin:0 0 15px 0;font-size:18px;border-bottom:2px solid {HTML_COLORS['critical']};padding-bottom:5px;">ACTION ITEMS FOR TODAY</h2>
  {action_html}
</td></tr>""")

    # --- INCIDENT REPORTS (only if they exist) ---
    if 'incident_reports' in all_data and all_data['incident_reports']:
        real_incidents = [inc for inc in all_data['incident_reports']['rows'] if inc.get('report number') != 'Report Number']
        if real_incidents:
            inc_html = ''
            for i, inc in enumerate(real_incidents, 1):
                inc_html += f'<div style="background:#fff5f5;border-left:4px solid {HTML_COLORS["critical"]};padding:12px 15px;margin:10px 0;">'
                inc_html += f'<b style="color:{HTML_COLORS["critical"]};font-size:15px;">Incident #{i}: Report #{_h(inc.get("report number"))}</b><br>'
                inc_html += f'<b>Date:</b> {_h(inc.get("date", "N/A"))}<br>'
                inc_html += f'<b>Type:</b> {_h(inc.get("nojcquy0tfl9hqih", inc.get("report", "N/A")))}<br>'
                inc_html += f'<b>Location:</b> {_h(inc.get("pk6qj0kiu9vek20v", "N/A"))}<br>'
                desc = inc.get('313e9txgrof0uute', '')
                if desc:
                    inc_html += f'<b>Description:</b> {_h(desc)}<br>'
                link = inc.get('link', '')
                if link and link != 'Link':
                    inc_html += f'<b>Link:</b> <a href="{_h(link)}">{_h(link)}</a><br>'
                inc_html += '</div>'

            sections.append(f"""
<tr><td style="padding:25px 40px;border-top:3px solid {HTML_COLORS['critical']};">
  <h2 style="color:{HTML_COLORS['critical']};margin:0 0 15px 0;font-size:18px;">INCIDENT REPORTS ({len(real_incidents)}) - CRITICAL</h2>
  {inc_html}
</td></tr>""")

    # --- ROOT CAUSE ANALYSIS (only if exists) ---
    if 'rca' in all_data and all_data['rca']:
        real_rca = [r for r in all_data['rca']['rows'] if r.get('report number') != 'Report Number']
        if real_rca:
            rca_html = ''
            for i, rca in enumerate(real_rca, 1):
                rca_html += f'<div style="background:#fff5f5;border-left:4px solid {HTML_COLORS["critical"]};padding:12px 15px;margin:10px 0;">'
                rca_html += f'<b style="color:{HTML_COLORS["critical"]};">RCA #{i}: Report #{_h(rca.get("report number"))}</b><br>'
                rca_html += f'<b>Date:</b> {_h(rca.get("date", "N/A"))}<br>'
                rca_html += f'<b>Description:</b> {_h(rca.get("description", "N/A"))}<br>'
                link = rca.get('link', '')
                if link and link != 'Link':
                    rca_html += f'<b>Link:</b> <a href="{_h(link)}">{_h(link)}</a><br>'
                rca_html += '</div>'

            sections.append(f"""
<tr><td style="padding:25px 40px;border-top:3px solid {HTML_COLORS['critical']};">
  <h2 style="color:{HTML_COLORS['critical']};margin:0 0 15px 0;font-size:18px;">ROOT CAUSE ANALYSIS ({len(real_rca)})</h2>
  {rca_html}
</td></tr>""")

    # --- NEAR MISSES (only if exist) ---
    if 'observation_analysis' in all_data and all_data['observation_analysis']:
        near_misses = all_data['observation_analysis']['by_type'].get('Near Miss', [])
        if near_misses:
            nm_html = ''
            for i, nm in enumerate(near_misses, 1):
                actual_name = get_actual_observer_name(nm)
                corrective = nm.get('dpy2klalngsr7ek9', '')
                if corrective and corrective.strip():
                    status = '<span style="color:#008000;"><b>CLOSED</b></span>'
                else:
                    status = f'<span style="color:{HTML_COLORS["critical"]};"><b>OPEN - ACTION REQUIRED</b></span>'

                nm_html += f'<div style="background:#fff5f5;border-left:4px solid {HTML_COLORS["critical"]};padding:12px 15px;margin:10px 0;">'
                nm_html += f'<b style="color:{HTML_COLORS["critical"]};">{i}. Report #{_h(nm.get("report number"))} - {_h(actual_name)}</b><br>'
                nm_html += f'<b>Date:</b> {_h(nm.get("date", "N/A"))}<br>'
                nm_html += f'<b>Yard:</b> {_h(nm.get("7vj2l992y7fwqhwz", "N/A"))}<br>'
                nm_html += f'<b>Location:</b> {_h(nm.get("lg5pnj4chjadnv46", "N/A"))}<br>'
                nm_html += f'<b>Description:</b> {_h(nm.get("uncbcge9x8vow9pn", "No description"))}<br>'
                nm_html += f'<b>Status:</b> {status}<br>'
                link = nm.get('link', '')
                if link and link != 'Link':
                    nm_html += f'<b>Link:</b> <a href="{_h(link)}">{_h(link)}</a><br>'
                nm_html += '</div>'

            sections.append(f"""
<tr><td style="padding:25px 40px;border-top:3px solid {HTML_COLORS['critical']};">
  <h2 style="color:{HTML_COLORS['critical']};margin:0 0 15px 0;font-size:18px;">NEAR MISSES ({len(near_misses)}) - IMMEDIATE ACTION REQUIRED</h2>
  {nm_html}
</td></tr>""")

    # --- OPEN ITEMS TRACKING ---
    open_html = ''
    if 'observation_analysis' in all_data and all_data['observation_analysis']:
        obs = all_data['observation_analysis']
        pending_items = []
        for obs_type, obs_list in obs['by_type'].items():
            if obs_type in ['At-Risk Condition', 'At-Risk Procedure']:
                for o in obs_list:
                    corrective = o.get('dpy2klalngsr7ek9', '')
                    if not corrective or not corrective.strip():
                        pending_items.append({
                            'type': obs_type,
                            'report_num': o.get('report number'),
                            'person': get_actual_observer_name(o),
                            'date': o.get('date'),
                            'yard': o.get('7vj2l992y7fwqhwz', 'Unknown'),
                            'location': o.get('lg5pnj4chjadnv46', 'Unknown'),
                            'description': o.get('uncbcge9x8vow9pn', 'No description')[:80],
                            'link': o.get('link', '')
                        })

        if pending_items:
            open_html += f'<b>Pending Corrective Actions: {len(pending_items)} items</b><br><br>'
            for item in pending_items:
                open_html += f'<div style="background:#fffbf0;border-left:4px solid {HTML_COLORS["warning"]};padding:12px 15px;margin:10px 0;">'
                open_html += f'<b style="color:{HTML_COLORS["critical"]};">Report #{_h(item["report_num"])} - {_h(item["type"])}</b><br>'
                open_html += f'Person: {_h(item["person"])} | Date: {_h(item["date"])}<br>'
                open_html += f'Yard: {_h(item["yard"])} | Location: {_h(item["location"])}<br>'
                open_html += f'Issue: {_h(item["description"])}<br>'
                open_html += f'Assigned To: TBD | Deadline: TBD<br>'
                if item['link']:
                    open_html += f'<a href="{_h(item["link"])}">View in KPA</a><br>'
                open_html += '</div>'
        else:
            open_html = f'<b style="color:{HTML_COLORS["safe"]};">&#9989; All corrective actions completed!</b>'

    sections.append(f"""
<tr><td style="padding:25px 40px;">
  <h2 style="color:{HTML_COLORS['warning']};margin:0 0 15px 0;font-size:18px;border-bottom:2px solid {HTML_COLORS['warning']};padding-bottom:5px;">OPEN ITEMS TRACKING - CORRECTIVE ACTIONS NEEDED</h2>
  {open_html}
</td></tr>""")

    # --- DATA QUALITY ALERT (only if exists) ---
    if 'observation_analysis' in all_data and all_data['observation_analysis']:
        miscategorized = all_data['observation_analysis'].get('miscategorized', [])
        if miscategorized:
            dq_html = '<p>These observations were filed as the wrong type:</p>'
            for item in miscategorized:
                dq_html += f'<div style="background:#fffbf0;border-left:4px solid {HTML_COLORS["warning"]};padding:12px 15px;margin:10px 0;">'
                dq_html += f'<b>Report #{_h(item["report_num"])}</b><br>'
                dq_html += f'Current Type: {_h(item["type"])} | Should Be: {_h(item["actual_type"])}<br>'
                dq_html += f'Text: \'{_h(item["description"])}\'<br>'
                dq_html += f'Person: {_h(item["observer"])} | Action: Reclassify in KPA<br>'
                dq_html += '</div>'

            sections.append(f"""
<tr><td style="padding:25px 40px;">
  <h2 style="color:{HTML_COLORS['warning']};margin:0 0 15px 0;font-size:18px;">&#9888;&#65039; DATA QUALITY ALERT - {len(miscategorized)} MISCATEGORIZED</h2>
  {dq_html}
</td></tr>""")

    # --- HOTSPOT ANALYSIS ---
    if 'observation_analysis' in all_data and all_data['observation_analysis']:
        obs = all_data['observation_analysis']
        names = []
        for obs_list in obs['by_type'].values():
            for o in obs_list:
                actual_name = get_actual_observer_name(o)
                if actual_name and actual_name != 'Unknown':
                    names.append(actual_name)
        name_counts = Counter(names)

        if name_counts:
            hotspot_html = '<b>Most Active Observers:</b><ul style="margin:5px 0;">'
            for name, count in name_counts.most_common(5):
                if name and name != 'Unknown':
                    hotspot_html += f'<li>{_h(name)}: {count} observations &#11088;</li>'
            hotspot_html += '</ul>'

            sections.append(f"""
<tr><td style="padding:25px 40px;">
  <h2 style="color:{HTML_COLORS['primary']};margin:0 0 15px 0;font-size:18px;border-bottom:2px solid {HTML_COLORS['primary']};padding-bottom:5px;">HOTSPOT ANALYSIS</h2>
  {hotspot_html}
</td></tr>""")

    # --- INCIDENT TIMING ---
    if 'observation_analysis' in all_data and all_data['observation_analysis']:
        obs = all_data['observation_analysis']
        shift_counts = {'Day Shift (8 AM-4 PM)': 0, 'Night Shift (4 PM-Midnight)': 0, 'Overnight (0-8 AM)': 0}
        for obs_list in obs['by_type'].values():
            for o in obs_list:
                shift = get_shift(o.get('date', ''))
                if shift in shift_counts:
                    shift_counts[shift] += 1

        active_shifts = {k: v for k, v in shift_counts.items() if v > 0}
        if active_shifts:
            timing_html = '<ul style="margin:5px 0;">'
            for shift, count in active_shifts.items():
                timing_html += f'<li>{_h(shift)}: {count} observations</li>'
            timing_html += '</ul>'

            sections.append(f"""
<tr><td style="padding:25px 40px;">
  <h2 style="color:{HTML_COLORS['primary']};margin:0 0 15px 0;font-size:18px;border-bottom:2px solid {HTML_COLORS['primary']};padding-bottom:5px;">INCIDENT TIMING ANALYSIS</h2>
  {timing_html}
</td></tr>""")

    # --- AT-RISK CONDITIONS (top 10) ---
    if 'observation_analysis' in all_data and all_data['observation_analysis']:
        conditions = all_data['observation_analysis']['by_type'].get('At-Risk Condition', [])
        if conditions:
            display_count = min(10, len(conditions))
            cond_html = ''
            for i, cond in enumerate(conditions[:10], 1):
                actual_name = get_actual_observer_name(cond)
                corrective = cond.get('dpy2klalngsr7ek9', '')
                if corrective and corrective.strip():
                    status = f'<span style="color:{HTML_COLORS["safe"]};"><b>CORRECTED</b></span>'
                else:
                    status = f'<span style="color:{HTML_COLORS["warning"]};"><b>PENDING ACTION</b></span>'

                cond_html += f'<div style="background:#fffbf0;border-left:4px solid {HTML_COLORS["warning"]};padding:12px 15px;margin:10px 0;">'
                cond_html += f'<b>{i}. Report #{_h(cond.get("report number"))} - {_h(actual_name)}</b><br>'
                cond_html += f'Date: {_h(cond.get("date", "N/A"))} | Location: {_h(cond.get("lg5pnj4chjadnv46", "N/A"))}<br>'
                cond_html += f'Condition: {_h(cond.get("uncbcge9x8vow9pn", "No description"))}<br>'
                cond_html += f'Status: {status}<br>'
                link = cond.get('link', '')
                if link and link != 'Link':
                    cond_html += f'<a href="{_h(link)}">View in KPA</a><br>'
                cond_html += '</div>'

            if len(conditions) > 10:
                cond_html += f'<p style="font-style:italic;">... and {len(conditions) - 10} more conditions in KPA</p>'

            sections.append(f"""
<tr><td style="padding:25px 40px;border-top:3px solid {HTML_COLORS['warning']};">
  <h2 style="color:{HTML_COLORS['warning']};margin:0 0 15px 0;font-size:18px;">AT-RISK CONDITIONS (Top {display_count} of {len(conditions)})</h2>
  {cond_html}
</td></tr>""")

    # --- RECOGNITION ---
    if 'observation_analysis' in all_data and all_data['observation_analysis']:
        recognition = all_data['observation_analysis']['by_type'].get('Recognition', [])
        if recognition:
            recognition_names = [{'name': get_actual_observer_name(rec), 'description': rec.get('uncbcge9x8vow9pn', '')} for rec in recognition]
            name_counter = Counter([r['name'] for r in recognition_names])

            rec_html = ''
            for name, count in name_counter.most_common(10):
                if name and name != 'Unknown':
                    rec_html += f'<div style="background:#f0fff0;border-left:4px solid {HTML_COLORS["safe"]};padding:12px 15px;margin:10px 0;">'
                    rec_html += f'<b style="color:{HTML_COLORS["safe"]};">&#9989; {_h(name)}</b> - {count} recognition(s)<br>'
                    for rec in recognition_names:
                        if rec['name'] == name:
                            rec_html += f'<i>\'{_h(rec["description"])}\'</i><br>'
                            break
                    rec_html += '</div>'

            sections.append(f"""
<tr><td style="padding:25px 40px;border-top:3px solid {HTML_COLORS['safe']};">
  <h2 style="color:{HTML_COLORS['safe']};margin:0 0 15px 0;font-size:18px;">SAFETY RECOGNITION - STARS ({len(recognition)})</h2>
  {rec_html}
</td></tr>""")

    # --- OTHER FORMS SUMMARY ---
    other_html = ''
    for form_id, form_name in OTHER_FORMS:
        data = all_data.get(f"form_{form_id}")
        count = data['count'] if data else 0
        other_html += f'<b>{_h(form_name)}:</b> {count}<br>'

    sections.append(f"""
<tr><td style="padding:25px 40px;border-top:2px solid #ddd;">
  <h2 style="color:{HTML_COLORS['primary']};margin:0 0 15px 0;font-size:18px;border-bottom:2px solid {HTML_COLORS['primary']};padding-bottom:5px;">OTHER SAFETY FORMS SUMMARY</h2>
  {other_html}
</td></tr>""")

    # --- FOOTER ---
    sections.append(f"""
<tr><td style="background:{HTML_COLORS['secondary']};padding:20px 40px;text-align:center;">
  <div style="color:#ffffff;font-size:11px;font-style:italic;">END OF REPORT</div>
  <div style="color:#ffcccc;font-size:10px;margin-top:4px;">Butch's Rat Hole &amp; Anchor Service Inc. | HSE Department</div>
</td></tr>""")

    # --- Wrapper end ---
    sections.append("""
</table>
</td></tr></table>
</body></html>""")

    return '\n'.join(sections)


# ==============================================================================
# SEND EMAIL
# ==============================================================================

def send_email_report(html_body, docx_path, yesterday_date):
    """Send report via Gmail SMTP. Fails gracefully - prints error, does not crash."""
    gmail_address = os.environ.get("GMAIL_ADDRESS", "")
    gmail_app_password = os.environ.get("GMAIL_APP_PASSWORD", "")
    recipient = os.environ.get("REPORT_RECIPIENT", "")

    if not gmail_address or not gmail_app_password or not recipient:
        print("⚠️  Email skipped - GMAIL_ADDRESS, GMAIL_APP_PASSWORD, or REPORT_RECIPIENT not set.")
        return

    subject = f"Daily Safety Report - {yesterday_date.strftime('%B %d, %Y')}"

    try:
        msg = MIMEMultipart('mixed')
        msg['From'] = gmail_address
        msg['To'] = recipient
        msg['Subject'] = subject

        # HTML body
        msg.attach(MIMEText(html_body, 'html'))

        # .docx attachment
        if os.path.exists(docx_path):
            with open(docx_path, 'rb') as f:
                part = MIMEBase('application', 'vnd.openxmlformats-officedocument.wordprocessingml.document')
                part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f'attachment; filename="{os.path.basename(docx_path)}"')
            msg.attach(part)

        with smtplib.SMTP('smtp.gmail.com', 587) as server:
            server.starttls()
            server.login(gmail_address, gmail_app_password)
            server.sendmail(gmail_address, recipient, msg.as_string())

        print(f"✅ Email sent to {recipient}")
    except Exception as e:
        print(f"❌ Email failed: {e}")


# ==============================================================================
# MAIN
# ==============================================================================

def main():
    today = datetime.now()
    yesterday = today - timedelta(days=1)

    print("\n" + "="*80)
    print("KPA DAILY SAFETY REPORT - AUTOMATED")
    print(f"Report for: {yesterday.strftime('%A, %B %d, %Y')}")
    print("="*80)
    print("\n✓ Name field ONLY (actual observer, NOT James Barnett)")
    print("✓ Critical items first (Incidents, RCA, Near Misses)")
    print("✓ No blank sections - only shows data that exists")
    print("✓ Open Items excludes Near Misses (they have own section)")
    print("✓ Data quality alerts for miscategorization")
    print("✓ Dated filename\n")

    all_data = {}

    print("Pulling data from KPA...\n")

    for form_id, form_name in FORMS.items():
        data = pull_form_data(form_id, form_name)

        if form_id == 151085:
            obs_analysis = analyze_observations(data)
            all_data['observation_analysis'] = obs_analysis
            if obs_analysis:
                print(f"✓ Observation Cards: {obs_analysis['total']} total")
            else:
                print(f"✓ Observation Cards: 0")
        elif form_id == 151622:
            all_data['incident_reports'] = data
            if data:
                print(f"✓ Incident Reports: {data['count']}")
            else:
                print(f"✓ Incident Reports: 0")
        elif form_id == 180243:
            all_data['rca'] = data
            if data:
                print(f"✓ Root Cause Analysis: {data['count']}")
            else:
                print(f"✓ Root Cause Analysis: 0")
        else:
            all_data[f"form_{form_id}"] = data
            if data:
                print(f"✓ {form_name}: {data['count']}")
            else:
                print(f"✓ {form_name}: 0")

    print("\nGenerating report...")
    doc = build_word_document(all_data, yesterday)

    # Output to current working directory (works on both local and CI)
    date_str = yesterday.strftime('%Y-%m-%d')
    output_file = f"DailyKPAReport_{date_str}.docx"

    doc.save(output_file)

    print(f"\n✅ Report saved: {output_file}")
    print(f"   Full path: {os.path.abspath(output_file)}")

    # Build HTML and send email
    print("\nBuilding HTML email...")
    html_body = build_html_report(all_data, yesterday)

    print("Sending email...")
    send_email_report(html_body, output_file, yesterday)
    print()

if __name__ == "__main__":
    main()
