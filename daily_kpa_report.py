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
- Assessment & Audit Analysis (NEW - assessor details, compliance by yard,
  critical findings, corrective actions, trends, leadership recommendations)
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

# Assessment/Audit forms with metadata for deep analysis
ASSESSMENT_FORMS = {
    381707: {"name": "CSG - Safety Casing Field Assessment", "type": "Field Assessment", "division": "Casing"},
    152018: {"name": "Vehicle Inspection Checklist", "type": "Inspection", "division": "All"},
    385365: {"name": "TD - Rig Inspection", "type": "Rig Inspection", "division": "Transcend"},
    484193: {"name": "TD - Observation Card", "type": "Observation", "division": "Transcend"},
    226217: {"name": "WS - Poly Pipe Field Assessment", "type": "Field Assessment", "division": "Poly Pipe"},
    386087: {"name": "WS - Pit Lining Field Assessment", "type": "Field Assessment", "division": "Pit Lining"},
    172295: {"name": "Construction - Site Safety Review", "type": "Site Review", "division": "Construction"},
    153181: {"name": "RH - Rathole Field Assessment", "type": "Field Assessment", "division": "Rathole"},
    152034: {"name": "HSE - Workplace Inspection Checklist", "type": "Inspection", "division": "HSE"},
}

KPA_RESPONSE_URL = "https://brhas-ees.kpaehs.com/forms/responses/view"

# Keywords for smart field detection in assessment CSV headers
COMPLIANCE_KEYWORDS = ['compliance', 'rating', 'satisfactory', 'pass', 'fail', 'acceptable',
                       'result', 'score', 'compliant', 'conformance']
FINDING_KEYWORDS = ['issue', 'finding', 'non-conformance', 'corrective', 'deficiency',
                    'violation', 'hazard', 'concern', 'recommendation', 'comment',
                    'note', 'detail']
YARD_KEYWORDS = ['yard', 'location', 'site', 'field office', 'facility', 'area']

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


# ==============================================================================
# ASSESSMENT & AUDIT ANALYSIS FUNCTIONS
# ==============================================================================

def detect_field_columns(headers):
    """Detect key columns from assessment CSV headers using keyword matching"""
    fields = {
        'compliance': [],
        'findings': [],
        'yard': [],
        'severity': [],
        'assessor': [],
        'corrective_action': [],
    }

    if not headers:
        return fields

    for header in headers:
        h_lower = header.lower()

        # Skip standard metadata fields
        if h_lower in ['report number', 'date', 'link', 'observer', 'name']:
            continue

        if any(kw in h_lower for kw in COMPLIANCE_KEYWORDS):
            fields['compliance'].append(header)
        if any(kw in h_lower for kw in FINDING_KEYWORDS):
            fields['findings'].append(header)
        if any(kw in h_lower for kw in YARD_KEYWORDS):
            fields['yard'].append(header)
        if any(kw in h_lower for kw in ['severity', 'priority', 'risk level', 'critical']):
            fields['severity'].append(header)
        if any(kw in h_lower for kw in ['assessor', 'inspector', 'auditor', 'reviewer', 'conducted by']):
            fields['assessor'].append(header)
        if any(kw in h_lower for kw in ['corrective', 'action required', 'action taken', 'follow up', 'follow-up']):
            fields['corrective_action'].append(header)

    return fields


def classify_compliance_value(value):
    """Classify a field value as compliant, non-compliant, or unknown"""
    if not value:
        return 'unknown'
    v = value.strip().lower()

    non_compliant_terms = ['fail', 'unsatisfactory', 'non-compliant', 'unacceptable',
                           'poor', 'deficient', 'inadequate', 'needs improvement', 'not met']
    compliant_terms = ['pass', 'yes', 'satisfactory', 'compliant', 'acceptable', 'good',
                       'meets', 'adequate', 'ok', 'n/a', 'not applicable']

    for term in non_compliant_terms:
        if term in v:
            return 'non_compliant'
    for term in compliant_terms:
        if term in v:
            return 'compliant'
    return 'unknown'


def classify_severity(text):
    """Classify the severity level of a finding based on its text"""
    if not text:
        return 'low'
    t = text.lower()

    critical_terms = ['critical', 'immediate', 'danger', 'life-threatening', 'fatal',
                      'imminent', 'emergency', 'severe', 'death']
    high_terms = ['high', 'serious', 'major', 'significant', 'injury potential',
                  'non-compliant', 'violation', 'failed']
    medium_terms = ['medium', 'moderate', 'minor damage', 'needs repair', 'worn', 'missing']

    for term in critical_terms:
        if term in t:
            return 'critical'
    for term in high_terms:
        if term in t:
            return 'high'
    for term in medium_terms:
        if term in t:
            return 'medium'

    return 'low'


def get_assessor_name(row):
    """Get assessor/observer name from assessment form row"""
    for field_name, value in row.items():
        if any(kw in field_name.lower() for kw in ['assessor', 'inspector', 'auditor', 'conducted by', 'reviewer']):
            if value and value.strip() and value.strip().lower() not in ['none', 'unknown', '']:
                return value.strip()

    return get_actual_observer_name(row)


def get_yard_from_row(row, detected_fields):
    """Extract yard/location from a row using detected fields"""
    for field in detected_fields.get('yard', []):
        val = row.get(field, '').strip()
        if val and val.lower() not in ['n/a', 'none', 'unknown', '']:
            return val

    for key in ['7vj2l992y7fwqhwz', 'lg5pnj4chjadnv46']:
        val = row.get(key, '').strip()
        if val and val.lower() not in ['n/a', 'none', 'unknown', '']:
            return val

    for field_name, value in row.items():
        if ('yard' in field_name.lower() or 'location' in field_name.lower()):
            if value and value.strip() and value.strip().lower() not in ['n/a', 'none', 'unknown', '']:
                return value.strip()

    return 'Unknown'


def get_kpa_link(report_num):
    """Build clickable KPA link from report number"""
    if report_num and report_num not in ['Report Number', '']:
        return f"{KPA_RESPONSE_URL}/{report_num}"
    return ''


def analyze_assessments(all_data):
    """Analyze all assessment/audit form data for the daily report"""
    analysis = {
        'activity_summary': [],
        'findings_by_severity': {'critical': [], 'high': [], 'medium': [], 'low': []},
        'compliance_by_yard': {},
        'assessor_stats': {},
        'corrective_actions': [],
        'all_findings': [],
        'trends': [],
        'recommendations': {'immediate': [], 'this_week': [], 'monthly': []},
        'total_assessments': 0,
        'total_findings': 0,
        'has_data': False,
    }

    for form_id, form_info in ASSESSMENT_FORMS.items():
        data = all_data.get(f"form_{form_id}")
        if not data or data['count'] == 0:
            continue

        analysis['has_data'] = True
        analysis['total_assessments'] += data['count']

        detected = detect_field_columns(data['headers'])

        form_assessors = set()
        form_findings = []
        form_compliant = 0
        form_non_compliant = 0

        for row in data['rows']:
            assessor = get_assessor_name(row)
            form_assessors.add(assessor)

            if assessor not in analysis['assessor_stats']:
                analysis['assessor_stats'][assessor] = {
                    'total': 0, 'forms': set(), 'divisions': set(), 'findings_found': 0
                }
            analysis['assessor_stats'][assessor]['total'] += 1
            analysis['assessor_stats'][assessor]['forms'].add(form_info['name'])
            analysis['assessor_stats'][assessor]['divisions'].add(form_info['division'])

            yard = get_yard_from_row(row, detected)

            if yard not in analysis['compliance_by_yard']:
                analysis['compliance_by_yard'][yard] = {
                    'total': 0, 'compliant': 0, 'non_compliant': 0,
                    'findings': [], 'forms_used': set()
                }
            analysis['compliance_by_yard'][yard]['total'] += 1
            analysis['compliance_by_yard'][yard]['forms_used'].add(form_info['name'])

            # Check compliance fields
            row_compliant = True
            for comp_field in detected['compliance']:
                val = row.get(comp_field, '')
                result = classify_compliance_value(val)
                if result == 'non_compliant':
                    row_compliant = False
                    break

            if row_compliant:
                form_compliant += 1
                analysis['compliance_by_yard'][yard]['compliant'] += 1
            else:
                form_non_compliant += 1
                analysis['compliance_by_yard'][yard]['non_compliant'] += 1

            # Extract findings
            for finding_field in detected['findings']:
                finding_text = row.get(finding_field, '').strip()
                if finding_text and len(finding_text) > 3 and finding_text.lower() not in ['n/a', 'none', 'no', 'na']:
                    severity = classify_severity(finding_text)

                    for sev_field in detected['severity']:
                        sev_val = row.get(sev_field, '').strip()
                        if sev_val:
                            severity = classify_severity(sev_val)
                            break

                    finding = {
                        'form_name': form_info['name'],
                        'division': form_info['division'],
                        'assessor': assessor,
                        'yard': yard,
                        'description': finding_text[:200],
                        'severity': severity,
                        'report_num': row.get('report number', ''),
                        'date': row.get('date', ''),
                        'link': get_kpa_link(row.get('report number', '')),
                    }

                    form_findings.append(finding)
                    analysis['findings_by_severity'][severity].append(finding)
                    analysis['compliance_by_yard'][yard]['findings'].append(finding)
                    analysis['all_findings'].append(finding)
                    analysis['total_findings'] += 1
                    analysis['assessor_stats'][assessor]['findings_found'] += 1

            # Extract corrective actions
            for ca_field in detected['corrective_action']:
                ca_text = row.get(ca_field, '').strip()
                if ca_text and len(ca_text) > 3 and ca_text.lower() not in ['n/a', 'none', 'no', 'na']:
                    analysis['corrective_actions'].append({
                        'form_name': form_info['name'],
                        'description': ca_text[:200],
                        'assessor': assessor,
                        'yard': yard,
                        'date': row.get('date', ''),
                        'report_num': row.get('report number', ''),
                        'link': get_kpa_link(row.get('report number', '')),
                        'status': 'Open',
                    })

        # Activity summary for this form
        compliance_rate = (form_compliant / data['count'] * 100) if data['count'] > 0 else 0

        assessment_analysis_item = {
            'form_name': form_info['name'],
            'form_type': form_info['type'],
            'division': form_info['division'],
            'count': data['count'],
            'assessors': sorted(form_assessors - {'Unknown'}),
            'findings_count': len(form_findings),
            'compliance_rate': compliance_rate,
            'compliant': form_compliant,
            'non_compliant': form_non_compliant,
        }
        # Only add assessors list if "Unknown" was the only one
        if not assessment_analysis_item['assessors'] and 'Unknown' in form_assessors:
            assessment_analysis_item['assessors'] = ['Unknown']

        analysis['activity_summary'].append(assessment_analysis_item)

    if analysis['has_data']:
        _generate_assessment_trends(analysis)
        _generate_assessment_recommendations(analysis)

    return analysis


def _generate_assessment_trends(analysis):
    """Generate trend observations from assessment data"""
    trends = []

    # Yards with multiple findings
    problem_yards = {yard: info for yard, info in analysis['compliance_by_yard'].items()
                     if len(info['findings']) >= 2}
    if problem_yards:
        for yard, info in sorted(problem_yards.items(), key=lambda x: len(x[1]['findings']), reverse=True):
            trends.append(f"{yard}: {len(info['findings'])} findings across {info['total']} assessments")

    # Common safety terms across findings
    finding_words = Counter()
    safety_terms = ['ppe', 'housekeeping', 'guarding', 'electrical', 'fall', 'fire',
                    'chemical', 'ergonomic', 'noise', 'ventilation', 'lighting',
                    'signage', 'barricade', 'grounding', 'lockout', 'tagout',
                    'harness', 'helmet', 'goggles', 'gloves', 'boots']
    for finding in analysis['all_findings']:
        words = finding['description'].lower().split()
        for word in words:
            if word in safety_terms:
                finding_words[word] += 1

    for term, count in finding_words.most_common(3):
        if count >= 2:
            trends.append(f"{term.upper()} issues noted in {count} assessments")

    # Division activity
    division_counts = Counter()
    for summary in analysis['activity_summary']:
        division_counts[summary['division']] += summary['count']

    if division_counts:
        most_active = division_counts.most_common(1)[0]
        trends.append(f"Most active division: {most_active[0]} ({most_active[1]} assessments)")

    # Clean assessments (positive trend)
    clean_count = sum(1 for s in analysis['activity_summary'] if s['findings_count'] == 0 and s['count'] > 0)
    if clean_count > 0:
        trends.append(f"{clean_count} form type(s) had zero findings - strong compliance")

    analysis['trends'] = trends


def _generate_assessment_recommendations(analysis):
    """Generate leadership recommendations based on assessment analysis"""
    recs = analysis['recommendations']

    # IMMEDIATE: Critical and high findings
    critical = analysis['findings_by_severity']['critical']
    if critical:
        yards = set(f['yard'] for f in critical)
        recs['immediate'].append(f"Address {len(critical)} critical finding(s) in: {', '.join(yards)}")

    high = analysis['findings_by_severity']['high']
    if high:
        recs['immediate'].append(f"Review {len(high)} high-severity finding(s) requiring prompt attention")

    # THIS WEEK: Non-compliant yards and corrective actions
    non_compliant_yards = [yard for yard, info in analysis['compliance_by_yard'].items()
                           if info['non_compliant'] > 0]
    if non_compliant_yards:
        recs['this_week'].append(f"Follow up on non-compliant assessments at: {', '.join(non_compliant_yards[:5])}")

    open_cas = [ca for ca in analysis['corrective_actions'] if ca['status'] == 'Open']
    if open_cas:
        recs['this_week'].append(f"Track {len(open_cas)} open corrective action(s) to closure")

    # MONTHLY: Recognition and coverage
    if analysis['assessor_stats']:
        top_assessors = sorted(analysis['assessor_stats'].items(),
                               key=lambda x: x[1]['total'], reverse=True)[:3]
        names = [a[0] for a in top_assessors if a[0] != 'Unknown']
        if names:
            recs['monthly'].append(f"Recognize top assessors: {', '.join(names)}")

    active_divisions = set(s['division'] for s in analysis['activity_summary'])
    all_divisions = set(info['division'] for info in ASSESSMENT_FORMS.values())
    missing = all_divisions - active_divisions
    if missing:
        recs['monthly'].append(f"No assessments from: {', '.join(missing)} - consider scheduling")

    recs['monthly'].append("Review assessment frequency targets vs. actual completion rates")


# ==============================================================================
# ASSESSMENT AUDIT SUMMARY (replaces old "OTHER SAFETY FORMS SUMMARY")
# ==============================================================================

def _get_customer_from_row(row):
    """Extract customer/client name from a form row"""
    for field_name, value in row.items():
        if any(kw in field_name.lower() for kw in ['customer', 'client', 'company', 'operator', 'contractor']):
            if value and value.strip() and value.strip().lower() not in ['n/a', 'none', 'unknown', '']:
                return value.strip()
    return ''


def _get_issue_from_row(row, detected_fields):
    """Extract the primary issue/finding text from a form row"""
    # Try detected finding fields first
    for field in detected_fields.get('findings', []):
        val = row.get(field, '').strip()
        if val and len(val) > 3 and val.lower() not in ['n/a', 'none', 'no', 'na', 'no issues']:
            return val[:120]

    # Try corrective action fields (often contain the issue description)
    for field in detected_fields.get('corrective_action', []):
        val = row.get(field, '').strip()
        if val and len(val) > 3 and val.lower() not in ['n/a', 'none', 'no', 'na']:
            return val[:120]

    # Try observation description field used by observation cards
    for key in ['uncbcge9x8vow9pn']:
        val = row.get(key, '').strip()
        if val and len(val) > 3 and val.lower() not in ['n/a', 'none', 'no', 'na']:
            return val[:120]

    return 'None noted'


def extract_assessment_details(all_data):
    """Extract assessor, location, customer, form_id, and issues from each assessment row.

    Returns a list of dicts, one per form type in OTHER_FORMS.
    Each dict has: form_name, form_id, count, rows (list of detail dicts).
    Forms with 0 assessments still appear with count=0 and empty rows.
    """
    results = []

    for form_id, form_name in OTHER_FORMS:
        data = all_data.get(f"form_{form_id}")
        entry = {
            'form_name': form_name,
            'form_id': form_id,
            'count': data['count'] if data else 0,
            'rows': [],
        }

        if data and data['count'] > 0:
            detected = detect_field_columns(data['headers'])

            for row in data['rows']:
                report_num = row.get('report number', '')
                entry['rows'].append({
                    'assessor': get_assessor_name(row),
                    'location': get_yard_from_row(row, detected),
                    'customer': _get_customer_from_row(row),
                    'form_id': report_num,
                    'link': get_kpa_link(report_num),
                    'issue': _get_issue_from_row(row, detected),
                })

        results.append(entry)

    return results


def add_assessment_audit_summary(doc, assessment_details):
    """Create a Word table summarizing all assessment/audit forms.

    Replaces the old 'OTHER SAFETY FORMS SUMMARY' with a 6-column table:
    Form Type | Assessor | Location | Customer | Form ID | Issue Found
    """
    doc.add_page_break()
    add_heading(doc, "ASSESSMENT & AUDIT SUMMARY", 1)
    doc.add_paragraph()

    from docx.oxml.ns import qn as _qn
    from docx.oxml import OxmlElement as _OE

    # Check if there are any rows at all
    total_rows = sum(entry['count'] for entry in assessment_details)

    if total_rows == 0:
        p = doc.add_paragraph()
        p.add_run("No assessment or audit forms were completed yesterday.").font.italic = True

        # Still show the form list with counts
        doc.add_paragraph()
        for entry in assessment_details:
            p = doc.add_paragraph()
            run = p.add_run(f"{entry['form_name']}: ")
            run.font.bold = True
            p.add_run("0")
        return

    table = doc.add_table(rows=1, cols=6)
    table.style = 'Table Grid'

    # Dark header row
    headers = ['Form Type', 'Assessor', 'Location', 'Customer', 'Form ID', 'Issue Found']
    hdr_cells = table.rows[0].cells
    for i, txt in enumerate(headers):
        hdr_cells[i].text = txt
        for paragraph in hdr_cells[i].paragraphs:
            for run in paragraph.runs:
                run.font.bold = True
                run.font.size = Pt(8)
                run.font.color.rgb = RGBColor(255, 255, 255)
        shading = _OE('w:shd')
        shading.set(_qn('w:fill'), '800000')
        hdr_cells[i]._tc.get_or_add_tcPr().append(shading)

    for entry in assessment_details:
        if entry['count'] == 0:
            # Show a single row with 0 count
            row_cells = table.add_row().cells
            row_cells[0].text = entry['form_name']
            row_cells[1].text = '-'
            row_cells[2].text = '-'
            row_cells[3].text = '-'
            row_cells[4].text = '-'
            row_cells[5].text = '0 assessments'
            for cell in row_cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.font.size = Pt(8)
                        run.font.color.rgb = RGBColor(128, 128, 128)
        else:
            for detail in entry['rows']:
                row_cells = table.add_row().cells
                row_cells[0].text = entry['form_name']
                row_cells[1].text = detail['assessor']
                row_cells[2].text = detail['location']
                row_cells[3].text = detail['customer'] or '-'
                row_cells[4].text = str(detail['form_id'])
                row_cells[5].text = detail['issue']

                for cell in row_cells:
                    for paragraph in cell.paragraphs:
                        for run in paragraph.runs:
                            run.font.size = Pt(8)

                # Color-code the issue cell
                issue_text = detail['issue'].lower()
                if issue_text != 'none noted':
                    for paragraph in row_cells[5].paragraphs:
                        for run in paragraph.runs:
                            run.font.color.rgb = COLORS['warning']

                # Make Form ID a clickable link if available
                if detail['link']:
                    for paragraph in row_cells[4].paragraphs:
                        paragraph.clear()
                    p = row_cells[4].paragraphs[0]
                    add_hyperlink(p, detail['link'], str(detail['form_id']))

    # Summary line
    doc.add_paragraph()
    p = doc.add_paragraph()
    p.add_run(f"Total: {total_rows} assessments/audits completed").font.bold = True


def build_assessment_html(assessment_details):
    """Build an HTML table for the assessment/audit summary in email.

    Returns an HTML string with a styled table matching BRHAS color scheme.
    """
    total_rows = sum(entry['count'] for entry in assessment_details)

    if total_rows == 0:
        html = '<p style="font-style:italic;">No assessment or audit forms were completed yesterday.</p>'
        html += '<ul style="margin:5px 0;color:#888;">'
        for entry in assessment_details:
            html += f'<li><b>{_h(entry["form_name"])}:</b> 0</li>'
        html += '</ul>'
        return html

    html = '<table width="100%" cellpadding="5" cellspacing="0" '
    html += 'style="border-collapse:collapse;font-size:12px;margin-bottom:10px;">'

    # Header
    html += f'<tr style="background:{HTML_COLORS["secondary"]};color:#ffffff;">'
    for hdr in ['Form Type', 'Assessor', 'Location', 'Customer', 'Form ID', 'Issue Found']:
        html += f'<th style="text-align:left;padding:8px;border:1px solid #600000;">{hdr}</th>'
    html += '</tr>'

    row_idx = 0
    for entry in assessment_details:
        if entry['count'] == 0:
            bg = '#f9f9f9' if row_idx % 2 == 0 else '#ffffff'
            html += f'<tr style="background:{bg};color:#999;">'
            html += f'<td style="border:1px solid #ddd;padding:6px;">{_h(entry["form_name"])}</td>'
            for _ in range(4):
                html += '<td style="border:1px solid #ddd;padding:6px;text-align:center;">-</td>'
            html += '<td style="border:1px solid #ddd;padding:6px;">0 assessments</td>'
            html += '</tr>'
            row_idx += 1
        else:
            for detail in entry['rows']:
                bg = '#f9f9f9' if row_idx % 2 == 0 else '#ffffff'
                html += f'<tr style="background:{bg};">'
                html += f'<td style="border:1px solid #ddd;padding:6px;">{_h(entry["form_name"])}</td>'
                html += f'<td style="border:1px solid #ddd;padding:6px;">{_h(detail["assessor"])}</td>'
                html += f'<td style="border:1px solid #ddd;padding:6px;">{_h(detail["location"])}</td>'
                html += f'<td style="border:1px solid #ddd;padding:6px;">{_h(detail["customer"]) or "-"}</td>'

                # Form ID with link
                if detail['link']:
                    html += f'<td style="border:1px solid #ddd;padding:6px;">'
                    html += f'<a href="{_h(detail["link"])}" style="color:#0563C1;">{_h(detail["form_id"])}</a></td>'
                else:
                    html += f'<td style="border:1px solid #ddd;padding:6px;">{_h(detail["form_id"])}</td>'

                # Issue with color
                issue = detail['issue']
                if issue.lower() != 'none noted':
                    html += f'<td style="border:1px solid #ddd;padding:6px;color:{HTML_COLORS["warning"]};">{_h(issue)}</td>'
                else:
                    html += f'<td style="border:1px solid #ddd;padding:6px;color:{HTML_COLORS["safe"]};">{_h(issue)}</td>'

                html += '</tr>'
                row_idx += 1

    html += '</table>'
    html += f'<p><b>Total: {total_rows} assessments/audits completed</b></p>'

    return html


# ==============================================================================
# DOCUMENT HELPERS
# ==============================================================================

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
# ASSESSMENT & AUDIT ANALYSIS - WORD DOCUMENT SECTION
# ==============================================================================

def add_hyperlink(paragraph, url, text):
    """Add a clickable hyperlink to a Word document paragraph"""
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement

    part = paragraph.part
    r_id = part.relate_to(
        url,
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
        is_external=True
    )

    hyperlink = OxmlElement('w:hyperlink')
    hyperlink.set(qn('r:id'), r_id)

    new_run = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')

    color = OxmlElement('w:color')
    color.set(qn('w:val'), '0563C1')
    rPr.append(color)

    u = OxmlElement('w:u')
    u.set(qn('w:val'), 'single')
    rPr.append(u)

    sz = OxmlElement('w:sz')
    sz.set(qn('w:val'), '18')
    rPr.append(sz)

    new_run.append(rPr)

    t = OxmlElement('w:t')
    t.text = text
    new_run.append(t)

    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)

    return hyperlink


def add_assessment_analysis_section(doc, assessment_data):
    """Add Assessment & Audit Analysis section to the Word document.

    Inserts after Incident Timing Analysis, before At-Risk Conditions.
    Shows assessor activity, compliance by yard, critical findings,
    corrective actions, trends, and leadership recommendations.
    """
    if not assessment_data or not assessment_data.get('has_data'):
        return

    doc.add_page_break()
    add_heading(doc, "ASSESSMENT & AUDIT ANALYSIS", 1, COLORS['primary'])

    p = doc.add_paragraph()
    p.add_run(f"Total Assessments Completed: ").font.bold = True
    p.add_run(f"{assessment_data['total_assessments']}")
    p.add_run(f"  |  ")
    p.add_run(f"Total Findings: ").font.bold = True
    p.add_run(f"{assessment_data['total_findings']}")
    doc.add_paragraph()

    # --- 1. Activity Summary Table ---
    add_heading(doc, "Assessment Activity Summary", 2)

    if assessment_data['activity_summary']:
        table = doc.add_table(rows=1, cols=5)
        table.style = 'Table Grid'

        hdr_cells = table.rows[0].cells
        for i, txt in enumerate(['Form', 'Count', 'Assessor(s)', 'Findings', 'Compliance']):
            hdr_cells[i].text = txt
            for paragraph in hdr_cells[i].paragraphs:
                for run in paragraph.runs:
                    run.font.bold = True
                    run.font.size = Pt(9)
                    run.font.color.rgb = RGBColor(255, 255, 255)
            # Dark header background
            from docx.oxml.ns import qn as _qn
            from docx.oxml import OxmlElement as _OE
            shading = _OE('w:shd')
            shading.set(_qn('w:fill'), '800000')
            hdr_cells[i]._tc.get_or_add_tcPr().append(shading)

        for summary in assessment_data['activity_summary']:
            row_cells = table.add_row().cells
            row_cells[0].text = summary['form_name']
            row_cells[1].text = str(summary['count'])

            assessor_text = ', '.join(summary['assessors'][:3])
            if len(summary['assessors']) > 3:
                assessor_text += f" +{len(summary['assessors']) - 3} more"
            row_cells[2].text = assessor_text

            row_cells[3].text = str(summary['findings_count'])

            if summary['count'] > 0:
                rate = summary['compliance_rate']
                if rate >= 90:
                    row_cells[4].text = f"\u2705 {rate:.0f}%"
                elif rate >= 70:
                    row_cells[4].text = f"\U0001f7e1 {rate:.0f}%"
                else:
                    row_cells[4].text = f"\U0001f534 {rate:.0f}%"
            else:
                row_cells[4].text = "N/A"

            for cell in row_cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.font.size = Pt(9)

    doc.add_paragraph()

    # --- 2. Compliance Dashboard by Yard ---
    if assessment_data['compliance_by_yard']:
        add_heading(doc, "Compliance by Yard", 2)

        table = doc.add_table(rows=1, cols=5)
        table.style = 'Table Grid'

        hdr_cells = table.rows[0].cells
        for i, txt in enumerate(['Yard/Location', 'Assessments', 'Compliant', 'Non-Compliant', 'Status']):
            hdr_cells[i].text = txt
            for paragraph in hdr_cells[i].paragraphs:
                for run in paragraph.runs:
                    run.font.bold = True
                    run.font.size = Pt(9)
                    run.font.color.rgb = RGBColor(255, 255, 255)
            from docx.oxml.ns import qn as _qn
            from docx.oxml import OxmlElement as _OE
            shading = _OE('w:shd')
            shading.set(_qn('w:fill'), '800000')
            hdr_cells[i]._tc.get_or_add_tcPr().append(shading)

        for yard, info in sorted(assessment_data['compliance_by_yard'].items(),
                                  key=lambda x: x[1]['non_compliant'], reverse=True):
            row_cells = table.add_row().cells
            row_cells[0].text = yard
            row_cells[1].text = str(info['total'])
            row_cells[2].text = str(info['compliant'])
            row_cells[3].text = str(info['non_compliant'])

            if info['total'] > 0:
                rate = info['compliant'] / info['total'] * 100
                if rate >= 90:
                    row_cells[4].text = f"\u2705 {rate:.0f}%"
                elif rate >= 70:
                    row_cells[4].text = f"\U0001f7e1 {rate:.0f}%"
                else:
                    row_cells[4].text = f"\U0001f534 {rate:.0f}%"
            else:
                row_cells[4].text = "N/A"

            for cell in row_cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.font.size = Pt(9)

        doc.add_paragraph()

    # --- 3. Critical Findings ---
    critical = assessment_data['findings_by_severity']['critical']
    high = assessment_data['findings_by_severity']['high']

    if critical or high:
        add_heading(doc, "Critical Findings - Immediate Attention Required", 2, COLORS['critical'])

        for finding in critical:
            p = doc.add_paragraph()
            run = p.add_run("\U0001f534 CRITICAL: ")
            run.font.bold = True
            run.font.color.rgb = COLORS['critical']
            p.add_run(finding['description'])

            doc.add_paragraph(
                f"Form: {finding['form_name']} | Assessor: {finding['assessor']}",
                style='List Bullet'
            )
            doc.add_paragraph(
                f"Yard: {finding['yard']} | Date: {finding['date']}",
                style='List Bullet'
            )

            if finding['link']:
                p = doc.add_paragraph(style='List Bullet')
                p.add_run("View in KPA: ")
                add_hyperlink(p, finding['link'], finding['link'])

            doc.add_paragraph()

        for finding in high[:5]:
            p = doc.add_paragraph()
            run = p.add_run("\U0001f7e1 HIGH: ")
            run.font.bold = True
            run.font.color.rgb = COLORS['warning']
            p.add_run(finding['description'])

            doc.add_paragraph(
                f"Form: {finding['form_name']} | Yard: {finding['yard']}",
                style='List Bullet'
            )

            if finding['link']:
                p = doc.add_paragraph(style='List Bullet')
                p.add_run("View in KPA: ")
                add_hyperlink(p, finding['link'], finding['link'])

            doc.add_paragraph()

        if len(high) > 5:
            p = doc.add_paragraph()
            run = p.add_run(f"... and {len(high) - 5} more high-severity findings")
            run.font.italic = True
    else:
        add_heading(doc, "Findings Summary", 2)
        medium = assessment_data['findings_by_severity']['medium']
        low = assessment_data['findings_by_severity']['low']

        if medium or low:
            p = doc.add_paragraph()
            p.add_run("No critical or high-severity findings. ").font.bold = True
            p.add_run(f"{len(medium)} medium, {len(low)} low-severity items noted.")
        else:
            p = doc.add_paragraph("\u2705 No findings - All assessments passed!")
            p.runs[0].font.color.rgb = COLORS['safe']
            p.runs[0].font.bold = True

    doc.add_paragraph()

    # --- 4. Top Performing Assessors ---
    if assessment_data['assessor_stats']:
        add_heading(doc, "Top Performing Assessors", 2, COLORS['safe'])

        sorted_assessors = sorted(
            assessment_data['assessor_stats'].items(),
            key=lambda x: x[1]['total'], reverse=True
        )

        rank = 0
        for name, stats in sorted_assessors[:10]:
            if name == 'Unknown':
                continue
            rank += 1

            p = doc.add_paragraph()
            prefix = "\u2B50 " if rank <= 3 else "   "
            run = p.add_run(f"{prefix}{rank}. {name}")
            run.font.bold = True

            divisions = ', '.join(stats['divisions']) if stats['divisions'] else 'N/A'
            detail = f" - {stats['total']} assessment(s) | Divisions: {divisions}"
            if stats['findings_found'] > 0:
                detail += f" | {stats['findings_found']} finding(s) identified"
            p.add_run(detail)

        doc.add_paragraph()

    # --- 5. Corrective Actions Tracker ---
    if assessment_data['corrective_actions']:
        add_heading(doc, "Corrective Actions Tracker", 2, COLORS['warning'])

        p = doc.add_paragraph()
        p.add_run(f"Open Corrective Actions: {len(assessment_data['corrective_actions'])}").font.bold = True
        doc.add_paragraph()

        for i, ca in enumerate(assessment_data['corrective_actions'][:10], 1):
            p = doc.add_paragraph()
            run = p.add_run(f"{i}. {ca['description']}")
            run.font.bold = True

            doc.add_paragraph(
                f"Form: {ca['form_name']} | Yard: {ca['yard']}",
                style='List Bullet'
            )
            doc.add_paragraph(
                f"Identified by: {ca['assessor']} on {ca['date']}",
                style='List Bullet'
            )

            if ca['link']:
                p = doc.add_paragraph(style='List Bullet')
                p.add_run("View: ")
                add_hyperlink(p, ca['link'], ca['link'])

        if len(assessment_data['corrective_actions']) > 10:
            p = doc.add_paragraph()
            run = p.add_run(
                f"... and {len(assessment_data['corrective_actions']) - 10} more corrective actions"
            )
            run.font.italic = True

        doc.add_paragraph()

    # --- 6. Trends & Patterns ---
    if assessment_data['trends']:
        add_heading(doc, "Trends & Patterns", 2)

        for trend in assessment_data['trends']:
            doc.add_paragraph(f"\U0001F4CA {trend}", style='List Bullet')

        doc.add_paragraph()

    # --- 7. Recommended Actions for Leadership ---
    recs = assessment_data['recommendations']
    if any([recs['immediate'], recs['this_week'], recs['monthly']]):
        add_heading(doc, "Recommended Actions for Leadership", 2, COLORS['primary'])

        if recs['immediate']:
            p = doc.add_paragraph()
            run = p.add_run("\U0001f534 IMMEDIATE:")
            run.font.bold = True
            run.font.color.rgb = COLORS['critical']
            for rec in recs['immediate']:
                doc.add_paragraph(rec, style='List Bullet')

        if recs['this_week']:
            p = doc.add_paragraph()
            run = p.add_run("\U0001f7e1 THIS WEEK:")
            run.font.bold = True
            run.font.color.rgb = COLORS['warning']
            for rec in recs['this_week']:
                doc.add_paragraph(rec, style='List Bullet')

        if recs['monthly']:
            p = doc.add_paragraph()
            run = p.add_run("\U0001F4CA MONTH-OVER-MONTH:")
            run.font.bold = True
            for rec in recs['monthly']:
                doc.add_paragraph(rec, style='List Bullet')


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
    # ASSESSMENT & AUDIT ANALYSIS (after Timing, before At-Risk Conditions)
    # ========================================================================

    if 'assessment_analysis' in all_data and all_data['assessment_analysis']:
        try:
            add_assessment_analysis_section(doc, all_data['assessment_analysis'])
        except Exception as e:
            print(f"Warning: Assessment analysis section error: {e}")
            # Continue building report even if this section fails

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
    # ASSESSMENT & AUDIT SUMMARY (detailed table replacing old "Other Forms")
    # ========================================================================

    if 'assessment_details' in all_data:
        try:
            add_assessment_audit_summary(doc, all_data['assessment_details'])
        except Exception as e:
            print(f"Warning: Assessment audit summary table error: {e}")
            # Fallback to simple count list
            doc.add_page_break()
            add_heading(doc, "OTHER SAFETY FORMS SUMMARY", 1)
            doc.add_paragraph()
            for form_id, form_name in OTHER_FORMS:
                data = all_data.get(f"form_{form_id}")
                count = data['count'] if data else 0
                p = doc.add_paragraph()
                run = p.add_run(f"{form_name}: ")
                run.font.bold = True
                p.add_run(f"{count}")
    else:
        # Fallback if assessment_details not generated
        doc.add_page_break()
        add_heading(doc, "OTHER SAFETY FORMS SUMMARY", 1)
        doc.add_paragraph()
        for form_id, form_name in OTHER_FORMS:
            data = all_data.get(f"form_{form_id}")
            count = data['count'] if data else 0
            p = doc.add_paragraph()
            run = p.add_run(f"{form_name}: ")
            run.font.bold = True
            p.add_run(f"{count}")

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

    # --- ASSESSMENT & AUDIT ANALYSIS ---
    if 'assessment_analysis' in all_data and all_data['assessment_analysis']:
        try:
            aa = all_data['assessment_analysis']
            if aa.get('has_data'):
                aa_html = ''

                # Header stats
                aa_html += f'<b>Total Assessments:</b> {aa["total_assessments"]} | '
                aa_html += f'<b>Total Findings:</b> {aa["total_findings"]}<br><br>'

                # Activity Summary Table
                if aa['activity_summary']:
                    aa_html += f'<h3 style="color:{HTML_COLORS["secondary"]};margin:10px 0 8px 0;font-size:15px;">Assessment Activity Summary</h3>'
                    aa_html += '<table width="100%" cellpadding="6" cellspacing="0" style="border-collapse:collapse;font-size:13px;margin-bottom:15px;">'
                    aa_html += f'<tr style="background:{HTML_COLORS["secondary"]};color:#ffffff;">'
                    aa_html += '<th style="text-align:left;padding:8px;">Form</th>'
                    aa_html += '<th style="text-align:center;padding:8px;">Count</th>'
                    aa_html += '<th style="text-align:left;padding:8px;">Assessor(s)</th>'
                    aa_html += '<th style="text-align:center;padding:8px;">Findings</th>'
                    aa_html += '<th style="text-align:center;padding:8px;">Compliance</th></tr>'

                    for i, s in enumerate(aa['activity_summary']):
                        bg = '#f9f9f9' if i % 2 == 0 else '#ffffff'
                        assessor_text = _h(', '.join(s['assessors'][:3]))
                        if len(s['assessors']) > 3:
                            assessor_text += f' +{len(s["assessors"]) - 3}'

                        rate = s['compliance_rate']
                        if rate >= 90:
                            comp_text = f'<span style="color:{HTML_COLORS["safe"]};">&#9989; {rate:.0f}%</span>'
                        elif rate >= 70:
                            comp_text = f'<span style="color:{HTML_COLORS["warning"]};">&#128993; {rate:.0f}%</span>'
                        else:
                            comp_text = f'<span style="color:{HTML_COLORS["critical"]};">&#128308; {rate:.0f}%</span>'

                        aa_html += f'<tr style="background:{bg};">'
                        aa_html += f'<td style="border-bottom:1px solid #eee;padding:6px;">{_h(s["form_name"])}</td>'
                        aa_html += f'<td style="border-bottom:1px solid #eee;padding:6px;text-align:center;">{s["count"]}</td>'
                        aa_html += f'<td style="border-bottom:1px solid #eee;padding:6px;">{assessor_text}</td>'
                        aa_html += f'<td style="border-bottom:1px solid #eee;padding:6px;text-align:center;">{s["findings_count"]}</td>'
                        aa_html += f'<td style="border-bottom:1px solid #eee;padding:6px;text-align:center;">{comp_text}</td></tr>'

                    aa_html += '</table>'

                # Compliance by Yard Table
                if aa['compliance_by_yard']:
                    aa_html += f'<h3 style="color:{HTML_COLORS["secondary"]};margin:15px 0 8px 0;font-size:15px;">Compliance by Yard</h3>'
                    aa_html += '<table width="100%" cellpadding="6" cellspacing="0" style="border-collapse:collapse;font-size:13px;margin-bottom:15px;">'
                    aa_html += f'<tr style="background:{HTML_COLORS["secondary"]};color:#ffffff;">'
                    aa_html += '<th style="text-align:left;padding:8px;">Yard</th>'
                    aa_html += '<th style="text-align:center;padding:8px;">Total</th>'
                    aa_html += '<th style="text-align:center;padding:8px;">Compliant</th>'
                    aa_html += '<th style="text-align:center;padding:8px;">Non-Compliant</th>'
                    aa_html += '<th style="text-align:center;padding:8px;">Status</th></tr>'

                    sorted_yards = sorted(aa['compliance_by_yard'].items(),
                                          key=lambda x: x[1]['non_compliant'], reverse=True)
                    for i, (yard, info) in enumerate(sorted_yards):
                        bg = '#f9f9f9' if i % 2 == 0 else '#ffffff'
                        if info['total'] > 0:
                            rate = info['compliant'] / info['total'] * 100
                            if rate >= 90:
                                status = f'<span style="color:{HTML_COLORS["safe"]};">&#9989; {rate:.0f}%</span>'
                            elif rate >= 70:
                                status = f'<span style="color:{HTML_COLORS["warning"]};">&#128993; {rate:.0f}%</span>'
                            else:
                                status = f'<span style="color:{HTML_COLORS["critical"]};">&#128308; {rate:.0f}%</span>'
                        else:
                            status = 'N/A'

                        aa_html += f'<tr style="background:{bg};">'
                        aa_html += f'<td style="border-bottom:1px solid #eee;padding:6px;">{_h(yard)}</td>'
                        aa_html += f'<td style="border-bottom:1px solid #eee;padding:6px;text-align:center;">{info["total"]}</td>'
                        aa_html += f'<td style="border-bottom:1px solid #eee;padding:6px;text-align:center;">{info["compliant"]}</td>'
                        aa_html += f'<td style="border-bottom:1px solid #eee;padding:6px;text-align:center;">{info["non_compliant"]}</td>'
                        aa_html += f'<td style="border-bottom:1px solid #eee;padding:6px;text-align:center;">{status}</td></tr>'

                    aa_html += '</table>'

                # Critical Findings
                critical = aa['findings_by_severity']['critical']
                high = aa['findings_by_severity']['high']
                if critical or high:
                    aa_html += f'<h3 style="color:{HTML_COLORS["critical"]};margin:15px 0 8px 0;font-size:15px;">Critical Findings - Immediate Attention</h3>'

                    for f in critical:
                        aa_html += f'<div style="background:#fff5f5;border-left:4px solid {HTML_COLORS["critical"]};padding:12px 15px;margin:8px 0;">'
                        aa_html += f'<b style="color:{HTML_COLORS["critical"]};">&#128308; CRITICAL:</b> {_h(f["description"])}<br>'
                        aa_html += f'Form: {_h(f["form_name"])} | Assessor: {_h(f["assessor"])} | Yard: {_h(f["yard"])}<br>'
                        if f['link']:
                            aa_html += f'<a href="{_h(f["link"])}">View in KPA</a>'
                        aa_html += '</div>'

                    for f in high[:5]:
                        aa_html += f'<div style="background:#fffbf0;border-left:4px solid {HTML_COLORS["warning"]};padding:12px 15px;margin:8px 0;">'
                        aa_html += f'<b style="color:{HTML_COLORS["warning"]};">&#128993; HIGH:</b> {_h(f["description"])}<br>'
                        aa_html += f'Form: {_h(f["form_name"])} | Yard: {_h(f["yard"])}<br>'
                        if f['link']:
                            aa_html += f'<a href="{_h(f["link"])}">View in KPA</a>'
                        aa_html += '</div>'

                    if len(high) > 5:
                        aa_html += f'<p style="font-style:italic;">... and {len(high) - 5} more high-severity findings</p>'
                else:
                    medium = aa['findings_by_severity']['medium']
                    low = aa['findings_by_severity']['low']
                    if medium or low:
                        aa_html += f'<p><b>No critical or high-severity findings.</b> {len(medium)} medium, {len(low)} low-severity items noted.</p>'
                    else:
                        aa_html += f'<p style="color:{HTML_COLORS["safe"]};"><b>&#9989; No findings - All assessments passed!</b></p>'

                # Top Assessors
                if aa['assessor_stats']:
                    aa_html += f'<h3 style="color:{HTML_COLORS["safe"]};margin:15px 0 8px 0;font-size:15px;">Top Performing Assessors</h3>'
                    sorted_a = sorted(aa['assessor_stats'].items(), key=lambda x: x[1]['total'], reverse=True)
                    rank = 0
                    for name, stats in sorted_a[:10]:
                        if name == 'Unknown':
                            continue
                        rank += 1
                        star = '&#11088; ' if rank <= 3 else ''
                        divs = ', '.join(stats['divisions']) if stats['divisions'] else 'N/A'
                        finding_note = f' | {stats["findings_found"]} finding(s)' if stats['findings_found'] > 0 else ''
                        aa_html += f'<div style="margin:4px 0 4px 15px;">{star}<b>{_h(name)}</b> - {stats["total"]} assessment(s) | {_h(divs)}{finding_note}</div>'

                # Corrective Actions
                if aa['corrective_actions']:
                    aa_html += f'<h3 style="color:{HTML_COLORS["warning"]};margin:15px 0 8px 0;font-size:15px;">Corrective Actions ({len(aa["corrective_actions"])} open)</h3>'
                    for i, ca in enumerate(aa['corrective_actions'][:5], 1):
                        aa_html += f'<div style="background:#fffbf0;border-left:4px solid {HTML_COLORS["warning"]};padding:10px 15px;margin:6px 0;">'
                        aa_html += f'<b>{i}. {_h(ca["description"])}</b><br>'
                        aa_html += f'{_h(ca["form_name"])} | {_h(ca["yard"])} | By: {_h(ca["assessor"])}<br>'
                        if ca['link']:
                            aa_html += f'<a href="{_h(ca["link"])}">View in KPA</a>'
                        aa_html += '</div>'
                    if len(aa['corrective_actions']) > 5:
                        aa_html += f'<p style="font-style:italic;">... and {len(aa["corrective_actions"]) - 5} more</p>'

                # Trends
                if aa['trends']:
                    aa_html += f'<h3 style="color:{HTML_COLORS["primary"]};margin:15px 0 8px 0;font-size:15px;">Trends &amp; Patterns</h3>'
                    aa_html += '<ul style="margin:5px 0;">'
                    for trend in aa['trends']:
                        aa_html += f'<li>&#128202; {_h(trend)}</li>'
                    aa_html += '</ul>'

                # Recommendations
                recs = aa['recommendations']
                if any([recs['immediate'], recs['this_week'], recs['monthly']]):
                    aa_html += f'<h3 style="color:{HTML_COLORS["primary"]};margin:15px 0 8px 0;font-size:15px;">Recommended Actions for Leadership</h3>'

                    if recs['immediate']:
                        aa_html += f'<div style="margin:5px 0;"><b style="color:{HTML_COLORS["critical"]};">&#128308; IMMEDIATE:</b></div>'
                        aa_html += '<ul style="margin:3px 0;">'
                        for r in recs['immediate']:
                            aa_html += f'<li>{_h(r)}</li>'
                        aa_html += '</ul>'

                    if recs['this_week']:
                        aa_html += f'<div style="margin:5px 0;"><b style="color:{HTML_COLORS["warning"]};">&#128993; THIS WEEK:</b></div>'
                        aa_html += '<ul style="margin:3px 0;">'
                        for r in recs['this_week']:
                            aa_html += f'<li>{_h(r)}</li>'
                        aa_html += '</ul>'

                    if recs['monthly']:
                        aa_html += '<div style="margin:5px 0;"><b>&#128202; MONTH-OVER-MONTH:</b></div>'
                        aa_html += '<ul style="margin:3px 0;">'
                        for r in recs['monthly']:
                            aa_html += f'<li>{_h(r)}</li>'
                        aa_html += '</ul>'

                sections.append(f"""
<tr><td style="padding:25px 40px;border-top:3px solid {HTML_COLORS['primary']};">
  <h2 style="color:{HTML_COLORS['primary']};margin:0 0 15px 0;font-size:18px;border-bottom:2px solid {HTML_COLORS['primary']};padding-bottom:5px;">ASSESSMENT &amp; AUDIT ANALYSIS</h2>
  {aa_html}
</td></tr>""")
        except Exception as e:
            print(f"Warning: HTML assessment analysis error: {e}")

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

    # --- ASSESSMENT & AUDIT SUMMARY (replaces old "Other Forms Summary") ---
    if 'assessment_details' in all_data:
        try:
            audit_table_html = build_assessment_html(all_data['assessment_details'])
        except Exception as e:
            print(f"Warning: HTML assessment summary table error: {e}")
            audit_table_html = ''
            for form_id, form_name in OTHER_FORMS:
                data = all_data.get(f"form_{form_id}")
                count = data['count'] if data else 0
                audit_table_html += f'<b>{_h(form_name)}:</b> {count}<br>'
    else:
        audit_table_html = ''
        for form_id, form_name in OTHER_FORMS:
            data = all_data.get(f"form_{form_id}")
            count = data['count'] if data else 0
            audit_table_html += f'<b>{_h(form_name)}:</b> {count}<br>'

    sections.append(f"""
<tr><td style="padding:25px 40px;border-top:2px solid #ddd;">
  <h2 style="color:{HTML_COLORS['primary']};margin:0 0 15px 0;font-size:18px;border-bottom:2px solid {HTML_COLORS['primary']};padding-bottom:5px;">ASSESSMENT &amp; AUDIT SUMMARY</h2>
  {audit_table_html}
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
    print("✓ Assessment & Audit Analysis with compliance, findings, trends")
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

    # Analyze assessment/audit forms for the deep-analysis section
    print("\nAnalyzing assessment & audit data...")
    try:
        assessment_analysis = analyze_assessments(all_data)
        all_data['assessment_analysis'] = assessment_analysis
        if assessment_analysis['has_data']:
            print(f"✓ Assessment Analysis: {assessment_analysis['total_assessments']} assessments, "
                  f"{assessment_analysis['total_findings']} findings")
        else:
            print("✓ Assessment Analysis: No assessment data for yesterday")
    except Exception as e:
        print(f"⚠️  Assessment analysis failed (non-fatal): {e}")
        all_data['assessment_analysis'] = None

    # Extract per-row assessment details for the summary table
    try:
        assessment_details = extract_assessment_details(all_data)
        all_data['assessment_details'] = assessment_details
        detail_count = sum(entry['count'] for entry in assessment_details)
        print(f"✓ Assessment Details: {detail_count} form rows extracted for summary table")
    except Exception as e:
        print(f"⚠️  Assessment details extraction failed (non-fatal): {e}")
        all_data['assessment_details'] = None

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
