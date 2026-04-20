"""
DAILY YTD SAFETY STATS
=======================
All-company YTD stats: man-hours from payroll, incidents from KPA.
Produces output/ytd_stats.json with per-company TRIR, DART, days since recordable.

Usage:
    KPA_API_TOKEN=... python daily_ytd_stats.py
"""

import calendar
import csv
import json
import os
import sys
from datetime import datetime, date, timezone
from io import StringIO

import requests

# ==============================================================================
# CONFIG
# ==============================================================================

API_TOKEN = os.environ.get("KPA_API_TOKEN")
API_BASE = "https://api.kpaehs.com/v1"
INCIDENT_FORM_ID = 151622
INC_TYPE_HASH = "nojcquy0tfl9hqih"
SVC_LINE_HASH = "sha7vur5q2l6d6gq"
COMPANY_HASH = "lsx3msa0w9n9edb4"

# ==============================================================================
# MAN-HOURS FROM PAYROLL (Q1 2026 Man Hours.xlsx -- Regular + OT, all employees)
# Source: 2026 Q1 Man Hours.xlsx -- includes salary employees
# Total Q1: 807,718 man-hours across all companies
# ==============================================================================

Q1_MAN_HOURS = {
    "BRHAS": {
        "total": 545770,
        "by_division": {
            "Casing":        310907,
            "Rathole":       102556,
            "Poly Pipe":      42774,
            "Corporate":      30911,
            "Pit Lining":     15142,
            "Fabrication":     6964,
            "Construction":    7918,
            "Shop":            7661,
            "Anchors":         7675,
            "Fencing":         5701,
            "Environmental":   4527,
            "Drilling Tools":  3032,
        },
        # Monthly breakdown (from 1st Quarter Work Hours Report, prorated to Q1 Man Hours totals)
        # March is used for prorating future months
        "monthly": {
            "2026-01": 178567,  # ~32.7% of Q1
            "2026-02": 170786,  # ~31.3% of Q1
            "2026-03": 196417,  # ~36.0% of Q1
        },
    },
    "Permian": {
        "total": 91987,
        "by_division": {"Rentals": 91987},
        "monthly": {"2026-01": 29818, "2026-02": 29482, "2026-03": 32687},
    },
    "BTI": {
        "total": 78685,
        "by_division": {"Trucking": 70536, "Shop": 8149},
        "monthly": {"2026-01": 26632, "2026-02": 26958, "2026-03": 25095},
    },
    "Transcend": {
        "total": 63831,
        "by_division": {
            "TD Rig 20": 15415, "TD Rig 4": 13086, "TD Rig 18": 12949,
            "TD Rig 16": 11813, "Drilling": 9979, "TD Rig 12": 588,
        },
        "monthly": {"2026-01": 20182, "2026-02": 19621, "2026-03": 24028},
    },
    "Valor": {
        "total": 16649,
        "by_division": {"Downhole Tools": 16649},
        "monthly": {"2026-01": 5326, "2026-02": 5611, "2026-03": 5712},
    },
}

# Map KPA incident company names -> dashboard company names
KPA_COMPANY_MAP = {
    "Butch's Rat Hole & Anchor Service Inc": "BRHAS",
    "Butch's Rat Hole & Anchor Service, Inc.": "BRHAS",
    "Butch's Rat Hole & Anchor Service": "BRHAS",
    "Transcend Drilling Company Inc": "Transcend",
    "Transcend Drilling Company, Inc.": "Transcend",
    "Transcend Drilling": "Transcend",
    "Butch's Trucking Inc": "BTI",
    "Butch's Trucking": "BTI",
    "Valor Energy Services": "Valor",
    "Valor Energy Services LLC": "Valor",
    "Permian Equipment Rentals": "Permian",
    "Permian Equipment Rentals LLC": "Permian",
}

# Map KPA service line -> dashboard division (for BRHAS)
KPA_SVC_TO_DIVISION = {
    "Casing": "Casing",
    "Rat Hole": "Rathole", "Rathole": "Rathole",
    "Poly Pipe": "Poly Pipe",
    "Pit Lining": "Pit Lining",
    "Anchor": "Anchors", "Anchors": "Anchors",
    "Construction": "Construction", "Civil": "Construction",
    "Environmental": "Environmental", "Containment": "Environmental",
    "Fencing": "Fencing",
}


# ==============================================================================
# MAN-HOURS COMPUTATION
# ==============================================================================

def compute_ytd_hours(today):
    """Compute YTD man-hours per company using payroll data.

    For Q1 months (Jan-Mar): use actual payroll totals.
    For current month beyond Q1: prorate March hours by days elapsed.
    """
    year = today.year
    result = {}

    for company, data in Q1_MAN_HOURS.items():
        monthly = data["monthly"]
        march_hours = monthly.get("2026-03", 0)
        company_ytd = 0.0
        months_detail = []

        for m in range(1, today.month + 1):
            key = f"{year}-{m:02d}"
            if key in monthly:
                if year == today.year and m == today.month:
                    # Current month within Q1 -- prorate
                    days_in_month = calendar.monthrange(year, m)[1]
                    fraction = today.day / days_in_month
                    prorated = round(monthly[key] * fraction)
                    company_ytd += prorated
                    months_detail.append(f"{key} (prorated: {prorated:,}h)")
                else:
                    company_ytd += monthly[key]
                    months_detail.append(key)
            else:
                # Month beyond Q1 -- estimate from March
                if m == today.month:
                    days_in_month = calendar.monthrange(year, m)[1]
                    fraction = today.day / days_in_month
                    prorated = round(march_hours * fraction)
                    company_ytd += prorated
                    months_detail.append(f"{key} (est: {prorated:,}h)")
                else:
                    company_ytd += march_hours
                    months_detail.append(f"{key} (est: {march_hours:,}h)")

        # Prorate division hours proportionally to match YTD company total
        q1_total = data["total"]
        scale = company_ytd / q1_total if q1_total > 0 else 1.0
        prorated_divs = {}
        for div_name, div_q1 in data["by_division"].items():
            prorated_divs[div_name] = round(div_q1 * scale)

        result[company] = {
            "total": round(company_ytd),
            "by_division": prorated_divs,
            "months": months_detail,
        }

    return result


# ==============================================================================
# LIVE KPA INCIDENT PULL
# ==============================================================================

def pull_all_incidents():
    """Pull all incidents from KPA form 151622. Returns list of row dicts."""
    if not API_TOKEN:
        return None

    payload = {
        "token": API_TOKEN,
        "form_id": INCIDENT_FORM_ID,
        "format": "csv",
    }
    try:
        resp = requests.post(f"{API_BASE}/responses.flat", json=payload, timeout=60)
    except Exception as e:
        print(f"  Warning: KPA incident pull failed: {e}")
        return None

    if not resp.text or not resp.text.strip():
        return []

    reader = csv.DictReader(StringIO(resp.text))
    rows = list(reader)
    # First row is labels header, skip it
    if rows and rows[0].get("date", "") == "Date":
        rows = rows[1:]
    return rows


def parse_incident_date(date_str):
    """Parse KPA date string to date object. Returns None on failure."""
    try:
        return datetime.strptime(date_str[:10], "%Y-%m-%d").date()
    except (ValueError, TypeError):
        return None


def normalize_company(raw_company):
    """Map KPA company name to dashboard company name."""
    raw = raw_company.strip()
    if raw in KPA_COMPANY_MAP:
        return KPA_COMPANY_MAP[raw]
    # Fuzzy fallback
    lower = raw.lower()
    if "butch" in lower and "truck" in lower:
        return "BTI"
    if "transcend" in lower:
        return "Transcend"
    if "valor" in lower:
        return "Valor"
    if "permian" in lower:
        return "Permian"
    if "butch" in lower:
        return "BRHAS"
    return raw


def compute_days_since(incidents, today):
    """Compute days-since-recordable at company and service-line level.

    Returns:
        overall_last_date: date of most recent recordable across all data
        by_company: {company: {last_date, days_since, ytd_count, service_lines: {svc: {last_date, days_since}}}}
        ytd_count: number of recordables in current year
    """
    recordables = []
    for row in incidents:
        inc_type = row.get(INC_TYPE_HASH, "")
        if "Recordable" not in inc_type:
            continue
        d = parse_incident_date(row.get("date", ""))
        if not d:
            continue
        raw_co = row.get(COMPANY_HASH, "").strip() or "Unknown"
        recordables.append({
            "date": d,
            "company": normalize_company(raw_co),
            "service_line": row.get(SVC_LINE_HASH, "").strip() or "Unknown",
        })

    # YTD count (current year only)
    ytd_count = sum(1 for r in recordables if r["date"].year == today.year)

    # Overall last recordable
    overall_last = max((r["date"] for r in recordables), default=None)

    # By company -> by service line
    by_company = {}
    for r in recordables:
        co = r["company"]
        sl = r["service_line"]
        is_ytd = r["date"].year == today.year

        if co not in by_company:
            by_company[co] = {"last_date": r["date"], "ytd_count": 0, "service_lines": {}}
        else:
            if r["date"] > by_company[co]["last_date"]:
                by_company[co]["last_date"] = r["date"]

        if is_ytd:
            by_company[co]["ytd_count"] += 1

        if sl not in by_company[co]["service_lines"]:
            by_company[co]["service_lines"][sl] = {"last_date": r["date"], "ytd_count": 0}
        else:
            if r["date"] > by_company[co]["service_lines"][sl]["last_date"]:
                by_company[co]["service_lines"][sl]["last_date"] = r["date"]

        if is_ytd:
            by_company[co]["service_lines"][sl]["ytd_count"] += 1

    # Compute days_since for each level
    for co, co_data in by_company.items():
        co_data["days_since"] = (today - co_data["last_date"]).days
        co_data["last_date"] = co_data["last_date"].isoformat()
        for sl, sl_data in co_data["service_lines"].items():
            sl_data["days_since"] = (today - sl_data["last_date"]).days
            sl_data["last_date"] = sl_data["last_date"].isoformat()

    return overall_last, by_company, ytd_count


# ==============================================================================
# MAIN
# ==============================================================================

def main():
    today = date.today()

    print("=" * 50)
    print("  DAILY YTD SAFETY STATS (ALL COMPANIES)")
    print("=" * 50)

    print(f"\n[1/3] Computing YTD man-hours from payroll...")
    hours_by_company = compute_ytd_hours(today)
    ytd_hours = sum(c["total"] for c in hours_by_company.values())
    print(f"  Total YTD man-hours: {ytd_hours:,.0f}")
    for co, data in sorted(hours_by_company.items()):
        print(f"    {co}: {data['total']:,.0f}")

    print(f"\n[2/3] Pulling incidents from KPA (live)...")
    incidents = pull_all_incidents()

    if incidents is None:
        print("  Skipped (no API token or request failed)")
        overall_last = None
        by_company = {}
        ytd_rec = 0
    else:
        overall_last, by_company, ytd_rec = compute_days_since(incidents, today)
        print(f"  Total incidents pulled: {len(incidents)}")
        print(f"  YTD recordables (live): {ytd_rec}")

        if overall_last:
            print(f"  Last recordable (all companies): {overall_last.isoformat()} ({(today - overall_last).days} days ago)")
        for co, co_data in sorted(by_company.items()):
            print(f"    {co}: {co_data['last_date']} ({co_data['days_since']} days, {co_data['ytd_count']} YTD)")
            for sl, sl_data in sorted(co_data["service_lines"].items()):
                print(f"      {sl}: {sl_data['last_date']} ({sl_data['days_since']} days)")

    print(f"\n[3/3] Computing YTD stats...")

    # Overall stats
    ytd_trir = round(ytd_rec * 200_000 / ytd_hours, 2) if ytd_hours > 0 else None
    days_rec = (today - overall_last).days if overall_last else None
    days_lti = days_rec  # All recordables treated as DART for now
    ytd_dart_cases = ytd_rec
    ytd_dart = round(ytd_dart_cases * 200_000 / ytd_hours, 2) if ytd_hours > 0 else None

    # Per-company stats
    stats_by_company = {}
    for co, hrs_data in hours_by_company.items():
        co_hrs = hrs_data["total"]
        co_inc = by_company.get(co, {})
        co_rec = co_inc.get("ytd_count", 0)
        co_trir = round(co_rec * 200_000 / co_hrs, 2) if co_hrs > 0 else 0.0
        co_dart = co_trir  # All recordables = DART cases
        co_days = co_inc.get("days_since", None)
        co_last = co_inc.get("last_date", None)

        # Normalize service line names to dashboard division names
        raw_svc = co_inc.get("service_lines", {})
        norm_svc = {}
        for sl_name, sl_data in raw_svc.items():
            mapped = KPA_SVC_TO_DIVISION.get(sl_name, sl_name)
            if mapped in norm_svc:
                # Merge: keep most recent date, sum ytd_count
                existing = norm_svc[mapped]
                if sl_data.get("last_date", "") > existing.get("last_date", ""):
                    existing["last_date"] = sl_data["last_date"]
                    existing["days_since"] = sl_data["days_since"]
                existing["ytd_count"] = existing.get("ytd_count", 0) + sl_data.get("ytd_count", 0)
            else:
                norm_svc[mapped] = dict(sl_data)

        stats_by_company[co] = {
            "man_hours": co_hrs,
            "recordables": co_rec,
            "trir": co_trir,
            "dart": co_dart,
            "days_since_recordable": co_days,
            "last_recordable_date": co_last,
            "by_division": hrs_data["by_division"],
            "service_lines": norm_svc,
        }

    json_data = {
        "report_date": today.isoformat(),
        "generated_at": datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S"),
        "ytd_trir": ytd_trir,
        "ytd_dart": ytd_dart,
        "ytd_dart_cases": ytd_dart_cases,
        "ytd_recordables": ytd_rec,
        "ytd_man_hours": round(ytd_hours),
        "days_since_lti": days_lti,
        "days_since_recordable": days_rec,
        "last_lti_date": overall_last.isoformat() if overall_last else None,
        "last_recordable_date": overall_last.isoformat() if overall_last else None,
        "by_company": by_company,
        "stats_by_company": stats_by_company,
        "man_hours_source": "payroll_q1_2026",
    }

    os.makedirs("output", exist_ok=True)
    out = os.path.join("output", "ytd_stats.json")
    with open(out, "w", encoding="utf-8") as f:
        json.dump(json_data, f, indent=2, default=str)

    print(f"\n  YTD TRIR: {ytd_trir}")
    print(f"  YTD DART: {ytd_dart}")
    print(f"  Days since recordable: {days_rec}")
    print(f"  JSON written: {out}")
    for co, st in sorted(stats_by_company.items()):
        print(f"    {co}: {st['man_hours']:,}h, TRIR={st['trir']}, rec={st['recordables']}, days={st['days_since_recordable']}")
    print("  Done.")


if __name__ == "__main__":
    main()
