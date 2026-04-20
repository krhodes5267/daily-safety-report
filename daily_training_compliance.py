"""
DAILY TRAINING COMPLIANCE TRACKER
===================================
Pulls ALL company employee training compliance from KPA and writes
output/training_compliance.json with per-company and per-yard breakdowns.

Usage:
    KPA_API_TOKEN=... python daily_training_compliance.py
"""

import json
import os
import sys
import time
from collections import Counter
from datetime import datetime, timezone

import requests

# ==============================================================================
# CONFIG
# ==============================================================================

API_TOKEN = os.environ.get("KPA_API_TOKEN")
if not API_TOKEN:
    print("ERROR: KPA_API_TOKEN environment variable is not set.")
    sys.exit(1)

API_BASE = "https://api.kpaehs.com/v1"

# LOB ID -> dashboard company mapping
LOB_TO_COMPANY = {
    # BRHAS divisions
    "6009f696823a6201bbc9b056": "BRHAS",  # Casing
    "6009f696823a6201bbc9b04a": "BRHAS",  # Rathole
    "6009f696823a6201bbc9b051": "BRHAS",  # Anchors
    "6009f696823a6201bbc9b053": "BRHAS",  # Pit Lining
    "6009f696823a6201bbc9b04d": "BRHAS",  # Poly Pipe
    "609452db13b68a005495157f": "BRHAS",  # Fencing
    "6148833d9e0bc6004e070765": "BRHAS",  # Construction/Civil
    "6152f95cb365e00022f9ecb5": "BRHAS",  # Environmental/Containment
    "6009f696823a6201bbc9b055": "BRHAS",  # Shop
    "5d24f5d8adf0d700172f96aa": "BRHAS",  # Corporate
    "6009f696823a6201bbc9b05f": "BRHAS",  # Fabrication (under Rathole)
    "65c0eca2cd54c200664ff534": "BRHAS",  # Water Trucking (under Rathole)
    # BTI
    "6009f696823a6201bbc9b05b": "BTI",    # Trucking
    # Transcend
    "6009f696823a6201bbc9b058": "Transcend",  # Drilling (general)
    "653a5c9abadfcb0013cf96d9": "Transcend",  # TD RIG 16
    "642d7f30ab6b2d516d3c69d8": "Transcend",  # TD RIG 18
    "65c0ed11323ea80063d7f38c": "Transcend",  # TD RIG 2
    "65c0ed07f2772b002b4bc0a0": "Transcend",  # TD RIG 4
    "65d62c63e73e7d0059a1bb1a": "Transcend",  # TD Rig 12
    "66426f23067e2c00252514fc": "Transcend",  # TD Rig 20
    "645a3565888c9541c069075a": "Transcend",  # TD RIG 32
    # Valor
    "6009f696823a6201bbc9b04b": "Valor",  # Downhole Tools
    "6009f696823a6201bbc9b052": "BRHAS",  # Drilling Tools (BRHAS)
    # Permian
    "5d166c18efd5700017316462": "Permian",  # Rentals (PER)
}

# LOB ID -> division (all companies)
LOB_TO_DIVISION = {
    # BRHAS
    "6009f696823a6201bbc9b056": "Casing",
    "6009f696823a6201bbc9b04a": "Rathole",
    "6009f696823a6201bbc9b051": "Anchors",
    "6009f696823a6201bbc9b053": "Pit Lining",
    "6009f696823a6201bbc9b04d": "Poly Pipe",
    "609452db13b68a005495157f": "Fencing",
    "6148833d9e0bc6004e070765": "Construction",
    "6152f95cb365e00022f9ecb5": "Environmental",
    "6009f696823a6201bbc9b055": "Shop",
    "5d24f5d8adf0d700172f96aa": "Corporate",
    "6009f696823a6201bbc9b05f": "Rathole",      # Fabrication -> Rathole
    "65c0eca2cd54c200664ff534": "Rathole",      # Water Trucking -> Rathole
    # BTI
    "6009f696823a6201bbc9b05b": "Trucking",
    # Transcend
    "6009f696823a6201bbc9b058": "Drilling",
    "653a5c9abadfcb0013cf96d9": "TD Rig 16",
    "642d7f30ab6b2d516d3c69d8": "TD Rig 18",
    "65c0ed11323ea80063d7f38c": "TD Rig 2",
    "65c0ed07f2772b002b4bc0a0": "TD Rig 4",
    "65d62c63e73e7d0059a1bb1a": "TD Rig 12",
    "66426f23067e2c00252514fc": "TD Rig 20",
    "645a3565888c9541c069075a": "TD Rig 32",
    # Valor
    "6009f696823a6201bbc9b04b": "Downhole Tools",
    "6009f696823a6201bbc9b052": "Drilling Tools",
    # Permian
    "5d166c18efd5700017316462": "Rentals",
}

CASING_LOB_ID = "6009f696823a6201bbc9b056"

# KPA field office ID -> location mapping (all companies)
FO_YARD_MAP = {
    # Casing
    "671017668ee2a10019b2f7f0": "Midland",       # Midland Yukon
    "5d166f31efd5700017316be4": "Midland",        # Midland
    "5d2cf0da6f00c900179d969b": "Kilgore",
    "6009f55901f3bb0142271514": "Hobbs",
    "5d166ec1d57b5c00178cfab0": "Jourdanton",
    "5cddbce7cc6e850017e270a1": "Bryan",
    "6009f55901f3bb0142271515": "Laredo",
    "6009f55901f3bb014227151c": "San Angelo",
    "671017898ee2a10019b2fc9a": "Midland",        # Overhead BRHAS
    # BRHAS other
    "6710171f58245c0012f8101a": "Midland",        # Midland 1788 BRHAS
    "671016ec58245c0012f803ba": "Levelland",      # Levelland Yard BRHAS
    "6009f55901f3bb0142271517": "Levelland",      # Levelland
    "671017772ff29400126b36a4": "Odessa",         # Odessa CR 100 BRHAS
    "67d99c42fd81f039fee14511": "Odessa",         # Odessa 24th Street
    "6009f55901f3bb014227151e": "Seminole",
    "6009f55901f3bb0142271510": "Barstow",
    "6009f55901f3bb0142271519": "North Dakota",
    "6009f55901f3bb014227151a": "Pennsylvania",
    "671016c658245c0012f7fcca": "North Dakota",   # Dickinson, ND
    "5d2cf09b6f00c900179d9625": "Oklahoma",       # Shawnee, OK
    "5d166f00289e22001747cf5a": "Pennsylvania",    # Towanda, PA
    "628247029a986d003f82db67": "Ohio",            # Wintersville, OH
    "65c0eebc323ea80063d80acd": "Levelland",      # WTX Cement -> Levelland
    "6009f55901f3bb014227151f": "Jourdanton",     # PLEASANTON - BRHAS
    "6009f55901f3bb014227151d": "Seguin",
    "6009f55901f3bb0142271512": "Corporate",      # CORPORATE (Butch's)
    "5d237a2c081403001783219f": "Corporate",
    "600ade638a5ff10031d6e00c": "Midland",        # CASING MIDLAND APPLICATION
    "6009f55901f3bb0142271518": "Midland",        # MIDLAND (Butch's)
    "6009f55901f3bb0142271523": "Levelland",      # WATER TRUCKS LEVELLAND
    # BTI
    "6710172e58245c0012f81325": "Midland",        # Midland BTI
    "671016fcebfbea001969d58f": "Levelland",      # Levelland Yard BTI
    "671017988ccbba001993842e": "Midland",        # Overhead BTI
    # Transcend
    "67101751fd1319001257edbf": "Midland",        # Midland Transcend
    "671017a6ebfbea001969e8f8": "Midland",        # Overhead Transcend
    "6009f55901f3bb0142271520": "Midland",        # Transcend
    # Valor
    "6710170b8ee2a10019b2e7e0": "Levelland",     # Levelland Yard Valor
    "6009f55901f3bb0142271521": "Levelland",      # VALOR LEVELLAND
    # Permian
    "6710173e58245c0012f81659": "Midland",        # Midland PER
    "6913ad0e135b6c0010ec2c77": "Lubbock",        # Lubbock PER
    "6009f55901f3bb014227151b": "Midland",        # Permian
}

YARD_ORDER = ["Midland", "Bryan", "Kilgore", "Hobbs", "Jourdanton", "Laredo"]
COMPANY_ORDER = ["BRHAS", "BTI", "Transcend", "Permian", "Valor"]


# ==============================================================================
# KPA API HELPERS
# ==============================================================================

def call_kpa_json_single(endpoint, data_key):
    url = f"{API_BASE}/{endpoint}"
    payload = {"token": API_TOKEN}
    try:
        r = requests.post(url, json=payload, timeout=120)
        data = json.loads(r.text.strip())
        return data.get(data_key, [])
    except Exception as e:
        print(f"  KPA API error ({endpoint}): {e}")
        return []


def call_kpa_json_paginated(endpoint, data_key="employees", max_pages=50):
    all_rows = []
    page = 1
    rate_limit_retries = 0
    while True:
        payload = {"token": API_TOKEN, "limit": 500, "page": page}
        url = f"{API_BASE}/{endpoint}"
        try:
            r = requests.post(url, json=payload, timeout=120)
            text = r.text.strip()
            if r.status_code == 429 or "rate_limit" in text:
                rate_limit_retries += 1
                if rate_limit_retries > 5:
                    print(f"  WARNING: Rate limited {rate_limit_retries} times, stopping.")
                    break
                print(f"  Rate limited, waiting 30s (attempt {rate_limit_retries}/5)...")
                time.sleep(30)
                continue
            rate_limit_retries = 0
            data = json.loads(text)
        except Exception as e:
            print(f"  KPA API error ({endpoint}): {e}")
            break
        items = data.get(data_key, [])
        if not items:
            break
        all_rows.extend(items)
        last_page = data.get("paging", {}).get("last_page", 1)
        if page >= last_page or page >= max_pages:
            break
        page += 1
        time.sleep(1.5)
    return all_rows


# ==============================================================================
# USER CLASSIFICATION
# ==============================================================================

def classify_user(user):
    """Determine company and division for a KPA user based on LOB IDs."""
    lobs = user.get("lineOfBusiness_id", [])
    if not isinstance(lobs, list):
        lobs = [lobs] if lobs else []

    company = None
    division = None

    for lob_id in lobs:
        co = LOB_TO_COMPANY.get(lob_id)
        if co:
            company = co
            div = LOB_TO_DIVISION.get(lob_id)
            if div:
                division = div
            break

    return company, division


def get_user_yard(user):
    """Get Casing yard from field office IDs."""
    fo_ids = user.get("fieldOffice_id", [])
    if isinstance(fo_ids, str):
        fo_ids = [fo_ids]
    elif not isinstance(fo_ids, list):
        fo_ids = []
    for fid in fo_ids:
        yard = FO_YARD_MAP.get(fid, "")
        if yard:
            return yard
    return "Unassigned"


# ==============================================================================
# MAIN
# ==============================================================================

def main():
    print("=" * 50)
    print("  DAILY TRAINING COMPLIANCE TRACKER (ALL COMPANIES)")
    print("=" * 50)

    # Step 1: Get all employees
    print("\n[1/4] Fetching KPA users...")
    all_users = call_kpa_json_single("users.list", "users")

    user_info = {}  # uid -> {name, company, division, yard}
    company_counts = Counter()

    for u in all_users:
        uid = u.get("id", "")
        if not uid or u.get("terminationDate"):
            continue
        first = u.get("firstname", "")
        last = u.get("lastname", "")
        company, division = classify_user(u)
        if not company:
            continue  # Skip users not in any known LOB

        yard = get_user_yard(u)

        # Valor is all Levelland
        if company == "Valor":
            yard = "Levelland"

        user_info[uid] = {
            "name": f"{first} {last}".strip(),
            "company": company,
            "division": division or "",
            "yard": yard,
        }
        company_counts[company] += 1

    print(f"  Classified employees: {len(user_info)} (from {len(all_users)} total)")
    for co in COMPANY_ORDER:
        if company_counts.get(co, 0) > 0:
            print(f"    {co}: {company_counts[co]}")

    # Step 2: Get training programs
    print("\n[2/4] Fetching training programs...")
    programs = call_kpa_json_single("trainings.v2.list", "trainings")
    training_lookup = {}
    training_created = {}
    for p in programs:
        tid = p.get("id")
        name = p.get("title", p.get("name", ""))
        if tid is not None and name:
            training_lookup[tid] = name
            created_ms = p.get("created_on")
            if created_ms:
                training_created[tid] = created_ms
    print(f"  Training programs: {len(training_lookup)}")

    # Step 3: Get employee training status
    print("\n[3/4] Fetching training employee status...")
    all_status = call_kpa_json_paginated(
        "training-employee-status.list", data_key="employees"
    )
    # Filter to classified users only
    known_status = [r for r in all_status if r.get("m_user_id") in user_info]
    print(f"  Training records: {len(known_status)} (from {len(all_status)} total)")

    # Step 4: Process
    print("\n[4/4] Processing compliance data...")
    now_ms = int(datetime.now(tz=timezone.utc).timestamp() * 1000)
    employees = []

    for row in known_status:
        uid = row.get("m_user_id", "")
        info = user_info.get(uid)
        if not info:
            continue

        incomplete_ids = row.get("incomplete_training_ids", []) or []
        complete_ids = row.get("complete_training_ids", []) or []

        incomplete_names = [training_lookup.get(tid, f"Program #{tid}")
                            for tid in incomplete_ids]

        total = len(incomplete_ids) + len(complete_ids)
        pct = round(len(complete_ids) / total * 100) if total > 0 else 100

        status = "Complete"
        if pct < 100:
            status = "Overdue" if row.get("status") == "overdue" else "In Progress"

        max_days_since = 0
        for tid in incomplete_ids:
            created_ms = training_created.get(tid, 0)
            if created_ms > 0:
                days = (now_ms - created_ms) // (1000 * 86400)
                if days > max_days_since:
                    max_days_since = days

        employees.append({
            "employee_name": info["name"],
            "company": info["company"],
            "division": info["division"],
            "yard": info["yard"],
            "percent_complete": pct,
            "incomplete_count": len(incomplete_ids),
            "complete_count": len(complete_ids),
            "total_assigned": total,
            "incomplete_training_names": incomplete_names[:5],
            "status": status,
            "days_since_assignment": max_days_since,
        })

    # Overall stats
    total_emp = len(employees)
    compliant = sum(1 for e in employees if e["percent_complete"] >= 100)
    overdue = sum(1 for e in employees if e["status"] == "Overdue")
    overall_pct = round(compliant / total_emp * 100, 1) if total_emp > 0 else 0

    # By company
    by_company = {}
    for co in COMPANY_ORDER:
        co_emps = [e for e in employees if e["company"] == co]
        co_compliant = sum(1 for e in co_emps if e["percent_complete"] >= 100)
        by_company[co] = {
            "total": len(co_emps),
            "compliant": co_compliant,
            "pct": round(co_compliant / len(co_emps) * 100, 1) if co_emps else 0,
        }

    # By division (BRHAS subdivisions)
    by_division = {}
    brhas_divisions = set(e["division"] for e in employees if e["company"] == "BRHAS" and e["division"])
    for div in sorted(brhas_divisions):
        div_emps = [e for e in employees if e["division"] == div]
        div_compliant = sum(1 for e in div_emps if e["percent_complete"] >= 100)
        by_division[div] = {
            "total": len(div_emps),
            "compliant": div_compliant,
            "pct": round(div_compliant / len(div_emps) * 100, 1) if div_emps else 0,
        }

    # By yard (Casing only, for backward compatibility)
    by_yard = {}
    for y in YARD_ORDER:
        yard_emps = [e for e in employees if e["division"] == "Casing" and e["yard"] == y]
        yard_compliant = sum(1 for e in yard_emps if e["percent_complete"] >= 100)
        if yard_emps:
            by_yard[y] = {
                "total": len(yard_emps),
                "compliant": yard_compliant,
                "pct": round(yard_compliant / len(yard_emps) * 100, 1),
            }

    # By company -> location (for all companies)
    by_company_location = {}
    for e in employees:
        co = e["company"]
        loc = e["yard"] or "Unassigned"
        if co not in by_company_location:
            by_company_location[co] = {}
        if loc not in by_company_location[co]:
            by_company_location[co][loc] = {"total": 0, "compliant": 0}
        by_company_location[co][loc]["total"] += 1
        if e["percent_complete"] >= 100:
            by_company_location[co][loc]["compliant"] += 1
    # Compute pct
    for co in by_company_location:
        for loc in by_company_location[co]:
            d = by_company_location[co][loc]
            d["pct"] = round(d["compliant"] / d["total"] * 100, 1) if d["total"] > 0 else 0

    # By division -> location (BRHAS subdivisions)
    by_division_location = {}
    for e in employees:
        if e["company"] != "BRHAS" or not e["division"]:
            continue
        div = e["division"]
        loc = e["yard"] or "Unassigned"
        if div not in by_division_location:
            by_division_location[div] = {}
        if loc not in by_division_location[div]:
            by_division_location[div][loc] = {"total": 0, "compliant": 0}
        by_division_location[div][loc]["total"] += 1
        if e["percent_complete"] >= 100:
            by_division_location[div][loc]["compliant"] += 1
    for div in by_division_location:
        for loc in by_division_location[div]:
            d = by_division_location[div][loc]
            d["pct"] = round(d["compliant"] / d["total"] * 100, 1) if d["total"] > 0 else 0

    # By location (cross-company aggregate for "All Companies + specific yard")
    by_location = {}
    for e in employees:
        loc = e["yard"] or "Unassigned"
        if loc not in by_location:
            by_location[loc] = {"total": 0, "compliant": 0}
        by_location[loc]["total"] += 1
        if e["percent_complete"] >= 100:
            by_location[loc]["compliant"] += 1
    for loc in by_location:
        d = by_location[loc]
        d["pct"] = round(d["compliant"] / d["total"] * 100, 1) if d["total"] > 0 else 0

    # Non-compliant employees sorted by worst first (all companies)
    non_compliant = [e for e in employees if e["percent_complete"] < 100]
    non_compliant.sort(key=lambda x: (x["percent_complete"], -x["days_since_assignment"]))

    # Casing yard distribution
    casing_emps = [e for e in employees if e["division"] == "Casing"]
    yard_dist = Counter(e["yard"] for e in casing_emps)

    json_data = {
        "report_date": datetime.now(timezone.utc).strftime("%Y-%m-%d"),
        "generated_at": datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S"),
        "overall_pct": overall_pct,
        "total_employees": total_emp,
        "compliant_count": compliant,
        "overdue_count": overdue,
        "by_company": by_company,
        "by_division": by_division,
        "by_yard": by_yard,
        "by_company_location": by_company_location,
        "by_division_location": by_division_location,
        "by_location": by_location,
        "non_compliant_employees": non_compliant,
        "headcount_by_yard": dict(yard_dist),
    }

    os.makedirs("output", exist_ok=True)
    out = os.path.join("output", "training_compliance.json")
    with open(out, "w", encoding="utf-8") as f:
        json.dump(json_data, f, indent=2, default=str)

    print(f"\n  Overall compliance: {overall_pct}%")
    print(f"  Total employees: {total_emp}")
    print(f"  Compliant: {compliant}")
    print(f"  Overdue: {overdue}")
    print(f"  Non-compliant: {len(non_compliant)}")
    for co in COMPANY_ORDER:
        cd = by_company.get(co, {})
        print(f"    {co}: {cd.get('total', 0)} emps, {cd.get('pct', 0)}% compliant")
    print(f"  JSON written: {out}")
    print("  Done.")


if __name__ == "__main__":
    main()
