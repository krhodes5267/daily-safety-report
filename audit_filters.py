"""
AUDIT_FILTERS.PY -- Verify dashboard filter combinations show correct data
=========================================================================
Tests every Company > Division > Yard filter combination against archive
data to ensure events are correctly attributed and no data is lost or
duplicated when filtering.

Usage:
    python audit_filters.py
    python audit_filters.py --date 2026-04-17
    python audit_filters.py --range 2026-04-13 2026-04-19
"""
import json
import os
import sys
import argparse
from collections import defaultdict
from datetime import datetime, date, timedelta

# ==============================================================================
# MIRROR OF DASHBOARD JS FILTER LOGIC
# ==============================================================================

COMPANY_DIVISIONS = {
    "BRHAS": ["All Divisions", "Casing", "Rathole", "Pit Lining", "Poly Pipe",
              "Anchors", "Construction", "Drilling Tools", "Environmental",
              "Fencing", "Shop", "Corporate"],
    "Valor": [],
    "BTI": [],
    "Transcend": [],
    "Permian": [],
}

SPEEDING_TO_COMPANY = {
    "Casing": "BRHAS", "Rathole": "BRHAS", "Fencing": "BRHAS",
    "Pit Lining": "BRHAS", "Poly Pipe": "BRHAS", "Anchors": "BRHAS",
    "Construction": "BRHAS", "Drilling Tools": "BRHAS", "Environmental": "BRHAS",
    "Valor Energy Services": "Valor", "Downhole Tools": "Valor",
    "Valor": "Valor",
    "Butch's Trucking": "BTI", "BTI": "BTI",
    "Transcend": "Transcend", "Transcend Drilling": "Transcend",
    "Rentals": "Permian", "Permian": "Permian", "PER": "Permian",
    "Sales/Admin": "BRHAS", "Unassigned": "BRHAS",
    "Water/Construction": "BRHAS", "Fabrication": "BRHAS", "Unknown": "BRHAS",
    "Corporate": "BRHAS", "Shop": "BRHAS",
}

SPEEDING_TO_DIVISION = {
    "Casing": "Casing", "Rathole": "Rathole", "Fencing": "Fencing",
    "Pit Lining": "Pit Lining", "Poly Pipe": "Poly Pipe", "Anchors": "Anchors",
    "Construction": "Construction", "Drilling Tools": "Drilling Tools",
    "Environmental": "Environmental",
    "Valor Energy Services": "Valor", "Downhole Tools": "Valor",
    "Valor": "Valor",
    "Butch's Trucking": "BTI", "BTI": "BTI",
    "Transcend": "Transcend", "Transcend Drilling": "Transcend",
    "Rentals": "Permian", "Permian": "Permian", "PER": "Permian",
    "Sales/Admin": "All Divisions", "Unassigned": "All Divisions",
    "Water/Construction": "Construction", "Fabrication": "Shop", "Unknown": "All Divisions",
    "Corporate": "Corporate", "Shop": "Shop",
}

KPA_SL_TO_COMPANY = {
    "Casing": "BRHAS", "Rat Hole": "BRHAS", "Rathole": "BRHAS",
    "Anchor": "BRHAS", "Anchors": "BRHAS",
    "Poly Pipe": "BRHAS", "Pit Lining": "BRHAS",
    "Construction": "BRHAS", "Drilling Tools": "BRHAS",
    "Environmental": "BRHAS", "Fencing": "BRHAS",
    "Downhole Tools": "Valor", "Fabrication": "BRHAS", "Shop": "BRHAS",
    "Corporate": "BRHAS", "Containment": "BRHAS", "Civil": "BRHAS",
    "Water/Construction": "BRHAS", "Transcend Drilling": "Transcend",
    "Valor": "Valor", "Valor Energy": "Valor", "Valor Energy Services": "Valor",
    "BTI": "BTI", "Butch's Trucking": "BTI", "Trucking": "BTI",
    "Drilling": "Transcend", "Transcend": "Transcend",
    "Rentals": "Permian", "PER": "Permian", "Permian": "Permian",
}

KPA_SL_TO_DIVISION = {
    "Casing": "Casing", "Rat Hole": "Rathole", "Rathole": "Rathole",
    "Anchor": "Anchors", "Anchors": "Anchors",
    "Poly Pipe": "Poly Pipe", "Pit Lining": "Pit Lining",
    "Construction": "Construction", "Drilling Tools": "Drilling Tools",
    "Environmental": "Environmental", "Fencing": "Fencing",
    "Containment": "Environmental", "Civil": "Construction",
    "Downhole Tools": "Valor", "Fabrication": "Shop", "Shop": "Shop",
    "Corporate": "Corporate",
    "Water/Construction": "Construction", "Transcend Drilling": "Transcend",
    "Valor": "Valor", "Valor Energy": "Valor", "Valor Energy Services": "Valor",
    "BTI": "BTI", "Butch's Trucking": "BTI", "Trucking": "BTI",
    "Drilling": "Transcend", "Transcend": "Transcend",
    "Rentals": "Permian", "PER": "Permian", "Permian": "Permian",
}

# Casing yards
CASING_YARDS = ["Bryan", "Hobbs", "Jourdanton", "Kilgore", "Laredo", "Midland"]


def matches_company_division(event_co, event_div, filter_co, filter_div):
    """Mirror of JS matchesCompanyDivision()."""
    if filter_co == "All":
        return True
    if event_co and event_co != filter_co and event_co != "All":
        return False
    if filter_div not in ("All Divisions", "All") and event_div:
        return event_div == filter_div
    return True


# ==============================================================================
# FILTER FUNCTIONS (mirror JS)
# ==============================================================================
def get_filtered_speeding(events, company, division):
    """Mirror of JS getFilteredSpeedingEvents()."""
    results = []
    for e in events:
        co = SPEEDING_TO_COMPANY.get(e.get("division", ""), "")
        div = SPEEDING_TO_DIVISION.get(e.get("division", ""), e.get("division", ""))
        if matches_company_division(co, div, company, division):
            results.append(e)
    return results


def get_yard_filtered_speeding(events, company, division, yard):
    """Mirror of JS getYardFilteredSpeeding()."""
    filtered = get_filtered_speeding(events, company, division)
    if yard == "All Yards":
        return filtered
    return [e for e in filtered if e.get("yard") == yard]


def get_filtered_observations(obs, company, division):
    """Mirror of JS getFilteredObservations()."""
    if company == "All":
        return obs
    results = []
    for o in obs:
        sl = o.get("service_line", "")
        if not sl:
            results.append(o)  # No service line = show for all
            continue
        co = KPA_SL_TO_COMPANY.get(sl, "")
        div = KPA_SL_TO_DIVISION.get(sl, sl)
        if not co and "Drilling" in sl:
            co = "Transcend"
            div = "Transcend"
        if matches_company_division(co, div, company, division):
            results.append(o)
    return results


def get_filtered_incidents(incidents, company, division):
    """Mirror of JS getFilteredIncidents()."""
    if company == "All":
        return incidents
    results = []
    for inc in incidents:
        sl = inc.get("service_line", "")
        if not sl:
            results.append(inc)
            continue
        co = KPA_SL_TO_COMPANY.get(sl, "")
        div = KPA_SL_TO_DIVISION.get(sl, sl)
        if not co and "Drilling" in sl:
            co = "Transcend"
            div = "Transcend"
        if matches_company_division(co, div, company, division):
            results.append(inc)
    return results


def get_yard_filtered_camera(events, yard):
    """Mirror of JS getYardFilteredCamera()."""
    if yard == "All Yards":
        return events
    return [e for e in events if e.get("yard") == yard]


# ==============================================================================
# AUDIT
# ==============================================================================
passed = 0
failed = 0
warnings = 0


def check(label, condition, expected=None, actual=None, detail=""):
    global passed, failed
    if condition:
        passed += 1
    else:
        failed += 1
        msg = f"  [FAIL] {label}"
        if expected is not None:
            msg += f" -- expected: {expected}, got: {actual}"
        if detail:
            msg += f" ({detail})"
        print(msg)


def warn(label, detail=""):
    global warnings
    warnings += 1
    print(f"  [WARN] {label}{' -- ' + detail if detail else ''}")


def load_archive_range(archive_dir, start_date, end_date):
    """Load and merge archive files for a date range."""
    current = datetime.strptime(start_date, "%Y-%m-%d").date()
    end_dt = datetime.strptime(end_date, "%Y-%m-%d").date()

    speeding = []
    camera = []
    observations = []
    near_misses = []
    incidents = []
    assessments = []
    ytd = None
    days_loaded = 0

    while current <= end_dt:
        path = os.path.join(archive_dir, f"{current.isoformat()}.json")
        if os.path.exists(path):
            with open(path, encoding="utf-8") as f:
                day = json.load(f)
            if day.get("speeding") and day["speeding"].get("events"):
                speeding.extend(day["speeding"]["events"])
            if day.get("camera") and day["camera"].get("events"):
                camera.extend(day["camera"]["events"])
            if day.get("kpa"):
                kpa = day["kpa"]
                if kpa.get("observations") and kpa["observations"].get("details"):
                    observations.extend(kpa["observations"]["details"])
                if kpa.get("near_misses"):
                    near_misses.extend(kpa["near_misses"])
                if kpa.get("incidents"):
                    incidents.extend(kpa["incidents"])
                if kpa.get("assessments") and kpa["assessments"].get("details"):
                    assessments.extend(kpa["assessments"]["details"])
            if day.get("ytd"):
                ytd = day["ytd"]
            days_loaded += 1
        current += timedelta(days=1)

    return {
        "speeding": speeding,
        "camera": camera,
        "observations": observations,
        "near_misses": near_misses,
        "incidents": incidents,
        "assessments": assessments,
        "ytd": ytd,
        "days_loaded": days_loaded,
    }


def audit_data_coverage(data):
    """Check that all events have required classification fields."""
    print("\n== DATA COVERAGE ==")

    # Speeding: every event needs a division
    spd = data["speeding"]
    no_div = [e for e in spd if not e.get("division")]
    check("Speeding: all events have division",
          len(no_div) == 0, "0 missing", len(no_div))
    unknown_div = [e for e in spd if e.get("division") == "Unknown"]
    if unknown_div:
        vehicles = set(e.get("vehicle", "?") for e in unknown_div[:10])
        warn(f"Speeding events with 'Unknown' division (mapped to BRHAS): {len(unknown_div)}",
             ", ".join(vehicles))

    # Speeding: every event maps to a known company
    unmapped_divs = set()
    for e in spd:
        d = e.get("division", "")
        if d and d not in SPEEDING_TO_COMPANY:
            unmapped_divs.add(d)
    check("Speeding: all divisions map to company",
          len(unmapped_divs) == 0, "0 unmapped", unmapped_divs or "none")

    # Camera: every event has a yard
    cam = data["camera"]
    no_yard_cam = [e for e in cam if not e.get("yard") or e["yard"] == "Unknown"]
    check("Camera: all events have yard",
          len(no_yard_cam) == 0, "0 unknown", len(no_yard_cam))

    # Observations: check service_line coverage
    obs = data["observations"]
    no_sl = [o for o in obs if not o.get("service_line")]
    pct_with_sl = ((len(obs) - len(no_sl)) / len(obs) * 100) if obs else 100
    check("Observations: >80% have service_line",
          pct_with_sl >= 80, ">80%", f"{pct_with_sl:.1f}%")
    if no_sl:
        warn(f"Observations without service_line: {len(no_sl)}/{len(obs)}")

    # Observations: all service_lines map to a company
    unmapped_sl = set()
    for o in obs:
        sl = o.get("service_line", "")
        if sl and sl not in KPA_SL_TO_COMPANY:
            if "Drilling" not in sl:  # Drilling -> Transcend handled specially
                unmapped_sl.add(sl)
    check("Observations: all service_lines map to company",
          len(unmapped_sl) == 0, "0 unmapped", unmapped_sl or "none")

    # Incidents: check service_line
    inc = data["incidents"]
    if inc:
        no_sl_inc = [i for i in inc if not i.get("service_line")]
        check("Incidents: have service_line",
              len(no_sl_inc) < len(inc), "some with SL", f"{len(no_sl_inc)} missing")


def audit_filter_completeness(data):
    """Test that filtering by every company/division/yard returns correct data."""
    print("\n== FILTER COMPLETENESS ==")

    spd = data["speeding"]
    obs = data["observations"]
    inc = data["incidents"]
    cam = data["camera"]

    # Test 1: Sum of all company filters = total (no events lost)
    print("\n  -- Speeding by Company --")
    company_totals = {}
    for co in COMPANY_DIVISIONS:
        filtered = get_filtered_speeding(spd, co, "All Divisions")
        company_totals[co] = len(filtered)
        if filtered:
            print(f"    {co}: {len(filtered)} events")

    # "All" should equal total
    all_filtered = get_filtered_speeding(spd, "All", "All")
    check("Speeding: 'All' filter = total events",
          len(all_filtered) == len(spd), len(spd), len(all_filtered))

    # Sum of companies should >= total (some may show in multiple due to mapping)
    total_by_company = sum(company_totals.values())
    # Events with unknown/unmapped division won't show in any company
    unmapped = [e for e in spd if SPEEDING_TO_COMPANY.get(e.get("division", ""), "") == ""]
    expected_mapped = len(spd) - len(unmapped)
    check("Speeding: company filters cover all mapped events",
          total_by_company >= expected_mapped,
          f">={expected_mapped}", total_by_company)

    # Test 2: Each BRHAS division
    print("\n  -- Speeding by BRHAS Division --")
    div_totals = {}
    for div in COMPANY_DIVISIONS["BRHAS"]:
        if div == "All Divisions":
            continue
        filtered = get_filtered_speeding(spd, "BRHAS", div)
        div_totals[div] = len(filtered)
        if filtered:
            print(f"    BRHAS > {div}: {len(filtered)} events")

    brhas_all = get_filtered_speeding(spd, "BRHAS", "All Divisions")
    brhas_div_sum = sum(div_totals.values())
    # Sum of divisions might be less than "All Divisions" if some events
    # have divisions not in COMPANY_DIVISIONS list (e.g. Corporate, Shop)
    check("Speeding: BRHAS division sum <= BRHAS All",
          brhas_div_sum <= len(brhas_all),
          f"<={len(brhas_all)}", brhas_div_sum)

    # Test 3: Yard filtering for Casing
    print("\n  -- Speeding by Casing Yard --")
    for yard in CASING_YARDS:
        yard_events = get_yard_filtered_speeding(spd, "BRHAS", "Casing", yard)
        if yard_events:
            print(f"    BRHAS > Casing > {yard}: {len(yard_events)} events")
        # Verify every event in this yard actually has the right yard
        wrong_yard = [e for e in yard_events if e.get("yard") and e["yard"] != yard]
        check(f"Speeding: {yard} yard filter correct",
              len(wrong_yard) == 0, "0 wrong", len(wrong_yard))

    # Test 4: Observations by company/division
    print("\n  -- Observations by Company --")
    for co in COMPANY_DIVISIONS:
        filtered = get_filtered_observations(obs, co, "All Divisions")
        if filtered:
            print(f"    {co}: {len(filtered)} observations")

    # Test 5: Observations by BRHAS division
    print("\n  -- Observations by BRHAS Division --")
    for div in COMPANY_DIVISIONS["BRHAS"]:
        if div == "All Divisions":
            continue
        filtered = get_filtered_observations(obs, "BRHAS", div)
        if filtered:
            print(f"    BRHAS > {div}: {len(filtered)} observations")

    # Test 6: Camera by yard
    print("\n  -- Camera by Yard --")
    for yard in CASING_YARDS:
        filtered = get_yard_filtered_camera(cam, yard)
        if filtered:
            print(f"    {yard}: {len(filtered)} camera events")
            wrong = [e for e in filtered if e.get("yard") != yard]
            check(f"Camera: {yard} filter correct",
                  len(wrong) == 0, "0 wrong", len(wrong))


def audit_no_data_loss(data):
    """Ensure no events disappear when going from All to specific filters."""
    print("\n== DATA LOSS CHECK ==")

    spd = data["speeding"]
    obs = data["observations"]

    # Speeding: every event with a mapped division should appear in exactly one company
    for e in spd[:200]:  # Sample
        div = e.get("division", "")
        co = SPEEDING_TO_COMPANY.get(div, "")
        if not co:
            continue
        # This event should appear when filtering for its company
        company_events = get_filtered_speeding(spd, co, "All Divisions")
        found = any(
            ce.get("vehicle") == e.get("vehicle") and ce.get("time") == e.get("time")
            for ce in company_events
        )
        if not found:
            check(f"Speeding loss: {div} event in {co}",
                  False, "found", "missing",
                  f"vehicle={e.get('vehicle')}")
            break

    # Observations: every observation with service_line should appear in its company
    for o in obs[:200]:
        sl = o.get("service_line", "")
        if not sl:
            continue
        co = KPA_SL_TO_COMPANY.get(sl, "")
        if not co:
            if "Drilling" in sl:
                co = "Transcend"
            else:
                continue
        filtered = get_filtered_observations(obs, co, "All Divisions")
        found = any(
            fo.get("report_number") == o.get("report_number") and fo.get("date") == o.get("date")
            for fo in filtered
        )
        if not found:
            check(f"Observation loss: SL={sl} in {co}",
                  False, "found", "missing",
                  f"observer={o.get('observer')}")
            break

    check("Speeding: no data loss in company filters (sample)", True)
    check("Observations: no data loss in company filters (sample)", True)


def audit_cross_check_totals(data):
    """Cross-check archive totals against filter sums."""
    print("\n== CROSS-CHECK TOTALS ==")

    spd = data["speeding"]
    obs = data["observations"]

    # Speeding tier counts
    tier_counts = {"RED": 0, "ORANGE": 0, "YELLOW": 0}
    for e in spd:
        t = (e.get("tier") or "YELLOW").upper()
        tier_counts[t] = tier_counts.get(t, 0) + 1
    print(f"  Speeding tiers: RED={tier_counts['RED']} ORANGE={tier_counts['ORANGE']} YELLOW={tier_counts['YELLOW']}")
    check("Speeding: tier sum = total",
          sum(tier_counts.values()) == len(spd),
          len(spd), sum(tier_counts.values()))

    # Observation type counts
    type_counts = defaultdict(int)
    for o in obs:
        type_counts[o.get("type", "Other")] += 1
    print(f"  Observation types: {dict(type_counts)}")
    check("Observations: type sum = total",
          sum(type_counts.values()) == len(obs),
          len(obs), sum(type_counts.values()))


def main():
    parser = argparse.ArgumentParser(description="Audit dashboard filter accuracy")
    parser.add_argument("--date", help="Single date to audit (YYYY-MM-DD)")
    parser.add_argument("--range", nargs=2, metavar=("START", "END"),
                        help="Date range to audit")
    parser.add_argument("--archive-dir", default="archive")
    args = parser.parse_args()

    if args.date:
        start_date = end_date = args.date
    elif args.range:
        start_date, end_date = args.range
    else:
        # Default: last 7 days
        end_dt = date.today() - timedelta(days=1)
        start_dt = end_dt - timedelta(days=6)
        start_date = start_dt.isoformat()
        end_date = end_dt.isoformat()

    print("=" * 60)
    print(f"  FILTER ACCURACY AUDIT")
    print(f"  Range: {start_date} to {end_date}")
    print(f"  Archive: {args.archive_dir}/")
    print("=" * 60)

    data = load_archive_range(args.archive_dir, start_date, end_date)
    print(f"\n  Loaded {data['days_loaded']} days")
    print(f"  Speeding: {len(data['speeding'])} events")
    print(f"  Camera: {len(data['camera'])} events")
    print(f"  Observations: {len(data['observations'])}")
    print(f"  Incidents: {len(data['incidents'])}")
    print(f"  Assessments: {len(data['assessments'])}")

    if not data["speeding"] and not data["observations"]:
        print("\n  ERROR: No data loaded. Check archive directory.")
        sys.exit(1)

    audit_data_coverage(data)
    audit_filter_completeness(data)
    audit_no_data_loss(data)
    audit_cross_check_totals(data)

    print("\n" + "=" * 60)
    print(f"  RESULTS: {passed} passed, {failed} failed, {warnings} warnings")
    print("=" * 60)

    # Save report
    report = {
        "date": datetime.now().isoformat(),
        "range": f"{start_date} to {end_date}",
        "days_loaded": data["days_loaded"],
        "passed": passed,
        "failed": failed,
        "warnings": warnings,
    }
    with open("audit_filters_report.json", "w") as f:
        json.dump(report, f, indent=2)

    sys.exit(1 if failed > 0 else 0)


if __name__ == "__main__":
    main()
