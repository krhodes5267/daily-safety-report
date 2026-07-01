"""
Universal Safety Recap -- Text Data Exporter.

Formats collected data into a plain-text summary for Claude Web HTML generation.
Mirrors the print_data_summary() pattern from standalone recap scripts.
ASCII-only output (Windows cp1252 safe).
"""

import os
from calendar import month_name

from .config import (
    DIVISIONS,
    FORM_FIELD_MAP,
    NAICS_DART_BENCHMARK,
    NAICS_TRIR_BENCHMARK,
)
from .data_fetcher import KMH_TO_MPH

# ---------------------------------------------------------------------------
# KPA incident form field hash -> human-readable label
# ---------------------------------------------------------------------------
INC_FIELD_LABELS = {
    "lsx3msa0w9n9edb4": "Company",
    "sha7vur5q2l6d6gq": "Service Line",
    "pk6qj0kiu9vek20v": "District",
    "9ohdd2lwvl7p0oc6": "Location",
    "g5w4b0uh15wxqykt": "Employee Involved",
    "55gg4nkoemnnfo2a": "Employee Name",
    "nojcquy0tfl9hqih": "Incident Classification",
    "ke38yo6wx6s2es2h": "Job Title",
    "eo0bl6w5ouq06s86": "Date/Time of Incident",
    "cngln49ixwebmnkx": "Time Reported",
    "w997tirq97oenvuz": "Supervisor",
    "0uywiyo2waa2grw1": "Business Unit",
    "313e9txgrof0uute": "Description",
    "r6s5m9o6u002isqj": "Drug/Alcohol Test Required",
    "opjml94mahirshxx": "Drug/Alcohol Test Performed",
    "e6l0cwglzkh745sb": "Days Away From Work",
    "3z9b59fvhkikd9f8": "Injury Type",
    "qymn43fkdqj1hf2j": "Body Part",
    "x93shv5j2r4w96r6": "Treatment Type",
}

# Observation form field hash -> human-readable label
OBS_FIELD_LABELS = {
    "t5187momol3em85v": "Company",
    "64c7upqkyt79zhh1": "Service Line",
    "7vj2l992y7fwqhwz": "District",
    "lg5pnj4chjadnv46": "Location",
    "bff8m4x6xbc033kg": "Type",
    "0kc57oj2zkse21o3": "Observer Name",
    "uncbcge9x8vow9pn": "Description",
    "dpy2klalngsr7ek9": "Corrective Action",
    "vxew6ukynemxwvjr": "Customer",
}

# Fields to skip in detail output
SKIP_FIELDS = {
    "link", "kpa_link", "id", "response_id", "form_id",
    "select-surface", "select-intermediate", "select-long",
    "select-other", "select-yes", "select-no", "select-n/a",
    "updated_time", "version", "duration", "latitude", "longitude",
    "surrogate", "score-percent", "score-possible", "score-earned",
}

# Non-recordable incident classification keywords (substring match)
NON_RECORDABLE_KEYWORDS = [
    "near miss", "near-miss",
    "property damage",
    "first aid",
    "environmental",
    "equipment damage",
    "report only",
    "vehicle accident", "at-fault vehicle",
    "not-at-fault vehicle", "non-injury vehicle",
]

# Camera coaching statuses that count as dismissed
CAMERA_DISMISSED_STATUSES = {
    "uncoachable",
}

# Minimum severity level for camera events (high or above)
CAMERA_MIN_SEVERITY = {"high", "critical"}


def export_text_data(division_key, data, output_dir=None):
    """Export collected data as a plain-text summary file.

    Args:
        division_key: str (e.g., "per", "casing")
        data: dict returned by DataCollector.collect_monthly()
        output_dir: directory for output file (default: ~/Downloads/)

    Returns:
        str: path to generated .txt file
    """
    cfg = data.get("division", DIVISIONS.get(division_key, {}))
    year = data.get("year", 0)
    month = data.get("month", 0)
    period_str = f"{month_name[month]} {year}" if month else "Unknown"

    lines = []
    w = lines.append  # shorthand

    w("=" * 80)
    w(f"{cfg.get('display_name', division_key).upper()} MONTHLY HSE DATA SUMMARY")
    w("=" * 80)
    w(f"Period: {period_str}")
    w(f"Month: {year}-{month:02d}")
    w("")

    # -- Division Info --
    w("DIVISION INFO")
    w("-" * 80)
    w(f"Name: {cfg.get('display_name', division_key)}")
    w(f"Company: {cfg.get('company', 'BRHAS')}")
    w(f"Manager: {cfg.get('manager', 'TBD')}")
    w(f"Safety Rep: {cfg.get('safety_rep', 'TBD')}")

    man_hours = data.get("man_hours", {})
    mh_headcount = man_hours.get("headcount", len(man_hours.get("by_employee", [])))
    total_hours = man_hours.get("total_hours", 0)

    # Training roster (KPA)
    _, training_rows = data.get("training_status", ([], []))
    kpa_headcount = len(training_rows)

    # Use man-hours headcount as primary (active employees with timecards);
    # flag discrepancy if KPA roster differs significantly
    headcount = mh_headcount if mh_headcount > 0 else kpa_headcount
    w(f"Headcount: {headcount}")
    if mh_headcount > 0 and kpa_headcount > 0 and abs(mh_headcount - kpa_headcount) > 5:
        w(f"  NOTE: Man-hours roster = {mh_headcount}, KPA training roster = {kpa_headcount}")
    w(f"Monthly Man Hours: {total_hours:,.1f}")
    w("")

    # -- Observations --
    _, obs_rows = data.get("observations_current", ([], []))
    _, hse_obs_rows = data.get("hse_observations_current", ([], []))
    all_obs = obs_rows + hse_obs_rows

    # -- Incidents --
    _, inc_rows = data.get("incidents_current", ([], []))

    # -- Previous month data --
    _, prev_obs = data.get("observations_prev", ([], []))
    _, prev_hse_obs = data.get("hse_observations_prev", ([], []))
    prev_obs_total = len(prev_obs) + len(prev_hse_obs)
    _, prev_inc = data.get("incidents_prev", ([], []))

    # -- Lagging Indicators --
    w("LAGGING INDICATORS")
    w("-" * 80)

    # Determine recordable incidents
    recordable = 0
    dart_cases = 0
    for inc in inc_rows:
        inc_type = ""
        for key in ("nojcquy0tfl9hqih", "type", "Incident Classification"):
            val = inc.get(key, "")
            if val:
                inc_type = str(val).strip().lower()
                break
        is_non_recordable = any(kw in inc_type for kw in NON_RECORDABLE_KEYWORDS)
        if inc_type and not is_non_recordable:
            recordable += 1
            # Check for days away / restricted / transferred for DART
            days_away = 0
            for dk in ("e6l0cwglzkh745sb", "Days Away From Work"):
                dv = inc.get(dk, "")
                if dv:
                    try:
                        days_away = int(float(str(dv).strip()))
                    except (ValueError, TypeError):
                        pass
                    break
            if days_away > 0:
                dart_cases += 1

    # TRIR/DART use YTD man-hours (monthly hours * months elapsed in year)
    ytd_hours = total_hours * month if total_hours > 0 else 0
    trir = round(recordable * 200000 / ytd_hours, 2) if ytd_hours > 0 else 0
    dart = round(dart_cases * 200000 / ytd_hours, 2) if ytd_hours > 0 else 0

    w(f"Total Incidents Reported: {len(inc_rows)} (Previous Month: {len(prev_inc)})")
    w(f"OSHA Recordable Incidents: {recordable}")
    w(f"YTD Man Hours: {ytd_hours:,.0f} (Monthly: {total_hours:,.0f} x {month} months)")
    w(f"TRIR (per 200k hrs): {trir:.2f} (Industry Benchmark: {NAICS_TRIR_BENCHMARK})")
    w(f"DART Rate (per 200k hrs): {dart:.2f} (Industry Benchmark: {NAICS_DART_BENCHMARK})")
    w("")

    # -- Leading Indicators --
    w("LEADING INDICATORS")
    w("-" * 80)
    obs_rate = round(len(all_obs) / headcount, 2) if headcount > 0 else 0

    # Training compliance -- use headcount as denominator for consistency
    compliant = sum(1 for r in training_rows if r.get("percent_complete", 0) >= 100)
    overdue = sum(1 for r in training_rows if r.get("status", "").lower() == "overdue")
    train_pct = round(compliant / kpa_headcount * 100, 1) if kpa_headcount > 0 else 0

    w(f"Observations: {len(all_obs)} ({obs_rate} per employee) (Previous Month: {prev_obs_total})")
    w(f"Training Compliance: {train_pct:.1f}%")
    w(f"  Total Employees (KPA): {kpa_headcount}")
    w(f"  Compliant: {compliant}")
    w(f"  Non-Compliant: {kpa_headcount - compliant}")
    w(f"  Overdue: {overdue}")
    w("")

    # -- Fleet Metrics (if applicable) --
    has_fleet = bool(cfg.get("motive_group_ids"))
    if has_fleet:
        _render_fleet_text(w, data, cfg)
    else:
        w("NOTE: No Motive fleet data available for this division.")
        w("")

    # -- Vehicle Inspections (Pre/Post Trip) --
    _, insp_rows = data.get("vehicle_inspections_current", ([], []))
    _, prev_insp_rows = data.get("vehicle_inspections_prev", ([], []))
    if insp_rows or prev_insp_rows:
        w("VEHICLE INSPECTIONS (PRE/POST TRIP)")
        w("-" * 80)
        w(f"Total Inspections: {len(insp_rows)} (Previous Month: {len(prev_insp_rows)})")
        # Count by status if available
        status_counts = {}
        for insp in insp_rows:
            status = str(insp.get("status", insp.get("result", "Completed"))).strip()
            if not status:
                status = "Completed"
            status_counts[status] = status_counts.get(status, 0) + 1
        if status_counts:
            for s, c in sorted(status_counts.items(), key=lambda x: x[1], reverse=True):
                w(f"  {s}: {c}")
        w("")

    # -- Observation Summary (by type and by rig/location) --
    if all_obs:
        w("OBSERVATION SUMMARY")
        w("-" * 80)
        w(f"Total: {len(all_obs)} (Previous Month: {prev_obs_total})")
        obs_rate_str = f"{obs_rate:.2f}" if isinstance(obs_rate, float) else str(obs_rate)
        w(f"Rate: {obs_rate_str} per employee")
        w("")

        # By type
        by_type = {}
        for obs in all_obs:
            obs_type = _get_field(obs, "bff8m4x6xbc033kg", "type", "Type")
            if obs_type:
                # Observations can have multiple types comma-separated
                for t in obs_type.split(","):
                    t = t.strip()
                    if t:
                        by_type[t] = by_type.get(t, 0) + 1

        if by_type:
            w("By Type:")
            for t, c in sorted(by_type.items(), key=lambda x: x[1], reverse=True):
                pct = f"{(c / len(all_obs) * 100):.0f}%"
                w(f"  {t}: {c} ({pct})")
            w("")

        # By rig/location (useful for rig-based divisions like Transcend)
        rig_field = cfg.get("obs_fields", {}).get("rig")
        if rig_field:
            by_rig = {}
            for obs in all_obs:
                rig = str(obs.get(rig_field, "")).strip()
                if not rig:
                    rig = "Unknown"
                by_rig[rig] = by_rig.get(rig, 0) + 1
            if by_rig:
                w("By Rig:")
                for r, c in sorted(by_rig.items(), key=lambda x: x[1], reverse=True):
                    w(f"  {r}: {c}")
                w("")

        # Near Misses and Good Catches callout
        near_misses = []
        good_catches = []
        for obs in all_obs:
            obs_type = _get_field(obs, "bff8m4x6xbc033kg", "type", "Type").lower()
            if "near miss" in obs_type:
                near_misses.append(obs)
            if "good catch" in obs_type:
                good_catches.append(obs)

        if near_misses or good_catches:
            w(f"Near Misses: {len(near_misses)}")
            w(f"Good Catches: {len(good_catches)}")
            w("")

    # -- Observation Details --
    w("OBSERVATION DETAILS")
    w("-" * 80)
    if all_obs:
        for i, obs in enumerate(all_obs, 1):
            w(f"Observation #{i}:")
            date = obs.get("date", "N/A")
            if date and len(date) > 10:
                date = date[:10]
            observer = obs.get("observer", obs.get("_observer", "Unknown"))
            w(f"  Date: {date}")
            w(f"  Observer: {observer}")

            # Use field map for readable output
            obs_type = _get_field(obs, "bff8m4x6xbc033kg", "type", "Type")
            obs_loc = _get_field(obs, "lg5pnj4chjadnv46", "location", "Location")
            obs_desc = _get_field(obs, "uncbcge9x8vow9pn", "description", "Description")
            obs_action = _get_field(obs, "dpy2klalngsr7ek9", "action", "Corrective Action")

            w(f"  Type: {obs_type or 'N/A'}")
            w(f"  Location: {obs_loc or 'N/A'}")
            w(f"  Description: {obs_desc or 'N/A'}")
            if obs_action:
                w(f"  Corrective Action: {obs_action}")
            w("")
    else:
        w("  No observations recorded")
        w("")

    # -- Incident Details --
    w("INCIDENT DETAILS")
    w("-" * 80)
    if inc_rows:
        for i, inc in enumerate(inc_rows, 1):
            w(f"Incident #{i}:")
            for key, val in inc.items():
                if not val:
                    continue
                val_str = str(val).strip()
                if val_str.startswith("http://") or val_str.startswith("https://"):
                    continue
                if key.lower() in SKIP_FIELDS:
                    continue
                label = INC_FIELD_LABELS.get(key, OBS_FIELD_LABELS.get(key, key))
                w(f"  {label}: {val_str}")
            w("")
    else:
        w("  No incidents recorded")
        w("")

    # -- Assessment Details (deduplicated by response_id) --
    _, assess_rows = data.get("assessments_current", ([], []))
    if assess_rows:
        _render_assessments_text(w, assess_rows)

    # -- JSA Summary --
    _, jsa_rows = data.get("jsas_current", ([], []))
    _, prev_jsa_rows = data.get("jsas_prev", ([], []))
    _, jsa_review_rows = data.get("jsa_reviews_current", ([], []))
    if jsa_rows or prev_jsa_rows:
        _render_jsa_summary(w, jsa_rows, prev_jsa_rows, jsa_review_rows, cfg)

    # -- Rig Inspections --
    _, rig_insp_rows = data.get("rig_inspections_current", ([], []))
    _, prev_rig_insp_rows = data.get("rig_inspections_prev", ([], []))
    if rig_insp_rows or prev_rig_insp_rows:
        _render_rig_inspections_text(w, rig_insp_rows, prev_rig_insp_rows, cfg)

    # -- Vehicle Inspection Details (enhanced) --
    if insp_rows:
        _render_vehicle_inspection_details(w, insp_rows, cfg)

    # -- Training Compliance Details --
    w("TRAINING COMPLIANCE")
    w("-" * 80)
    if kpa_headcount > 0:
        w(f"Overall Compliance: {train_pct:.1f}%")
        w(f"Total Employees: {kpa_headcount}")
        w(f"Compliant: {compliant}")
        w(f"Non-Compliant: {kpa_headcount - compliant}")
        w(f"Overdue: {overdue}")
        w("")

        non_compliant = [
            r for r in training_rows
            if r.get("percent_complete", 0) < 100
        ]
        non_compliant.sort(
            key=lambda x: len(x.get("incomplete_training_names", [])),
            reverse=True,
        )

        if non_compliant:
            w(f"NON-COMPLIANT EMPLOYEES ({len(non_compliant)}):")
            w("-" * 80)
            for emp in non_compliant:
                name = emp.get("employee_name", "Unknown")
                pct = emp.get("percent_complete", 0)
                status = emp.get("status", "Unknown")
                incomplete = emp.get("incomplete_training_names", [])
                w(f"\nEmployee: {name}")
                w(f"Completion: {pct}%")
                w(f"Status: {status}")
                w(f"Incomplete Programs ({len(incomplete)}):")
                for prog in incomplete:
                    w(f"  - {prog}")
        else:
            w("All employees are compliant")
    else:
        w("  Training compliance data unavailable")

    w("")
    w("=" * 80)
    w("END OF DATA SUMMARY")
    w("=" * 80)

    # Write to file
    # Default: Desktop/Monthly HSE Recaps/{Month Year}/
    if not output_dir:
        month_folder = f"{month_name[month]} {year}"
        output_dir = os.path.join(
            os.path.expanduser("~"), "Desktop", "Monthly HSE Recaps", month_folder
        )
    os.makedirs(output_dir, exist_ok=True)

    display = cfg.get("display_name", division_key)
    safe_name = display.replace(" ", "_").replace("/", "-").replace("(", "").replace(")", "").replace("'", "")
    filename = f"{safe_name}_{month_name[month]}_{year}_HSE_Data.txt"
    filepath = os.path.join(output_dir, filename)

    text = "\n".join(lines)
    with open(filepath, "w", encoding="utf-8") as f:
        f.write(text)

    print(f"\n  Data summary saved: {filepath}")
    return filepath


# ===========================================================================
# Fleet Metrics -- Full Breakdown
# ===========================================================================

def _render_fleet_text(w, data, cfg):
    """Render complete fleet metrics section with severity, by-yard, camera."""
    w("FLEET METRICS")
    w("-" * 80)
    trips = data.get("ifta_trips", [])
    speeding = data.get("speeding_events", [])
    camera = data.get("camera_events", [])
    vehicles = data.get("vehicles", {})

    # -- Mileage by vehicle --
    by_vehicle = {}
    total_miles = 0.0
    for trip in trips:
        t = trip.get("ifta_trip_report", trip)
        vid = str(t.get("vehicle", {}).get("number", t.get("vehicle_number", "")))
        miles = _safe_float(t.get("total_miles", t.get("distance", 0)))
        total_miles += miles
        if vid:
            by_vehicle[vid] = by_vehicle.get(vid, 0) + miles

    active_vehicles = len(by_vehicle)  # vehicles with actual mileage
    total_in_group = len(vehicles)

    w(f"Active Vehicles (with mileage): {active_vehicles}")
    w(f"Total Vehicles in Motive Group: {total_in_group}")
    w(f"Fleet Utilization: {(active_vehicles / total_in_group * 100):.0f}%" if total_in_group > 0 else "Fleet Utilization: N/A")
    w(f"Total Mileage: {total_miles:,.1f} miles")
    avg_per = total_miles / active_vehicles if active_vehicles > 0 else 0
    w(f"Avg Miles per Active Vehicle: {avg_per:,.0f}")
    w("")

    # -- Mileage by yard (multi-yard divisions) --
    yards_cfg = cfg.get("yards", {})
    yard_order = yards_cfg.get("order", [])
    if yard_order:
        w("MILEAGE BY YARD")
        w("-" * 80)
        by_yard_miles = {y: {"miles": 0.0, "vehicles": 0} for y in yard_order}
        by_yard_miles["Other"] = {"miles": 0.0, "vehicles": 0}

        for vid, miles in by_vehicle.items():
            v = vehicles.get(vid, {})
            yard = v.get("yard", "")
            bucket = yard if yard in by_yard_miles else "Other"
            by_yard_miles[bucket]["miles"] += miles
            by_yard_miles[bucket]["vehicles"] += 1

        w(f"{'Yard':<20} {'Miles':>12} {'Vehicles':>10} {'Avg/Vehicle':>12} {'% of Total':>10}")
        w("-" * 66)
        for yard in yard_order:
            d = by_yard_miles[yard]
            avg = f"{d['miles']/d['vehicles']:,.0f}" if d["vehicles"] > 0 else "0"
            pct = f"{(d['miles']/total_miles*100):.1f}%" if total_miles > 0 else "0%"
            w(f"{yard:<20} {d['miles']:>12,.0f} {d['vehicles']:>10} {avg:>12} {pct:>10}")
        if by_yard_miles["Other"]["miles"] > 0:
            d = by_yard_miles["Other"]
            avg = f"{d['miles']/d['vehicles']:,.0f}" if d["vehicles"] > 0 else "0"
            pct = f"{(d['miles']/total_miles*100):.1f}%"
            w(f"{'Other':<20} {d['miles']:>12,.0f} {d['vehicles']:>10} {avg:>12} {pct:>10}")
        w("")

    # -- Speeding Analysis --
    w("SPEEDING ANALYSIS")
    w("-" * 80)

    # Separate valid vs dismissed speeding events
    valid_speeding = []
    dismissed_speeding = []
    for evt in speeding:
        e = evt.get("speeding_event", evt) if isinstance(evt, dict) else evt
        status = str(e.get("status", "reported")).strip().lower()
        if status == "invalid":
            dismissed_speeding.append(evt)
        else:
            valid_speeding.append(evt)

    w(f"Total Speeding Events: {len(valid_speeding)} valid ({len(dismissed_speeding)} dismissed)")

    if valid_speeding:
        # Severity tiers
        tiers = {"Critical (20+ mph over)": 0, "High (15-19 mph over)": 0, "Medium (<15 mph over)": 0}
        by_driver = {}
        by_speed_vehicle = {}
        max_over_by_driver = {}

        for evt in valid_speeding:
            e = evt.get("speeding_event", evt) if isinstance(evt, dict) else evt

            over_kph = _safe_float(e.get("avg_over_speed_in_kph", 0))
            if over_kph > 0:
                over = over_kph * KMH_TO_MPH
            else:
                speed_kph = _safe_float(e.get("max_vehicle_speed", 0))
                limit_kph = _safe_float(e.get("max_posted_speed_limit_in_kph", 0))
                if speed_kph == 0 and limit_kph == 0:
                    speed_kph = _safe_float(e.get("speed", 0))
                    limit_kph = _safe_float(e.get("posted_speed_limit", 0))
                over = (speed_kph - limit_kph) * KMH_TO_MPH

            if over <= 0:
                continue

            if over >= 20:
                tiers["Critical (20+ mph over)"] += 1
            elif over >= 15:
                tiers["High (15-19 mph over)"] += 1
            else:
                tiers["Medium (<15 mph over)"] += 1

            driver = _extract_driver(e)
            if driver:
                by_driver[driver] = by_driver.get(driver, 0) + 1
                if over > max_over_by_driver.get(driver, 0):
                    max_over_by_driver[driver] = over

            vid = _extract_vehicle_num(e)
            if vid:
                by_speed_vehicle[vid] = by_speed_vehicle.get(vid, 0) + 1

        w("")
        w("Severity Breakdown:")
        for tier, count in tiers.items():
            pct = f"{(count/len(valid_speeding)*100):.0f}%" if len(valid_speeding) > 0 else "0%"
            w(f"  {tier}: {count} ({pct})")

        # Rate per 10k miles
        if total_miles > 0:
            rate = len(valid_speeding) / (total_miles / 10000)
            w(f"\nSpeeding Rate: {rate:.1f} events per 10,000 miles")

        # By yard (multi-yard)
        if yard_order:
            w("")
            w("Speeding by Yard:")
            by_yard_speed = {y: {"total": 0, "critical": 0, "high": 0, "medium": 0} for y in yard_order}
            by_yard_speed["Other"] = {"total": 0, "critical": 0, "high": 0, "medium": 0}

            for evt in valid_speeding:
                e = evt.get("speeding_event", evt) if isinstance(evt, dict) else evt
                vid = _extract_vehicle_num(e)
                yard = vehicles.get(vid, {}).get("yard", "")
                bucket = yard if yard in by_yard_speed else "Other"

                over_kph = _safe_float(e.get("avg_over_speed_in_kph", 0))
                if over_kph > 0:
                    over = over_kph * KMH_TO_MPH
                else:
                    speed_kph = _safe_float(e.get("max_vehicle_speed", 0))
                    limit_kph = _safe_float(e.get("max_posted_speed_limit_in_kph", 0))
                    if speed_kph == 0 and limit_kph == 0:
                        speed_kph = _safe_float(e.get("speed", 0))
                        limit_kph = _safe_float(e.get("posted_speed_limit", 0))
                    over = (speed_kph - limit_kph) * KMH_TO_MPH

                by_yard_speed[bucket]["total"] += 1
                if over >= 20:
                    by_yard_speed[bucket]["critical"] += 1
                elif over >= 15:
                    by_yard_speed[bucket]["high"] += 1
                else:
                    by_yard_speed[bucket]["medium"] += 1

            w(f"  {'Yard':<20} {'Total':>8} {'Critical':>10} {'High':>8} {'Medium':>8}")
            w(f"  {'-'*56}")
            for yard in yard_order:
                d = by_yard_speed[yard]
                if d["total"] > 0:
                    w(f"  {yard:<20} {d['total']:>8} {d['critical']:>10} {d['high']:>8} {d['medium']:>8}")
            if by_yard_speed["Other"]["total"] > 0:
                d = by_yard_speed["Other"]
                w(f"  {'Other':<20} {d['total']:>8} {d['critical']:>10} {d['high']:>8} {d['medium']:>8}")

        # Top offenders with max mph over
        if by_driver:
            w("")
            w("Top Speeding Offenders:")
            w(f"  {'Driver':<30} {'Events':>8} {'Max MPH Over':>14}")
            w(f"  {'-'*54}")
            for name, count in sorted(by_driver.items(), key=lambda x: x[1], reverse=True)[:15]:
                max_over = max_over_by_driver.get(name, 0)
                w(f"  {name:<30} {count:>8} {max_over:>13.0f}")
    w("")

    # -- Camera / DriveCam Events --
    w("CAMERA/DRIVECAM EVENTS")
    w("-" * 80)

    # Filter camera events: high severity only, exclude dismissed
    filtered_camera = []
    dismissed_count = 0
    low_severity_count = 0
    for evt in camera:
        e = evt.get("driver_performance_event", evt) if isinstance(evt, dict) else evt
        coaching_status = str(e.get("coaching_status", "")).strip().lower()
        meta = e.get("metadata", {}) or {}
        severity = str(meta.get("severity", "")).strip().lower()

        if coaching_status in CAMERA_DISMISSED_STATUSES:
            dismissed_count += 1
        elif severity not in CAMERA_MIN_SEVERITY:
            low_severity_count += 1
        else:
            filtered_camera.append(evt)

    w(f"Total Camera Events: {len(filtered_camera)} (High Severity)")
    exclusions = []
    if dismissed_count > 0:
        exclusions.append(f"{dismissed_count} dismissed")
    if low_severity_count > 0:
        exclusions.append(f"{low_severity_count} low/medium severity")
    if exclusions:
        w(f"  (Excluded: {', '.join(exclusions)})")

    if filtered_camera:
        by_type = {}
        by_cam_driver = {}
        drowsiness_events = []
        coaching_counts = {}

        for evt in filtered_camera:
            e = evt.get("driver_performance_event", evt) if isinstance(evt, dict) else evt
            event_type = str(e.get("type", e.get("event_type", "Unknown"))).strip()
            by_type[event_type] = by_type.get(event_type, 0) + 1

            # Track coaching status
            cs = str(e.get("coaching_status", "unknown")).strip()
            coaching_counts[cs] = coaching_counts.get(cs, 0) + 1

            driver = _extract_driver(e)
            if driver:
                by_cam_driver[driver] = by_cam_driver.get(driver, 0) + 1

            # Track drowsiness/fatigue separately
            if any(kw in event_type.lower() for kw in ("drows", "fatigue", "yawn", "sleep")):
                drowsiness_events.append(e)

        w("")
        w("Events by Type:")
        for etype, count in sorted(by_type.items(), key=lambda x: x[1], reverse=True):
            pct = f"{(count/len(filtered_camera)*100):.0f}%" if len(filtered_camera) > 0 else "0%"
            w(f"  {etype}: {count} ({pct})")

        w("")
        w("Coaching Status:")
        for cs, count in sorted(coaching_counts.items(), key=lambda x: x[1], reverse=True):
            w(f"  {cs}: {count}")

        # Drowsiness/fatigue callout
        if drowsiness_events:
            w("")
            w(f"*** LIFE-SAFETY ALERT: {len(drowsiness_events)} DROWSINESS/FATIGUE EVENT(S) ***")
            for evt in drowsiness_events:
                driver = _extract_driver(evt)
                vid = _extract_vehicle_num(evt)
                dt = evt.get("start_time", evt.get("date", "Unknown"))
                w(f"  Driver: {driver or 'Unknown'}, Vehicle: {vid or 'Unknown'}, Date: {dt}")

        # Top camera offenders
        if by_cam_driver:
            w("")
            w("Top Camera Event Offenders:")
            w(f"  {'Driver':<30} {'Events':>8}")
            w(f"  {'-'*40}")
            for name, count in sorted(by_cam_driver.items(), key=lambda x: x[1], reverse=True)[:10]:
                w(f"  {name:<30} {count:>8}")

        # By yard
        if yard_order:
            by_yard_cam = {y: 0 for y in yard_order}
            by_yard_cam["Other"] = 0
            for evt in filtered_camera:
                e = evt.get("driver_performance_event", evt) if isinstance(evt, dict) else evt
                vid = _extract_vehicle_num(e)
                yard = vehicles.get(vid, {}).get("yard", "")
                bucket = yard if yard in by_yard_cam else "Other"
                by_yard_cam[bucket] += 1
            w("")
            w("Camera Events by Yard:")
            for yard in yard_order:
                if by_yard_cam[yard] > 0:
                    w(f"  {yard}: {by_yard_cam[yard]}")
            if by_yard_cam["Other"] > 0:
                w(f"  Other: {by_yard_cam['Other']}")
    w("")

    # -- Top vehicles by miles --
    if by_vehicle:
        w("TOP VEHICLES BY MILES")
        w("-" * 80)
        w(f"{'Vehicle':<15} {'Driver':<25} {'Miles':>10} {'% of Fleet':>10}")
        w("-" * 62)
        sorted_v = sorted(by_vehicle.items(), key=lambda x: x[1], reverse=True)[:15]
        for vid, miles in sorted_v:
            driver = vehicles.get(vid, {}).get("driver", "Unknown")
            pct = f"{(miles/total_miles*100):.1f}%" if total_miles > 0 else "0%"
            w(f"{vid:<15} {driver:<25} {miles:>10,.0f} {pct:>10}")
        w("")


# ===========================================================================
# Assessment Details -- Deduplicated
# ===========================================================================

def _render_assessments_text(w, assess_rows):
    """Render assessment details, grouped by report number.

    KPA responses.flat returns line-items for repeatable sections.
    Each row shares a 'report number' but header fields (date, observer)
    are often only on the first/parent row. We group by report number,
    merge all non-empty fields, and list employees assessed.
    """
    from collections import defaultdict, OrderedDict

    # Group rows by report number
    by_report = OrderedDict()
    for row in assess_rows:
        rn = row.get("report number", row.get("report_number", "unknown"))
        if rn not in by_report:
            by_report[rn] = []
        by_report[rn].append(row)

    w("ASSESSMENT DETAILS")
    w("-" * 80)
    w(f"Total Assessments: {len(by_report)} unique reports (from {len(assess_rows)} line-item rows)")
    w("")

    # Collect fields that indicate deficiencies
    deficiency_count = 0

    for i, (report_num, group) in enumerate(by_report.items(), 1):
        w(f"Assessment #{i} (Report #{report_num}):")
        w(f"  Line Items: {len(group)}")

        # Merge header fields from all rows (take first non-empty value)
        merged = {}
        for row in group:
            for k, v in row.items():
                if v and k not in merged:
                    merged[k] = v

        # Prefer conducted date (actual job date) over submission date
        conducted_raw = merged.get("tm4zqob5uficucju", "")
        if conducted_raw:
            # Format: "6/11/2026 1:14 PM" -> "2026-06-11"
            from datetime import datetime as _dt
            for fmt in ("%m/%d/%Y %I:%M %p", "%m/%d/%Y"):
                try:
                    date = _dt.strptime(conducted_raw.strip(), fmt).strftime("%Y-%m-%d")
                    break
                except ValueError:
                    continue
            else:
                date = conducted_raw.strip()
        else:
            date = merged.get("date", "")
        observer = merged.get("observer", merged.get("_observer", ""))
        score = merged.get("score-percent", "")
        district = _get_field(merged, "7vj2l992y7fwqhwz", "district", "District")
        report_name = merged.get("report", "")

        if date:
            w(f"  Date: {date[:10] if len(date) > 10 else date}")
        if observer:
            w(f"  Assessor: {observer}")
        if report_name:
            w(f"  Form: {report_name}")
        if district:
            w(f"  Yard/District: {district}")
        if score:
            w(f"  Score: {str(score).rstrip('%')}%")

        # Collect employee names from repeatable section fields
        # These are typically in fields like 'hashA-hashB' (compound hash = repeatable field)
        employees = []
        for row in group:
            for k, v in row.items():
                if v and "-" in k and len(k) > 20:
                    # Compound hash field from repeatable section
                    val = str(v).strip()
                    if val and val not in employees and not val.startswith("http"):
                        employees.append(val)

        if employees:
            w(f"  Employees Assessed: {', '.join(employees)}")

        # Check for deficiency markers: select-no fields, followup-ids, notes
        has_deficiency = False
        followup_fields = []
        notes_fields = []
        for row in group:
            for k, v in row.items():
                if not v:
                    continue
                v_str = str(v).strip()
                if k == "select-no" and v_str:
                    has_deficiency = True
                if k.endswith("-followups") and v_str:
                    followup_fields.append(v_str)
                    has_deficiency = True
                if k.endswith("-followup-ids") and v_str:
                    has_deficiency = True
                if k.endswith("-notes") and v_str and len(v_str) > 2:
                    notes_fields.append(v_str)

        if has_deficiency:
            deficiency_count += 1
            w("  *** DEFICIENCY IDENTIFIED ***")
            for ff in followup_fields[:5]:
                w(f"    Followup: {ff[:200]}")
        if notes_fields:
            for nf in notes_fields[:5]:
                w(f"  Note: {nf[:200]}")

        # Show non-empty content fields (skip metadata/empty hash fields)
        content_shown = 0
        for row in group:
            for k, v in row.items():
                if not v:
                    continue
                v_str = str(v).strip()
                if v_str.startswith("http"):
                    continue
                if k.lower() in SKIP_FIELDS:
                    continue
                if k in ("report number", "report_number", "date", "observer",
                         "_observer", "score-percent", "7vj2l992y7fwqhwz",
                         "report", "select-no", "select-yes", "select-n/a",
                         "select-safety", "select-hse", "select-quality",
                         "select-dot", "select-spider", "select-slips",
                         "select-flush mount spider", "select-surface",
                         "select-intermediate", "select-long", "select-other",
                         "parentrepnum", "parentlink"):
                    continue
                if k.endswith(("-followups", "-followup-ids", "-notes",
                               "-attachments", "-lat", "-lon")):
                    continue
                # Skip compound hash fields already shown as employees
                if "-" in k and len(k) > 20:
                    continue
                # Show remaining content fields
                label = INC_FIELD_LABELS.get(k, OBS_FIELD_LABELS.get(k, k))
                if content_shown < 10:  # cap per assessment to avoid noise
                    w(f"  {label}: {v_str[:200]}")
                    content_shown += 1
        w("")

    w(f"Assessments with deficiencies: {deficiency_count}/{len(by_report)}")
    w("")


# ===========================================================================
# JSA Summary
# ===========================================================================

def _render_jsa_summary(w, jsa_rows, prev_jsa_rows, review_rows, cfg):
    """Render JSA log summary with totals, by-rig breakdown, and review rate."""
    w("JSA LOG SUMMARY")
    w("-" * 80)
    w(f"Total JSAs: {len(jsa_rows)} (Previous Month: {len(prev_jsa_rows)})")
    w("")

    if not jsa_rows:
        w("  No JSAs recorded")
        w("")
        return

    # By rig (using district/rig field)
    rig_field = cfg.get("jsa_fields", {}).get("district", "25dzncbqyxgx39xq")
    by_rig = {}
    by_observer = {}
    for row in jsa_rows:
        rig = str(row.get(rig_field, "")).strip()
        if not rig:
            rig = "Unknown"
        by_rig[rig] = by_rig.get(rig, 0) + 1

        observer = row.get("observer", row.get("_observer", "Unknown"))
        if observer:
            by_observer[observer] = by_observer.get(observer, 0) + 1

    if by_rig:
        w("JSAs by Rig/District:")
        for r, c in sorted(by_rig.items(), key=lambda x: x[1], reverse=True):
            w(f"  {r}: {c}")
        w("")

    # JSA review rate
    if review_rows:
        review_rate = round(len(review_rows) / len(jsa_rows) * 100, 1) if jsa_rows else 0
        w(f"JSA Reviews: {len(review_rows)} ({review_rate:.1f}% review rate)")
        w("")

    # Top JSA submitters
    if by_observer:
        w("Top JSA Submitters:")
        w(f"  {'Name':<30} {'Count':>8}")
        w(f"  {'-'*40}")
        for name, count in sorted(by_observer.items(), key=lambda x: x[1], reverse=True)[:10]:
            w(f"  {name:<30} {count:>8}")
        w("")


# ===========================================================================
# Rig Inspections
# ===========================================================================

def _render_rig_inspections_text(w, insp_rows, prev_insp_rows, cfg):
    """Render rig inspection summary with scores and deficiency themes."""
    from collections import OrderedDict

    w("RIG INSPECTIONS")
    w("-" * 80)
    w(f"Total Inspections: {len(insp_rows)} (Previous Month: {len(prev_insp_rows)})")
    w("")

    if not insp_rows:
        w("  No rig inspections recorded")
        w("")
        return

    # Group by report number (same pattern as assessments)
    by_report = OrderedDict()
    for row in insp_rows:
        rn = row.get("report number", row.get("report_number", "unknown"))
        if rn not in by_report:
            by_report[rn] = []
        by_report[rn].append(row)

    scores = []
    for i, (report_num, group) in enumerate(by_report.items(), 1):
        merged = {}
        for row in group:
            for k, v in row.items():
                if v and k not in merged:
                    merged[k] = v

        date = merged.get("date", "")
        observer = merged.get("observer", merged.get("_observer", ""))
        score_pct = merged.get("score-percent", "")
        report_name = merged.get("report", "")

        w(f"Inspection #{i} (Report #{report_num}):")
        if date:
            w(f"  Date: {date[:10] if len(date) > 10 else date}")
        if observer:
            w(f"  Inspector: {observer}")
        if report_name:
            w(f"  Form: {report_name}")
        if score_pct:
            clean_score = str(score_pct).rstrip("%")
            w(f"  Score: {clean_score}%")
            try:
                scores.append(float(clean_score))
            except (ValueError, TypeError):
                pass

        # Check for deficiencies (select-no fields)
        has_deficiency = False
        for row in group:
            for k, v in row.items():
                if k == "select-no" and v:
                    has_deficiency = True
                    break
        if has_deficiency:
            w("  *** DEFICIENCY IDENTIFIED ***")
        w("")

    if scores:
        avg_score = sum(scores) / len(scores)
        w(f"Average Score: {avg_score:.1f}%")
        w(f"Score Range: {min(scores):.0f}% - {max(scores):.0f}%")
        w("")


# ===========================================================================
# Vehicle Inspection Details (Enhanced)
# ===========================================================================

def _render_vehicle_inspection_details(w, insp_rows, cfg):
    """Render detailed vehicle inspection records with scores, drivers, deficiencies."""
    from collections import OrderedDict

    # Check if division has specific inspection field hashes (like BTI)
    insp_fields = cfg.get("inspection_fields", {})

    w("VEHICLE INSPECTION DETAILS")
    w("-" * 80)

    # Group by report number
    by_report = OrderedDict()
    for row in insp_rows:
        rn = row.get("report number", row.get("report_number", "unknown"))
        if rn not in by_report:
            by_report[rn] = []
        by_report[rn].append(row)

    scores = []
    deficiency_themes = {}

    for i, (report_num, group) in enumerate(by_report.items(), 1):
        merged = {}
        for row in group:
            for k, v in row.items():
                if v and k not in merged:
                    merged[k] = v

        date = merged.get("date", "")
        observer = merged.get("observer", merged.get("_observer", ""))
        score_pct = merged.get("score-percent", "")

        # Use division-specific field hashes if available
        driver = ""
        truck = ""
        if insp_fields:
            driver = str(merged.get(insp_fields.get("driver", ""), "")).strip()
            truck = str(merged.get(insp_fields.get("truck", ""), "")).strip()

        w(f"Inspection #{i}:")
        if date:
            w(f"  Date: {date[:10] if len(date) > 10 else date}")
        if observer:
            w(f"  Inspector: {observer}")
        if driver:
            w(f"  Driver: {driver}")
        if truck:
            w(f"  Truck: {truck}")
        if score_pct:
            clean_score = str(score_pct).rstrip("%")
            w(f"  Score: {clean_score}%")
            try:
                scores.append(float(clean_score))
            except (ValueError, TypeError):
                pass

        # Check for deficiencies
        has_deficiency = False
        followup_fields = []
        for row in group:
            for k, v in row.items():
                if not v:
                    continue
                v_str = str(v).strip()
                if k == "select-no" and v_str:
                    has_deficiency = True
                if k.endswith("-followups") and v_str:
                    followup_fields.append(v_str)
                    has_deficiency = True
                    # Track deficiency themes
                    for theme in v_str.split(","):
                        theme = theme.strip()
                        if theme:
                            deficiency_themes[theme] = deficiency_themes.get(theme, 0) + 1

        if has_deficiency:
            w("  *** DEFICIENCY IDENTIFIED ***")
            for ff in followup_fields[:3]:
                w(f"    Finding: {ff[:200]}")
        w("")

    if scores:
        avg_score = sum(scores) / len(scores)
        w(f"Average Score: {avg_score:.1f}%")
        w(f"Score Range: {min(scores):.0f}% - {max(scores):.0f}%")
        w(f"Inspections with Scores: {len(scores)}/{len(by_report)}")
        w("")

    if deficiency_themes:
        w("Top Deficiency Themes:")
        for theme, count in sorted(deficiency_themes.items(), key=lambda x: x[1], reverse=True)[:10]:
            w(f"  {theme}: {count}")
        w("")


# ===========================================================================
# Helpers
# ===========================================================================

def _get_field(row, hash_key, alt_key, label=None):
    """Get a field value trying hash key first, then alt key."""
    val = row.get(hash_key, "")
    if not val:
        val = row.get(alt_key, "")
    if not val and label:
        val = row.get(label, "")
    return str(val).strip() if val else ""


def _safe_float(val):
    try:
        return float(val or 0)
    except (ValueError, TypeError):
        return 0.0


def _extract_driver(event_dict):
    d = event_dict.get("driver", {})
    if isinstance(d, dict):
        return f"{d.get('first_name', '')} {d.get('last_name', '')}".strip()
    return ""


def _extract_vehicle_num(event_dict):
    v = event_dict.get("vehicle", {})
    if isinstance(v, dict):
        return str(v.get("number", ""))
    return ""
