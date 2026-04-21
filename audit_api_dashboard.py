"""
API + DASHBOARD RENDERING AUDIT
================================
Comprehensive audit that:
1. Calls the Render API for multiple date ranges
2. Validates API data quality (division mapping, tier classification, counts)
3. Opens dashboard in Playwright, selects date ranges, reads rendered values
4. Compares dashboard display vs API response

Usage:
    python audit_api_dashboard.py
"""

import json
import os
import sys
import time
from datetime import date, timedelta

import requests

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(line_buffering=True)

API_URL = "https://brhas-safety-api.onrender.com"
DASH_URL = "http://localhost:8888/dashboard.html"
RESULTS = []


def check(name, passed, expected, actual, detail=""):
    status = "PASS" if passed else "FAIL"
    icon = "[PASS]" if passed else "[FAIL]"
    RESULTS.append({"check": name, "status": status, "expected": expected,
                    "actual": actual, "detail": detail})
    msg = f"  {icon} {name}"
    if not passed:
        msg += f" -- expected: {expected}, got: {actual}"
        if detail:
            msg += f" ({detail})"
    print(msg)


# =============================================================================
# PART 1: API DATA QUALITY AUDIT
# =============================================================================

def audit_api_data_quality():
    """Validate API response data quality across multiple date ranges."""
    print("\n" + "=" * 60)
    print("  PART 1: API DATA QUALITY")
    print("=" * 60)

    today = date.today()
    ranges = {
        "yesterday": (today - timedelta(days=1), today - timedelta(days=1)),
        "last_week": (today - timedelta(days=today.weekday() + 7),
                      today - timedelta(days=today.weekday() + 1)),
        "this_month": (today.replace(day=1), today),
    }

    for label, (start, end) in ranges.items():
        print(f"\n-- Date range: {label} ({start} to {end}) --")

        try:
            r = requests.get(f"{API_URL}/api/data",
                           params={"start": str(start), "end": str(end)},
                           timeout=120)
            data = r.json()
        except Exception as e:
            check(f"API fetch {label}", False, "JSON", str(e)[:80])
            continue

        check(f"API response {label}", r.status_code == 200, 200, r.status_code)

        # --- Speeding checks ---
        spd = data.get("speeding") or {}
        spd_events = spd.get("events", [])
        spd_total = spd.get("total_events", 0)
        spd_summary = spd.get("summary", {})
        spd_by_div = spd.get("by_division", {})

        check(f"Speeding events list matches total ({label})",
              len(spd_events) == spd_total, spd_total, len(spd_events))

        # Tier sum
        tier_sum = sum(spd_summary.get(t, 0) for t in ["red", "orange", "yellow"])
        check(f"Speeding tier sum ({label})",
              tier_sum == spd_total, spd_total, tier_sum)

        # Division mapping -- how many are "Unknown"?
        unknown_count = spd_by_div.get("Unknown", 0)
        unknown_pct = (unknown_count / spd_total * 100) if spd_total > 0 else 0
        check(f"Speeding division mapped ({label})",
              unknown_pct < 20,
              "< 20% unknown", f"{unknown_pct:.0f}% unknown ({unknown_count}/{spd_total})",
              "API should resolve vehicle -> division via group IDs")

        # Tier classification -- check individual events
        if spd_events:
            tier_errors = 0
            for e in spd_events[:50]:  # Sample first 50
                speed = e.get("speed", e.get("recorded_speed_mph", 0))
                posted = e.get("posted_speed", e.get("posted_limit_mph", 0))
                over = e.get("overspeed", e.get("mph_over_limit", speed - posted))
                tier = e.get("tier", e.get("severity", ""))
                if over >= 20 or speed >= 90:
                    exp = "RED"
                elif over >= 15:
                    exp = "ORANGE"
                else:
                    exp = "YELLOW"
                if tier.upper() != exp:
                    tier_errors += 1
            check(f"Speeding tier classification ({label})",
                  tier_errors == 0, "0 errors", f"{tier_errors} errors")

        # --- KPA checks ---
        kpa = data.get("kpa") or {}
        obs = kpa.get("observations", [])
        inc = kpa.get("incidents", [])
        nm = kpa.get("near_misses", [])
        assess = kpa.get("assessments", [])

        # Reasonable count check (multi-day range should have some data)
        days_in_range = (end - start).days + 1
        if days_in_range >= 5:
            check(f"KPA observations reasonable ({label})",
                  len(obs) >= 1, ">= 1", len(obs),
                  f"{days_in_range} days should have observations")

        # --- YTD checks ---
        ytd = data.get("ytd") or {}
        trir = ytd.get("ytd_trir")
        man_hours = ytd.get("ytd_man_hours", 0)
        recordables = ytd.get("ytd_recordables", 0)

        if trir is not None and man_hours > 0:
            expected_trir = recordables * 200000 / man_hours
            check(f"TRIR formula ({label})",
                  abs(expected_trir - trir) < 0.05,
                  f"{expected_trir:.2f}", f"{trir:.2f}")

        check(f"Man-hours positive ({label})",
              man_hours > 0, "> 0", man_hours)

        # --- CA checks ---
        cas = data.get("cas") or {}
        ca_list = cas.get("corrective_actions", [])
        ca_total = cas.get("total_open", 0)
        ca_overdue = cas.get("total_overdue", 0)

        check(f"CA list matches total ({label})",
              len(ca_list) == ca_total, ca_total, len(ca_list))
        check(f"CA overdue <= total ({label})",
              ca_overdue <= ca_total, f"<= {ca_total}", ca_overdue)


# =============================================================================
# PART 2: DASHBOARD RENDERING AUDIT (Playwright)
# =============================================================================

def audit_dashboard_rendering():
    """Compare what the dashboard displays vs what the API returns."""
    print("\n" + "=" * 60)
    print("  PART 2: DASHBOARD RENDERING vs API")
    print("=" * 60)

    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        print("  [SKIP] Playwright not installed")
        return

    today = date.today()
    yesterday = today - timedelta(days=1)

    # Get API data for "yesterday" to compare
    try:
        r = requests.get(f"{API_URL}/api/data",
                        params={"start": str(yesterday), "end": str(yesterday)},
                        timeout=120)
        api_data = r.json()
    except Exception as e:
        check("API fetch for rendering comparison", False, "JSON", str(e)[:80])
        return

    api_spd = api_data.get("speeding", {})
    api_kpa = api_data.get("kpa", {})
    api_ytd = api_data.get("ytd", {})
    api_cas = api_data.get("cas", {})

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page(viewport={"width": 1920, "height": 1080})

        # Collect console errors
        console_errors = []
        page.on("console", lambda m: console_errors.append(m.text)
                if m.type == "error" else None)

        print("\n  Loading dashboard...")
        page.goto(DASH_URL, wait_until="networkidle", timeout=30000)
        page.wait_for_timeout(3000)

        # Select "yesterday" date range
        print("  Selecting 'yesterday' date range...")
        page.select_option("#dateRangeSelect", "yesterday")
        page.wait_for_timeout(8000)  # Wait for API fetch + render

        # Check for JS errors during API fetch
        api_errors = [e for e in console_errors if "API" in e or "fetch" in e or "Error" in e]
        check("No JS errors during API fetch",
              len(api_errors) == 0, "0 errors", f"{len(api_errors)} errors",
              "; ".join(api_errors[:3]) if api_errors else "")

        # Check API banner -- info banner is OK (shows date range), error banner is not
        banner_text = page.evaluate('''() => {
            var b = document.getElementById('apiBanner');
            return b && b.style.display !== 'none' ? b.textContent : null;
        }''')
        is_info_banner = banner_text and "Showing data for" in banner_text
        is_error_banner = banner_text and ("unavailable" in banner_text or "offline" in banner_text)
        check("API banner OK (info, not error)",
              not is_error_banner, "info or hidden",
              "error: " + banner_text if is_error_banner else "ok")

        # --- Read Overview tab values ---
        print("\n  Reading Overview tab...")
        page.click('.nav-item[data-tab="overview"]')
        page.wait_for_timeout(2000)

        overview_vals = page.evaluate('''() => {
            var result = {};
            // Try to find TRIR value
            var cards = document.querySelectorAll('.card-num, .kpi-value, .metric-value');
            result.cardTexts = [];
            cards.forEach(function(c) {
                if (c.offsetParent !== null) {
                    result.cardTexts.push(c.textContent.trim());
                }
            });
            // Get all visible text in exec summary
            var exec = document.getElementById('execSummarySection');
            result.execText = exec ? exec.innerText.substring(0, 2000) : '';
            return result;
        }''')

        # Check TRIR appears somewhere on the overview tab (any section)
        overview_full = page.evaluate('''() => {
            var tab = document.querySelector('[data-tab-content="overview"]');
            if (!tab) {
                var tabs = document.querySelectorAll('.tab-content');
                for (var i = 0; i < tabs.length; i++) {
                    if (tabs[i].style.display !== 'none') { tab = tabs[i]; break; }
                }
            }
            return tab ? tab.innerText : '';
        }''')
        api_trir = api_ytd.get("ytd_trir")
        if api_trir is not None:
            trir_strs = [f"{api_trir:.2f}", f"{api_trir:.1f}", str(api_trir)]
            found = any(t in overview_full for t in trir_strs)
            check("TRIR displayed on overview tab",
                  found, trir_strs[0],
                  "found" if found else f"not found (searched: {trir_strs})")

        # --- Read Fleet tab values ---
        print("\n  Reading Fleet tab...")
        page.click('.nav-item[data-tab="fleet"]')
        page.wait_for_timeout(2000)

        fleet_vals = page.evaluate('''() => {
            var result = {};
            var fleetTab = document.querySelector('[data-tab-content="fleet"]');
            if (!fleetTab) fleetTab = document.querySelector('.tab-content:not([style*="display: none"])');
            result.fleetText = fleetTab ? fleetTab.innerText.substring(0, 3000) : '';

            // Look for speeding section
            var spdSection = document.getElementById('speedingSection');
            result.speedingText = spdSection ? spdSection.innerText.substring(0, 1000) : '';

            // Look for camera section
            var camSection = document.getElementById('cameraSummarySection');
            result.cameraText = camSection ? camSection.innerText.substring(0, 1000) : '';

            return result;
        }''')

        fleet_text = fleet_vals.get("fleetText", "")
        api_spd_total = api_spd.get("total_events", 0)

        # Check speeding total appears
        check("Speeding total displayed on Fleet tab",
              str(api_spd_total) in fleet_text,
              str(api_spd_total), "found" if str(api_spd_total) in fleet_text else "not found",
              f"Fleet tab text length: {len(fleet_text)}")

        # --- Read CAs tab values ---
        print("\n  Reading CAs tab...")
        page.click('.nav-item[data-tab="cas"]')
        page.wait_for_timeout(2000)

        ca_vals = page.evaluate('''() => {
            var result = {};
            var casTab = document.querySelector('.tab-content:not([style*="display: none"])');
            result.casText = casTab ? casTab.innerText.substring(0, 3000) : '';
            return result;
        }''')

        cas_text = ca_vals.get("casText", "")

        # Dashboard filters CAs through isSafetyCA -- just verify CAs tab has data
        # Look for "TOTAL OPEN" and "OVERDUE" labels with any number
        has_total = "TOTAL OPEN" in cas_text
        has_overdue = "OVERDUE" in cas_text
        check("CAs tab shows total open count",
              has_total, "TOTAL OPEN label present",
              "found" if has_total else "not found")
        check("CAs tab shows overdue count",
              has_overdue, "OVERDUE label present",
              "found" if has_overdue else "not found")

        # Verify numbers are reasonable (dashboard filters, so count may differ from API)
        ca_numbers = page.evaluate('''() => {
            var sec = document.getElementById('casDedicatedSection');
            if (!sec) return {};
            var cards = sec.querySelectorAll('.card-num, .kpi-value');
            var nums = [];
            cards.forEach(function(c) { nums.push(c.textContent.trim()); });
            return {numbers: nums, text: sec.innerText.substring(0, 500)};
        }''')
        ca_nums = ca_numbers.get("numbers", [])
        check("CAs tab has numeric values",
              len(ca_nums) > 0, "> 0 values", f"{len(ca_nums)} values: {ca_nums[:6]}")

        # --- Read Field tab values ---
        print("\n  Reading Field tab...")
        page.click('.nav-item[data-tab="field"]')
        page.wait_for_timeout(2000)

        field_vals = page.evaluate('''() => {
            var result = {};
            var tab = document.querySelector('.tab-content:not([style*="display: none"])');
            result.fieldText = tab ? tab.innerText.substring(0, 3000) : '';
            return result;
        }''')

        field_text = field_vals.get("fieldText", "")
        api_obs_count = len(api_kpa.get("observations", []))

        check("Observations count displayed on Field tab",
              str(api_obs_count) in field_text,
              str(api_obs_count), "found" if str(api_obs_count) in field_text else "not found")

        # --- Check for NaN/undefined/null in visible text ---
        print("\n  Checking for rendering errors...")
        page.click('.nav-item[data-tab="overview"]')
        page.wait_for_timeout(1000)

        tabs = ["overview", "fleet", "field", "training", "cas", "detail"]
        for tab in tabs:
            page.click(f'.nav-item[data-tab="{tab}"]')
            page.wait_for_timeout(800)
            bad_text = page.evaluate('''(tabName) => {
                var content = document.querySelector('.tab-content:not([style*="display: none"])');
                if (!content) return [];
                var text = content.innerText;
                var issues = [];
                if (text.match(/\\bNaN\\b/)) issues.push('NaN');
                if (text.match(/\\bundefined\\b/)) issues.push('undefined');
                if (text.match(/\\bnull\\b/) && !text.match(/no.*null|null.*no/i)) issues.push('null');
                if (text.match(/\\[object/)) issues.push('[object]');
                if (text.match(/26A0/)) issues.push('26A0 raw hex');
                return issues;
            }''', tab)
            check(f"No rendering errors on {tab} tab",
                  len(bad_text) == 0, "clean", ", ".join(bad_text) if bad_text else "clean")

        # Take screenshot for reference
        page.click('.nav-item[data-tab="overview"]')
        page.wait_for_timeout(1000)
        page.screenshot(path="C:/Users/krhod/daily-safety-report/audit_api_dashboard.png",
                       full_page=False)
        print("\n  Screenshot saved: audit_api_dashboard.png")

        browser.close()


# =============================================================================
# MAIN
# =============================================================================

def main():
    print("=" * 60)
    print("  API + DASHBOARD RENDERING AUDIT")
    print(f"  Date: {date.today()}")
    print(f"  API: {API_URL}")
    print(f"  Dashboard: {DASH_URL}")
    print("=" * 60)

    audit_api_data_quality()
    audit_dashboard_rendering()

    # Summary
    print("\n" + "=" * 60)
    passed = sum(1 for r in RESULTS if r["status"] == "PASS")
    failed = sum(1 for r in RESULTS if r["status"] == "FAIL")
    total = len(RESULTS)
    print(f"  RESULTS: {passed}/{total} passed, {failed} failed")
    if failed > 0:
        print(f"\n  FAILURES:")
        for r in RESULTS:
            if r["status"] == "FAIL":
                print(f"    - {r['check']}: expected {r['expected']}, got {r['actual']}")
                if r["detail"]:
                    print(f"      {r['detail']}")
    print("=" * 60)

    # Save report
    report = {
        "audit_date": str(date.today()),
        "total_checks": total, "passed": passed, "failed": failed,
        "pass_rate": f"{passed/total*100:.1f}%" if total > 0 else "N/A",
        "checks": RESULTS,
    }
    with open("audit_api_dashboard_report.json", "w") as f:
        json.dump(report, f, indent=2)
    print(f"  Report saved: audit_api_dashboard_report.json")

    return 0 if failed == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
