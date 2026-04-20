"""
Automated Dashboard Audit Script
Tests every valid filter combination across all 6 tabs.
Catches JS errors, NaN/undefined, empty sections, broken charts.
Generates HTML report with screenshots for failures.
"""

import json
import os
import sys
import time
import html as html_lib
from datetime import datetime
from playwright.sync_api import sync_playwright

# Force unbuffered output on Windows
sys.stdout.reconfigure(line_buffering=True)

DASHBOARD_URL = "http://localhost:8888/dashboard.html"
OUTPUT_DIR = os.path.join(os.path.dirname(__file__), "audit_output")
SCREENSHOT_DIR = os.path.join(OUTPUT_DIR, "screenshots")

# Valid filter mappings (from dashboard.html)
COMPANY_DIVISIONS = {
    "All": [],
    "BRHAS": ["All Divisions", "Casing", "Rathole", "Pit Lining", "Poly Pipe",
              "Anchors", "Construction", "Drilling Tools", "Environmental",
              "Fencing", "Shop", "Corporate"],
    "Valor": [],
    "BTI": [],
    "Transcend": [],
    "Permian": [],
}

DIVISION_YARDS = {
    "Casing": ["All Yards", "Midland", "Bryan", "Kilgore", "Hobbs", "Jourdanton", "Laredo"],
    "Rathole": ["All Yards", "Midland", "Levelland", "Barstow", "Jourdanton",
                "Ohio", "Pennsylvania", "Oklahoma", "North Dakota"],
    "Poly Pipe": ["All Yards", "Midland"],
    "Pit Lining": ["All Yards", "Midland"],
    "Anchors": ["All Yards", "Midland", "Seminole", "Levelland"],
    "Construction": ["All Yards", "Odessa"],
    "Environmental": ["All Yards", "Midland"],
    "Fencing": ["All Yards", "Midland"],
    "Drilling Tools": ["All Yards", "Midland"],
    "Shop": ["All Yards", "Midland", "Levelland"],
    "Corporate": ["All Yards", "Corporate"],
}

TABS = ["overview", "fleet", "field", "training", "cas", "detail"]
# Only test "today" -- other date ranges are disabled when API is offline (local testing)
# The API disables non-today options with "(API offline)" when brhas-safety-api is unreachable
DATE_RANGES = ["today"]


def build_combos():
    """Generate all valid (company, division, yard, dateRange) tuples."""
    combos = []
    for company, divisions in COMPANY_DIVISIONS.items():
        if not divisions:
            # No division filter for this company
            for dr in DATE_RANGES:
                combos.append((company, "All Divisions", "All Yards", dr))
        else:
            for div in divisions:
                yards = DIVISION_YARDS.get(div, ["All Yards"])
                if div == "All Divisions":
                    yards = ["All Yards"]
                for yard in yards:
                    for dr in DATE_RANGES:
                        combos.append((company, div, yard, dr))
    return combos


def check_issues(page, tab):
    """Run all issue detection checks on the current tab view."""
    issues = []

    result = page.evaluate("""(tab) => {
        const issues = [];
        const tabEl = document.getElementById('tab-' + tab);
        if (!tabEl) {
            issues.push({type: 'MISSING_TAB', msg: 'Tab element #tab-' + tab + ' not found'});
            return issues;
        }

        // Check for NaN/undefined in visible card numbers
        const cards = tabEl.querySelectorAll('.card-num');
        cards.forEach((card, i) => {
            const text = card.textContent.trim();
            if (text === 'NaN' || text === 'undefined' || text === 'null') {
                issues.push({type: 'BAD_VALUE', msg: 'Card #' + i + ' shows "' + text + '"'});
            }
        });

        // Check for broken canvas (0 dimensions)
        const canvases = tabEl.querySelectorAll('canvas');
        canvases.forEach((c, i) => {
            const rect = c.getBoundingClientRect();
            if (rect.width === 0 || rect.height === 0) {
                issues.push({type: 'BROKEN_CHART', msg: 'Canvas #' + i + ' has 0 dimensions (' + rect.width + 'x' + rect.height + ')'});
            }
        });

        // Check for empty section containers (rendered but blank)
        // Only flag PRIMARY sections (direct children of the tab panel), not sub-sections
        // Sub-sections (like cameraByYardSection) are cleared by parent render functions when
        // the parent already displays a "no data" message, so they're expected to be empty.
        const primarySections = [
            'execSummarySection', 'portfolioScorecardSection', 'leadingIndicatorsSection',
            'trirTrendSection', 'redFlagSection', 'yardScorecardSection',
            'fleetKPISection', 'cameraSummarySection', 'speedingSummarySection',
            'obsCultureSection', 'bbsGoalsSection', 'assessAccountSection',
            'trainingDedicatedSection', 'casDedicatedSection', 'caAgingSection', 'nmToCASection',
            'incidentsSection', 'nearMissSection', 'weekendSpotlightSection', 'investigationTrackerSection'
        ];
        primarySections.forEach((id) => {
            const s = tabEl.querySelector('#' + id);
            if (s && s.innerHTML.trim() === '') {
                issues.push({type: 'EMPTY_SECTION', msg: 'Section #' + id + ' is empty'});
            }
        });

        // Check section titles for empty text
        const titles = tabEl.querySelectorAll('.section-title, h2');
        titles.forEach((t, i) => {
            if (t.textContent.trim() === '') {
                issues.push({type: 'EMPTY_TITLE', msg: 'Section title #' + i + ' is empty'});
            }
        });

        // Check for "undefined" or "NaN" appearing in general text
        const allText = tabEl.innerText || '';
        // Exclude expected uses
        const lines = allText.split('\\n');
        lines.forEach((line, i) => {
            const trimmed = line.trim();
            if (!trimmed) return;
            // Skip lines that are just labels or expected text
            if (trimmed === 'undefined' || /\bundefined\b/.test(trimmed)) {
                if (trimmed.indexOf('unavailable') === -1 && trimmed.indexOf('not defined') === -1) {
                    issues.push({type: 'UNDEFINED_TEXT', msg: 'Line ' + i + ': "' + trimmed.substring(0, 80) + '"'});
                }
            }
            if (/\bNaN\b/.test(trimmed) && trimmed.indexOf('NaN') !== -1) {
                issues.push({type: 'NAN_TEXT', msg: 'Line ' + i + ': "' + trimmed.substring(0, 80) + '"'});
            }
        });

        // Check for horizontal overflow
        const scrollW = tabEl.scrollWidth;
        const clientW = tabEl.clientWidth;
        if (scrollW > clientW + 20) {
            issues.push({type: 'OVERFLOW', msg: 'Tab content overflows: ' + scrollW + 'px > ' + clientW + 'px'});
        }

        return issues;
    }""", tab)

    if result:
        issues.extend(result)

    return issues


def run_audit():
    os.makedirs(SCREENSHOT_DIR, exist_ok=True)

    combos = build_combos()
    total_checks = len(combos) * len(TABS)
    print(f"Audit starting: {len(combos)} filter combos x {len(TABS)} tabs = {total_checks} checks")

    results = []
    console_errors = []
    pass_count = 0
    fail_count = 0
    start_time = time.time()

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(viewport={"width": 1920, "height": 1080})
        page = context.new_page()

        # Collect console errors
        page.on("console", lambda msg: console_errors.append({
            "type": msg.type,
            "text": msg.text,
            "time": datetime.now().isoformat()
        }) if msg.type == "error" else None)

        # Collect page errors
        page_errors = []
        page.on("pageerror", lambda err: page_errors.append(str(err)))

        # Navigate and wait for data load
        print("Loading dashboard...")
        page.goto(DASHBOARD_URL, wait_until="networkidle")
        page.wait_for_timeout(2000)  # Extra time for renders
        print("Dashboard loaded.")

        for ci, (company, division, yard, date_range) in enumerate(combos):
            combo_label = f"{company} > {division} > {yard} > {date_range}"
            print(f"  [{ci+1}/{len(combos)}] {combo_label}")

            # Clear page errors for this combo
            page_errors.clear()

            try:
                # Set company
                page.select_option("#companySelect", company)
                page.wait_for_timeout(300)

                # Set division (if available in dropdown)
                div_options = page.evaluate("""() => {
                    const sel = document.getElementById('divisionSelect');
                    return Array.from(sel.options).map(o => o.value);
                }""")
                if division in div_options:
                    page.select_option("#divisionSelect", division)
                    page.wait_for_timeout(300)

                # Set yard (if available)
                loc_el = page.query_selector("#locationSelect")
                if loc_el:
                    loc_options = page.evaluate("""() => {
                        const sel = document.getElementById('locationSelect');
                        if (!sel) return [];
                        return Array.from(sel.options).map(o => o.value);
                    }""")
                    if yard in loc_options:
                        page.select_option("#locationSelect", yard)
                        page.wait_for_timeout(200)

                # Set date range
                page.select_option("#dateRangeSelect", date_range)
                page.wait_for_timeout(500)  # Wait for re-render

            except Exception as e:
                results.append({
                    "combo": combo_label,
                    "tab": "SETUP",
                    "issues": [{"type": "SETUP_ERROR", "msg": str(e)}],
                    "screenshot": None
                })
                fail_count += 1
                continue

            # Test each tab
            for tab in TABS:
                try:
                    # Click tab
                    tab_btn = page.query_selector(f'.nav-item[data-tab="{tab}"]')
                    if tab_btn:
                        tab_btn.click()
                        page.wait_for_timeout(400)

                    issues = check_issues(page, tab)

                    # Add any page errors
                    for pe in page_errors:
                        issues.append({"type": "PAGE_ERROR", "msg": pe[:200]})

                    screenshot_path = None
                    if issues:
                        fail_count += 1
                        # Take screenshot
                        ss_name = f"{ci}_{tab}_{company}_{division}_{yard}_{date_range}.png"
                        ss_name = ss_name.replace(" ", "_").replace("/", "_")
                        screenshot_path = os.path.join(SCREENSHOT_DIR, ss_name)
                        page.screenshot(path=screenshot_path, full_page=False)
                    else:
                        pass_count += 1

                    results.append({
                        "combo": combo_label,
                        "tab": tab,
                        "issues": issues,
                        "screenshot": screenshot_path
                    })

                except Exception as e:
                    fail_count += 1
                    results.append({
                        "combo": combo_label,
                        "tab": tab,
                        "issues": [{"type": "EXCEPTION", "msg": str(e)[:200]}],
                        "screenshot": None
                    })

                page_errors.clear()

        browser.close()

    elapsed = time.time() - start_time

    # Generate reports
    report = {
        "timestamp": datetime.now().isoformat(),
        "elapsed_seconds": round(elapsed, 1),
        "total_combos": len(combos),
        "total_checks": len(combos) * len(TABS),
        "pass_count": pass_count,
        "fail_count": fail_count,
        "pass_rate": round(pass_count / max(pass_count + fail_count, 1) * 100, 1),
        "console_errors": console_errors[:50],
        "results": results,
    }

    json_path = os.path.join(OUTPUT_DIR, "audit_report.json")
    with open(json_path, "w") as f:
        json.dump(report, f, indent=2)

    # Generate HTML report
    generate_html_report(report)

    print(f"\n{'='*60}")
    print(f"AUDIT COMPLETE")
    print(f"  Time: {elapsed:.1f}s")
    print(f"  Combos tested: {len(combos)}")
    print(f"  Total checks: {pass_count + fail_count}")
    print(f"  PASS: {pass_count}  FAIL: {fail_count}")
    print(f"  Pass rate: {report['pass_rate']}%")
    print(f"  Console errors: {len(console_errors)}")
    print(f"  Report: {os.path.join(OUTPUT_DIR, 'audit_report.html')}")
    print(f"{'='*60}")

    return report


def generate_html_report(report):
    """Generate an HTML audit report."""
    failures = [r for r in report["results"] if r["issues"]]

    # Group failures by issue type
    issue_types = {}
    for r in failures:
        for issue in r["issues"]:
            t = issue["type"]
            if t not in issue_types:
                issue_types[t] = []
            issue_types[t].append({"combo": r["combo"], "tab": r["tab"], "msg": issue["msg"]})

    html = f"""<!DOCTYPE html>
<html><head><meta charset="utf-8">
<title>Dashboard Audit Report</title>
<style>
body {{ font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif; max-width: 1200px; margin: 0 auto; padding: 20px; background: #f8fafc; }}
h1 {{ color: #0f172a; }}
.summary {{ display: flex; gap: 16px; margin: 20px 0; }}
.stat {{ background: white; border-radius: 12px; padding: 20px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); text-align: center; flex: 1; }}
.stat-num {{ font-size: 32px; font-weight: 700; }}
.stat-lbl {{ font-size: 13px; color: #64748b; margin-top: 4px; }}
.pass {{ color: #059669; }}
.fail {{ color: #dc2626; }}
.section {{ background: white; border-radius: 12px; padding: 20px; margin: 16px 0; box-shadow: 0 1px 3px rgba(0,0,0,0.1); }}
table {{ width: 100%; border-collapse: collapse; font-size: 13px; }}
th {{ text-align: left; padding: 8px 12px; border-bottom: 2px solid #e2e8f0; color: #475569; }}
td {{ padding: 8px 12px; border-bottom: 1px solid #f1f5f9; }}
tr:hover {{ background: #f8fafc; }}
.badge {{ display: inline-block; padding: 2px 8px; border-radius: 4px; font-size: 11px; font-weight: 600; }}
.badge-red {{ background: #fef2f2; color: #dc2626; }}
.badge-green {{ background: #f0fdf4; color: #059669; }}
.badge-yellow {{ background: #fffbeb; color: #d97706; }}
img {{ max-width: 100%; border: 1px solid #e2e8f0; border-radius: 8px; margin: 8px 0; }}
details {{ margin: 4px 0; }}
summary {{ cursor: pointer; font-weight: 600; color: #334155; }}
.console-err {{ background: #1e1e1e; color: #f87171; padding: 12px; border-radius: 8px; font-family: monospace; font-size: 12px; max-height: 200px; overflow-y: auto; }}
</style>
</head><body>
<h1>Dashboard Audit Report</h1>
<p style="color:#64748b">Generated: {report['timestamp']} | Duration: {report['elapsed_seconds']}s</p>

<div class="summary">
    <div class="stat"><div class="stat-num">{report['total_checks']}</div><div class="stat-lbl">Total Checks</div></div>
    <div class="stat"><div class="stat-num pass">{report['pass_count']}</div><div class="stat-lbl">Passed</div></div>
    <div class="stat"><div class="stat-num fail">{report['fail_count']}</div><div class="stat-lbl">Failed</div></div>
    <div class="stat"><div class="stat-num {'pass' if report['pass_rate'] >= 95 else 'fail'}">{report['pass_rate']}%</div><div class="stat-lbl">Pass Rate</div></div>
    <div class="stat"><div class="stat-num {'pass' if len(report['console_errors']) == 0 else 'fail'}">{len(report['console_errors'])}</div><div class="stat-lbl">Console Errors</div></div>
</div>
"""

    # Issue type summary
    if issue_types:
        html += '<div class="section"><h2>Issues by Type</h2><table><thead><tr><th>Type</th><th>Count</th><th>Sample</th></tr></thead><tbody>'
        for itype, items in sorted(issue_types.items(), key=lambda x: -len(x[1])):
            sample = html_lib.escape(items[0]["msg"][:100])
            html += f'<tr><td><span class="badge badge-red">{html_lib.escape(itype)}</span></td><td>{len(items)}</td><td>{sample}</td></tr>'
        html += '</tbody></table></div>'

    # Console errors
    if report["console_errors"]:
        html += '<div class="section"><h2>Console Errors</h2><div class="console-err">'
        for ce in report["console_errors"][:20]:
            html += html_lib.escape(ce["text"][:200]) + "<br>"
        html += '</div></div>'

    # Detailed failures
    if failures:
        html += f'<div class="section"><h2>Failures ({len(failures)})</h2>'
        for r in failures:
            combo_esc = html_lib.escape(r["combo"])
            html += f'<details><summary>{combo_esc} | <span class="badge badge-red">{html_lib.escape(r["tab"])}</span></summary>'
            html += '<ul>'
            for issue in r["issues"]:
                html += f'<li><strong>{html_lib.escape(issue["type"])}:</strong> {html_lib.escape(issue["msg"][:200])}</li>'
            html += '</ul>'
            if r["screenshot"]:
                rel_path = os.path.relpath(r["screenshot"], OUTPUT_DIR).replace("\\", "/")
                html += f'<img src="{rel_path}" alt="Screenshot" loading="lazy">'
            html += '</details>'
        html += '</div>'

    # All passes (collapsed)
    passes = [r for r in report["results"] if not r["issues"]]
    if passes:
        html += f'<div class="section"><details><summary><span class="badge badge-green">PASSED ({len(passes)})</span> - Click to expand</summary><ul>'
        for r in passes:
            html += f'<li>{html_lib.escape(r["combo"])} | {r["tab"]}</li>'
        html += '</ul></details></div>'

    html += '</body></html>'

    html_path = os.path.join(OUTPUT_DIR, "audit_report.html")
    with open(html_path, "w", encoding="utf-8") as f:
        f.write(html)


if __name__ == "__main__":
    run_audit()
