# Casing Safety Dashboard — Project Guide

## Overview
Safety dashboard for BRHAS (Butch's Rat Hole & Anchor Service) and its sister companies. Deployed to GitHub Pages with daily automated updates via GitHub Actions.

**Live dashboard:** https://krhodes5267.github.io/daily-safety-report/
**Repo:** https://github.com/krhodes5267/daily-safety-report

## Architecture
```
Data Sources → Daily Scripts → output/*.json → archive_today.py → archive/YYYY-MM-DD.json
                                                                         ↓
                                                 dashboard.html ← loads archive via fetch()
```
- **Novara Flex API** (novaraflex.com, formerly KPA): Observations, incidents, assessments, training, corrective actions
- **Motive API** (gomotive.com): Speeding events, camera events, device status, vehicle data
- **GitHub Action**: `daily-casing-dashboard.yml` runs at 5 AM CT, generates data, archives, deploys

## API Keys (local testing)
- `KPA_API_TOKEN=plnuF2fKcYuS07JoQHgOweB33Y8dldzdI`
- `MOTIVE_API_KEY=8d3dd502-36c0-47c4-ade3-a1fbbef0c05c`
- **KPA API CRITICAL:** Always use `requests.post(url, json=payload)` — requires JSON body, NOT query params. Token goes in JSON body.

## Key Files
| File | Purpose |
|------|---------|
| `dashboard.html` | Main dashboard (~2400 lines), all JS inline |
| `api_config.py` | Single source of truth: form IDs, group IDs, field hashes, man-hours |
| `dashboard_api.py` | Flask API for Render (live fallback), data extraction logic |
| `backfill_archive.py` | Bulk historical backfill (fetch-once-partition design) |
| `archive_today.py` | Daily archiver (runs in GitHub Action) |
| `audit_filters.py` | Filter accuracy audit script |
| `archive/` | 110+ daily JSON snapshots (2026 YTD) |

## Filter Chain
Dashboard filters: **Company → Division → Yard**

Filter maps in `dashboard.html` (lines ~1478-1521):
- `SPEEDING_TO_COMPANY` / `SPEEDING_TO_DIVISION` — Motive division names → company/division
- `KPA_SL_TO_COMPANY` / `KPA_SL_TO_DIVISION` — KPA service_line values → company/division
- `SINGLE_YARD_DIVISIONS` — divisions with only one yard (for yard inference)
- `COMPANY_DIVISIONS` — dropdown options per company

## KPA Field Hashes
| Hash | Field |
|------|-------|
| `bff8m4x6xbc033kg` | Observation type |
| `uncbcge9x8vow9pn` | Observation description |
| `lg5pnj4chjadnv46` | Observation location (rig/site name) |
| `7vj2l992y7fwqhwz` | Observation yard/district |
| `sha7vur5q2l6d6gq` | Service line |
| `lsx3msa0w9n9edb4` | Company |

## Running Commands
```bash
# Filter audit (verify all company/division/yard combos)
python audit_filters.py --range 2026-04-01 2026-04-20

# Backfill archive (after data extraction changes)
KPA_API_TOKEN=... MOTIVE_API_KEY=... python backfill_archive.py --start 2026-01-01 --end 2026-04-20

# Archive today's output
python archive_today.py
```

## Windows Notes
- Use ASCII-only in print() (cp1252 encoding) — no Unicode em-dashes or arrows
- Use forward slashes in Bash paths (`/c/Users/krhod/...`)

---

# Division Audit Protocol

## Philosophy
**Each division/company is its own entity.** Treat each like it has its own safety director. What matters to a trucking company director is completely different from a casing safety director. Custom analysis per division.

## Division Profiles

### BRHAS — Casing (Safety Director view)
- **Yards:** Bryan, Hobbs, Jourdanton, Kilgore, Laredo, Midland (6 active, San Angelo CLOSED)
- **KPA footprint:** Heavy — ~2000 observations YTD, 6 incidents, 1500+ assessments (form 381707 field assessment + 229645 pre/post trip)
- **Motive footprint:** Moderate speeding, heavy camera (all 6 yards)
- **KPA service_line values:** `Casing`
- **Motive groups:** 167175, 169090, 169092, 186740, 169091, 186739 (+ 186741 San Angelo)
- **What their safety director wants:**
  - Yard-vs-yard comparison (who's leading, who's falling behind)
  - Observation trends by yard (target per yard, are they hitting it?)
  - Observation quality — type distribution (too many Recognitions = padding?)
  - Assessment completion rate by yard
  - Speeding rate per vehicle per yard
  - Camera event trends by yard
  - Miles per event KPIs
  - Incident trends and severity
  - Top repeat offenders (speeding + camera)

### BRHAS — Rathole (Safety Director view)
- **Yards:** Midland, Levelland, Barstow, Jourdanton, Ohio, Pennsylvania, Oklahoma, North Dakota
- **KPA footprint:** Light observations (~20 YTD), but has dedicated **Rathole Field Assessments** (form 153181, ~30 YTD) — this is their main KPA tool
- **Motive footprint:** HEAVY speeding (1997 events YTD, #1 division)
- **KPA service_line values:** `Rat Hole`, `Rathole`
- **Motive groups:** 167176, 220453, 274965, 186742, 307752, 341789, 308775, 351218, 265996-266028, 290471
- **What their safety director wants:**
  - Assessment completion by yard (especially remote yards — OH, PA, OK, ND)
  - Which remote yards have ZERO safety activity? Flag them
  - Speeding rate by yard (miles per event)
  - Are out-of-state yards getting regular assessments?
  - Repeat speeding offenders across yards
  - Total miles driven by yard
  - Miles per speeding event KPI

### BRHAS — Poly Pipe
- **Yards:** Midland (main), Bryan, Hobbs, Jourdanton (sub-yards added recently)
- **KPA footprint:** Moderate (~37 obs YTD), assessments via form 226217
- **Motive groups:** 167180, 296017, 296020, 296036, 296040
- **KPA service_line values:** `Poly Pipe`
- **What matters:** Sub-yard growth tracking, are new yards submitting observations?

### BRHAS — Pit Lining (Water Solutions)
- **Yard:** Midland only
- **KPA footprint:** Light (~23 obs YTD), form 386087
- **KPA service_line values:** `Pit Lining`
- **What matters:** Small crew observation rate, assessment completion

### BRHAS — Anchors
- **Yard:** Midland only
- **KPA footprint:** Very light (~7 obs YTD)
- **KPA service_line values:** `Anchor`, `Anchors`
- **What matters:** Are they submitting any observations at all? Flag if zero activity weeks

### BRHAS — Construction
- **Yard:** Odessa (mapped as Midland in some Motive groups)
- **KPA footprint:** Light (~11 obs YTD), form 172295
- **KPA service_line values:** `Construction`, `Civil`, `Water/Construction`
- **What matters:** Are `Civil` and `Water/Construction` routing correctly?

### BRHAS — Environmental
- **Yard:** Midland only
- **KPA footprint:** Very light
- **KPA service_line values:** `Environmental`, `Containment`
- **What matters:** Is `Containment` routing here correctly?

### BRHAS — Fencing
- **Yard:** Midland only
- **KPA service_line values:** `Fencing`
- **What matters:** Basic activity check

### BRHAS — Drilling Tools
- **Yard:** Midland only
- **KPA service_line values:** `Drilling Tools`, `Downhole Tools` (note: Downhole Tools also maps to Valor in some contexts)
- **What matters:** Verify Drilling Tools vs Downhole Tools routing (BRHAS vs Valor)

### BRHAS — Shop/Fabrication
- **Yards:** Midland + Levelland (LL-FAB, LL-SHOP vehicles)
- **KPA service_line values:** `Shop`, `Fabrication`
- **Motive:** LL-FAB and LL-SHOP prefixed vehicles
- **What matters:** Are Levelland fab shop vehicles being captured?

### BRHAS — Corporate
- **Yard:** Midland
- **KPA service_line values:** `Corporate`
- **Motive group:** 265988 (WTC/admin vehicles)
- **What matters:** Minimal — just verify sales/admin vehicles route here

### Valor Energy Services (Separate company)
- **Yard:** Levelland
- **KPA footprint:** Observations under `Valor`, `Valor Energy`, `Valor Energy Services`
- **KPA service_line values:** `Valor`, `Valor Energy`, `Valor Energy Services`, `Downhole Tools`
- **Motive groups:** 167178, 265985
- **What matters:**
  - Zero cross-contamination with BRHAS
  - Observation rate for a small crew
  - Speeding and camera should ONLY show Valor vehicles

### BTI / Butch's Trucking (Separate company — TRUCKING)
- **Yard:** Midland
- **KPA footprint:** Minimal observations. Main KPA tool is **Vehicle Inspection Checklists** (form 152018) — Bernard's field truck audits (~47 YTD)
- **Motive footprint:** 335 speeding events YTD — trucking is ALL about Motive
- **KPA service_line values:** `BTI`, `Butch's Trucking`, `Trucking`
- **Motive groups:** 186743, 265989
- **What a trucking safety director wants:**
  - **Total miles driven** (critical KPI for trucking)
  - **Miles per speeding event** (fleet safety rate)
  - **Miles per camera event**
  - **FMCSA SAFER score** (look up at https://safer.fmcsa.dot.gov)
  - Speeding severity distribution (how many reds vs yellows?)
  - Vehicle inspection completion rate
  - DOT inspection results if available
  - Repeat offender drivers
  - Hours of service compliance (if available from Motive)
  - Camera coaching completion rate

### Transcend Drilling (Separate company — SPUDDER RIGS)
- **Yard:** Midland
- **KPA footprint:** Significant — 151 observations YTD broken by rig (`Drilling`, `Drilling - Rig 4`, `Drilling - Rig 16`, `Drilling - Rig 18`, `Drilling - Rig 20`), plus TD Rig Inspections (form 385365) and TD Observations (form 484193)
- **Motive groups:** 247035, 265986
- **KPA service_line values:** `Drilling`, `Transcend`, `Transcend Drilling`, `Drilling - Rig *`
- **What a rig safety director wants:**
  - **Rig-by-rig comparison** (obs count, assessment count, incident rate)
  - Which rigs are filing observations? Which are silent?
  - Rig inspection completion rate (form 385365)
  - Observation type quality by rig
  - Speeding for rig move crews
  - Incident trends by rig

### Permian Equipment Rentals (Separate company)
- **Yard:** Midland
- **KPA footprint:** Minimal
- **KPA service_line values:** `Rentals`, `PER`, `Permian`
- **Motive groups:** 265984
- **What matters:** Small operation — basic activity check, verify no cross-contamination

---

## Audit Steps (per division)

### Step 1: Data Inventory
- Count speeding, camera, observations, incidents, assessments for last 30 days
- Verify Motive group IDs all map correctly
- Verify KPA service_line values all route correctly
- Check yard field populates for multi-yard divisions

### Step 2: Data Quality
- Check for duplicates, missing fields, misrouted data
- Verify severity tiers calculate correctly
- Check observation type distribution (quality indicator)

### Step 3: Custom Analysis (per division profile above)
- Build the metrics that matter for THAT division's safety director
- Calculate KPIs (miles per event, obs per headcount, assessment completion rate)
- Identify trends, gaps, and red flags

### Step 4: Fix Issues
- Fix filter maps in dashboard.html + audit_filters.py
- Re-run backfill if archive data changed
- Re-run audit_filters.py to verify

### Step 5: Update division_audit.md with results

## After Each Division
Commit changes and update `division_audit.md` tracker.

---

# Monthly HSE Recap Generation

## Overview
The `generate_recap.py` script produces plain-text HSE data summaries for each division. These text files are then used as input to create formatted PDF reports (done separately via Claude Web or manual process).

## Key Files
| File | Purpose |
|------|---------|
| `generate_recap.py` | CLI entry point for monthly recap generation |
| `safety_recap/config.py` | Division configs: Motive group IDs, KPA form IDs, sections, field hashes |
| `safety_recap/data_fetcher.py` | `DataCollector` class - fetches from KPA API, Motive API, man-hours Excel |
| `safety_recap/text_exporter.py` | Formats collected data into plain-text `.txt` summaries |

## How to Run

### Generate a single division
```bash
cd C:/Users/krhod/daily-safety-report
KPA_API_TOKEN=plnuF2fKcYuS07JoQHgOweB33Y8dldzdI MOTIVE_API_KEY=8d3dd502-36c0-47c4-ade3-a1fbbef0c05c py -3 -u generate_recap.py --division transcend --month 2026-06 --data-only
```

### Generate all divisions
```bash
KPA_API_TOKEN=plnuF2fKcYuS07JoQHgOweB33Y8dldzdI MOTIVE_API_KEY=8d3dd502-36c0-47c4-ade3-a1fbbef0c05c py -3 -u generate_recap.py --division all --month 2026-06 --data-only
```

### With man-hours file
```bash
KPA_API_TOKEN=plnuF2fKcYuS07JoQHgOweB33Y8dldzdI MOTIVE_API_KEY=8d3dd502-36c0-47c4-ade3-a1fbbef0c05c py -3 -u generate_recap.py --division all --month 2026-06 --data-only --man-hours "C:/Users/krhod/Downloads/WorkMonthHoursReport_2026-06_2026-06.xlsx"
```

### Important flags
- `--division [key|all]` — Division key from config.py DIVISIONS dict, or `all`
- `--month YYYY-MM` — Report month
- `--data-only` — Text-only output (no HTML/PDF)
- `--man-hours [path]` — Path to man-hours Excel file

## Output Location
- Default: `~/Desktop/Monthly HSE Recaps/{Month Year}/`
- Copy finished files to `~/Downloads/` for easy access

## Windows-Specific Notes
- **Must use `py -3`** (not `python`) — Windows Python launcher
- **Use `-u` flag** for unbuffered output so you can see progress
- **API env vars inline:** `KPA_API_TOKEN=... MOTIVE_API_KEY=... py -3 -u generate_recap.py ...`
- **KPA rate limiting:** 1.5s delay between calls, 30/60/90s backoff on 429. Run divisions ONE AT A TIME to avoid excessive rate limiting
- **10-minute bash timeout:** For long runs, use `run_in_background` and check with `TaskOutput`

## Division List & Data Sources

### Divisions WITH Motive fleet data (KPA + Motive):
| Division Key | Display Name | Motive Group IDs |
|---|---|---|
| `casing` | Casing | 167175, 169090, 169092, 186740, 169091, 186739 |
| `rathole` | Rathole | 167176, 220453, 274965, 186742, 307752, 341789, 308775, 351218, 265996-266028, 290471 |
| `bti` | Butch's Trucking (BTI) | 265989 |
| `transcend` | Transcend Drilling | 265986 |
| `valor` | Valor Energy Services | 167178, 265985 |
| `poly_pipe` | Poly Pipe | 265993, 296040, 296036, 296017, 296020 |
| `anchors` | Anchors | 265982 |
| `construction` | Construction | 265983 |
| `environmental` | Environmental | 265987 |
| `fencing` | Fencing | 265991 |
| `pit_lining` | Pit Lining (Water Solutions) | 265992 |

### Divisions KPA-ONLY (no Motive):
| Division Key | Display Name | Reason |
|---|---|---|
| `per` | Permian Equipment Rentals | No fleet tracking needed |
| `hutchs` | Hutch's | KPA only |
| `shop` | Shop/Fabrication | KPA only |
| `downhole` | Downhole Tools | KPA only |

### Divisions to SKIP during monthly generation:
- `hutchs`, `shop`, `downhole` — skip entirely (KPA-only, minimal data)
- `casing` — often already done separately

## Text Exporter Sections (what each .txt file contains)

### Standard sections (all divisions):
1. **Division Info** — name, company, manager, safety rep, headcount, man-hours
2. **Lagging Indicators** — incidents, OSHA recordables, TRIR, DART
3. **Leading Indicators** — observations (count + rate), training compliance
4. **Observation Summary** — by type breakdown, by rig (if applicable), near misses/good catches count
5. **Observation Details** — individual observation records
6. **Incident Details** — individual incident records
7. **Assessment Details** — field assessments grouped by report, with scores and deficiency flags
8. **Training Compliance** — overall %, non-compliant employee list with incomplete programs

### Fleet sections (divisions with Motive data):
9. **Fleet Metrics** — active vehicles, utilization, total mileage, avg per vehicle
10. **Speeding Analysis** — total events, severity breakdown (critical/high/medium), rate per 10k miles, top offenders
11. **Camera/DriveCam Events** — by type, drowsiness alerts, top offenders
12. **Top Vehicles by Miles** — top 15 vehicles with driver and mileage

### Division-specific sections:
13. **JSA Log Summary** — total JSAs, by rig/district, review rate, top submitters (Transcend, BTI)
14. **Rig Inspections** — individual inspections with scores, deficiency flags, average score (Transcend)
15. **Vehicle Inspection Details** — Bernard's truck audits with driver, truck#, score, deficiency themes (BTI)
16. **Vehicle Inspections (Pre/Post Trip)** — count summary (BTI)
17. **Mileage by Yard** — yard-level fleet breakdown (multi-yard divisions like Casing, Rathole)

## Known Issues & Fixes Applied

### Division-specific forms bypass company filter
- **Problem:** `_fetch_form()` applies `_filter_rows()` which checks company/service_line fields. Division-specific forms (like Transcend's rig inspection form 385365) don't have these fields, so all rows get silently dropped.
- **Fix:** Division-specific form fetches (rig_inspections, assessments) call `kpa.get_form_responses()` directly instead of `_fetch_form()`, bypassing the company/service_line filter.
- **Affected forms:** 385365 (TD - Rig Inspection), 206674 (TD - Transcend Field Assessment), 225539 (TD - Management Audit)

### Score double-percent
- **Problem:** KPA `score-percent` field value already includes `%`, and renderers were appending another `%`.
- **Fix:** `str(score).rstrip('%')` before appending `%` in all score display lines.

### Man-hours Excel format mismatch
- **Problem:** 2026 man-hours files use a single "Detail" sheet with all companies, format `WorkMonthReport_YYYY-MM_YYYY-MM`. The parser's `parse_2026()` method handles this but needs the correct sheet name.
- **Status:** Parser auto-detects format via `has_detail_sheet()`. May show 0.0 hours if the Excel sheet doesn't have the expected "Detail" tab or if department/co-code filtering doesn't match.

### KPA header echo rows
- **Problem:** KPA API sometimes returns a header echo row (date="Date", observer="Observer") that slips through date filtering.
- **Impact:** Minor — may inflate total counts by 1. Individual record rendering skips these since they have no useful data.

## Comparison to PDF Reports

### Transcend PDF sections vs text data:
| PDF Section | Text File Section | Status |
|---|---|---|
| Executive Summary | Lagging + Leading Indicators | Done |
| Observations (by type, by rig) | Observation Summary + Details | Done |
| Near Misses & Good Catches | Observation Summary callout | Done |
| JSA Log (by rig, review rate) | JSA Log Summary | Done |
| KPA Form Activity | Not rendered | Gap - would need all-forms query |
| Rig Inspections & Field Assessments | Rig Inspections + Assessments | Done |
| Incidents | Incident Details | Done |
| Training Compliance | Training Compliance | Done |

### BTI PDF sections vs text data:
| PDF Section | Text File Section | Status |
|---|---|---|
| Executive Summary | Lagging + Leading Indicators | Done |
| Fleet Mileage (top trucks) | Top Vehicles by Miles | Done |
| Speeding Summary (by driver) | Speeding Analysis | Done |
| DOT Inspection Scorecard | N/A | Gap - external data source |
| Bernard's Vehicle Inspections | Vehicle Inspection Details | Done |
| Incidents | Incident Details | Done |
| Training Compliance | Training Compliance | Done |

## Step-by-Step: Running Monthly Recaps for a New Month

1. **Get the man-hours file** from Downloads (format: `WorkMonthHoursReport_YYYY-MM_YYYY-MM.xlsx`)
2. **Run divisions one at a time** to avoid KPA rate limiting:
   ```bash
   cd C:/Users/krhod/daily-safety-report
   # Run each division individually:
   KPA_API_TOKEN=plnuF2fKcYuS07JoQHgOweB33Y8dldzdI MOTIVE_API_KEY=8d3dd502-36c0-47c4-ade3-a1fbbef0c05c py -3 -u generate_recap.py --division poly_pipe --month 2026-06 --data-only --man-hours "C:/Users/krhod/Downloads/WorkMonthHoursReport_2026-06_2026-06.xlsx"
   ```
3. **Order of divisions to run:** anchors, poly_pipe, pit_lining, construction, environmental, fencing, rathole, bti, transcend, valor, per
4. **Skip:** casing (done separately), hutchs, shop, downhole (KPA-only, minimal)
5. **Output lands in:** `~/Desktop/Monthly HSE Recaps/{Month YYYY}/`
6. **Copy to Downloads:** `cp "C:/Users/krhod/Desktop/Monthly HSE Recaps/June 2026/"*.txt "C:/Users/krhod/Downloads/"`
7. **Verify each file** has the expected sections (fleet data for Motive divisions, observations, training, etc.)
