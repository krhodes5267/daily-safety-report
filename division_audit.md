# Division Audit Tracker

**Started:** 2026-04-20
**Last Updated:** 2026-04-20
**Total Issues Found:** 28
**Total Issues Fixed:** 2

---

## Progress

| # | Division | Company | Status | Issues | Notes |
|---|----------|---------|--------|--------|-------|
| 1 | Casing | BRHAS | COMPLETE | 2 fixed, 5 warn | 6 yards, all have data, yard enrichment fixed |
| 2 | Rathole | BRHAS | COMPLETE | 5 warn | 8 yards, 3 remote yards SILENT, 98.3% unknown drivers |
| 3 | Poly Pipe | BRHAS | COMPLETE | 3 warn | 100% unknown yard, sub-yards not in speeding data |
| 4 | Pit Lining | BRHAS | COMPLETE | 2 warn | 100% unknown yard, 2 NFMVAs |
| 5 | Anchors | BRHAS | COMPLETE | 2 warn | 100% unknown yard, 0 assessments |
| 6 | Construction | BRHAS | COMPLETE | 1 warn | 12% unknown yard (30/247), Levelland routing |
| 7 | Environmental | BRHAS | COMPLETE | 2 critical | ZERO KPA activity, worst mi/event ratio |
| 8 | Fencing | BRHAS | COMPLETE | 2 critical | ZERO KPA activity, 20.6% RED rate |
| 9 | Drilling Tools | BRHAS | COMPLETE | 1 critical | ZERO ACTIVITY across all systems |
| 10 | Shop/Fabrication | BRHAS | COMPLETE | 1 warn | Minimal, 3 speeding events, Levelland only |
| 11 | Corporate | BRHAS | COMPLETE | 1 info | ZERO ACTIVITY, expected for admin |
| 12 | Valor | Valor | COMPLETE | 2 critical | ZERO KPA activity, 251 speeding, no cross-contamination |
| 13 | BTI | BTI | COMPLETE | 3 warn | 93% known drivers (BEST), 0 obs, 0 assessments in archive |
| 14 | Transcend | Transcend | COMPLETE | 2 warn | Rig 18 nearly silent (1 obs), good obs culture otherwise |
| 15 | Permian | Permian | COMPLETE | 1 warn | 3 incidents but 0 obs/assessments |

---

## Division Summaries

### 1. Casing (BRHAS)
**Status:** COMPLETE
**Audited:** 2026-04-20
```
Yards: Bryan, Hobbs, Jourdanton, Kilgore, Laredo, Midland (all 6 have data -- PASS)
Motive Groups: 7 mapped (167175, 169090, 169092, 186740, 169091, 186739, 186741)
KPA Service Lines: Casing
Assessment Forms: 381707 (Field Assessment=93), 229645 (Pre/Post Trip=1414)
YTD Data: 274 speeding, 105 camera, 2011 obs, 15 incidents, 1507 assessments, 899K miles

Speeding by Yard: Midland 134, Laredo 50, Kilgore 36, Hobbs 25, Bryan 22, Jourdanton 7
Speeding by Tier: RED 45, ORANGE 103, YELLOW 126
Camera by Yard: Midland 60, Laredo 19, Kilgore 12, Hobbs 6, Bryan 3, San Angelo 3, Jourdanton 2
Obs by Yard: Midland 1531, Bryan 178, Hobbs 91, Unknown 74, Jourdanton 73, Laredo 41, Kilgore 21
KPIs: 3,284 mi/speeding, 8,569 mi/camera, 18.6 obs/day

Issues Found:
  1. FIXED: Speeding yard=Unknown for all Casing vehicles (numeric+C names have no yard prefix)
     -> Built vehicle-to-yard lookup from Motive /v1/vehicles API using CASING_GROUP_IDS
     -> Added build_vehicle_lookup() + enrich_speeding_yards() to backfill_archive.py
     -> Added _build_vehicle_yards() to archive_today.py for daily pipeline
     -> 346 of 4502 total speeding events enriched (all Casing)
  2. FIXED: Camera audit showed 0 events (audit script checked for CSG prefix, but Casing
     vehicles use {number}C format). All camera events in archive are already Casing-only.
  3. WARN: Hobbs obs quality -- 85.7% Recognition type (78/91), possibly padding
  4. WARN: 3 San Angelo camera events (yard should be closed)
  5. WARN: 73 speeding events with Unknown driver (26.6%)
  6. WARN: Obs declining monthly: Jan 594 -> Feb 529 -> Mar 465 -> Apr 423
  7. WARN: 74 observations with empty yard field (3.7%)
```

### 2. Rathole (BRHAS)
**Status:** COMPLETE
**Audited:** 2026-04-20
```
Yards: Midland, Levelland, Barstow, Jourdanton, Ohio, Pennsylvania, Oklahoma, North Dakota
KPA Service Lines: Rat Hole, Rathole
Assessment Forms: 153181 (RH - Rathole Field Assessment)
YTD Data: 1997 speeding, 0 camera, 20 obs, 1 incident, 30 assessments, 1.10M miles

Speeding by Yard: Levelland 873, Unknown 322, Pennsylvania 264, Jourdanton 244,
                   Midland 144, Barstow 100, North Dakota 47, Oklahoma 3
Speeding by Tier: RED 288, ORANGE 728, YELLOW 981
Speeding Monthly: Jan 781, Feb 622, Mar 329, Apr 265 (IMPROVING -66%)

Observations: 20 total (Levelland 10, Jourdanton 6, Unknown 2, Barstow 1, ND 1)
  Types: Recognition 8, At-Risk Condition 5, At-Risk Procedure 5, Suggestion 1, At-Risk Behavior 1

Assessments: 30 total (RH - Rathole Field Assessment)
  Monthly: Jan 10, Feb 10, Mar 5, Apr 5 (DECLINING -50%)

Incidents: 1 (2026-03-02: crew rigging up auger, pin injury)

Mileage: 1,097,546 miles
  Monthly: Jan 264K, Feb 308K, Mar 338K, Apr 188K
KPIs: 550 mi/speeding event (VERY LOW), 0.2 obs/day

Remote Yard Coverage:
  Midland: 144 speed, 0 obs, 0 assess -- !! NO KPA ACTIVITY
  Levelland: 873 speed, 10 obs, 0 assess -- OK (but light)
  Barstow: 100 speed, 1 obs, 0 assess -- !! BARELY ACTIVE
  Jourdanton: 244 speed, 6 obs, 0 assess -- OK (but light)
  Ohio: 0 speed, 0 obs, 0 assess -- !! ZERO ACTIVITY (may be inactive)
  Pennsylvania: 264 speed, 0 obs, 0 assess -- !! NO KPA ACTIVITY
  Oklahoma: 3 speed, 0 obs, 0 assess -- !! NO KPA ACTIVITY
  North Dakota: 47 speed, 1 obs, 0 assess -- !! BARELY ACTIVE

Issues Found:
  1. WARN: 98.3% unknown drivers (1963/1997) -- only 6 named drivers in YTD data
  2. WARN: 322 speeding with Unknown yard (16.1%) -- need vehicle lookup expansion
  3. WARN: Pennsylvania has 264 speeding events but ZERO KPA activity
  4. WARN: Midland has 144 speeding events but ZERO observations or assessments
  5. WARN: Ohio yard has ZERO activity across all systems (possibly inactive?)
```

### 3. Poly Pipe (BRHAS)
**Status:** COMPLETE
**Audited:** 2026-04-20
```
Yards: Midland (main), Bryan, Hobbs, Jourdanton (sub-yards)
KPA Service Lines: Poly Pipe
Assessment Forms: 226217 (WS - Poly Pipe Field Assessment)
YTD Data: 350 speeding, 0 camera, 37 obs, 3 incidents, 15 assessments, 443K miles

Speeding: 350 events (RED 49, ORANGE 139, YELLOW 162)
  All 350 have Unknown yard (100%)
  Monthly: Jan 184, Feb 82, Mar 58, Apr 26 (IMPROVING -86%)
  All 350 have Unknown driver (100%)

Observations: 37 total (all Midland)
  Types: At-Risk Condition 19, At-Risk Procedure 6, Recognition 6, At-Risk Behavior 4, Suggestion 2
  Good quality mix -- not padding with Recognitions

Assessments: 15 total (WS - Poly Pipe Field Assessment)

Incidents: 3
  2026-01-29: crew connecting poly line
  2026-02-05: transferring poly pipe between spools
  2026-04-14: NAFMVA - poly tech driving

Mileage: 443,227 miles
  Monthly: Jan 119K, Feb 113K, Mar 134K, Apr 77K
KPIs: 1,266 mi/speeding event

Issues Found:
  1. WARN: 100% Unknown yard on speeding -- need vehicle lookup for Poly Pipe groups
  2. WARN: 100% Unknown driver on speeding
  3. WARN: All 37 obs are Midland only -- sub-yards (Bryan/Hobbs/Jourdanton) submitting zero
```

### 4. Pit Lining (BRHAS)
**Status:** COMPLETE
**Audited:** 2026-04-20
```
Yards: Midland only
KPA Service Lines: Pit Lining
Assessment Forms: 386087 (WS - Pit Lining Field Assessment)
YTD Data: 182 speeding, 0 camera, 23 obs, 2 incidents, 8 assessments, 144K miles

Speeding: 182 events (RED 13, ORANGE 19, YELLOW 150)
  All 182 have Unknown yard (100%) -- single yard so not critical
  Monthly: Jan 91, Feb 26, Mar 53, Apr 12 (volatile but trending down)
  All 182 have Unknown driver (100%)

Observations: 23 total (all Midland)
  Types: At-Risk Procedure 11, At-Risk Condition 8, At-Risk Behavior 2, Recognition 1, Suggestion 1
  Excellent quality mix -- only 1 Recognition out of 23 (4.3%)

Assessments: 8 total (WS - Pit Lining Field Assessment)

Incidents: 2 (both NFMVAs in January)
  2026-01-06: NAFMVA - driving Unit #2337 west bound
  2026-01-11: NAFMVA - shop mechanic driving Unit #2277

Mileage: 143,794 miles
  Monthly: Jan 31K, Feb 27K, Mar 58K, Apr 28K
KPIs: 790 mi/speeding event

Issues Found:
  1. WARN: 100% Unknown driver on speeding (single yard so yard unknown is expected)
  2. WARN: 2 NFMVAs in first 2 weeks of January -- check if root cause addressed
```

### 5. Anchors (BRHAS)
**Status:** COMPLETE
**Audited:** 2026-04-20
```
Yards: Midland only (also Levelland 2 obs, Seminole 1 obs)
KPA Service Lines: Anchor, Anchors
Assessment Forms: none division-specific
YTD Data: 55 speeding, 0 camera, 7 obs, 1 incident, 0 assessments, 93K miles

Speeding: 55 events (RED 6, ORANGE 18, YELLOW 31)
  Monthly: Jan 17, Feb 20, Mar 9, Apr 9
  Top: Unknown 24, Marshall Navarrette 11, Shawn Anderson 7

Observations: 7 total (Midland 4, Levelland 2, Seminole 1)
  Types: Recognition 4, At-Risk Condition 2, At-Risk Procedure 1
  57% Recognition -- small sample but low quality indicator

Assessments: 0 (no division-specific assessment form)

Incidents: 1 (2026-02-02: employee setting anchors, multi-well pad)

Mileage: 92,656 miles
  Monthly: Jan 22K, Feb 28K, Mar 27K, Apr 16K
KPIs: 1,685 mi/speeding event

Issues Found:
  1. WARN: ZERO assessments YTD -- no dedicated assessment form for Anchors
  2. WARN: Only 7 observations in ~4 months -- extremely low engagement
```

### 6. Construction (BRHAS)
**Status:** COMPLETE
**Audited:** 2026-04-20
```
Yards: Odessa/Midland + Levelland
KPA Service Lines: Construction, Civil, Water/Construction
Assessment Forms: 172295 (Construction - Site Safety Review)
YTD Data: 247 speeding, 0 camera, 11 obs, 1 incident, 9 assessments, 205K miles

Speeding: 247 events (RED 22, ORANGE 39, YELLOW 186)
  By Yard: Levelland 217 (87.9%), Unknown 30 (12.1%)
  Monthly: Jan 48, Feb 90, Mar 78, Apr 31
  Top: Unknown 143, Michael Reyna Jr 29, Smokey Oliver 25

Observations: 11 total (all Midland)
  Types: At-Risk Condition 5, At-Risk Procedure 2, At-Risk Behavior 2, Recognition 1, Suggestion 1

Assessments: 9 (Construction - Site Safety Review)

Incidents: 1 (2026-01-28: first aid, 3rd party injury)

Mileage: 205,429 miles
  Monthly: Jan 36K, Feb 61K, Mar 70K, Apr 38K
KPIs: 832 mi/speeding event

Issues Found:
  1. WARN: 30 speeding events (12%) have Unknown yard -- 217 route to Levelland correctly
     NOTE: Construction routing to Levelland is correct (Water/Construction division maps there)
```

### 7. Environmental (BRHAS)
**Status:** COMPLETE
**Audited:** 2026-04-20
```
Yards: Midland only
KPA Service Lines: Environmental, Containment
Assessment Forms: none division-specific
YTD Data: 193 speeding, 0 camera, 0 obs, 0 incidents, 0 assessments, 59K miles

Speeding: 193 events (RED 26, ORANGE 46, YELLOW 121)
  Monthly: Jan 59, Feb 67, Mar 48, Apr 19
  All 193 Unknown driver, all Unknown yard

Mileage: 59,319 miles
  Monthly: Jan 10K, Feb 20K, Mar 18K, Apr 11K
KPIs: 307 mi/speeding event (WORST ratio company-wide)

Issues Found:
  1. CRITICAL: ZERO KPA activity (0 obs, 0 incidents, 0 assessments) -- complete safety blind spot
  2. CRITICAL: 307 miles per speeding event is worst in company -- 1 event every 307 miles
```

### 8. Fencing (BRHAS)
**Status:** COMPLETE
**Audited:** 2026-04-20
```
Yards: Midland only
KPA Service Lines: Fencing
Assessment Forms: none division-specific
YTD Data: 107 speeding, 0 camera, 0 obs, 1 incident, 0 assessments, 61K miles

Speeding: 107 events (RED 22, ORANGE 33, YELLOW 52)
  Monthly: Jan 33, Feb 24, Mar 29, Apr 21
  RED rate: 20.6% (HIGHEST in company, average is ~12%)
  All 107 Unknown driver, all Unknown yard

Incidents: 1 (2026-04-09: Steven Ortega/Germin Olivas at Todd Micro Grid)

Mileage: 60,577 miles
  Monthly: Jan 17K, Feb 15K, Mar 17K, Apr 13K
KPIs: 566 mi/speeding event

Issues Found:
  1. CRITICAL: ZERO KPA activity (0 obs, 0 assessments) despite active incident
  2. CRITICAL: 20.6% RED speeding rate -- worst in company (avg ~12%)
```

### 9. Drilling Tools (BRHAS)
**Status:** COMPLETE
**Audited:** 2026-04-20
```
Yards: Midland
KPA Service Lines: Drilling Tools
Motive Groups: 329823
YTD Data: 0 speeding, 0 camera, 0 obs, 0 incidents, 0 assessments, 0 miles

Issues Found:
  1. CRITICAL: ZERO ACTIVITY across ALL systems -- division may be inactive or
     all activity routing to Valor (Downhole Tools). Verify with operations.
```

### 10. Shop/Fabrication (BRHAS)
**Status:** COMPLETE
**Audited:** 2026-04-20
```
Yards: Levelland (LL-FAB, LL-SHOP vehicles)
KPA Service Lines: Shop, Fabrication
YTD Data: 3 speeding, 0 camera, 0 obs, 0 incidents, 0 assessments, 6.7K miles

Speeding: 3 events (RED 1, ORANGE 0, YELLOW 2) -- all Levelland
  Monthly: Jan 1, Mar 1, Apr 1

Mileage: 6,730 miles
  Monthly: Jan 1.5K, Feb 1.9K, Mar 2.2K, Apr 1.0K
KPIs: 2,243 mi/speeding event

Issues Found:
  1. WARN: No KPA activity but expected for shop/fab (primarily stationary work)
```

### 11. Corporate (BRHAS)
**Status:** COMPLETE
**Audited:** 2026-04-20
```
Yards: Midland
KPA Service Lines: Corporate
Motive Groups: 265988
YTD Data: 0 speeding, 0 camera, 0 obs, 0 incidents, 0 assessments, 0 miles

Issues Found:
  1. INFO: Zero activity -- expected for admin/sales. Low priority.
```

### 12. Valor Energy Services
**Status:** COMPLETE
**Audited:** 2026-04-20
```
Yards: Levelland
KPA Service Lines: Valor, Valor Energy, Valor Energy Services, Downhole Tools
Motive Groups: 167178, 265985
YTD Data: 251 speeding, 0 camera, 0 obs, 0 incidents, 0 assessments, 204K miles

Speeding: 251 events (RED 31, ORANGE 66, YELLOW 154)
  Monthly: Jan 105, Feb 70, Mar 51, Apr 25 (IMPROVING -76%)
  All 251 Unknown driver, all Unknown yard

Cross-Contamination Check: PASS
  No Downhole Tools observations in either Valor or Drilling Tools
  Valor and BRHAS speeding are properly separated by Motive group

Mileage: 204,089 miles
  Monthly: Jan 50K, Feb 52K, Mar 64K, Apr 37K
KPIs: 813 mi/speeding event

Issues Found:
  1. CRITICAL: ZERO KPA activity for an entire company -- no obs, no assessments, no incidents
  2. CRITICAL: 251 speeding events with no safety engagement at all
```

### 13. BTI / Butch's Trucking
**Status:** COMPLETE
**Audited:** 2026-04-20
```
Yards: Midland
KPA Service Lines: BTI, Butch's Trucking, Trucking
Motive Groups: 186743, 265989
YTD Data: 335 speeding, 0 camera, 0 obs, 1 incident, 0 assessments, 1.45M miles

Speeding: 335 events (RED 34, ORANGE 90, YELLOW 211)
  Monthly: Jan 162, Feb 74, Mar 67, Apr 32 (IMPROVING -80%)
  93% KNOWN drivers (BEST in company) -- only 23 unknown
  Top: Tre Scott 29, Brian Albert 18, Carlos Munoz 13, James Guzman 13, Adolfo Alvarez 13

Observations: 0 in archive (KPA service line filter may need verification)
Assessments: 0 in archive (form 152018 Vehicle Inspection Checklists route as "Shared")

Incidents: 1 (2026-03-17: employee tightening binder, gears broke)

Mileage: 1,450,464 miles (HIGHEST division)
  Monthly: Jan 339K, Feb 419K, Mar 430K, Apr 262K
  Average daily: 13,430 miles

Trucking KPIs:
  Miles per speeding event: 4,330
  Speeding per 100K miles: 23.1
  RED events per 100K miles: 2.3
  Unique speeding vehicles: 61

Data Quality:
  0 duplicate speeding events
  All 335 speeding events have Unknown yard (100%) -- single yard, not critical

Issues Found:
  1. WARN: 0 observations in archive -- BTI KPA service lines may not be captured, OR BTI
     doesn't submit observations. Verify against KPA directly.
  2. WARN: 0 assessments in archive -- form 152018 routes as "Shared" division, not "BTI".
     Bernard's vehicle inspections (~47 YTD per initial estimate) are not being captured.
  3. WARN: Tre Scott has 29 speeding events (8.7% of all BTI speeding) -- #1 repeat offender
```

### 14. Transcend Drilling
**Status:** COMPLETE
**Audited:** 2026-04-20
```
Yards: Midland
KPA Service Lines: Drilling, Transcend, Transcend Drilling, Drilling - Rig 4/16/18/20
Motive Groups: 247035, 265986
Assessment Forms: 385365 (TD - Rig Inspection), 484193 (TD - Observation Card)
YTD Data: 173 speeding, 0 camera, 151 obs, 1 incident, 5 assessments, 107K miles

Speeding: 173 events (RED 9, ORANGE 60, YELLOW 104)
  Monthly: Jan 78, Feb 38, Mar 35, Apr 22 (IMPROVING -72%)
  RED rate: 5.2% (LOWEST in company -- best)

Observations: 151 total (GOOD engagement for rig operation)
  By Rig: General (no rig) 60, Rig 4: 35, Rig 20: 29, Rig 16: 26, Rig 18: 1
  By Type: At-Risk Condition 60, Recognition 54, At-Risk Behavior 27, Suggestion 8, At-Risk Procedure 2
  Monthly: Jan 21, Feb 51, Mar 46, Apr 33 (IMPROVING after Jan ramp)

Rig Activity Check:
  Rig 4: 35 obs -- OK
  Rig 16: 26 obs -- OK
  Rig 18: 1 obs -- !! NEARLY SILENT
  Rig 20: 29 obs -- OK

Assessments: 5 total (TD - Rig Inspection)

Incidents: 1 (2026-03-23: doghouse door injury, Rig 18 yard)

Mileage: 107,356 miles
  Monthly: Jan 30K, Feb 30K, Mar 30K, Apr 17K (very consistent)
KPIs: 621 mi/speeding event, 1.4 obs/day

Issues Found:
  1. WARN: Rig 18 has only 1 observation YTD -- nearly silent. All other rigs have 26-35.
     Incident happened at Rig 18 (door injury) -- safety visibility gap at the rig with an incident.
  2. WARN: 60 observations (39.7%) have no rig specified (General/Drilling/Transcend) --
     could be office or yard obs, but reduces rig-level visibility
```

### 15. Permian Equipment Rentals
**Status:** COMPLETE
**Audited:** 2026-04-20
```
Yards: Midland
KPA Service Lines: Rentals, PER, Permian
Motive Groups: 265984
YTD Data: 5 speeding, 0 camera, 0 obs, 3 incidents, 0 assessments, 1.9K miles

Speeding: 5 events (RED 2, ORANGE 1, YELLOW 2)
  Monthly: Feb 3, Apr 2 -- very low volume

Incidents: 3 (HIGHEST per-capita)
  2026-02-07: employee operating unit 222, struck by panel
  2026-03-16: employee placing panel on generator
  2026-03-21: property damage, field hand loading trailer

Mileage: 1,861 miles (very low, equipment rental focus)

Issues Found:
  1. WARN: 3 incidents with 0 observations and 0 assessments -- highest incident-to-obs ratio
     in the company. Safety engagement is zero despite active incidents.
```

---

## Known Systemic Issues (fix before division audits)

### MUST FIX (blocks division audits)
- [x] **Mileage data not in archive** -- FIXED 2026-04-20. Added to backfill_archive.py (bulk IFTA fetch, 48,482 trips), archive_today.py (daily fetch), dashboard_api.py (fetch_mileage()), dashboard.html (mergeArchiveDays). YTD: 5.2M miles, 350 vehicles. Schema: `"mileage": {"total_miles": N, "by_division": {...}, "vehicle_count": N}`.
- [x] **Assessment form_name is empty in archive** -- FIXED 2026-04-20. Added FORM_NAME_MAP to backfill_archive.py mapping all 10 form IDs to human-readable names. Re-backfilled 110 archive files.

### SHOULD FIX
- [x] **182 observations YTD have empty service_line** -- FIXED 2026-04-20. 30 are Transcend (Rig N pattern in location, auto-inferred as "Drilling"). 149 are well-pad locations (likely Rathole, but can't confirm without KPA source fix). 3 have blank location. Added location-based Transcend inference to backfill_archive.py. Also expanded KPA_SVC_TO_DIVISION in api_config.py to cover all companies/divisions.
- [x] **Vehicle Inspection Checklists (form 152018) are "Shared"** -- FIXED 2026-04-20. Changed assessment routing: "Shared" forms now prefer service_line-based routing over hardcoded "Shared" division. Added BTI/Transcend/Valor/Permian to KPA_SVC_TO_DIVISION map so service_line values like "BTI", "Trucking" properly route.
- [x] **Camera events only cover ~30 days** -- DOCUMENTED 2026-04-20. Motive v2 camera API only retains ~30 days of events. This is a Motive platform limitation, not fixable via API. Historical camera data beyond 30 days is permanently unavailable. Dashboard archive files only contain camera data from the date they were archived forward. Backfills cannot recover historical camera events.
- [x] **83.6% of speeding events have Unknown driver** -- FIXED 2026-04-20. Added driver enrichment to build_vehicle_lookup() in backfill_archive.py -- indexes vehicle_drivers by both full and short vehicle number. Driver data comes from Motive /v1/vehicles current_driver/permanent_driver fields. Note: many vehicles still have no assigned driver in Motive (configuration issue at fleet level, not an API limitation).
- [x] **Non-Casing divisions have 100% Unknown yard on speeding** -- FIXED 2026-04-20. Expanded build_vehicle_lookup() in backfill_archive.py and _build_vehicle_yards() in archive_today.py from CASING_GROUP_IDS to GROUP_ID_MAP. Now maps all Motive group IDs (Rathole, Poly Pipe, Anchors, BTI, Transcend, Valor, etc.) to yard names.

### NICE TO HAVE
- [ ] **BTI SAFER score** -- look up at https://safer.fmcsa.dot.gov for Butch's Trucking Inc. Manual lookup, add to BTI profile.
- [ ] **Transcend rig-by-rig obs routing** -- service_lines like `Drilling - Rig 4` are all mapping to `Transcend` company but could be split into per-rig views in the dashboard.

---

## Company-Wide Red Flags (from deep dive)

### CRITICAL -- Safety Blind Spots
1. **Environmental** -- 193 speeding events, 307 mi/event (worst), ZERO KPA activity
2. **Fencing** -- 107 speeding events, 20.6% RED rate (worst), ZERO KPA activity, 1 incident
3. **Valor** -- 251 speeding events, entire separate COMPANY with ZERO safety engagement
4. **Permian** -- 3 incidents with ZERO observations or assessments

### HIGH -- Data Gaps
5. **Rathole Pennsylvania** -- 264 speeding events with ZERO KPA activity (remote yard)
6. **Rathole Midland** -- 144 speeding events with ZERO observations (home base!)
7. **BTI assessments** -- ~47 vehicle inspections routing as "Shared", not captured in BTI data
8. **Transcend Rig 18** -- 1 observation YTD + 1 incident (safety gap at incident rig)

### MEDIUM -- Trends
9. **Casing speeding worsening** -- +103% Jan-to-Mar (66 -> 134)
10. **Casing obs declining** -- -29% Jan-to-Apr (594 -> 423)
11. **Rathole assessment volume dropping** -- -50% (10/mo -> 5/mo)
12. **Company-wide driver ID crisis** -- 83.6% unknown, prevents accountability

---

## Audit Order
1. ~~Fix systemic issues~~ (mileage, form_name) -- DONE
2. ~~Casing~~ -- COMPLETE
3. ~~Rathole~~ -- COMPLETE
4. ~~BTI~~ -- COMPLETE
5. ~~Transcend~~ -- COMPLETE
6. ~~Poly Pipe~~ -- COMPLETE
7. ~~Valor~~ -- COMPLETE
8. ~~Remaining BRHAS divisions~~ -- ALL COMPLETE
9. ~~Permian~~ -- COMPLETE

**ALL 15 DIVISIONS AUDITED**

## How to Resume
Open Claude Code in `C:\Users\krhod\daily-safety-report` and say:
```
Continue the division audit. Check division_audit.md for current progress and CLAUDE.md for division profiles. Start with the next pending item.
```

## Session Log
- **2026-04-20**: Created CLAUDE.md with per-division safety profiles. Created division_audit.md tracker. Identified 5 systemic issues (2 must-fix: mileage data + assessment form names). Mileage pattern found in casing_monthly_recap.py lines 436-530 using Motive IFTA endpoint. All archive data re-backfilled with KPA yard field (OBS_YARD_HASH). Filter maps fixed for Water/Construction, Fabrication, Unknown divisions. Filter audit passing 25/25.
- **2026-04-20 (session 2)**: Fixed systemic issues (mileage + form_name). Started Casing audit. Discovered speeding yard=Unknown for all Casing vehicles (numeric+C names). Built vehicle-to-yard lookup from Motive /v1/vehicles API (build_vehicle_lookup + enrich_speeding_yards in backfill_archive.py, _build_vehicle_yards in archive_today.py). Also normalized vehicle number matching (short-form + full-form keys). Re-backfilled 110 archive files. Casing audit COMPLETE: all 6 yards have data, 4 issues found (2 fixed, 5 warnings documented).
- **2026-04-20 (session 3)**: Completed ALL 15 division audits. Ran dedicated scripts for Rathole, BTI, Transcend + bulk audit for remaining 11 divisions. Key findings: 4 CRITICAL safety blind spots (Environmental, Fencing, Valor, Permian have zero KPA engagement despite active speeding/incidents), 83.6% company-wide unknown drivers, Rathole has 3 remote yards with zero KPA activity (PA, OK, Midland), BTI assessments routing as "Shared" not captured. 28 total issues found, 2 fixed (Casing yard enrichment). Deep dive completed with vehicle-level analysis, monthly trends, cross-division comparison.
