# Division Audit Tracker

**Started:** 2026-04-20
**Last Updated:** 2026-04-20
**Total Issues Found:** 0
**Total Issues Fixed:** 0

---

## Progress

| # | Division | Company | Status | Issues | Notes |
|---|----------|---------|--------|--------|-------|
| 1 | Casing | BRHAS | COMPLETE | 4 issues | 6 yards, all have data, yard enrichment fixed |
| 2 | Rathole | BRHAS | PENDING | - | 8 yards across states, heavy speeding |
| 3 | Poly Pipe | BRHAS | PENDING | - | New sub-yards (Bryan/Hobbs/Jourdanton) |
| 4 | Pit Lining | BRHAS | PENDING | - | Single yard, small crew |
| 5 | Anchors | BRHAS | PENDING | - | Single yard, very light KPA |
| 6 | Construction | BRHAS | PENDING | - | Odessa, check Civil/WTC routing |
| 7 | Environmental | BRHAS | PENDING | - | Check Containment routing |
| 8 | Fencing | BRHAS | PENDING | - | Single yard, basic check |
| 9 | Drilling Tools | BRHAS | PENDING | - | Verify vs Valor Downhole Tools |
| 10 | Shop/Fabrication | BRHAS | PENDING | - | Midland + Levelland fab |
| 11 | Corporate | BRHAS | PENDING | - | Admin/sales vehicles |
| 12 | Valor | Valor | PENDING | - | Levelland, check cross-contamination |
| 13 | BTI | BTI | PENDING | - | Trucking focus: miles, SAFER, KPIs |
| 14 | Transcend | Transcend | PENDING | - | Spudder rigs, rig-by-rig analysis |
| 15 | Permian | Permian | PENDING | - | Small, basic check |

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
**Status:** PENDING
**Audited:** -
```
Yards: Midland, Levelland, Barstow, Jourdanton, Ohio, Pennsylvania, Oklahoma, North Dakota
Motive Groups: 16+ mapped (167176, 220453, 274965, 186742, 307752, 341789, 308775, 351218, 265996-266028, 290471)
KPA Service Lines: Rat Hole, Rathole
Assessment Forms: 153181 (Rathole Field Assessment)
YTD Data: ~1997 speeding, ~0 camera, ~20 obs, ~0 incidents, ~30 assessments
Issues Found: -
Issues Fixed: -
Status: -
```

### 3. Poly Pipe (BRHAS)
**Status:** PENDING
**Audited:** -
```
Yards: Midland (main), Bryan, Hobbs, Jourdanton (sub-yards)
Motive Groups: 5 mapped (167180, 296017, 296020, 296036, 296040)
KPA Service Lines: Poly Pipe
Assessment Forms: 226217 (Poly Pipe Field Assessment)
YTD Data: ~350 speeding, ~0 camera, ~37 obs, ~1 incident, ~15 assessments
Issues Found: -
Issues Fixed: -
Status: -
```

### 4. Pit Lining (BRHAS)
**Status:** PENDING
**Audited:** -
```
Yards: Midland
Motive Groups: 2 mapped (167179, 265992)
KPA Service Lines: Pit Lining
Assessment Forms: 386087 (Pit Lining Field Assessment)
YTD Data: ~182 speeding, ~0 camera, ~23 obs, ~0 incidents, ~8 assessments
Issues Found: -
Issues Fixed: -
Status: -
```

### 5. Anchors (BRHAS)
**Status:** PENDING
**Audited:** -
```
Yards: Midland
Motive Groups: 2 mapped (167177, 265982)
KPA Service Lines: Anchor, Anchors
Assessment Forms: none division-specific
YTD Data: ~55 speeding, ~0 camera, ~7 obs, ~0 incidents, ~0 assessments
Issues Found: -
Issues Fixed: -
Status: -
```

### 6. Construction (BRHAS)
**Status:** PENDING
**Audited:** -
```
Yards: Odessa (mapped as Midland in some Motive groups)
Motive Groups: 2 mapped (247032, 265983)
KPA Service Lines: Construction, Civil, Water/Construction
Assessment Forms: 172295 (Construction - Site Safety Review)
YTD Data: ~247 speeding (incl Water/Construction), ~0 camera, ~11 obs, ~0 incidents, ~9 assessments
Issues Found: -
Issues Fixed: -
Status: -
```

### 7. Environmental (BRHAS)
**Status:** PENDING
**Audited:** -
```
Yards: Midland
Motive Groups: 2 mapped (186738, 265987)
KPA Service Lines: Environmental, Containment
Assessment Forms: none division-specific
YTD Data: ~193 speeding, ~0 camera, ~0 obs, ~0 incidents, ~0 assessments
Issues Found: -
Issues Fixed: -
Status: -
```

### 8. Fencing (BRHAS)
**Status:** PENDING
**Audited:** -
```
Yards: Midland
Motive Groups: 2 mapped (220456, 265991)
KPA Service Lines: Fencing
Assessment Forms: none division-specific
YTD Data: ~107 speeding, ~0 camera, ~0 obs, ~1 incident, ~0 assessments
Issues Found: -
Issues Fixed: -
Status: -
```

### 9. Drilling Tools (BRHAS)
**Status:** PENDING
**Audited:** -
```
Yards: Midland
Motive Groups: 1 mapped (329823)
KPA Service Lines: Drilling Tools (note: Downhole Tools -> Valor)
Assessment Forms: none division-specific
YTD Data: ~0 speeding, ~0 camera, ~0 obs, ~0 incidents, ~0 assessments
Issues Found: -
Issues Fixed: -
Status: -
```

### 10. Shop/Fabrication (BRHAS)
**Status:** PENDING
**Audited:** -
```
Yards: Midland, Levelland (LL-FAB, LL-SHOP vehicles)
Motive Groups: 0 dedicated (vehicles show as Unknown with LL-FAB/LL-SHOP prefix)
KPA Service Lines: Shop, Fabrication
Assessment Forms: none division-specific
YTD Data: ~3 speeding, ~0 camera, ~0 obs, ~0 incidents, ~0 assessments
Issues Found: -
Issues Fixed: -
Status: -
```

### 11. Corporate (BRHAS)
**Status:** PENDING
**Audited:** -
```
Yards: Midland
Motive Groups: 1 mapped (265988)
KPA Service Lines: Corporate
Assessment Forms: none
YTD Data: Sales/admin vehicles only
Issues Found: -
Issues Fixed: -
Status: -
```

### 12. Valor Energy Services
**Status:** PENDING
**Audited:** -
```
Yards: Levelland
Motive Groups: 2 mapped (167178, 265985)
KPA Service Lines: Valor, Valor Energy, Valor Energy Services, Downhole Tools
Assessment Forms: none division-specific
YTD Data: ~251 speeding, ~0 camera, ~0 obs, ~0 incidents, ~0 assessments
Issues Found: -
Issues Fixed: -
Status: -
```

### 13. BTI / Butch's Trucking
**Status:** PENDING
**Audited:** -
```
Yards: Midland
Motive Groups: 2 mapped (186743, 265989)
KPA Service Lines: BTI, Butch's Trucking, Trucking
Assessment Forms: 152018 (Vehicle Inspection Checklist — Bernard's truck audits)
YTD Data: ~335 speeding, ~0 camera, ~0 obs, ~0 incidents, ~47 vehicle inspections
SAFER Score: TBD (look up at safer.fmcsa.dot.gov)
Issues Found: -
Issues Fixed: -
Status: -
```

### 14. Transcend Drilling
**Status:** PENDING
**Audited:** -
```
Yards: Midland
Motive Groups: 2 mapped (247035, 265986)
KPA Service Lines: Drilling, Transcend, Transcend Drilling, Drilling - Rig 4/16/18/20
Assessment Forms: 385365 (TD - Rig Inspection), 484193 (TD - Observation Card)
YTD Data: ~173 speeding, ~0 camera, ~151 obs (by rig), ~1 incident, ~5 rig inspections
Issues Found: -
Issues Fixed: -
Status: -
```

### 15. Permian Equipment Rentals
**Status:** PENDING
**Audited:** -
```
Yards: Midland
Motive Groups: 1 mapped (265984)
KPA Service Lines: Rentals, PER, Permian
Assessment Forms: none
YTD Data: ~5 speeding, ~0 camera, ~0 obs, ~1 incident, ~0 assessments
Issues Found: -
Issues Fixed: -
Status: -
```

---

## Known Systemic Issues (fix before division audits)

### MUST FIX (blocks division audits)
- [x] **Mileage data not in archive** — FIXED 2026-04-20. Added to backfill_archive.py (bulk IFTA fetch, 48,482 trips), archive_today.py (daily fetch), dashboard_api.py (fetch_mileage()), dashboard.html (mergeArchiveDays). YTD: 5.2M miles, 350 vehicles. Schema: `"mileage": {"total_miles": N, "by_division": {...}, "vehicle_count": N}`.
- [x] **Assessment form_name is empty in archive** — FIXED 2026-04-20. Added FORM_NAME_MAP to backfill_archive.py mapping all 10 form IDs to human-readable names. Re-backfilled 110 archive files.

### SHOULD FIX
- [ ] **182 observations YTD have empty service_line** — can't route to any division. Investigate: are these from a form that doesn't have the SL field? Could some be inferred from observer name or location?
- [ ] **Vehicle Inspection Checklists (form 152018) are "Shared"** — need to determine which are BTI (Bernard's truck audits) vs other divisions. May need to check observer or vehicle field.
- [ ] **Camera events only cover ~30 days** — Motive v2 API retention limit. Not fixable via API. Document as known limitation.

### NICE TO HAVE
- [ ] **BTI SAFER score** — look up at https://safer.fmcsa.dot.gov for Butch's Trucking Inc. Manual lookup, add to BTI profile.
- [ ] **Transcend rig-by-rig obs routing** — service_lines like `Drilling - Rig 4` are all mapping to `Transcend` company but could be split into per-rig views in the dashboard.

## Audit Order
1. **Fix systemic issues** (mileage, form_name) — these affect ALL divisions
2. **Casing** — largest, most complex, most data
3. **Rathole** — second largest, multi-state, heavy speeding
4. **BTI** — trucking company, completely different KPIs (mileage-centric)
5. **Transcend** — spudder rigs, rig-by-rig analysis
6. **Poly Pipe** — new sub-yards to verify
7. **Valor** — separate company, cross-contamination check
8. **Remaining BRHAS divisions** (Pit Lining, Anchors, Construction, Environmental, Fencing, Drilling Tools, Shop, Corporate)
9. **Permian** — smallest, last

## How to Resume
Open Claude Code in `C:\Users\krhod\daily-safety-report` and say:
```
Continue the division audit. Check division_audit.md for current progress and CLAUDE.md for division profiles. Start with the next pending item.
```

## Session Log
- **2026-04-20**: Created CLAUDE.md with per-division safety profiles. Created division_audit.md tracker. Identified 5 systemic issues (2 must-fix: mileage data + assessment form names). Mileage pattern found in casing_monthly_recap.py lines 436-530 using Motive IFTA endpoint. All archive data re-backfilled with KPA yard field (OBS_YARD_HASH). Filter maps fixed for Water/Construction, Fabrication, Unknown divisions. Filter audit passing 25/25.
- **2026-04-20 (session 2)**: Fixed systemic issues (mileage + form_name). Started Casing audit. Discovered speeding yard=Unknown for all Casing vehicles (numeric+C names). Built vehicle-to-yard lookup from Motive /v1/vehicles API (build_vehicle_lookup + enrich_speeding_yards in backfill_archive.py, _build_vehicle_yards in archive_today.py). Also normalized vehicle number matching (short-form + full-form keys). Re-backfilled 110 archive files. Casing audit COMPLETE: all 6 yards have data, 4 issues found (2 fixed, 5 warnings documented).
