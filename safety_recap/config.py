"""
Universal Safety Recap -- Division Configurations & Constants.

All 15 divisions, shared form IDs, Motive group mappings, BRHAS branding,
and section toggles live here. One source of truth for the entire system.
"""

from docx.shared import RGBColor

# ---------------------------------------------------------------------------
# BRHAS Brand Colors
# ---------------------------------------------------------------------------
BRAND_RED = RGBColor(0xC0, 0x00, 0x00)       # #C00000 -- primary Butch's red
DARK_BLUE = RGBColor(0x1F, 0x38, 0x64)       # #1F3864
MEDIUM_BLUE = RGBColor(0x2E, 0x75, 0xB6)     # #2E75B6
DARK_RED = RGBColor(0x8B, 0x00, 0x00)        # #8B0000
GREEN = RGBColor(0x00, 0x80, 0x00)           # #008000
BLACK = RGBColor(0x00, 0x00, 0x00)
GRAY = RGBColor(0x4D, 0x4D, 0x4D)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
LIGHT_GRAY_HEX = "F2F2F2"
CALIBRI = "Calibri"

# ---------------------------------------------------------------------------
# NAICS 213112 Benchmarks (BLS - Support Activities for Oil and Gas)
# ---------------------------------------------------------------------------
NAICS_TRIR_BENCHMARK = 2.09   # Industry average TRIR
NAICS_DART_BENCHMARK = 1.10   # Industry average DART

# ---------------------------------------------------------------------------
# KPA API
# ---------------------------------------------------------------------------
KPA_BASE_URL = "https://api.novaraflex.com/v1"
KPA_CALL_DELAY = 1.5  # seconds between paginated calls

# KPA endpoints used
KPA_ENDPOINTS = {
    "responses": "responses.flat",
    "followups": "followups.list",
    "users": "users.list",
    "completed_trainings": "completedtrainings.v2.list",
    "trainings": "trainings.v2.list",
    "training_status": "training-employee-status.list",
    "establishments": "establishments.list",
    "field_offices": "fieldoffices.list",
    "lines_of_business": "linesofbusiness.list",
}

# ---------------------------------------------------------------------------
# KPA Line of Business ID -> division mapping (for user/training filtering)
# From linesofbusiness.list API endpoint
# ---------------------------------------------------------------------------
LOB_ID_MAP = {
    "6009f696823a6201bbc9b051": "anchors",       # Anchor
    "6009f696823a6201bbc9b056": "casing",         # Casing
    "6009f696823a6201bbc9b04a": "rathole",        # Rat Hole
    "6009f696823a6201bbc9b053": "pit_lining",     # Pit Lining
    "6009f696823a6201bbc9b04d": "poly_pipe",      # Poly Pipe
    "6009f696823a6201bbc9b05b": "bti",            # Trucking
    "6009f696823a6201bbc9b058": "transcend",      # Drilling (general)
    "653a5c9abadfcb0013cf96d9": "transcend",      # TD RIG 16
    "642d7f30ab6b2d516d3c69d8": "transcend",      # TD RIG 18
    "65c0ed11323ea80063d7f38c": "transcend",      # TD RIG 2
    "65c0ed07f2772b002b4bc0a0": "transcend",      # TD RIG 4
    "65d62c63e73e7d0059a1bb1a": "transcend",      # TD Rig 12
    "66426f23067e2c00252514fc": "transcend",      # TD Rig 20
    "645a3565888c9541c069075a": "transcend",      # TD RIG 32
    "6009f696823a6201bbc9b04b": "downhole",       # DownHole Tools
    "6009f696823a6201bbc9b052": "downhole",       # Drilling Tools
    "609452db13b68a005495157f": "fencing",         # Fencing
    "6009f696823a6201bbc9b055": "shop",            # Shop
    "65c0ebf92eb8010021bcc281": "hutchs",          # Hutch's Oilfield Supply
    "5d166c18efd5700017316462": "per",             # Rentals (PER)
    "5d24f5d8adf0d700172f96aa": "corporate",       # Corporate
    "6148833d9e0bc6004e070765": "construction",    # Civil
    "6152f95cb365e00022f9ecb5": "environmental",   # Containment
    "6009f696823a6201bbc9b05f": "rathole",         # Fabrication
    "65c0eca2cd54c200664ff534": "rathole",         # Water Trucking
}

# Reverse: division_key -> set of LOB IDs
DIVISION_LOB_IDS = {}
for _lob_id, _div_key in LOB_ID_MAP.items():
    DIVISION_LOB_IDS.setdefault(_div_key, set()).add(_lob_id)

# ---------------------------------------------------------------------------
# Shared HSE Form IDs (used by ALL divisions)
# ---------------------------------------------------------------------------
SHARED_FORMS = {
    "observation":          151085,   # HSE - Observation Card
    "incident":             151622,   # HSE - Incident Reporting
    "root_cause_analysis":  180243,   # HSE - Root Cause Analysis Report
    "five_why":             180035,   # HSE - 5 Why
    "jsa_log":              170742,   # HSE - JSA Log
    "jsa_review":           367333,   # HSE - JSA Review Log
    "vehicle_inspection":   152018,   # HSE - Vehicle Inspection Checklist
    "non_dot_pretrip":      151213,   # HSE - Non-DOT Pre-Trip Inspection
    "workplace_inspection": 152034,   # HSE - Workplace Inspection Checklist
    "safety_meeting":       180746,   # HSE - Safety Meeting Sign In Log
    "tailgate_talk":        161692,   # HSE - Tailgate Talk
    "hot_work_permit":      267566,   # HSE - Hot Work Permit
    "fall_protection":      180935,   # HSE - Fall Protection Inspection
}

# ---------------------------------------------------------------------------
# Motive API
# ---------------------------------------------------------------------------
MOTIVE_BASE_V1 = "https://api.gomotive.com/v1"
MOTIVE_BASE_V2 = "https://api.gomotive.com/v2"

# Master Motive group ID -> (division_key, yard) mapping
MOTIVE_GROUP_MAP = {
    # Casing yards
    167175: ("casing", "Midland"),
    169090: ("casing", "Bryan"),
    169092: ("casing", "Kilgore"),
    186740: ("casing", "Hobbs"),
    169091: ("casing", "Jourdanton"),
    186739: ("casing", "Laredo"),
    186741: ("casing", "San Angelo"),
    186746: ("casing", ""),           # Parent group

    # Rathole yards
    266026: ("rathole", "Midland"),
    266025: ("rathole", "Levelland"),
    266024: ("rathole", "Barstow"),
    265996: ("rathole", "Jourdanton"),
    290472: ("rathole", "Jourdanton"),
    265998: ("rathole", "Oklahoma"),
    266028: ("rathole", "Ohio"),
    266027: ("rathole", "Pennsylvania"),
    265997: ("rathole", "North Dakota"),
    265988: ("rathole", ""),          # Parent group

    # Single-group divisions
    265989: ("bti", ""),
    265986: ("transcend", ""),
    265985: ("valor", ""),

    # Poly Pipe
    265993: ("poly_pipe", ""),
    296040: ("poly_pipe", ""),   # Poly Crew
    296036: ("poly_pipe", ""),   # Poly OM
    296017: ("poly_pipe", ""),   # Pumps & Gens
    296020: ("poly_pipe", ""),   # Supervisors

    # Anchors
    265982: ("anchors", ""),

    # Construction
    265983: ("construction", ""),

    # Environmental
    265987: ("environmental", ""),

    # Fencing
    265991: ("fencing", ""),

    # Pit Lining
    265992: ("pit_lining", ""),
}

# ---------------------------------------------------------------------------
# Logo file mapping (in daily-safety-report/logos/)
# ---------------------------------------------------------------------------
LOGO_MAP = {
    "brhas":        "Butchs.jpg",
    "casing":       "Butchs.jpg",
    "rathole":      "Butchs.jpg",
    "anchors":      "Butchs.jpg",
    "poly_pipe":    "Butchs.jpg",
    "pit_lining":   "Butchs.jpg",
    "construction": "Butchs.jpg",
    "environmental":"Butchs.jpg",
    "fencing":      "Butchs.jpg",
    "downhole":     "Butchs.jpg",
    "shop":         "Butchs.jpg",
    "bti":          "ButchTrucking.jpg",
    "transcend":    "Transcend.jpg",
    "valor":        "Valor.jpg",
    "per":          "Permian.jpg",
    "hutchs":       "Hutchs.png",
}

# ---------------------------------------------------------------------------
# Yard Info (multi-yard divisions)
# ---------------------------------------------------------------------------
CASING_YARDS = {
    "order": ["Midland", "Bryan", "Kilgore", "Hobbs", "Jourdanton", "Laredo", "San Angelo"],
    "info": {
        "Midland":    {"safety_reps": "Michael Hancock & Michael Salazar", "manager": "Richie Bentley"},
        "Bryan":      {"safety_reps": "Justin Conrad",                     "manager": "Danny Lohse"},
        "Kilgore":    {"safety_reps": "James Barnett (J.P.)",              "manager": "Frankie Balderas"},
        "Hobbs":      {"safety_reps": "Allen Batts",                       "manager": "Clifton Eaves"},
        "Jourdanton": {"safety_reps": "Joey Speyrer",                      "manager": "Enrique Flores"},
        "Laredo":     {"safety_reps": "Joey Speyrer",                      "manager": "Chris Jacobo"},
        "San Angelo": {"safety_reps": "Michael Hancock & Michael Salazar", "manager": "Jeremy Jones"},
    },
}

RATHOLE_YARDS = {
    "order": ["Midland", "Levelland", "Barstow", "Jourdanton", "Oklahoma", "North Dakota", "Ohio", "Pennsylvania"],
    "info": {
        "Midland":       {"safety_reps": "John Snodgrass", "manager": ""},
        "Levelland":     {"safety_reps": "John Snodgrass", "manager": ""},
        "Barstow":       {"safety_reps": "John Snodgrass", "manager": ""},
        "Jourdanton":    {"safety_reps": "Joey Speyrer",   "manager": ""},
        "Oklahoma":      {"safety_reps": "",                "manager": ""},
        "North Dakota":  {"safety_reps": "",                "manager": ""},
        "Ohio":          {"safety_reps": "",                "manager": ""},
        "Pennsylvania":  {"safety_reps": "",                "manager": ""},
    },
}

# ---------------------------------------------------------------------------
# Transcend Rig Normalization
# ---------------------------------------------------------------------------
TRANSCEND_RIG_CANONICAL = {
    "drilling - rig 20": "Rig 20",
    "drilling - rig 18": "Rig 18",
    "drilling - rig 16": "Rig 16",
    "drilling - rig 4":  "Rig 4",
    "drilling - rig 2":  "Rig 2",
    "drilling":          "Field Ops",
    "rig 20": "Rig 20",
    "rig 18": "Rig 18",
    "rig 16": "Rig 16",
    "rig 4":  "Rig 4",
    "rig 2":  "Rig 2",
}

TRANSCEND_DISTRICT_TO_RIG = {
    "transcend rig 20":    "Rig 20",
    "transcend rig 18":    "Rig 18",
    "transcend rig 4":     "Rig 4",
    "transcend field ops": "Field Ops",
    "carlsbad":            "Field Ops",
    "midland":             "Field Ops",
    "jal":                 "Field Ops",
}

# ---------------------------------------------------------------------------
# KPA Form Field Hash -> Semantic Name Mapping
# These vary by form. Used by section renderers for location/yard extraction.
# ---------------------------------------------------------------------------
FORM_FIELD_MAP = {
    # Observation Card (151085)
    "observation": {
        "company":       "t5187momol3em85v",
        "service_line":  "64c7upqkyt79zhh1",
        "district":      "7vj2l992y7fwqhwz",  # Yard / district
        "location":      "lg5pnj4chjadnv46",  # Well site / job location
        "type":          "bff8m4x6xbc033kg",
        "name":          "0kc57oj2zkse21o3",
        "description":   "uncbcge9x8vow9pn",
        "action":        "dpy2klalngsr7ek9",
        "customer":      "vxew6ukynemxwvjr",
    },
    # Incident Reporting (151622)
    "incident": {
        "company":       "lsx3msa0w9n9edb4",
        "service_line":  "sha7vur5q2l6d6gq",
        "district":      "pk6qj0kiu9vek20v",  # Yard / district
        "location":      "9ohdd2lwvl7p0oc6",  # Incident location description
        "type":          "nojcquy0tfl9hqih",
        "employee":      "55gg4nkoemnnfo2a",
        "description":   "313e9txgrof0uute",
        "supervisor":    "w997tirq97oenvuz",
    },
    # JSA Log (170742)
    "jsa": {
        "company":       "6zyx6l5f244mk0v5",
        "service_line":  "77ykc2bzrss3qvxy",
        "district":      "25dzncbqyxgx39xq",
        "activity":      "axalt9p7igo67qbn",
        "supervisor":    "72hkeik43b6w36og",
    },
    # Vehicle Inspection (152018)
    "vehicle_inspection": {
        "service_line":  "hxy6pwclvjke1sln",
    },
}

# ---------------------------------------------------------------------------
# Risk Theme Keywords (for observation categorization)
# ---------------------------------------------------------------------------
RISK_THEMES = [
    ("Winter Weather / Ice",    r"ice|freeze|freez|frozen|cold|slip|icy|snow|icicle|resbal"),
    ("Third-Party Management",  r"third.?party|trucker|truck.?driver|spotter|backing.?up|3.?party|vendor|contractor"),
    ("PPE Compliance",          r"ppe|hard.?hat|helmet|safety.?glass|gloves|goggles|hi.?vis|fr |first.?aid.?kit|vest|steel.?toe|boot|hearing|ear.?plug"),
    ("Equipment Deficiency",    r"missing|broken|crack|damage|worn|rip|tear|defect|loose|leak|uncovered|inoperable|malfunction"),
    ("Housekeeping / Trip",     r"housekeep|trip|walkway|debris|clean|board|hose.*walk|clutter|organize"),
    ("Fall Protection",         r"fall|tie.?off|hand.?rail|guard.?rail|ladder|scaffold|harness|lanyard|anchor.?point"),
    ("Lifting / Rigging",       r"lift|casing.*being|boom|crane|sling|rigging|underneath|struck.?by|pinch|crush"),
    ("Signage / Compliance",    r"signage|sign|barricade|barrier|flagging|label"),
    ("Electrical / Grounding",  r"ground|grounded|grounding|electrical|shock|energiz|de.?energiz|lockout|loto"),
    ("Visibility / Weather",    r"visib|wind|dust|dirt.*air|fog|dark|light"),
]

# ---------------------------------------------------------------------------
# All default sections enabled per emphasis style
# ---------------------------------------------------------------------------
CORE_SECTIONS = [
    "executive_summary",
    "incidents",
    "observations",
    "training",
    "corrective_actions",
    "assessments",
]

FLEET_SECTIONS = [
    "fleet_mileage",
    "speeding",
    "camera_events",
    "vehicle_inspections",
]

# ---------------------------------------------------------------------------
# 15 Division Configurations
# ---------------------------------------------------------------------------
DIVISIONS = {
    # -----------------------------------------------------------------------
    # CASING -- Multi-yard, fleet-heavy
    # -----------------------------------------------------------------------
    "casing": {
        "display_name": "Casing Division",
        "company": "BRHAS",
        "manager": "Ken Mattern",
        "safety_rep": "Safety Team",
        "emphasis": "multi_yard",
        "locations": ["Midland", "Bryan", "Kilgore", "Hobbs", "Jourdanton", "Laredo", "San Angelo"],
        "yards": CASING_YARDS,

        # KPA filters
        "kpa_company_filter": "Butch's",
        "kpa_service_line": "Casing",
        "kpa_service_line_alt": ["casing"],  # lowercase variants in data

        # Man-hours
        "man_hours_sheet": "BRHAS",
        "man_hours_division": "Casing",

        # Motive
        "motive_group_ids": [167175, 169090, 169092, 186740, 169091, 186739, 186741, 186746],

        # Sections
        "sections": {
            "executive_summary": True,
            "incidents": True,
            "observations": True,
            "training": True,
            "corrective_actions": True,
            "assessments": True,
            "fleet_mileage": True,
            "speeding": True,
            "camera_events": True,
            "vehicle_inspections": True,
            "jsas": True,
            "equipment_inspections": True,
        },

        # Division-specific form IDs
        "form_ids": {
            "field_assessment":     381707,
            "management_audit":     225522,
            "supervisor_checklist": 329638,
            "prepost_trip":         229645,
            "workplace_inspection": 187860,
            "competency_evals":     [302943, 304908, 404870, 304913],
            "bump_test":            476428,
            "backing_eval":         472040,
        },
    },

    # -----------------------------------------------------------------------
    # RATHOLE -- Multi-yard, fleet-heavy
    # -----------------------------------------------------------------------
    "rathole": {
        "display_name": "Rathole Division",
        "company": "BRHAS",
        "manager": "Division Management",
        "safety_rep": "John Snodgrass",
        "emphasis": "multi_yard",
        "locations": ["Midland", "Levelland", "Barstow", "Jourdanton", "Oklahoma", "North Dakota", "Ohio", "Pennsylvania"],
        "yards": RATHOLE_YARDS,

        "kpa_company_filter": "Butch's",
        "kpa_service_line": "Rathole",
        "kpa_service_line_alt": ["rathole", "rat hole"],

        "man_hours_sheet": "BRHAS",
        "man_hours_division": "Rathole",

        "motive_group_ids": [266026, 266025, 266024, 265996, 290472, 265998, 266028, 266027, 265997, 265988],

        "sections": {
            "executive_summary": True,
            "incidents": True,
            "observations": True,
            "training": True,
            "corrective_actions": True,
            "assessments": True,
            "fleet_mileage": True,
            "speeding": True,
            "vehicle_inspections": True,
            "jsas": True,
            "equipment_inspections": True,
            "emergency_drills": True,
            "permits": True,
        },

        "form_ids": {
            "field_assessment": 153181,
            "management_audit": 509407,
            "site_eval":        487635,
            "skidsteer":        153815,
            "auger_track_rig":  153175,
            "backhoe":          366178,
            "ground_disturbance": 501863,
            "jsa":              409005,
            "emergency_drill":  [459864, 459865],
        },
    },

    # -----------------------------------------------------------------------
    # ANCHORS
    # -----------------------------------------------------------------------
    "anchors": {
        "display_name": "Anchors Division",
        "company": "BRHAS",
        "manager": "Kayla Magana",
        "safety_rep": "John Snodgrass",
        "emphasis": "standard",
        "locations": [],

        "kpa_company_filter": "Butch's",
        "kpa_service_line": "Anchors",
        "kpa_service_line_alt": ["anchors", "anchor"],

        "man_hours_sheet": "BRHAS",
        "man_hours_division": "Anchors",

        "motive_group_ids": [265982],

        "sections": {
            "executive_summary": True,
            "incidents": True,
            "observations": True,
            "training": True,
            "corrective_actions": True,
            "assessments": True,
            "fleet_mileage": True,
            "speeding": True,
            "jsas": True,
        },

        "form_ids": {},
    },

    # -----------------------------------------------------------------------
    # POLY PIPE
    # -----------------------------------------------------------------------
    "poly_pipe": {
        "display_name": "Poly Pipe Division",
        "company": "BRHAS",
        "manager": "Mathew Garcia",
        "safety_rep": "Jose Romero",
        "emphasis": "standard",
        "locations": [],

        "kpa_company_filter": "Butch's",
        "kpa_service_line": "Poly Pipe",
        "kpa_service_line_alt": ["poly pipe", "polypipe"],

        "man_hours_sheet": "BRHAS",
        "man_hours_division": "Poly Pipe",

        "motive_group_ids": [265993, 296040, 296036, 296017, 296020],

        "sections": {
            "executive_summary": True,
            "incidents": True,
            "observations": True,
            "training": True,
            "corrective_actions": True,
            "assessments": True,
            "fleet_mileage": True,
            "speeding": True,
            "jsas": True,
        },

        "form_ids": {
            "field_assessment": 226217,
        },
    },

    # -----------------------------------------------------------------------
    # PIT LINING
    # -----------------------------------------------------------------------
    "pit_lining": {
        "display_name": "Pit Lining Division",
        "company": "BRHAS",
        "manager": "Josh Jacobs",
        "safety_rep": "Jose Romero",
        "emphasis": "standard",
        "locations": [],

        "kpa_company_filter": "Butch's",
        "kpa_service_line": "Pit Lining",
        "kpa_service_line_alt": ["pit lining"],

        "man_hours_sheet": "BRHAS",
        "man_hours_division": "Pit Lining",

        "motive_group_ids": [265992],

        "sections": {
            "executive_summary": True,
            "incidents": True,
            "observations": True,
            "training": True,
            "corrective_actions": True,
            "assessments": True,
            "fleet_mileage": True,
            "speeding": True,
            "jsas": True,
        },

        "form_ids": {
            "field_assessment": 386087,
        },
    },

    # -----------------------------------------------------------------------
    # CONSTRUCTION
    # -----------------------------------------------------------------------
    "construction": {
        "display_name": "Construction Division",
        "company": "BRHAS",
        "manager": "Robert Travis",
        "safety_rep": "Jose Romero",
        "emphasis": "standard",
        "locations": [],

        "kpa_company_filter": "Butch's",
        "kpa_service_line": "Construction",
        "kpa_service_line_alt": ["construction"],

        "man_hours_sheet": "BRHAS",
        "man_hours_division": "Construction",

        "motive_group_ids": [265983],

        "sections": {
            "executive_summary": True,
            "incidents": True,
            "observations": True,
            "training": True,
            "corrective_actions": True,
            "assessments": True,
            "fleet_mileage": True,
            "speeding": True,
            "jsas": True,
            "permits": True,
            "equipment_inspections": True,
        },

        "form_ids": {
            "site_review":        172295,
            "hazard_assessment":  193753,
            "aerial_lift":        193702,
            "pretrip":            337757,
            "jsa":                337763,
            "one_call":           337771,
        },
    },

    # -----------------------------------------------------------------------
    # ENVIRONMENTAL
    # -----------------------------------------------------------------------
    "environmental": {
        "display_name": "Environmental Division",
        "company": "BRHAS",
        "manager": "Joshua Arp",
        "safety_rep": "John Snodgrass",
        "emphasis": "standard",
        "locations": [],

        "kpa_company_filter": "Butch's",
        "kpa_service_line": "Environmental",
        "kpa_service_line_alt": ["environmental"],

        "man_hours_sheet": "BRHAS",
        "man_hours_division": "Environmental",

        "motive_group_ids": [265987],

        "sections": {
            "executive_summary": True,
            "incidents": True,
            "observations": True,
            "training": True,
            "corrective_actions": True,
            "assessments": True,
            "fleet_mileage": True,
            "speeding": True,
            "jsas": True,
        },

        "form_ids": {},
    },

    # -----------------------------------------------------------------------
    # FENCING
    # -----------------------------------------------------------------------
    "fencing": {
        "display_name": "Fencing Division",
        "company": "BRHAS",
        "manager": "Josh Jacobs",
        "safety_rep": "John Snodgrass",
        "emphasis": "standard",
        "locations": [],

        "kpa_company_filter": "Butch's",
        "kpa_service_line": "Fencing",
        "kpa_service_line_alt": ["fencing"],

        "man_hours_sheet": "BRHAS",
        "man_hours_division": "Fencing",

        "motive_group_ids": [265991],

        "sections": {
            "executive_summary": True,
            "incidents": True,
            "observations": True,
            "training": True,
            "corrective_actions": True,
            "assessments": True,
            "fleet_mileage": True,
            "speeding": True,
        },

        "form_ids": {},
    },

    # -----------------------------------------------------------------------
    # DOWNHOLE TOOLS
    # -----------------------------------------------------------------------
    "downhole": {
        "display_name": "Downhole Tools",
        "company": "BRHAS/Valor",
        "manager": "Division Management",
        "safety_rep": "John Snodgrass",
        "emphasis": "standard",
        "locations": [],

        "kpa_company_filter": "Butch's",
        "kpa_service_line": "Downhole Tools",
        "kpa_service_line_alt": ["downhole tools", "downhole"],

        "man_hours_sheet": "BRHAS",
        "man_hours_division": "Downhole Tools",

        "motive_group_ids": [],

        "sections": {
            "executive_summary": True,
            "incidents": True,
            "observations": True,
            "training": True,
            "corrective_actions": True,
            "assessments": True,
        },

        "form_ids": {},
    },

    # -----------------------------------------------------------------------
    # BTI (Butch's Trucking) -- Fleet-heavy emphasis
    # -----------------------------------------------------------------------
    "bti": {
        "display_name": "Butch's Trucking (BTI)",
        "company": "BTI",
        "manager": "Bernard Bradley",
        "safety_rep": "Kelly Rhodes",
        "emphasis": "fleet_heavy",
        "locations": [],

        "kpa_company_filter": "Butch's Trucking",
        "kpa_service_line": "BTI",
        "kpa_service_line_alt": ["bti", "butch's trucking", "trucking"],

        "man_hours_sheet": "BTI",
        "man_hours_division": None,  # entire sheet is BTI

        "motive_group_ids": [265989],

        "sections": {
            "executive_summary": True,
            "incidents": True,
            "observations": True,
            "training": True,
            "corrective_actions": True,
            "assessments": True,
            "fleet_mileage": True,
            "speeding": True,
            "vehicle_inspections": True,
            "dot_inspections": True,
            "jsas": True,
        },

        "form_ids": {
            "field_assessment":    227486,
            "vehicle_inspection":  152018,
            "jsa":                 227487,
            "four_gas_meter":      439122,
            "driver_competency":   376390,
        },

        # BTI-specific: Bernard Bradley's inspection field hashes
        "inspection_fields": {
            "driver":    "o4anp0lsn0j6sl4h",
            "truck":     "shwrxgdeo3liaukp",
            "score":     "score-percent",
            "company":   "ge09m6h1ne6po6x9",
            "inspector": "537yjjovkn8vtqex",
            "location":  "jovgwqv8n1rkj6x7",
            "trailer":   "y9nj1tqxtcandgfr",
            "truck_make":"28oshxapsgaqdq2u",
        },
    },

    # -----------------------------------------------------------------------
    # TRANSCEND DRILLING -- Rig-based emphasis
    # -----------------------------------------------------------------------
    "transcend": {
        "display_name": "Transcend Drilling",
        "company": "Transcend",
        "manager": "Division Management",
        "safety_rep": "Kelly Rhodes",
        "emphasis": "rig_based",
        "locations": [],

        "kpa_company_filter": "Transcend",
        "kpa_service_line": "Transcend",
        "kpa_service_line_alt": ["transcend", "transcend drilling", "drilling"],

        "man_hours_sheet": "Transcend",
        "man_hours_division": None,

        "motive_group_ids": [265986],

        "sections": {
            "executive_summary": True,
            "incidents": True,
            "observations": True,
            "training": True,
            "corrective_actions": True,
            "assessments": True,
            "fleet_mileage": True,
            "speeding": True,
            "jsas": True,
            "rig_inspections": True,
            "emergency_drills": True,
            "permits": True,
        },

        "form_ids": {
            "field_assessment":   206674,
            "management_audit":   225539,
            "pre_spud":           206675,
            "rig_inspection":     385365,
            "observation":        484193,
            "pre_tower_meeting":  328995,
            "hot_work":           336552,
            "confined_space":     336555,
            "loto":               336554,
            "emergency_drill":    [352961, 253784],
        },

        # Transcend observation card field hashes
        "obs_fields": {
            "company":     "t5187momol3em85v",
            "district":    "1tktpfww46umdo7t",
            "rig":         "614s26cxqzc5ne1e",
            "type":        "bff8m4x6xbc033kg",
            "name":        "0kc57oj2zkse21o3",
            "customer":    "vxew6ukynemxwvjr",
            "location":    "lg5pnj4chjadnv46",
            "description": "uncbcge9x8vow9pn",
            "action":      "dpy2klalngsr7ek9",
        },

        # JSA field hashes
        "jsa_fields": {
            "date_time":   "mxptkbxaxjp54hli",
            "company":     "6zyx6l5f244mk0v5",
            "service_line":"77ykc2bzrss3qvxy",
            "district":    "25dzncbqyxgx39xq",
            "customer":    "x2a6o6tc65c515f5",
            "location":    "8t848479ka9sivqb",
            "rig":         "n8mv4y9zkfawwhf7",
            "supervisor":  "72hkeik43b6w36og",
            "activity":    "axalt9p7igo67qbn",
            "topics":      "f1yva1wr1htct6at",
            "led_by":      "s7qd5nxgp7ikod85",
            "review_date": "ohn4nn97da4oxm9p",
            "reviewer":    "va1dzacufps1gh2t",
            "comments":    "0s8y7fkcqfcdtajc",
        },
    },

    # -----------------------------------------------------------------------
    # VALOR ENERGY SERVICES
    # -----------------------------------------------------------------------
    "valor": {
        "display_name": "Valor Energy Services",
        "company": "Valor",
        "manager": "Bobby Morris",
        "safety_rep": "John Snodgrass",
        "emphasis": "standard",
        "locations": [],

        "kpa_company_filter": "Valor",
        "kpa_service_line": "Valor",
        "kpa_service_line_alt": ["valor", "valor energy"],
        # Valor users are coded as "DownHole Tools" LOB in KPA but at Valor field offices
        "kpa_field_office_ids": [
            "6710170b8ee2a10019b2e7e0",  # Levelland Yard Valor
            "6009f55901f3bb0142271521",  # VALOR LEVELLAND
        ],

        "man_hours_sheet": "Valor",
        "man_hours_division": None,

        "motive_group_ids": [265985],

        "sections": {
            "executive_summary": True,
            "incidents": True,
            "observations": True,
            "training": True,
            "corrective_actions": True,
            "assessments": True,
            "fleet_mileage": True,
            "speeding": True,
            "jsas": True,
        },

        "form_ids": {},
    },

    # -----------------------------------------------------------------------
    # PERMIAN EQUIPMENT RENTALS (PER)
    # -----------------------------------------------------------------------
    "per": {
        "display_name": "Permian Equipment Rentals",
        "company": "PER",
        "manager": "Tate Lair",
        "safety_rep": "Kelly Rhodes",
        "emphasis": "standard",
        "locations": [],

        "kpa_company_filter": "Permian",
        "kpa_service_line": "Rentals",
        "kpa_service_line_alt": ["per", "permian equipment", "permian", "rentals", "rental"],

        "man_hours_sheet": "PER",
        "man_hours_division": None,

        "motive_group_ids": [],

        "sections": {
            "executive_summary": True,
            "incidents": True,
            "observations": True,
            "training": True,
            "corrective_actions": True,
            "assessments": True,
            "jsas": True,
        },

        "form_ids": {},
    },

    # -----------------------------------------------------------------------
    # HUTCH'S OILFIELD SUPPLY
    # -----------------------------------------------------------------------
    "hutchs": {
        "display_name": "Hutch's Oilfield Supply",
        "company": "Hutch's",
        "manager": "Division Management",
        "safety_rep": "Kelly Rhodes",
        "emphasis": "standard",
        "locations": [],

        "kpa_company_filter": "Hutch",
        "kpa_service_line": "Hutch's",
        "kpa_service_line_alt": ["hutch's", "hutchs", "hutch"],

        "man_hours_sheet": "Hutchs",
        "man_hours_division": None,

        "motive_group_ids": [],

        "sections": {
            "executive_summary": True,
            "incidents": True,
            "observations": True,
            "training": True,
            "corrective_actions": True,
            "assessments": True,
        },

        "form_ids": {},
    },

    # -----------------------------------------------------------------------
    # SHOP / MAINTENANCE
    # -----------------------------------------------------------------------
    "shop": {
        "display_name": "Shop / Maintenance",
        "company": "BRHAS",
        "manager": "Division Management",
        "safety_rep": "Safety Team",
        "emphasis": "standard",
        "locations": [],

        "kpa_company_filter": "Butch's",
        "kpa_service_line": "Shop",
        "kpa_service_line_alt": ["shop", "maintenance", "shop/maintenance"],

        "man_hours_sheet": "BRHAS",
        "man_hours_division": "Shop",

        "motive_group_ids": [],

        "sections": {
            "executive_summary": True,
            "incidents": True,
            "observations": True,
            "training": True,
            "corrective_actions": True,
            "assessments": True,
        },

        "form_ids": {},
    },
}

# ---------------------------------------------------------------------------
# Man-Hours Department Mapping (2026 "Detail" sheet Home Department values)
# Maps division key -> list of Home Department values to sum
# ---------------------------------------------------------------------------
MAN_HOURS_DEPARTMENTS = {
    "casing":        ["Casing"],
    "rathole":       ["Rat Hole", "Fabrication", "Water Trucking"],
    "anchors":       ["Anchor"],
    "poly_pipe":     ["Poly Pipe"],
    "pit_lining":    ["Pit Lining"],
    "construction":  ["Civil"],
    "environmental": ["Containment"],
    "fencing":       ["Fencing"],
    "downhole":      ["Downhole Tools", "Drilling Tools"],
    "shop":          ["Shop"],
    "bti":           ["Trucking"],
    "transcend":     ["Drilling", "TD Rig 4", "TD Rig 12", "TD Rig 16", "TD Rig 18", "TD Rig 20"],
    "valor":         [],  # Valor uses Co Code filtering -- handled separately
    "per":           ["Rentals"],
    "hutchs":        ["Hutch's Oilfield Supply"],
}

# Co Code mapping for companies that span their own payroll entity
MAN_HOURS_CO_CODES = {
    "per":       ["55BRH01"],
    "bti":       ["55BRH02"],
    "transcend": ["55BRH03"],
    "valor":     ["55BRH04"],
    "hutchs":    ["55BRH06"],
}


def get_division(key):
    """Get division config by key, raising KeyError if not found."""
    if key not in DIVISIONS:
        raise KeyError(f"Unknown division: {key!r}. Valid: {sorted(DIVISIONS.keys())}")
    return DIVISIONS[key]


def get_enabled_sections(division_key):
    """Return list of enabled section names for a division."""
    cfg = get_division(division_key)
    return [s for s, enabled in cfg["sections"].items() if enabled]


def get_motive_group_ids(division_key):
    """Return set of Motive group IDs for a division."""
    return set(get_division(division_key).get("motive_group_ids", []))


# Location normalization for man-hours Excel
LOCATION_NORMALIZE = {
    "midland yukon": "Midland",
    "midland 1788 brhas": "Midland",
    "overhead brhas": "Overhead",
    "overhead": "Overhead",
}


def normalize_location(raw_location):
    """Normalize location names from man-hours Excel to match yard names."""
    if not raw_location:
        return raw_location
    lower = raw_location.strip().lower()
    return LOCATION_NORMALIZE.get(lower, raw_location.strip())


def get_form_id(division_key, form_name):
    """Get a division-specific form ID, falling back to shared forms."""
    cfg = get_division(division_key)
    # Check division-specific first
    div_id = cfg.get("form_ids", {}).get(form_name)
    if div_id is not None:
        return div_id
    # Fall back to shared HSE forms
    return SHARED_FORMS.get(form_name)
