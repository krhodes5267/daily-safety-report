"""
API_CONFIG.PY -- Shared constants for dashboard_api.py
=====================================================
Extracted from daily report scripts. Single source of truth for
form IDs, group IDs, LOB IDs, hashes, and man-hours data.
"""

# ==============================================================================
# KPA API
# ==============================================================================
KPA_API_BASE = "https://api.novaraflex.com/v1"

# Form IDs
OBSERVATION_FORM_ID = 151085
INCIDENT_FORM_ID = 151622
CASING_ASSESSMENT_FORM_ID = 381707
CASING_PREPOST_FORM_ID = 229645
TD_RIG_INSPECTION_FORM_ID = 385365
TD_OBSERVATION_FORM_ID = 484193
POLY_PIPE_FORM_ID = 226217
PIT_LINING_FORM_ID = 386087
CONSTRUCTION_FORM_ID = 172295
RATHOLE_FORM_ID = 153181
VEHICLE_INSPECTION_FORM_ID = 152018
WORKPLACE_INSPECTION_FORM_ID = 152034

# KPA field hashes
OBS_TYPE_HASH = "bff8m4x6xbc033kg"
OBS_DESC_HASH = "uncbcge9x8vow9pn"
OBS_LOCATION_HASH = "lg5pnj4chjadnv46"
OBS_YARD_HASH = "7vj2l992y7fwqhwz"

INC_TYPE_HASH = "nojcquy0tfl9hqih"
INC_EMPLOYEE_HASH = "55gg4nkoemnnfo2a"
INC_DESC_HASH = "313e9txgrof0uute"
INC_LOCATION_HASH = "9ohdd2lwvl7p0oc6"
SVC_LINE_HASH = "sha7vur5q2l6d6gq"
COMPANY_HASH = "lsx3msa0w9n9edb4"

# Assessment form IDs for API date-range fetching
ASSESSMENT_FORM_IDS = [
    381707,  # CSG - Safety Casing Field Assessment
    229645,  # CSG - Pre/Post Trip Inspection
    385365,  # TD - Rig Inspection
    226217,  # WS - Poly Pipe Field Assessment
    386087,  # WS - Pit Lining Field Assessment
    172295,  # Construction - Site Safety Review
    153181,  # RH - Rathole Field Assessment
    152018,  # Vehicle Inspection Checklist
    152034,  # HSE - Workplace Inspection Checklist
]

SERVICE_LINE_HASHES = [
    "64c7upqkyt79zhh1", "sha7vur5q2l6d6gq", "68jbiriixiurowou",
    "hxy6pwclvjke1sln", "b5dqbf3qgxga92fi", "7cl5rryw636f6wbx",
    "service_line", "Service Line", "division", "Division",
]

# Form -> Division mapping
FORM_DIVISION_MAP = {
    381707: "Casing",
    229645: "Casing",
    385365: "Transcend",
    484193: "Transcend",
    226217: "Poly Pipe",
    386087: "Pit Lining",
    172295: "Construction",
    153181: "Rathole",
    152018: "Shared",
    152034: "Shared",
}

# Safety forms whitelist (for CA filtering)
SAFETY_FORMS = [
    "Observation Card", "CSG - JSA Review", "CSG - Safety Casing Field Assessment",
    "TD - Rig Inspection", "Incident Report", "Root Cause Analysis",
    "HSE - Workplace Inspection Checklist", "RH - Rathole Field Assessment",
    "Construction - Site Safety Review", "CSG - Workplace Inspection",
    "Tailgate Talk", "JSA Log", "TD - Observation Card",
    "Equipment Inspection", "Vehicle Inspection Checklist", "Non-DOT Pre-Trip Inspection",
    "Unknown (151673)", "Unknown (161108)", "Unknown (171710)", "Unknown (177981)",
]

# KPA company name normalization
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

# KPA service line -> dashboard division
KPA_SVC_TO_DIVISION = {
    "Casing": "Casing",
    "Rat Hole": "Rathole", "Rathole": "Rathole",
    "Poly Pipe": "Poly Pipe",
    "Pit Lining": "Pit Lining",
    "Anchor": "Anchors", "Anchors": "Anchors",
    "Construction": "Construction", "Civil": "Construction",
    "Water/Construction": "Construction",
    "Environmental": "Environmental", "Containment": "Environmental",
    "Fencing": "Fencing",
    "Drilling Tools": "Drilling Tools", "Downhole Tools": "Valor",
    "Shop": "Shop", "Fabrication": "Fabrication",
    "Corporate": "Corporate",
    # Non-BRHAS companies (map to their division name used in dashboard)
    "BTI": "BTI", "Butch's Trucking": "BTI", "Trucking": "BTI",
    "Drilling": "Transcend", "Transcend": "Transcend",
    "Transcend Drilling": "Transcend",
    "Drilling - Rig 4": "Transcend", "Drilling - Rig 12": "Transcend",
    "Drilling - Rig 16": "Transcend", "Drilling - Rig 18": "Transcend",
    "Drilling - Rig 20": "Transcend",
    "Valor": "Valor", "Valor Energy": "Valor",
    "Valor Energy Services": "Valor",
    "Rentals": "Permian", "PER": "Permian", "Permian": "Permian",
}

KPA_RESPONSE_URL = "https://brhas-ees.novaraflex.com/forms/responses/view"

# ==============================================================================
# MOTIVE API
# ==============================================================================
MOTIVE_BASE_V1 = "https://api.gomotive.com/v1"
MOTIVE_BASE_V2 = "https://api.gomotive.com/v2"

# Casing yard group IDs
CASING_GROUP_IDS = {
    167175: "Midland",
    169090: "Bryan",
    169092: "Kilgore",
    186740: "Hobbs",
    169091: "Jourdanton",
    186739: "Laredo",
    186741: "San Angelo",
}

# Full group ID -> (division, yard) map
GROUP_ID_MAP = {
    # Casing
    167175: ("Casing", "Midland"),
    169090: ("Casing", "Bryan"),
    169092: ("Casing", "Kilgore"),
    186740: ("Casing", "Hobbs"),
    169091: ("Casing", "Jourdanton"),
    186739: ("Casing", "Laredo"),
    186741: ("Casing", "San Angelo"),
    # Rathole
    167176: ("Rathole", "Midland"),
    220453: ("Rathole", "Levelland"),
    274965: ("Rathole", "Barstow"),
    186742: ("Rathole", "Jourdanton"),
    307752: ("Rathole", "Ohio"),
    341789: ("Rathole", "Pennsylvania"),
    308775: ("Rathole", "Oklahoma"),
    351218: ("Rathole", "North Dakota"),
    # Other BRHAS
    167180: ("Poly Pipe", "Midland"),
    167177: ("Anchors", "Midland"),
    247032: ("Construction", "Midland"),
    167179: ("Pit Lining", "Midland"),
    186738: ("Environmental", "Midland"),
    220456: ("Fencing", "Midland"),
    329823: ("Drilling Tools", "Midland"),
    # Non-BRHAS
    167178: ("Valor", "Midland"),
    186743: ("BTI", "Midland"),
    247035: ("Transcend", "Midland"),
    # --- New Motive groups (discovered 2026-04-20) ---
    # BRHAS divisions
    265982: ("Anchors", "Midland"),
    265983: ("Construction", "Midland"),
    265984: ("Permian", "Midland"),
    265985: ("Valor", "Levelland"),
    265986: ("Transcend", "Midland"),
    265987: ("Environmental", "Midland"),
    265988: ("Corporate", "Midland"),      # WTC / admin vehicles
    265989: ("BTI", "Midland"),
    265991: ("Fencing", "Midland"),
    265992: ("Pit Lining", "Midland"),
    265993: ("Poly Pipe", "Midland"),
    # Rathole sub-yards
    265996: ("Rathole", "Jourdanton"),
    265997: ("Rathole", "North Dakota"),
    265998: ("Rathole", "Oklahoma"),
    266024: ("Rathole", "Barstow"),
    266025: ("Rathole", "Levelland"),
    266026: ("Rathole", "Midland"),
    266027: ("Rathole", "Unknown"),         # TOW vehicles
    266028: ("Rathole", "Unknown"),         # WIN vehicles
    290471: ("Rathole", "Jourdanton"),
    # Poly Pipe sub-yards
    296017: ("Poly Pipe", "Bryan"),
    296020: ("Poly Pipe", "Hobbs"),
    296036: ("Poly Pipe", "Jourdanton"),
    296040: ("Poly Pipe", "Midland"),
    # Previously unmapped
    186746: ("Casing", "Unknown"),          # Vehicle 21356
}

# ==============================================================================
# MAN-HOURS (Q1 2026 payroll)
# ==============================================================================
Q1_MAN_HOURS = {
    "BRHAS": {
        "total": 545770,
        "by_division": {
            "Casing": 310907, "Rathole": 102556, "Poly Pipe": 42774,
            "Corporate": 30911, "Pit Lining": 15142, "Fabrication": 6964,
            "Construction": 7918, "Shop": 7661, "Anchors": 7675,
            "Fencing": 5701, "Environmental": 4527, "Drilling Tools": 3032,
        },
        "monthly": {"2026-01": 178567, "2026-02": 170786, "2026-03": 196417},
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

# ==============================================================================
# LOB (Line of Business) -> Company/Division (for training)
# ==============================================================================
LOB_TO_COMPANY = {
    "6009f696823a6201bbc9b056": "BRHAS",
    "6009f696823a6201bbc9b04a": "BRHAS",
    "6009f696823a6201bbc9b051": "BRHAS",
    "6009f696823a6201bbc9b053": "BRHAS",
    "6009f696823a6201bbc9b04d": "BRHAS",
    "609452db13b68a005495157f": "BRHAS",
    "6148833d9e0bc6004e070765": "BRHAS",
    "6152f95cb365e00022f9ecb5": "BRHAS",
    "6009f696823a6201bbc9b05b": "BTI",
    "6009f696823a6201bbc9b058": "Transcend",
    "6009f696823a6201bbc9b04b": "Valor",
    "5d166c18efd5700017316462": "Permian",
}

LOB_TO_DIVISION = {
    "6009f696823a6201bbc9b056": "Casing",
    "6009f696823a6201bbc9b04a": "Rathole",
    "6009f696823a6201bbc9b051": "Anchors",
    "6009f696823a6201bbc9b053": "Pit Lining",
    "6009f696823a6201bbc9b04d": "Poly Pipe",
    "609452db13b68a005495157f": "Fencing",
    "6148833d9e0bc6004e070765": "Construction",
    "6152f95cb365e00022f9ecb5": "Environmental",
}
