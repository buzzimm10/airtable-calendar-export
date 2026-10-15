import re

SHEET_COUNTRIES = [
    "Spain",
    "France (French)",
    "Poland (Polish)",
    "Portugal",
    "Italy (Italian)",
    "Romania (Romanian)",
    "Greece (Greek)",
    "Middle East (Arabic - MSA)",
    "Netherlands (Dutch)",
    "Belgium (French)",
    "Israel (Hebrew)",
    "Bulgaria (Bulgarian)",
    "Hungary (Hungarian)",
    "Czech R. (Czech)",
    "Sweden (Swedish)",
    "Denmark (Danish)",
    "Norway (Norwegian)",
]

CHANNEL_COUNTRY_MAP = {
    # Disney Channel variants
    "disney channel spain": {"Spain"},
    "disney channel españa": {"Spain"},
    "disney channel es": {"Spain"},
    "disney channel polska": {"Poland (Polish)"},
    "disney channel pl": {"Poland (Polish)"},
    "disney channel portugal": {"Portugal"},
    "disney channel pt": {"Portugal"},
    "disney channel italia": {"Italy (Italian)"},
    "disney channel it": {"Italy (Italian)"},
    "disney channel romania": {"Romania (Romanian)"},
    "disney channel romanian": {"Romania (Romanian)"},
    "disney channel greece": {"Greece (Greek)"},
    "disney channel gr": {"Greece (Greek)"},
    "disney channel netherlands": {"Netherlands (Dutch)"},
    "disney channel nl": {"Netherlands (Dutch)"},
    "disney channel belgium": {"Belgium (French)"},
    "disney channel be": {"Belgium (French)"},
    "disney channel israel": {"Israel (Hebrew)"},
    "disney channel il": {"Israel (Hebrew)"},
    "disney channel hungary": {"Hungary (Hungarian)"},
    "disney channel hu": {"Hungary (Hungarian)"},
    "disney channel czech r": {"Czech R. (Czech)"},
    "disney channel cz": {"Czech R. (Czech)"},
    "disney channel sweden": {"Sweden (Swedish)"},
    "disney channel se": {"Sweden (Swedish)"},
    "disney channel denmark": {"Denmark (Danish)"},
    "disney channel dk": {"Denmark (Danish)"},
    "disney channel norway": {"Norway (Norwegian)"},
    "disney channel no": {"Norway (Norwegian)"},
    # Disney Kids variants
    "disney kids france": {"France (French)"},
    "disney kids italia": {"Italy (Italian)"},
    "disney kids greece": {"Greece (Greek)"},
    "disney kids belgië-nederland": {"Netherlands (Dutch)"},
    "disney kids be": {"Belgium (French)"},    
    "disney kids belgium": {"Belgium (French)"},  
    "disney kids mena": {"Middle East (Arabic - MSA)"},
    "disney kids bulgaria": {"Bulgaria (Bulgarian)"},
    # Shared channels (Disney Romania, Disney Channel MENA)
    "disney romania": {"Romania (Romanian)"},
    "disney channel mena": {"Middle East (Arabic - MSA)"},
    "disney channel me": {"Middle East (Arabic - MSA)"},
    "disney channel arabia": {"Middle East (Arabic - MSA)"},
    # EMEA variant (all countries)
    "disney channel emea": set(SHEET_COUNTRIES),
}

COUNTRY_ALIASES = {
    "spain": "Spain",
    "es": "Spain",

    "france": "France (French)",
    "fr": "France (French)",
    "french": "France (French)",

    "poland": "Poland (Polish)",
    "pl": "Poland (Polish)",
    "polska": "Poland (Polish)",
    "polish": "Poland (Polish)",

    "portugal": "Portugal",
    "pt": "Portugal",

    "italy": "Italy (Italian)",
    "italia": "Italy (Italian)",
    "it": "Italy (Italian)",
    "italian": "Italy (Italian)",

    "romania": "Romania (Romanian)",
    "ro": "Romania (Romanian)",
    "romanian": "Romania (Romanian)",

    "greece": "Greece (Greek)",
    "gr": "Greece (Greek)",
    "greek": "Greece (Greek)",

    "mena": "Middle East (Arabic - MSA)",
    "middle east": "Middle East (Arabic - MSA)",
    "arabia": "Middle East (Arabic - MSA)",
    "arabic": "Middle East (Arabic - MSA)",
    "msa": "Middle East (Arabic - MSA)",

    "netherlands": "Netherlands (Dutch)",
    "nl": "Netherlands (Dutch)",
    "dutch": "Netherlands (Dutch)",

    "belgium": "Belgium (French)",
    "be": "Belgium (French)",
    "belgique": "Belgium (French)",

    "israel": "Israel (Hebrew)",
    "il": "Israel (Hebrew)",
    "hebrew": "Israel (Hebrew)",

    "bulgaria": "Bulgaria (Bulgarian)",
    "bg": "Bulgaria (Bulgarian)",
    "bulgarian": "Bulgaria (Bulgarian)",

    "hungary": "Hungary (Hungarian)",
    "hu": "Hungary (Hungarian)",
    "hungarian": "Hungary (Hungarian)",

    "czech": "Czech R. (Czech)",
    "cz": "Czech R. (Czech)",
    "czech republic": "Czech R. (Czech)",

    "sweden": "Sweden (Swedish)",
    "se": "Sweden (Swedish)",
    "swedish": "Sweden (Swedish)",

    "denmark": "Denmark (Danish)",
    "dk": "Denmark (Danish)",
    "danish": "Denmark (Danish)",

    "norway": "Norway (Norwegian)",
    "no": "Norway (Norwegian)",
    "norwegian": "Norway (Norwegian)",
}

COUNTRY_SHORT_NAMES = {
    "Spain": "Spain",
    "France (French)": "France",
    "Poland (Polish)": "Poland",
    "Portugal": "Portugal",
    "Italy (Italian)": "Italy",
    "Romania (Romanian)": "Romania",
    "Greece (Greek)": "Greece",
    "Middle East (Arabic - MSA)": "MENA",
    "Netherlands (Dutch)": "Netherlands",
    "Belgium (French)": "Belgium",
    "Israel (Hebrew)": "Israel",
    "Bulgaria (Bulgarian)": "Bulgaria",
    "Hungary (Hungarian)": "Hungary",
    "Czech R. (Czech)": "Czech Republic",
    "Sweden (Swedish)": "Sweden",
    "Denmark (Danish)": "Denmark",
    "Norway (Norwegian)": "Norway",
}

REGION_PRESETS = {
    "tier s": {"Spain"},
    "s": {"Spain"},
    "group s": {"Spain"},
    "a": {"Spain", "France (French)", "Poland (Polish)"},
    "tier a": {"Spain", "France (French)", "Poland (Polish)"},
    "group a": {"Spain", "France (French)", "Poland (Polish)"},
    "tier b": {
        "Portugal", "Italy (Italian)", "Romania (Romanian)", "Greece (Greek)",
        "Middle East (Arabic - MSA)", "Netherlands (Dutch)", "Belgium (French)",
        "Israel (Hebrew)", "Bulgaria (Bulgarian)"
    },
    "b": {
        "Portugal", "Italy (Italian)", "Romania (Romanian)", "Greece (Greek)",
        "Middle East (Arabic - MSA)", "Netherlands (Dutch)", "Belgium (French)",
        "Israel (Hebrew)", "Bulgaria (Bulgarian)"
    },
    "group b": {
        "Portugal", "Italy (Italian)", "Romania (Romanian)", "Greece (Greek)",
        "Middle East (Arabic - MSA)", "Netherlands (Dutch)", "Belgium (French)",
        "Israel (Hebrew)", "Bulgaria (Bulgarian)"
    },
    "tier c": {"Hungary (Hungarian)", "Czech R. (Czech)", "Sweden (Swedish)", "Denmark (Danish)", "Norway (Norwegian)"},
    "c": {"Hungary (Hungarian)", "Czech R. (Czech)", "Sweden (Swedish)", "Denmark (Danish)", "Norway (Norwegian)"},
    "group c": {"Hungary (Hungarian)", "Czech R. (Czech)", "Sweden (Swedish)", "Denmark (Danish)", "Norway (Norwegian)"},
    "mena": {"Middle East (Arabic - MSA)"},
    "nordics": {"Sweden (Swedish)", "Denmark (Danish)", "Norway (Norwegian)"},
}

EXCLUDE_NOTE_PATTERNS = [
    (re.compile(r"\bexcept france\b", re.I), "France (French)"),
    (re.compile(r"\bexcept mena\b", re.I), "Middle East (Arabic - MSA)"),
    (re.compile(r"\bexcept (spain|es)\b", re.I), "Spain"),
    (re.compile(r"\bexcept (poland|pl|polska)\b", re.I), "Poland (Polish)"),
    (re.compile(r"\bexcept (italy|it|italia)\b", re.I), "Italy (Italian)"),
    (re.compile(r"\bexcept (portugal|pt)\b", re.I), "Portugal"),
    (re.compile(r"\bexcept (romania|ro)\b", re.I), "Romania (Romanian)"),
    (re.compile(r"\bexcept (greece|gr)\b", re.I), "Greece (Greek)"),
    (re.compile(r"\bexcept (netherlands|nl|dutch)\b", re.I), "Netherlands (Dutch)"),
    (re.compile(r"\bexcept (belgium|be|belgique)\b", re.I), "Belgium (French)"),
    (re.compile(r"\bexcept (israel|il)\b", re.I), "Israel (Hebrew)"),
    (re.compile(r"\bexcept (bulgaria|bg)\b", re.I), "Bulgaria (Bulgarian)"),
    (re.compile(r"\bexcept (hungary|hu)\b", re.I), "Hungary (Hungarian)"),
    (re.compile(r"\bexcept (czech|cz)\b", re.I), "Czech R. (Czech)"),
    (re.compile(r"\bexcept (sweden|se)\b", re.I), "Sweden (Swedish)"),
    (re.compile(r"\bexcept (denmark|dk)\b", re.I), "Denmark (Danish)"),
    (re.compile(r"\bexcept (norway|no)\b", re.I), "Norway (Norwegian)"),
]
