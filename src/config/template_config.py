from pathlib import Path

TEMPLATE_CONFIG = {
    "Disney Channel": {
        "template_path": Path("templates/X_Social Media Calendar DC YouTube EMEA - month YYYY.xltx"),
        "full_calendar_sheet": "FULL CALENDAR YT DC EMEA",
        "monthly_overview_sheet": "MONTHLY OVERVIEW YT DC EMEA",
        "week_anchor_cells": ["D6", "D59", "D112", "D165", "D218"],
        "supports_monthly_overview": True,
    },
    "Disney Jr.": {
        "template_path": Path("templates/X_Social Media Calendar DJR YouTube EMEA - month YYYY.xltx"),
        "full_calendar_sheet": "FULL CALENDAR YT DJR EMEA",
        "monthly_overview_sheet": "MONTHLY OVERVIEW YT DJR EMEA",
        "week_anchor_cells": ["D6", "D59", "D112", "D165", "D218"],
        "supports_monthly_overview": True,
    },
}

def get_template_config(brand: str) -> dict:
    if brand not in TEMPLATE_CONFIG:
        raise ValueError(f"Unknown brand '{brand}'. Available: {', '.join(TEMPLATE_CONFIG)}")
    return TEMPLATE_CONFIG[brand]