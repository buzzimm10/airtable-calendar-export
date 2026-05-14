BRAND_CONFIG = {
    "disney channel": {
        "label": "Disney Channel",
        "template": "templates/X_Social Media Calendar DC YouTube EMEA - month YYYY.xltx",
        "exporter": "dc",
    },
    "disney jr.": {
        "label": "Disney Jr.",
        "template": "templates/X_Social Media Calendar DJR YouTube EMEA - month YYYY.xltx",
        "exporter": "djr",
    },
}


def normalize_brand(value: str) -> str:
    return str(value or "").strip().lower()


def get_brand_config(value: str) -> dict:
    key = normalize_brand(value)

    if key not in BRAND_CONFIG:
        raise ValueError(f"Unsupported brand: {value}")

    return BRAND_CONFIG[key]