import re
from typing import List, Optional

from config.countries import (
    COUNTRY_ALIASES,
    SHEET_COUNTRIES,
    COUNTRY_SHORT_NAMES,
    CHANNEL_COUNTRY_MAP,
    REGION_PRESETS,
)

from core.text_formatting import (
    normalize_str,
    lower_clean,

)

def extract_excluded_countries_from_notes(notes: Optional[str]) -> set:
    """
    Parse exclusions from strings like:
    - Except France
    - Except Italy, Portugal and Bulgaria
    - Except Italy / Portugal / Bulgaria
    - Except Portugal, Italy & Bulgaria
    - Except MENA
    """
    note_text = lower_clean(notes)
    if not note_text:
        return set()

    excluded = set()

    m = re.search(r"\bexcept\b(.+)", note_text, re.I)
    if not m:
        return excluded

    tail = m.group(1)

    # Normalize separators
    tail = tail.replace("&", ",")
    tail = tail.replace("/", ",")
    tail = re.sub(r"\band\b", ",", tail, flags=re.I)

    parts = [p.strip() for p in tail.split(",") if p.strip()]

    for part in parts:
        if part in COUNTRY_ALIASES:
            excluded.add(COUNTRY_ALIASES[part])
            continue

        for alias, canonical in COUNTRY_ALIASES.items():
            if alias in part:
                excluded.add(canonical)

    return excluded

def normalize_excluded_market_name(value: str) -> str:
    v = normalize_str(value)

    aliases = {
        "Middle East (Arabic - MSA)": "MENA",
        "France (French)": "France",
        "Spain": "Spain",
        "Portugal": "Portugal",
        "Italy (Italian)": "Italy",
        "Bulgaria (Bulgarian)": "Bulgaria",
    }

    return aliases.get(v, v)


def format_excluded_markets_comment(excluded_markets) -> str:
    """
    Converts Airtable Excluded Markets into:
    Except France
    Except MENA
    Except France, Spain
    """
    if not excluded_markets:
        return ""

    if isinstance(excluded_markets, list):
        items = excluded_markets
    else:
        items = [excluded_markets]

    cleaned = []
    for item in items:
        name = normalize_excluded_market_name(item)
        if name and name not in cleaned:
            cleaned.append(name)

    if not cleaned:
        return ""

    return "Except " + ", ".join(cleaned)

def collect_included_countries(youtube_channel, region, notes, excluded_markets=None, channel_id_to_name_map=None):
    """
    Determine included countries using:
    1) YouTube Channel (primary)
    2) Region / Tier (secondary)
    3) Notes exclusions (final override)

    If both channel and region exist, use intersection.
    """
    if channel_id_to_name_map is None:
        channel_id_to_name_map = {}

    included = set()

    # --------
    # Channel
    # --------
    channel_countries = set()

    if youtube_channel:
        if isinstance(youtube_channel, list):
            raw_channels = youtube_channel
        else:
            raw_channels = [youtube_channel]

        resolved_channels = []
        for item in raw_channels:
            item_norm = normalize_str(item)
            if item_norm in channel_id_to_name_map:
                resolved_channels.append(lower_clean(channel_id_to_name_map[item_norm]))
            else:
                resolved_channels.append(lower_clean(item_norm))

        channel_blob = " | ".join(resolved_channels)

        for key, countries in CHANNEL_COUNTRY_MAP.items():
            if key in channel_blob:
                channel_countries |= set(countries)

    # -------
    # Region
    # -------
    region_countries = set()

    if region:
        if isinstance(region, list):
            reg_text = " | ".join(lower_clean(x) for x in region)
        else:
            reg_text = lower_clean(region)

        parts = re.split(r"[,&+/|]+", reg_text)
        for part in parts:
            part = part.strip()
            if part in REGION_PRESETS:
                region_countries |= set(REGION_PRESETS[part])

    # Combine channel + region
    if channel_countries and region_countries:
        included = channel_countries & region_countries
    elif channel_countries:
        included = channel_countries
    elif region_countries:
        included = region_countries
    else:
        included = set()

    # Notes exclusions
    excluded = set()

    # New structured field
    if excluded_markets:
        raw_items = excluded_markets if isinstance(excluded_markets, list) else [excluded_markets]
        for item in raw_items:
            item_clean = lower_clean(item)
            if item_clean in COUNTRY_ALIASES:
                excluded.add(COUNTRY_ALIASES[item_clean])
            else:
                for alias, canonical in COUNTRY_ALIASES.items():
                    if alias in item_clean:
                        excluded.add(canonical)

    # Optional legacy fallback
    excluded |= extract_excluded_countries_from_notes(notes)

    included -= excluded
    return included, excluded

def resolve_region_countries(region_val) -> set:
    """
    Resolve Region/Tier field into canonical country set using REGION_PRESETS.
    """
    region_countries = set()

    if not region_val:
        return region_countries

    if isinstance(region_val, list):
        reg_text = " | ".join(lower_clean(x) for x in region_val)
    else:
        reg_text = lower_clean(region_val)

    # Clean list-like strings
    reg_text = reg_text.replace("[", "").replace("]", "").replace("'", "").replace('"', "")

    parts = re.split(r"[,&+/|]+", reg_text)
    for part in parts:
        part = part.strip()
        if part in REGION_PRESETS:
            region_countries |= set(REGION_PRESETS[part])

    return region_countries

def ordered_country_list(countries: set) -> List[str]:
    """
    Preserve the visual order from SHEET_COUNTRIES.
    """
    return [c for c in SHEET_COUNTRIES if c in countries]


def format_except_suffix(region_val, included_countries: set) -> str:
    """
    Build suffix like:
    (Except France)
    (Except Portugal, Italy, Bulgaria)
    based on region scope minus included countries.
    """
    region_scope = resolve_region_countries(region_val)
    if not region_scope:
        return ""

    excluded_from_scope = region_scope - included_countries
    if not excluded_from_scope:
        return ""

    ordered = ordered_country_list(excluded_from_scope)
    display_names = [COUNTRY_SHORT_NAMES.get(c, c) for c in ordered]

    if not display_names:
        return ""

    return f" (Except {', '.join(display_names)})"

def normalize_overview_region(region_val) -> str:
    """
    Format Region for MONTHLY OVERVIEW display.
    Examples:
    - ['Tier A', 'Tier B', 'Tier C'] -> TIER A, B & C
    - ['Tier A', 'Tier C'] -> TIER A & C
    - Spain Only -> SPAIN ONLY
    """
    if isinstance(region_val, list):
        raw_items = [normalize_str(x) for x in region_val if normalize_str(x)]
    else:
        text = normalize_str(region_val)
        if not text:
            return ""
        text = text.replace("[", "").replace("]", "").replace("'", "").replace('"', "")
        raw_items = [x.strip() for x in re.split(r"[|,/]+", text) if x.strip()]

    if not raw_items:
        return ""

    items_upper = [x.upper() for x in raw_items]

    # Spain Only special case
    if len(items_upper) == 1 and items_upper[0] == "SPAIN ONLY":
        return "SPAIN ONLY"

    # Normalize Tier labels
    tier_letters = []
    other_items = []
    for item in items_upper:
        m = re.match(r"^TIER\s+([ABC])$", item)
        if m:
            tier_letters.append(m.group(1))
        else:
            other_items.append(item)

    tier_letters = list(dict.fromkeys(tier_letters))  # preserve order, remove dupes

    tier_text = ""
    if tier_letters:
        if len(tier_letters) == 1:
            tier_text = f"TIER {tier_letters[0]}"
        elif len(tier_letters) == 2:
            tier_text = f"TIER {tier_letters[0]} & {tier_letters[1]}"
        else:
            tier_text = f"TIER {', '.join(tier_letters[:-1])} & {tier_letters[-1]}"

    if tier_text and other_items:
        return f"{tier_text} / {' / '.join(other_items)}"
    if tier_text:
        return tier_text
    return " / ".join(other_items)
