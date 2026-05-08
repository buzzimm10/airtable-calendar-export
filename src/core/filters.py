import re
from typing import Dict, List, Optional

from config.fields import AT_FIELDS
from core.dates import parse_date
from core.text_formatting import normalize_str


def filter_records_by_brand(records: List[Dict], brand: str) -> List[Dict]:
    target = (brand or "").strip().lower()
    if not target:
        return records

    filtered = []
    for rec in records:
        fields = rec.get("fields", rec)
        value = normalize_str(fields.get(AT_FIELDS["brand"]))

        if value and value.lower() == target:
            filtered.append(rec)

    return filtered


def filter_records_by_month(records: List[Dict], month: Optional[str]) -> List[Dict]:
    if not month:
        return records

    target = month.strip()

    target_year_month = None
    if re.match(r"^\d{4}-\d{2}$", target):
        y, m = target.split("-")
        target_year_month = (int(y), int(m))

    MONTH_NAMES = {
        "january": 1, "february": 2, "march": 3, "april": 4,
        "may": 5, "june": 6, "july": 7, "august": 8,
        "september": 9, "october": 10, "november": 11, "december": 12,
    }

    def parse_month_value(value: str):
        if not value:
            return None, None

        v = value.strip().lower()

        m = re.match(r"^(\d{4})-(\d{1,2})$", v)
        if m:
            return int(m.group(1)), int(m.group(2))

        if v.isdigit():
            n = int(v)
            if 1 <= n <= 12:
                return None, n

        if v in MONTH_NAMES:
            return None, MONTH_NAMES[v]

        v3 = v[:3]
        for name, number in MONTH_NAMES.items():
            if name[:3] == v3:
                return None, number

        return None, None

    target_year, target_month = parse_month_value(target)
    filtered = []

    for rec in records:
        fields = rec.get("fields", rec)

        rec_month_raw = normalize_str(fields.get(AT_FIELDS["month"]))
        rec_year, rec_month = parse_month_value(rec_month_raw)

        live = parse_date(fields.get(AT_FIELDS["live_date"]))
        if live:
            if rec_month is None:
                rec_month = live.month
            if rec_year is None:
                rec_year = live.year

        if target_year_month and live:
            if (live.year, live.month) == target_year_month:
                filtered.append(rec)
                continue

        if rec_month is None:
            continue

        if target_month and rec_month != target_month:
            continue

        if target_year and rec_year != target_year:
            continue

        filtered.append(rec)

    return filtered