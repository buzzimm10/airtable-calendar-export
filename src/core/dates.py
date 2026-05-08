from collections import Counter
from datetime import datetime, date, timedelta
from typing import Dict, List, Optional, Tuple

from config.fields import AT_FIELDS
from core.text_formatting import normalize_str


def parse_date(val) -> Optional[date]:
    if val is None:
        return None

    if isinstance(val, date) and not isinstance(val, datetime):
        return val

    if isinstance(val, datetime):
        return val.date()

    if isinstance(val, str):
        s = val.strip()

        for fmt in ("%Y-%m-%d", "%Y-%m-%dT%H:%M:%S.%fZ", "%Y-%m-%dT%H:%M:%SZ"):
            try:
                return datetime.strptime(s, fmt).date()
            except Exception:
                pass

        try:
            return datetime.strptime(s, "%A, %B %d, %Y").date()
        except Exception:
            pass

        try:
            return date.fromisoformat(s)
        except Exception:
            return None

    return None


def monday_of_week(d: date) -> date:
    return d - timedelta(days=d.weekday())


def resolve_target_month_year(records: List[Dict]) -> Tuple[Optional[int], Optional[int]]:
    counts = Counter()

    for rec in records:
        fields = rec.get("fields", rec)
        live = parse_date(fields.get(AT_FIELDS["live_date"]))
        if live:
            counts[(live.year, live.month)] += 1

    if not counts:
        return None, None

    return counts.most_common(1)[0][0]


def first_visible_monday_of_target_month(records: List[Dict]) -> Optional[date]:
    year, month = resolve_target_month_year(records)
    if not year or not month:
        return None

    first_day = date(year, month, 1)
    return monday_of_week(first_day)


def first_visible_monday(records: List[Dict]) -> Optional[date]:
    return first_visible_monday_of_target_month(records)


def full_calendar_anchor_dates_from_records(records: List[Dict], max_weeks: int = 5) -> Dict[str, date]:
    anchor_cells = ["D6", "D59", "D112", "D165", "D218"]

    first_monday = first_visible_monday_of_target_month(records)
    if not first_monday:
        return {}

    anchor_map = {}
    for i, cell in enumerate(anchor_cells[:max_weeks]):
        anchor_map[cell] = first_monday + timedelta(days=i * 7)

    return anchor_map


def month_name_from_records(records: List[Dict]) -> str:
    month_lookup = {
        1: "JANUARY", 2: "FEBRUARY", 3: "MARCH", 4: "APRIL",
        5: "MAY", 6: "JUNE", 7: "JULY", 8: "AUGUST",
        9: "SEPTEMBER", 10: "OCTOBER", 11: "NOVEMBER", 12: "DECEMBER",
    }

    for rec in records:
        fields = rec.get("fields", rec)

        raw_month = normalize_str(fields.get(AT_FIELDS["month"]))
        if raw_month:
            m = raw_month.strip().lower()

            # YYYY-MM
            if "-" in m:
                try:
                    month_number = int(m.split("-")[1])
                    if 1 <= month_number <= 12:
                        return month_lookup[month_number]
                except Exception:
                    pass

            names = {v.lower(): k for k, v in month_lookup.items()}
            if m in names:
                return month_lookup[names[m]]

            if m.isdigit():
                n = int(m)
                if 1 <= n <= 12:
                    return month_lookup[n]

        live = parse_date(fields.get(AT_FIELDS["live_date"]))
        if live:
            return month_lookup[live.month]

    return "MONTH"