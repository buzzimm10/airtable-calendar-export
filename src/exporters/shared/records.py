from collections import defaultdict
from typing import Dict, List, Tuple
from datetime import date

from config.fields import AT_FIELDS
from config.layouts import DAY_ORDER
from core.dates import parse_date, monday_of_week
from services.planner_service import asset_is_event, hub_role


def aggregate_by_visual_week_day_role(records: List[Dict], first_monday: date):
    buckets = defaultdict(list)

    for rec in records:
        fields = rec.get("fields", rec)

        if asset_is_event(fields):
            continue

        live_date = parse_date(fields.get(AT_FIELDS["live_date"]))
        if not live_date:
            continue

        delta_days = (live_date - first_monday).days
        if delta_days < 0:
            continue

        week_index = (delta_days // 7) + 1
        
        # --- EL FIX ESTÁ AQUÍ: Ahora permite hasta la semana 6 ---
        if week_index < 1 or week_index > 6:
            continue

        day_name = DAY_ORDER[live_date.weekday()]
        role = hub_role(fields.get(AT_FIELDS["owner_hub"]))

        if role not in {"Programming", "Mediaplanning"}:
            continue

        buckets[(week_index, day_name, role)].append(fields)

    return buckets