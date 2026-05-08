from datetime import date
from typing import Dict, List, Optional

from config.fields import AT_FIELDS
from core.dates import parse_date
from core.text_formatting import normalize_str, lower_clean, bullet_join, tc_in_out_str
from core.countries import format_excluded_markets_comment
from core.text_formatting import join_comments
import logging
import os
import logging
from collections import defaultdict
from datetime import date
from typing import Dict, List, Optional, Tuple

from dotenv import load_dotenv

from core.airtable_client import (
    fetch_airtable_records,
    fetch_channel_id_to_name_map,
    load_local_json,
)

from core.text_formatting import (
    normalize_str,
    lower_clean,
    bullet_join,
    tc_in_out_str,
    join_comments,
)

from core.countries import (
    format_excluded_markets_comment,
)

from config.fields import AT_FIELDS
from config.layouts import DAY_ORDER

from core.dates import (
    parse_date,
    monday_of_week,
)   

from core.filters import filter_records_by_brand, filter_records_by_month

class WeekBlock:
    def __init__(self, anchor_row: int, end_row: int, monday_date: date):
        self.anchor_row = anchor_row
        self.end_row = end_row
        self.monday_date = monday_date
        self.deadline_row: Optional[int] = None
        self.owner_row: Optional[int] = None
        self.property_row: Optional[int] = None
        self.content_type_row: Optional[int] = None
        self.title_row: Optional[int] = None
        self.synopsis_row: Optional[int] = None
        self.strategy_row: Optional[int] = None
        self.tx_row: Optional[int] = None
        self.tc_row: Optional[int] = None
        self.comments_row: Optional[int] = None
        self.rollout_country_rows: Dict[str, int] = {}
        self.status_country_rows: Dict[str, int] = {}

def hub_role(owner_hub: str) -> Optional[str]:
    h = lower_clean(owner_hub)
    if "uk" in h:
        return None
    if "media" in h or "creative" in h:
        return "Mediaplanning"
    return "Programming"

def asset_is_event(asset: Dict) -> bool:
    return lower_clean(asset.get(AT_FIELDS["content_type"])) == "event"


def earliest_deadline_for_day(records: List[Dict]) -> Optional[date]:
    deadlines = []
    for f in records:
        d = parse_date(f.get(AT_FIELDS["deadline_delivery"]))
        if d:
            deadlines.append(d)
    return min(deadlines) if deadlines else None

def load_and_filter_records(
    month: Optional[str],
    channel: str = "Disney Channel",
    json_path: Optional[str] = None,
    log_level: str = "INFO",
):
    load_dotenv()

    logging.basicConfig(
        level=getattr(logging, log_level.upper(), logging.INFO),
        format="%(levelname)s: %(message)s",
    )

    token = os.getenv("AIRTABLE_API_TOKEN")
    base_id = os.getenv("AIRTABLE_BASE_ID")
    table_id = os.getenv("AIRTABLE_TABLE_ID")
    view = os.getenv("AIRTABLE_VIEW")

    channel_id_to_name_map = {}

    if json_path:
        records = load_local_json(json_path)

        if token and base_id:
            channel_id_to_name_map = fetch_channel_id_to_name_map(token, base_id)
    else:
        if not token or not base_id or not table_id:
            raise RuntimeError("Missing Airtable config in .env")

        records = fetch_airtable_records(
            api_token=token,
            base_id=base_id,
            table_id=table_id,
            view=view,
        )

        channel_id_to_name_map = fetch_channel_id_to_name_map(token, base_id)

    logging.info(f"Fetched {len(records)} records from Airtable/JSON source")

    records = filter_records_by_brand(records, channel)
    logging.info(f"After brand filter '{channel}': {len(records)} records")

    records = filter_records_by_month(records, month)
    logging.info(f"After month filter '{month}': {len(records)} records")

    return records, channel_id_to_name_map

def build_comments_for_excel(record: Dict) -> str:
    notes = normalize_str(record.get(AT_FIELDS["notes"]))
    geof = normalize_str(record.get(AT_FIELDS["geofenced"]))

    excluded_comment = format_excluded_markets_comment(
        record.get(AT_FIELDS["excluded_markets"])
    )

    geofenced_comment = "Geofenced" if geof.lower() == "yes" else ""

    return join_comments(
        notes,
        excluded_comment,
        geofenced_comment,
    )

def assemble_text_for_single_asset(f: Dict) -> Dict[str, str]:
    ct = normalize_str(f.get(AT_FIELDS["content_type"]))
    if ct == "Exclusive":
        ct = "Exclusive Short"

    owner_role = hub_role(f.get(AT_FIELDS["owner_hub"]))

    prop_val = f.get(AT_FIELDS["property"])
    if isinstance(prop_val, list):
        prop = " / ".join(normalize_str(p) for p in prop_val if normalize_str(p))
    else:
        prop = normalize_str(prop_val)

    title = normalize_str(f.get(AT_FIELDS["youtube_title"]))
    desc = normalize_str(f.get(AT_FIELDS["youtube_description"]))

    strat_val = f.get(AT_FIELDS["strategy"])
    if isinstance(strat_val, list):
        strat = " / ".join(normalize_str(s) for s in strat_val if normalize_str(s))
    else:
        strat = normalize_str(strat_val)

    tx_id = normalize_str(f.get(AT_FIELDS["tx_id"])) or "—"
    tc_text = tc_in_out_str(f.get(AT_FIELDS["tc_in"]), f.get(AT_FIELDS["tc_out"]))
    comments = build_comments_for_excel(f)

    if owner_role != "Programming" or ct == "Vertical Short":
        ct = ct.upper()
        prop = prop.upper()

    if owner_role != "Programming":
        strat = ""

    return {
        "content_type": ct,
        "property": prop,
        "title": title,
        "synopsis": desc,
        "strategy": strat,
        "tx": tx_id,
        "tc": tc_text,
        "comments": comments,
    }

def assemble_texts_for_multi_assets(records: List[Dict], max_assets=2) -> Dict[str, str]:
    items = records[:max_assets]
    content_types, properties, titles, descs, strategies, txs, tcs, comments = [], [], [], [], [], [], [], []
    for f in items:
        single = assemble_text_for_single_asset(f)
        content_types.append(single["content_type"])
        properties.append(single["property"])
        titles.append(single["title"])
        descs.append(single["synopsis"])
        strategies.append(single["strategy"])
        txs.append(single["tx"])
        tcs.append(single["tc"])
        comments.append(single["comments"])
    return {
        "content_type": bullet_join(content_types),
        "property": bullet_join(properties),
        "title": bullet_join(titles),
        "synopsis": bullet_join(descs),
        "strategy": bullet_join(strategies),
        "tx": bullet_join(txs),
        "tc": bullet_join(tcs),
        "comments": bullet_join(comments),
    }

PRIORITY_ORDER = {
    "Current A": 1,
    "Current B": 2,
    "Current C": 3,
    "Tactical A": 4,
    "Tactical B": 5,
    "Evergreen": 6,
    "Tactical Opportunity": 7
}

def get_priority_sort_value(priority):
    if not priority:
        return 99
    return PRIORITY_ORDER.get(priority.strip(), 99)

def aggregate_by_week_day_and_role(records: List[Dict]) -> Dict[Tuple[date, str, str], List[Dict]]:
    buckets: Dict[Tuple[date, str, str], List[Dict]] = defaultdict(list)
    for rec in records:
        fields = rec.get("fields", rec)
        if asset_is_event(fields):
            continue
        live_date = parse_date(fields.get(AT_FIELDS["live_date"]))
        if not live_date:
            continue
        dow_name = (fields.get(AT_FIELDS["day_of_week"]) or "").strip()
        if dow_name not in DAY_ORDER:
            dow_name = DAY_ORDER[live_date.weekday()]
        role = hub_role(fields.get(AT_FIELDS["owner_hub"]))
        if role is None:
            continue
        mon = monday_of_week(live_date)
        key = (mon, dow_name, role)
        buckets[key].append(fields)

    def sort_key(f):
        priority = f.get(AT_FIELDS["priority"])
        return (
            get_priority_sort_value(priority),
            normalize_str(f.get(AT_FIELDS["asset_label"])),
            normalize_str(f.get(AT_FIELDS["youtube_title"])),
        )

    for k in buckets:
        buckets[k].sort(key=sort_key)
    return buckets
