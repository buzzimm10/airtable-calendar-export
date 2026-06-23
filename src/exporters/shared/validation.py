from collections import defaultdict
from dataclasses import fields
from config.fields import AT_FIELDS
from config.layouts import DAY_ORDER
from core.dates import parse_date
from services.planner_service import asset_is_event, hub_role

def get_record_label(fields):
    return (
        fields.get(AT_FIELDS.get("activity", "Activity"))
        or fields.get(AT_FIELDS["youtube_title"])
        or fields.get(AT_FIELDS["property"])
        or "Untitled"
    )

def run_basic_qa(records, max_assets_per_hub_per_day=2):
    issues = []

    grouped = defaultdict(list)

    for rec in records:
        fields = rec.get("fields", rec)

        if asset_is_event(fields):
            continue

        live_date = parse_date(fields.get(AT_FIELDS["live_date"]))
        if not live_date:
            issues.append("Missing or invalid Live Date")
            continue

        role = hub_role(fields.get(AT_FIELDS["owner_hub"]))
        day_name = DAY_ORDER[live_date.weekday()]

        grouped[(live_date, day_name, role)].append(fields)

        if not fields.get(AT_FIELDS["deadline_delivery"]):
            title = fields.get(AT_FIELDS["youtube_title"], "Untitled")
            issues.append(f"Missing deadline: {live_date} | {title}")

        if not fields.get(AT_FIELDS["region"]):
            title = fields.get(AT_FIELDS["youtube_title"], "Untitled")
            issues.append(f"Missing region: {live_date} | {title}")

#        if not fields.get(AT_FIELDS["youtube_channel"]):
#            title = fields.get(AT_FIELDS["youtube_title"], "Untitled")
#            issues.append(f"Missing YouTube Channel: {live_date} | {title}")

        if not fields.get(AT_FIELDS["content_type"]):
            issues.append(f"Missing Content Type: {live_date}")

        if not fields.get(AT_FIELDS["owner_hub"]):
            issues.append(f"Missing Owner Hub: {live_date}")


    for key, items in grouped.items():
        if len(items) > max_assets_per_hub_per_day:
            issues.append(f"Too many assets for {key}: {len(items)}")

    return issues