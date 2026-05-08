import json
import logging
from typing import Dict, List
import requests


def fetch_airtable_records(api_token, base_id, table_id, view=None, max_records=10000):
    url = f"https://api.airtable.com/v0/{base_id}/{table_id}"
    headers = {"Authorization": f"Bearer {api_token}"}
    params = {}

    if view:
        params["view"] = view

    records = []

    while True:
        r = requests.get(url, headers=headers, params=params, timeout=60)
        if r.status_code >= 400:
            logging.error("Airtable error %s: %s", r.status_code, r.text)
        r.raise_for_status()

        payload = r.json()
        records.extend(payload.get("records", []))

        offset = payload.get("offset")
        if not offset or len(records) >= max_records:
            break

        params["offset"] = offset

    return records


def load_local_json(path: str) -> List[Dict]:
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)

    if isinstance(data, dict) and "records" in data:
        return data["records"]

    if isinstance(data, list):
        return data

    raise ValueError("Unsupported JSON format for Airtable export.")


def fetch_channel_id_to_name_map(api_token, base_id) -> Dict[str, str]:
    channel_map = {}

    try:
        url = f"https://api.airtable.com/v0/meta/bases/{base_id}/tables"
        headers = {"Authorization": f"Bearer {api_token}"}

        r = requests.get(url, headers=headers, timeout=60)
        r.raise_for_status()

        tables = r.json().get("tables", [])
        channels_table_id = None

        for table in tables:
            if table.get("name", "").lower() == "channels":
                channels_table_id = table.get("id")
                break

        if not channels_table_id:
            logging.warning("Channels table not found; channel linking disabled")
            return {}

        url = f"https://api.airtable.com/v0/{base_id}/{channels_table_id}"
        params = {}

        while True:
            r = requests.get(url, headers=headers, params=params, timeout=60)
            r.raise_for_status()

            payload = r.json()

            for record in payload.get("records", []):
                rec_id = record.get("id", "")
                channel_name = record.get("fields", {}).get("Channel Name", "")
                if rec_id and channel_name:
                    channel_map[rec_id] = channel_name

            offset = payload.get("offset")
            if not offset:
                break

            params["offset"] = offset

    except Exception as e:
        logging.warning("Error loading channel map: %s", e)

    return channel_map