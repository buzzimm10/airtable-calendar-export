#!/usr/bin/env python3
"""
Populate the 'FULL CALENDAR YT DC EMEA' sheet of an existing Excel template
from Airtable (or local JSON) using openpyxl.

PATCHES (2026-03-25):
- Week label: write "WEEK <Week Number (Year)>" into merged cells B6, B59, B112, B165, B218 (per week block).
  Value inferred from Airtable records of that week; fallback to ISO week from Monday date.
- Countries by Hub: keep writing YES/NO in the correct hub column; uniform font
  (YES black, NO red; Italy always red).
- All prior behavior preserved: channel/month filters, merged-cell safe writing,
  dual Programming Fri/Sun, optional Mediaplanning, UK blank, real date headers,
  workbook/template preservation, STATUS CONFORMING mirroring with fallback.
"""

import argparse
import os
import re
import sys
import json
import logging
import math
from collections import defaultdict, Counter
from datetime import datetime, date, timedelta
from typing import Dict, List, Optional, Tuple

import requests
from copy import copy
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.cell.cell import Cell, MergedCell

from dotenv import load_dotenv
load_dotenv()  # carga .env del cwd (raíz del proyecto)

TARGET_SHEET = "FULL CALENDAR YT DC EMEA"

DAY_LAYOUT_LETTERS = {
    "Monday":    [("Programming", "D"), ("Programming", "E"), ("Mediaplanning", "F"), ("UK", "G")],
    "Tuesday":   [("Programming", "H"), ("Programming", "I"), ("Mediaplanning", "J"), ("UK", "K")],
    "Wednesday": [("Programming", "L"), ("Programming", "M"), ("Mediaplanning", "N"), ("UK", "O")],
    "Thursday":  [("Programming", "P"), ("Programming", "Q"), ("Mediaplanning", "R"), ("UK", "S")],
    "Friday":    [("Programming", "T"), ("Programming", "U"), ("Mediaplanning", "V"), ("UK", "W")],
    "Saturday":  [("Programming", "X"), ("Programming", "Y"), ("Mediaplanning", "Z"), ("UK", "AA")],
    "Sunday":    [("Programming", "AB"), ("Programming", "AC"), ("Mediaplanning", "AD"), ("UK", "AE")],
}
DAY_ORDER = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]

def _to_indices(layout_letters):
    d = {}
    for day, cols in layout_letters.items():
        d[day] = [(role, column_index_from_string(letter)) for role, letter in cols]
    return d

DAY_LAYOUT = _to_indices(DAY_LAYOUT_LETTERS)

BASE_DAY_TEMPLATE = {
    "Monday": {"programming": 1, "mediaplanning": 1, "uk": 1},
    "Tuesday": {"programming": 1, "mediaplanning": 1, "uk": 1},
    "Wednesday": {"programming": 1, "mediaplanning": 1, "uk": 1},
    "Thursday": {"programming": 1, "mediaplanning": 1, "uk": 1},
    "Friday": {"programming": 2, "mediaplanning": 1, "uk": 1},
    "Saturday": {"programming": 1, "mediaplanning": 1, "uk": 1},
    "Sunday": {"programming": 2, "mediaplanning": 1, "uk": 1},
}

# Country YES/NO row ranges per visible week in FULL CALENDAR
FULL_CALENDAR_COUNTRY_ROW_RANGES = {
    1: (14, 30),
    2: (67, 83),
    3: (120, 136),
    4: (173, 189),
    5: (226, 242),
}

# Formula column for weekly counts
FULL_CALENDAR_WEEKLY_COUNT_COL = 33  # AB

AT_FIELDS = {
    "property": "Property",
    "content_type": "Content Type",
    "live_date": "Live Date",
    "youtube_title": "YouTube Title",
    "youtube_description": "YouTube Description",
    "notes": "Notes",
    "tx_id": "TX ID / LINK",
    "tc_in": "TC IN",
    "tc_out": "TC OUT",
    "strategy": "Strategy",
    "region": "Region",
    "youtube_channel": "YouTube Channel",
    "geofenced": "GeoFenced?",
    "owner_hub": "Owner Hub",
    "day_of_week": "Day of Week",
    "priority": "Priority",
    "asset_label": "Asset Label",
    "activity": "Activity",
    "week_number": "Week Number (Year)",
    "month": "Month",
    "week_of_month": "Week of Month",
    "deadline_delivery": "Deadline Delivery",
    "channel": "Channel",
    "title": "YouTube Title",
}

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
    "a": {"Spain", "France (French)", "Poland (Polish)"},
    "b": {
        "Portugal", "Italy (Italian)", "Romania (Romanian)", "Greece (Greek)",
        "Middle East (Arabic - MSA)", "Netherlands (Dutch)", "Belgium (French)",
        "Israel (Hebrew)", "Bulgaria (Bulgarian)"
    },
    "c": {"Hungary (Hungarian)", "Czech R. (Czech)", "Sweden (Swedish)",
                "Denmark (Danish)", "Norway (Norwegian)"},
    "group a": {"Spain", "France (French)", "Poland (Polish)"},
    "group b": {
        "Portugal", "Italy (Italian)", "Romania (Romanian)", "Greece (Greek)",
        "Middle East (Arabic - MSA)", "Netherlands (Dutch)", "Belgium (French)",
        "Israel (Hebrew)", "Bulgaria (Bulgarian)"
    },
    "group c": {"Hungary (Hungarian)", "Czech R. (Czech)", "Sweden (Swedish)",
                "Denmark (Danish)", "Norway (Norwegian)"},
    "mena": {"Middle East (Arabic - MSA)"},
    "nordics": {"Sweden (Swedish)", "Denmark (Danish)", "Norway (Norwegian)"},
    "tier a": {"Spain", "France (French)", "Poland (Polish)"},
    "tier b": {
        "Portugal", "Italy (Italian)", "Romania (Romanian)", "Greece (Greek)",
        "Middle East (Arabic - MSA)", "Netherlands (Dutch)", "Belgium (French)",
        "Israel (Hebrew)", "Bulgaria (Bulgarian)"
    },
    "tier c": {"Hungary (Hungarian)", "Czech R. (Czech)", "Sweden (Swedish)",
                "Denmark (Danish)", "Norway (Norwegian)"},
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

MAX_ASSETS_PER_HUB_PER_DAY = 2
PRIORITY_ORDER = {
    "Current A": 1,
    "Current B": 2,
    "Current C": 3,
    "Tactical A": 4,
    "Tactical B": 5,
    "Evergreen": 6,
    "Tactical Opportunity": 7
}

MONTHLY_OVERVIEW_SHEET = "MONTHLY OVERVIEW YT DC EMEA"

OVERVIEW_DAY_COLS = {
    "Monday": 4,      # D
    "Tuesday": 5,     # E
    "Wednesday": 6,   # F
    "Thursday": 7,    # G
    "Friday": 8,      # H
    "Saturday": 9,    # I
    "Sunday": 10,     # J
}

OVERVIEW_WEEK_LAYOUT = {
    1: {"header_row": 4, "date_row": 5, "event_rows": (6, 7), "base_row": 8, "end_row": 24},
    2: {"header_row": 26, "date_row": 27, "event_rows": (28, 29), "base_row": 30, "end_row": 46},
    3: {"header_row": 48, "date_row": 49, "event_rows": (50, 51), "base_row": 52, "end_row": 68},
    4: {"header_row": 70, "date_row": 71, "event_rows": (72, 73), "base_row": 74, "end_row": 90},
    5: {"header_row": 92, "date_row": 93, "event_rows": (94, 95), "base_row": 96, "end_row": 112},
}

OVERVIEW_FILL_COLORS = {
    "Full Episode": "FFFF00",
    "Exclusive": "BDD7EE",
    "Lift": "A9D08E",
    "Musical Lift": "A9D08E",
    "Podcast": "92D050",
    "Magical Moment": "92D050",
    "Interactive": "92D050",
    "Editorial Short": "92D050",
    "Compilation": "F4B084",
    "Live Stream": "00B0F0",
    "Vertical Short": "FFB9D5",
}

OVERVIEW_DAY_LABELS = {
    "Monday": "MONDAY",
    "Tuesday": "TUESDAY",
    "Wednesday": "WEDNESDAY",
    "Thursday": "THURSDAY",
    "Friday": "FRIDAY",
    "Saturday": "SATURDAY",
    "Sunday": "SUNDAY",
}

OUT_OF_MONTH_FILL = "BFBFBF"

def get_priority_sort_value(priority):
    if not priority:
        return 99
    return PRIORITY_ORDER.get(priority.strip(), 99)

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

def normalize_str(s: Optional[str]) -> str:
    return str(s or "").strip()

def lower_clean(s: Optional[str]) -> str:
    return normalize_str(s).lower()

def tc_in_out_str(tc_in: Optional[str], tc_out: Optional[str]) -> str:
    tin = normalize_str(tc_in)
    tout = normalize_str(tc_out)

    tin_lines = [line.strip() for line in tin.splitlines() if line.strip()]
    tout_lines = [line.strip() for line in tout.splitlines() if line.strip()]

    # If either side has more than 1 line, use verbose paired format
    if len(tin_lines) > 1 or len(tout_lines) > 1:
        max_len = max(len(tin_lines), len(tout_lines))
        pairs = []
        for i in range(max_len):
            in_val = tin_lines[i] if i < len(tin_lines) else "—"
            out_val = tout_lines[i] if i < len(tout_lines) else "—"
            pairs.append(f"TC IN {in_val} / TC OUT {out_val}")
        return "\n".join(pairs)

    # Default simple 2-line format
    tin_val = tin_lines[0] if tin_lines else "—"
    tout_val = tout_lines[0] if tout_lines else "—"
    return f"{tin_val}\n{tout_val}"

def bullet_join(items: List[str]) -> str:
    cleaned = [normalize_str(x) for x in items if normalize_str(x)]
    return "\n".join(cleaned)

def get_writable_cell(ws, row: int, col: int, unmerge_if_needed: bool = False):
    """
    Return a writable Cell for a target position.
    If the position is inside a merged range, return the top-left cell.
    If the top-left is somehow still a MergedCell (can happen after insert_cols),
    optionally unmerge that range and return a normal writable cell.
    """
    for cr in list(ws.merged_cells.ranges):
        if cr.min_row <= row <= cr.max_row and cr.min_col <= col <= cr.max_col:
            top_left = ws.cell(row=cr.min_row, column=cr.min_col)

            if isinstance(top_left, MergedCell):
                if unmerge_if_needed:
                    ws.unmerge_cells(str(cr))
                    return ws.cell(row=cr.min_row, column=cr.min_col)
                raise RuntimeError(
                    f"Top-left of merged range {cr} is still a MergedCell at "
                    f"({cr.min_row}, {cr.min_col})"
                )

            return top_left

    cell = ws.cell(row=row, column=col)

    if isinstance(cell, MergedCell):
        if unmerge_if_needed:
            for cr in list(ws.merged_cells.ranges):
                if cr.min_row <= row <= cr.max_row and cr.min_col <= col <= cr.max_col:
                    ws.unmerge_cells(str(cr))
                    return ws.cell(row=cr.min_row, column=cr.min_col)
        raise RuntimeError(f"Cell {get_column_letter(col)}{row} is a MergedCell and could not be resolved.")

    return cell


def ensure_writable_cell(ws, row: int, col: int):
    """Return a real writable cell after unmerging any merged range covering the target."""
    merged_ranges = [cr for cr in list(ws.merged_cells.ranges)
                     if cr.min_row <= row <= cr.max_row and cr.min_col <= col <= cr.max_col]
    logging.debug(
        "ensure_writable_cell: target %s%d merged_ranges=%s",
        get_column_letter(col), row, merged_ranges,
    )
    for cr in merged_ranges:
        logging.debug(
            "ensure_writable_cell: target %s%d in merged range %s, unmerging",
            get_column_letter(col), row, cr,
        )
        ws.unmerge_cells(str(cr))

    cell = ws.cell(row=row, column=col)
    logging.debug(
        "ensure_writable_cell: after unmerge target %s%d cell_type=%s",
        get_column_letter(col), row, type(cell),
    )
    if isinstance(cell, MergedCell):
        logging.debug(
            "ensure_writable_cell: still MergedCell at %s%d after unmerge, returning top-left of active merged range",
            get_column_letter(col), row,
        )
        for cr in list(ws.merged_cells.ranges):
            if cr.min_row <= row <= cr.max_row and cr.min_col <= col <= cr.max_col:
                top_left = ws.cell(row=cr.min_row, column=cr.min_col)
                if isinstance(top_left, MergedCell):
                    logging.debug(
                        "ensure_writable_cell: top-left still MergedCell for range %s, unmerging it",
                        cr,
                    )
                    ws.unmerge_cells(str(cr))
                    return ws.cell(row=cr.min_row, column=cr.min_col)
                logging.debug(
                    "ensure_writable_cell: using top-left %s%d for merged range %s",
                    get_column_letter(cr.min_col), cr.min_row, cr,
                )
                return top_left

        logging.debug(
            "ensure_writable_cell: target %s%d is a stale MergedCell with no active merged range, replacing with fresh Cell",
            get_column_letter(col), row,
        )
        real_cell = Cell(ws, row=row, column=col)
        ws._cells[(row, col)] = real_cell
        return real_cell

    return cell

def cleanup_stale_merged_cells(ws, row_start: int = 1, row_end: Optional[int] = None, col_start: int = 1, col_end: Optional[int] = None):
    """
    Replace stale MergedCell objects that are no longer covered by an active merged range
    with normal writable Cell objects.
    """
    if row_end is None:
        row_end = ws.max_row
    if col_end is None:
        col_end = ws.max_column

    active_ranges = list(ws.merged_cells.ranges)

    def is_inside_active_merge(r: int, c: int) -> bool:
        for cr in active_ranges:
            if cr.min_row <= r <= cr.max_row and cr.min_col <= c <= cr.max_col:
                return True
        return False

    replaced = 0

    for row in range(row_start, row_end + 1):
        for col in range(col_start, col_end + 1):
            key = (row, col)
            cell_obj = ws._cells.get(key)

            if isinstance(cell_obj, MergedCell) and not is_inside_active_merge(row, col):
                ws._cells[key] = Cell(ws, row=row, column=col)
                replaced += 1

    if replaced:
        logging.info(
            "cleanup_stale_merged_cells: replaced %s stale merged cells in range %s:%s to %s:%s",
            replaced,
            get_column_letter(col_start), row_start,
            get_column_letter(col_end), row_end,
        )

def safe_set_value(ws, row: int, col: int, value, alignment=None, number_format=None, font=None, fill=None):
    cell = ensure_writable_cell(ws, row, col)
    cell.value = value
    if alignment is not None:
        cell.alignment = alignment
    if number_format is not None:
        cell.number_format = number_format
    if font is not None:
        cell.font = font
    if fill is not None:
        cell.fill = fill
    return cell


def force_unmerged_writable_cell(ws, row: int, col: int):
    """
    Unmerge any merged range that contains the target cell, then return a normal writable cell.
    This is intentionally aggressive and is meant for dynamic header/deadline cells after insert_cols().
    """
    return ensure_writable_cell(ws, row, col)

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
            logging.error(f"Airtable error {r.status_code}: {r.text}")
        r.raise_for_status()
        payload = r.json()
        batch = payload.get("records", [])
        records.extend(batch)
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
    """
    Fetch the Channels table and create a mapping of record IDs to Channel Names.
    e.g., {'rec1pxj5TXpateafv': 'Disney Channel Portugal', ...}
    """
    channel_map = {}
    try:
        # First, get the Channels table ID from metadata
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
        
        # Now fetch all channel records
        url = f"https://api.airtable.com/v0/{base_id}/{channels_table_id}"
        params = {}
        while True:
            r = requests.get(url, headers=headers, params=params, timeout=60)
            if r.status_code >= 400:
                logging.error(f"Airtable channels fetch error {r.status_code}: {r.text}")
                break
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
        
        logging.debug(f"Loaded {len(channel_map)} channel ID→Name mappings")
    except Exception as e:
        logging.warning(f"Error loading channel map: {e}")
    
    return channel_map

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

    def __repr__(self):
        return f"<WeekBlock row={self.anchor_row} monday={self.monday_date} end={self.end_row}>"

def cell_value(ws, row: int, col: int):
    return ws.cell(row=row, column=col).value

def get_anchor_dates_from_values_wb(template_path: str, anchors: List[str]) -> Dict[str, date]:
    wb_vals = load_workbook(template_path, data_only=True)
    if TARGET_SHEET not in wb_vals.sheetnames:
        raise RuntimeError(f"Target sheet '{TARGET_SHEET}' not found for reading anchors.")
    ws_vals = wb_vals[TARGET_SHEET]
    out = {}
    for addr in anchors:
        val = ws_vals[addr].value
        d = parse_date(val)
        if not d:
            logging.warning(
                f"Anchor {addr} value unreadable (got {val}). Open template in Excel to recalc/save, or pass --anchor-dates."
            )
        else:
            out[addr] = d
    return out

def find_week_blocks(ws_write, anchor_dates: Dict[str, date]) -> Dict[date, WeekBlock]:
    anchors = []
    for addr, mon_date in anchor_dates.items():
        cell = ws_write[addr]
        anchors.append((cell.row, mon_date, addr))
    anchors = [a for a in anchors if a[1] is not None]
    if not anchors:
        raise RuntimeError("No valid week anchors resolved.")
    anchors_sorted = sorted(anchors, key=lambda x: x[0])

    blocks: Dict[date, WeekBlock] = {}
    for i, (start_row, mon_date, _addr) in enumerate(anchors_sorted):
        end_row = ws_write.max_row if i == len(anchors_sorted) - 1 else anchors_sorted[i + 1][0] - 1
        blocks[mon_date] = WeekBlock(start_row, end_row, mon_date)

    for mon_date, block in blocks.items():
        in_status_conforming = False
        for r in range(block.anchor_row, block.end_row + 1):
            a1 = lower_clean(cell_value(ws_write, r, 1))
            b2 = normalize_str(cell_value(ws_write, r, 2))
            c3 = normalize_str(cell_value(ws_write, r, 3))
            b2_lower = b2.lower()
            c3_lower = c3.lower()

            # Status section starts at "status conforming" and ends at "comments" (in this template layout)
            if a1 == "status conforming" or b2_lower == "status conforming":
                in_status_conforming = True
            elif c3_lower == "comments" or a1 == "comments" or b2_lower == "comments":
                in_status_conforming = False

            if a1 == "owner" and block.owner_row is None:
                block.owner_row = r
            elif a1 == "property" and block.property_row is None:
                block.property_row = r
            elif a1 == "content type" and block.content_type_row is None:
                block.content_type_row = r
            elif a1 == "title" and block.title_row is None:
                block.title_row = r
            elif a1 == "synopsis" and block.synopsis_row is None:
                block.synopsis_row = r
            elif a1 == "planning strategy" and block.strategy_row is None:
                block.strategy_row = r
            elif a1 == "tx id / link" and block.tx_row is None:
                block.tx_row = r
            elif a1 == "tc in & out" and block.tc_row is None:
                block.tc_row = r
            elif a1 == "comments" and block.comments_row is None:
                block.comments_row = r
            elif a1 == "deadline" and block.deadline_row is None:
                block.deadline_row = r

            if b2 in ("Tier A", "Tier B", "Tier C") and c3 in SHEET_COUNTRIES:
                if in_status_conforming:
                    block.status_country_rows.setdefault(c3, r)
                else:
                    block.rollout_country_rows.setdefault(c3, r)
            elif in_status_conforming and c3 in SHEET_COUNTRIES and b2 not in ("Tier A", "Tier B", "Tier C"):
                # In status conforming section, countries without Tier label
                block.status_country_rows.setdefault(c3, r)

        # Fallback partial match if some rows were not found
        needed = {
            "owner_row": "owner",
            "property_row": "property",
            "content_type_row": "content type",
            "title_row": "title",
            "synopsis_row": "synopsis",
            "strategy_row": "planning strategy",
            "tx_row": "tx id",
            "tc_row": "tc in & out",
            "comments_row": "comments",
            "deadline_row": "deadline",
        }
        if any(getattr(block, k) is None for k in needed):
            for r in range(block.anchor_row, block.end_row + 1):
                row_text = " ".join([
                    lower_clean(cell_value(ws_write, r, 1)),
                    lower_clean(cell_value(ws_write, r, 2)),
                    lower_clean(cell_value(ws_write, r, 3)),
                ])
                for attr, needle in needed.items():
                    if getattr(block, attr) is None and needle in row_text:
                        setattr(block, attr, r)

    return blocks

def hub_role(owner_hub: str) -> Optional[str]:
    h = lower_clean(owner_hub)
    if "uk" in h:
        return None
    if "media" in h or "creative" in h:  # Creative Services treated as Mediaplanning column
        return "Mediaplanning"
    if "programming" in h or not h:
        return "Programming"
    return "Programming"

def asset_is_event(asset: Dict) -> bool:
    return lower_clean(asset.get(AT_FIELDS["content_type"])) == "event"

def collect_included_countries(youtube_channel, region, notes, channel_id_to_name_map=None):
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
    excluded = extract_excluded_countries_from_notes(notes)
    included -= excluded

    return included, excluded

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

def earliest_deadline_for_day(records: List[Dict]) -> Optional[date]:
    deadlines = []
    for f in records:
        d = parse_date(f.get(AT_FIELDS["deadline_delivery"]))
        if d:
            deadlines.append(d)
    if not deadlines:
        return None
    return min(deadlines)

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

    notes = normalize_str(f.get(AT_FIELDS["notes"]))
    geof = normalize_str(f.get(AT_FIELDS["geofenced"]))
    tx_id = normalize_str(f.get(AT_FIELDS["tx_id"])) or "—"
    tc_text = tc_in_out_str(f.get(AT_FIELDS["tc_in"]), f.get(AT_FIELDS["tc_out"]))
    comments = notes
    if geof.lower() == "yes":
        comments = (comments + " | " if comments else "") + "Geofenced"

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

def rebuild_full_calendar_week_merges(ws, block: WeekBlock):
    """
    Rebuild merges for:
    - WEEK # row: B:C
    - Live Date row: merged across each day group
    - Deadline row: merged across each day group
    """

    # 1) limpiar merges horizontales existentes dentro de las filas relevantes
    rows_to_clean = {block.anchor_row}
    if block.deadline_row:
        rows_to_clean.add(block.deadline_row)

    for cr in list(ws.merged_cells.ranges):
        if (
            cr.min_row == cr.max_row
            and (
                cr.min_row in rows_to_clean
                or (cr.min_row == block.anchor_row)
            )
        ):
            ws.unmerge_cells(str(cr))

    # 2) WEEK # -> B:C
    ws.merge_cells(
        start_row=block.anchor_row,
        start_column=2,
        end_row=block.anchor_row,
        end_column=3,
    )

    # 3) Live Date y Deadline por cada día
    for day_name in DAY_ORDER:
        day_cols = [col for _, col in DAY_LAYOUT[day_name]]
        start_col = min(day_cols)
        end_col = max(day_cols)

        # Live Date
        if end_col > start_col:
            ws.merge_cells(
                start_row=block.anchor_row,
                start_column=start_col,
                end_row=block.anchor_row,
                end_column=end_col,
            )

        # Deadline
        if block.deadline_row and end_col > start_col:
            ws.merge_cells(
                start_row=block.deadline_row,
                start_column=start_col,
                end_row=block.deadline_row,
                end_column=end_col,
            )

def write_to_day_role_columns(ws, block, day_name, role, records, enable_mediaplanning):
    if role == "Mediaplanning" and not enable_mediaplanning:
        return

    role_cols = [col for r, col in DAY_LAYOUT[day_name] if r == role]
    uk_cols = [col for r, col in DAY_LAYOUT[day_name] if r == "UK"]

    def _write(row_idx, col_idx, value):
        if not row_idx:
            return
        safe_set_value(
            ws,
            row_idx,
            col_idx,
            value,
            alignment=Alignment(horizontal='center', vertical='center', wrap_text=True),
        )

    if role == "Programming":
        # máximo 2 contenidos, uno por columna
        for idx, record in enumerate(records[:2]):
            col = role_cols[idx]
            single = assemble_text_for_single_asset(record)

            _write(block.content_type_row, col, single["content_type"])
            _write(block.property_row, col, single["property"])
            _write(block.title_row, col, single["title"])
            _write(block.synopsis_row, col, single["synopsis"])
            _write(block.strategy_row, col, single["strategy"])
            _write(block.tx_row, col, single["tx"])
            _write(block.tc_row, col, single["tc"])
            _write(block.comments_row, col, single["comments"])

    elif role == "Mediaplanning":
        # primer contenido en columna Mediaplanning
        # segundo contenido, si existe, en columna UK
        target_cols = []
        if role_cols:
            target_cols.append(role_cols[0])
        if uk_cols:
            target_cols.append(uk_cols[0])

        for idx, record in enumerate(records[:2]):
            if idx >= len(target_cols):
                break

            col = target_cols[idx]
            single = assemble_text_for_single_asset(record)

            _write(block.content_type_row, col, single["content_type"])
            _write(block.property_row, col, single["property"])
            _write(block.title_row, col, single["title"])
            _write(block.synopsis_row, col, single["synopsis"])
            _write(block.strategy_row, col, single["strategy"])
            _write(block.tx_row, col, single["tx"])
            _write(block.tc_row, col, single["tc"])
            _write(block.comments_row, col, single["comments"])

            # si el segundo contenido va en UK, cambiar cabecera OWNER a MEDIAPLANNING HUB
            if idx == 1 and block.owner_row and col in uk_cols:
                safe_set_value(
                    ws,
                    block.owner_row,
                    col,
                    "MEDIAPLANNING HUB",
                    alignment=Alignment(horizontal='center', vertical='center', wrap_text=True),
                )
def write_week_label(ws, block, week_number_value: str):
    """
    Write 'WEEK <N>' into B row of the anchor line.
    Safe version for dynamic layout.
    """
    safe_set_value(
        ws,
        block.anchor_row,
        2,
        f"WEEK {week_number_value}",
        alignment=Alignment(horizontal='center', vertical='center', wrap_text=True),
    )

def rewrite_full_calendar_owner_row(ws, block: WeekBlock, day_layout):
    """
    Fill OWNER row according to fixed columns in each day.
    """
    if not block.owner_row:
        return

    for day_name in DAY_ORDER:
        prog_cols = [col for role, col in day_layout[day_name] if role == "Programming"]
        media_cols = [col for role, col in day_layout[day_name] if role == "Mediaplanning"]
        uk_cols = [col for role, col in day_layout[day_name] if role == "UK"]

        for col in prog_cols:
            safe_set_value(
                ws,
                block.owner_row,
                col,
                "PROGRAMMING HUB",
                alignment=Alignment(horizontal='center', vertical='center', wrap_text=True),
            )

        for col in media_cols:
            safe_set_value(
                ws,
                block.owner_row,
                col,
                "MEDIAPLANNING HUB",
                alignment=Alignment(horizontal='center', vertical='center', wrap_text=True),
            )

        for col in uk_cols:
            safe_set_value(
                ws,
                block.owner_row,
                col,
                "YT TEAM UK",
                alignment=Alignment(horizontal='center', vertical='center', wrap_text=True),
            )

def shade_out_of_month_days_full_calendar(ws, blocks: Dict[date, WeekBlock], target_year: int, target_month: int):
    grey_fill = PatternFill(fill_type="solid", fgColor=OUT_OF_MONTH_FILL)

    for block in blocks.values():
        for idx, day_name in enumerate(DAY_ORDER):
            this_date = block.monday_date + timedelta(days=idx)

            if this_date.year == target_year and this_date.month == target_month:
                continue

            day_cols = [col for _, col in DAY_LAYOUT[day_name]]

            # no pintar las 2 últimas filas del bloque
            shade_until_row = max(block.anchor_row, block.end_row - 2)

            for col in day_cols:
                for row in range(block.anchor_row, shade_until_row + 1):
                    cell = ws.cell(row=row, column=col)
                    cell.fill = copy(grey_fill)

def set_full_calendar_column_widths(ws):
    """
    Adjust column widths for the fixed 4-column-per-day layout.
    Make Programming columns wider so days with 2 assets remain readable.
    """
    for day_name in DAY_ORDER:
        prog_cols = [col for role, col in DAY_LAYOUT[day_name] if role == "Programming"]
        media_cols = [col for role, col in DAY_LAYOUT[day_name] if role == "Mediaplanning"]
        uk_cols = [col for role, col in DAY_LAYOUT[day_name] if role == "UK"]

        for col in prog_cols:
            ws.column_dimensions[get_column_letter(col)].width = 18

        for col in media_cols:
            ws.column_dimensions[get_column_letter(col)].width = 16

        for col in uk_cols:
            ws.column_dimensions[get_column_letter(col)].width = 14

def auto_adjust_full_calendar_row_heights(ws, blocks: Dict[date, WeekBlock]):
    """
    Increase row heights for key content rows in FULL CALENDAR.
    """
    target_rows = set()

    for block in blocks.values():
        for row_idx in [
            block.property_row,
            block.content_type_row,
            block.title_row,
            block.synopsis_row,
            block.strategy_row,
            block.comments_row,
        ]:
            if row_idx:
                target_rows.add(row_idx)

    for row_idx in target_rows:
        ws.row_dimensions[row_idx].height = max(ws.row_dimensions[row_idx].height or 15, 36)

def write_week_deadlines_fixed(ws, block, per_day_records):
    if not block.deadline_row:
        return

    for day_name in DAY_ORDER:
        recs = per_day_records.get(day_name, [])
        if not recs:
            continue

        d = earliest_deadline_for_day(recs)
        if d:
            first_col = [col for _, col in DAY_LAYOUT[day_name]][0]
            safe_set_value(
                ws,
                block.deadline_row,
                first_col,
                d,
                alignment=Alignment(horizontal='center', vertical='center', wrap_text=True),
                number_format="dddd, dd mmmm yyyy",
            )

def write_countries_rollout_fixed(ws, block, per_day_role_records, channel_id_to_name_map=None):
    if channel_id_to_name_map is None:
        channel_id_to_name_map = {}

    for day_name in DAY_ORDER:
        # Programming: col1 y col2
        prog_cols = [col for r, col in DAY_LAYOUT[day_name] if r == "Programming"]
        prog_recs = per_day_role_records.get((day_name, "Programming"), [])

        for idx, record in enumerate(prog_recs[:2]):
            target_col = prog_cols[idx]
            included, _ = collect_included_countries(
                youtube_channel=record.get(AT_FIELDS["youtube_channel"]),
                region=record.get(AT_FIELDS["region"]),
                notes=record.get(AT_FIELDS["notes"]),
                channel_id_to_name_map=channel_id_to_name_map,
            )

            for country_name, rollout_row in block.rollout_country_rows.items():
                val = "YES" if country_name in included else "NO"
                safe_set_value(
                    ws, rollout_row, target_col, val,
                    alignment=Alignment(horizontal='center', vertical='center', wrap_text=True),
                    font=Font(color="FFFF0000" if (val == "NO" or "italy" in country_name.lower()) else "FF000000"),
                )

            for country_name, status_row in block.status_country_rows.items():
                val = "YES" if country_name in included else "NO"
                safe_set_value(
                    ws, status_row, target_col, val,
                    alignment=Alignment(horizontal='center', vertical='center', wrap_text=True),
                    font=Font(color="FFFF0000" if (val == "NO" or "italy" in country_name.lower()) else "FF000000"),
                )

        # Mediaplanning: col3 y opcional col4
        media_recs = per_day_role_records.get((day_name, "Mediaplanning"), [])
        media_cols = [col for r, col in DAY_LAYOUT[day_name] if r == "Mediaplanning"]
        uk_cols = [col for r, col in DAY_LAYOUT[day_name] if r == "UK"]
        target_cols = []
        if media_cols:
            target_cols.append(media_cols[0])
        if uk_cols:
            target_cols.append(uk_cols[0])

        for idx, record in enumerate(media_recs[:2]):
            if idx >= len(target_cols):
                break

            target_col = target_cols[idx]
            included, _ = collect_included_countries(
                youtube_channel=record.get(AT_FIELDS["youtube_channel"]),
                region=record.get(AT_FIELDS["region"]),
                notes=record.get(AT_FIELDS["notes"]),
                channel_id_to_name_map=channel_id_to_name_map,
            )

            for country_name, rollout_row in block.rollout_country_rows.items():
                val = "YES" if country_name in included else "NO"
                safe_set_value(
                    ws, rollout_row, target_col, val,
                    alignment=Alignment(horizontal='center', vertical='center', wrap_text=True),
                    font=Font(color="FFFF0000" if (val == "NO" or "italy" in country_name.lower()) else "FF000000"),
                )

            for country_name, status_row in block.status_country_rows.items():
                val = "YES" if country_name in included else "NO"
                safe_set_value(
                    ws, status_row, target_col, val,
                    alignment=Alignment(horizontal='center', vertical='center', wrap_text=True),
                    font=Font(color="FFFF0000" if (val == "NO" or "italy" in country_name.lower()) else "FF000000"),
                )

def write_full_calendar_start_date(ws, start_monday: date):
    """
    Only writes D6. The rest of the sheet updates via Excel formulas.
    """
    safe_set_value(
        ws,
        6,
        column_index_from_string("D"),
        start_monday,
        alignment=Alignment(horizontal='center', vertical='center', wrap_text=True),
        number_format="dddd, dd mmmm yyyy",
    )

def filter_records_by_channel(records: List[Dict], channel: str) -> List[Dict]:
    target = (channel or "").strip().lower()
    if not target:
        return records
    filtered = []
    for rec in records:
        fields = rec.get("fields", rec)
        ch = normalize_str(fields.get(AT_FIELDS["channel"]))
        if ch and ch.lower() == target:
            filtered.append(rec)
    return filtered

def filter_records_by_month(records: List[Dict], month: Optional[str]) -> List[Dict]:
    if not month:
        return records
    target = month.strip()

    MONTH_NAMES = {
        "january": 1, "february": 2, "march": 3, "april": 4,
        "may": 5, "june": 6, "july": 7, "august": 8,
        "september": 9, "october": 10, "november": 11, "december": 12
    }

    def month_to_number(value: str) -> Optional[int]:
        if not value:
            return None
        v = value.strip().lower()
        if v.isdigit():
            try:
                m = int(v)
                if 1 <= m <= 12:
                    return m
            except ValueError:
                return None
        if v in MONTH_NAMES:
            return MONTH_NAMES[v]
        # soportar abreviaturas (Jan, Feb, Mar...)
        v3 = v[:3]
        for k in MONTH_NAMES:
            if k[:3] == v3:
                return MONTH_NAMES[k]
        return None

    target_num = month_to_number(target)
    target_low = target.lower()

    filtered = []
    for rec in records:
        fields = rec.get("fields", rec)
        rec_month = normalize_str(fields.get(AT_FIELDS["month"]))
        rec_month_num = month_to_number(rec_month)

        if rec_month_num is None:
            # Fallback a live_date cuando Month no está presente o inválido
            live = parse_date(normalize_str(fields.get(AT_FIELDS["live_date"])))
            if live:
                rec_month_num = live.month

        if target_num is not None and rec_month_num is not None:
            if rec_month_num == target_num:
                filtered.append(rec)
                continue

        if rec_month.lower() == target_low:
            filtered.append(rec)
            continue

    return filtered

def month_name_from_records(records: List[Dict]) -> str:
    """
    Returns uppercase month name from records, fallback to current month field/value.
    """
    month_lookup = {
        1: "JANUARY", 2: "FEBRUARY", 3: "MARCH", 4: "APRIL",
        5: "MAY", 6: "JUNE", 7: "JULY", 8: "AUGUST",
        9: "SEPTEMBER", 10: "OCTOBER", 11: "NOVEMBER", 12: "DECEMBER",
    }

    for rec in records:
        fields = rec.get("fields", rec)
        raw_month = normalize_str(fields.get(AT_FIELDS["month"]))
        if raw_month:
            # If already text month
            m = raw_month.strip().lower()
            names = {v.lower(): k for k, v in month_lookup.items()}
            if m in names:
                return month_lookup[names[m]]
            # if numeric
            if m.isdigit():
                n = int(m)
                if 1 <= n <= 12:
                    return month_lookup[n]
        # fallback to live date
        live = parse_date(fields.get(AT_FIELDS["live_date"]))
        if live:
            return month_lookup[live.month]

    return "MONTH"

def first_visible_monday(records: List[Dict]) -> Optional[date]:
    """
    For MONTHLY OVERVIEW: use the Monday of the first visible week of the target month,
    not the earliest live date.
    """
    return first_visible_monday_of_target_month(records)

def resolve_target_month_year(records: List[Dict]) -> Tuple[Optional[int], Optional[int]]:
    """
    Determine target (year, month) from filtered records.
    Uses the most common live-date month/year.
    """
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

def full_calendar_anchor_dates_from_records(records: List[Dict], max_weeks: int = 5) -> Dict[str, date]:
    """
    Build FULL CALENDAR anchor dates from the target month, not from earliest live date.
    """
    anchor_cells = ["D6", "D59", "D112", "D165", "D218"]

    first_monday = first_visible_monday_of_target_month(records)
    if not first_monday:
        return {}

    anchor_map = {}
    for i, cell in enumerate(anchor_cells[:max_weeks]):
        anchor_map[cell] = first_monday + timedelta(days=i * 7)

    return anchor_map

def replace_month_in_title(text: str, month_name_upper: str) -> str:
    if not text:
        return text
    return text.replace("(MONTH)", month_name_upper).replace("MONTH", month_name_upper)

def update_full_calendar_month_title(ws, records: List[Dict]):
    """
    Update B3:Z3 merged title in FULL CALENDAR replacing MONTH with actual month.
    """
    month_name_upper = month_name_from_records(records)
    cell = get_writable_cell(ws, 3, 2)  # B3
    existing = normalize_str(cell.value)
    if existing:
        cell.value = replace_month_in_title(existing, month_name_upper)
    else:
        cell.value = month_name_upper
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

def overview_bucket(priority: str) -> str:
    p = normalize_str(priority)
    if p in {"Current A", "Current B", "Current C"}:
        return "priority"
    if p in {"Tactical A", "Tactical B", "Tactical Opportunity"}:
        return "tactical"
    if p == "Evergreen":
        return "evergreen"
    return "priority"


def overview_row_offset(content_type: str, priority: str) -> Optional[int]:
    ct = normalize_str(content_type)
    bucket = overview_bucket(priority)

    # Full Eps
    if ct == "Full Episode":
        return {
            "priority": 0,
            "tactical": 1,
            "evergreen": 2,
            "acquisitions": 3,
        }.get(bucket, 0)

    # Exclusive Shorts
    if ct == "Exclusive":
        return 4

    # Scene Lifts / Sneak Peeks / Videoclips
    if ct in {"Lift", "Musical Lift"}:
        return 5

    # Special Formats (Fillers)
    if ct in {"Podcast", "Magical Moment", "Interactive", "Editorial Short"}:
        return 6

    # Compilations
    if ct == "Compilation":
        return {
            "priority": 7,
            "tactical": 8,
            "evergreen": 9,
            "acquisitions": 10,
        }.get(bucket, 7)

    # Live Streams
    if ct == "Live Stream":
        return {
            "priority": 11,
            "tactical": 12,
            "evergreen": 13,
            "acquisitions": 14,
        }.get(bucket, 11)

    # Vertical Shorts
    if ct == "Vertical Short":
        return 15

    return None

def overview_visual_position(live_date: date, first_monday: date) -> Optional[Tuple[int, str]]:
    """
    Returns (visual_week_number, day_name) based on the visible calendar grid.
    Week numbers are 1..5 relative to first visible Monday.
    """
    if not live_date or not first_monday:
        return None

    delta_days = (live_date - first_monday).days
    if delta_days < 0:
        return None

    visual_week = (delta_days // 7) + 1
    if visual_week < 1 or visual_week > 5:
        return None

    day_name = DAY_ORDER[live_date.weekday()]
    return visual_week, day_name

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

def build_overview_lines(record: Dict, channel_id_to_name_map: Optional[Dict[str, str]] = None) -> Tuple[str, str, bool]:
    """
    Returns:
    line1, line2, geofenced
    """
    if channel_id_to_name_map is None:
        channel_id_to_name_map = {}

    activity = normalize_str(record.get(AT_FIELDS["activity"]))
    if not activity:
        activity = normalize_str(record.get(AT_FIELDS["asset_label"])) or normalize_str(record.get(AT_FIELDS["title"]))

    content_type = normalize_str(record.get(AT_FIELDS["content_type"]))

    if content_type == "Musical Lift":
        activity = f"♪ {activity}"
    elif content_type == "Lift":
        activity = f"▶ {activity}"
    elif content_type == "Exclusive":
        activity = f"★ {activity}"
    elif content_type == "Compilation":
        activity = f"COMP {activity}"
    elif content_type == "Live Stream":
        activity = f"LIVE {activity}"

    geofenced = lower_clean(record.get(AT_FIELDS["geofenced"])) == "yes"

    included_countries, _ = collect_included_countries(
        youtube_channel=record.get(AT_FIELDS["youtube_channel"]),
        region=record.get(AT_FIELDS["region"]),
        notes=record.get(AT_FIELDS["notes"]),
        channel_id_to_name_map=channel_id_to_name_map,
    )

    line2 = normalize_overview_region(record.get(AT_FIELDS["region"]))
    except_suffix = format_except_suffix(record.get(AT_FIELDS["region"]), included_countries)

    if line2:
        line2 = f"{line2}{except_suffix}"
    else:
        line2 = except_suffix.strip()

    return activity, line2, geofenced


def apply_overview_fill(cell, content_type: str):
    fill_color = OVERVIEW_FILL_COLORS.get(normalize_str(content_type))
    if fill_color:
        cell.fill = PatternFill(fill_type="solid", fgColor=fill_color)


def write_overview_cell(ws, row: int, col: int, line1: str, line2: str, content_type: str, geofenced: bool):
    cell = get_writable_cell(ws, row, col)

    new_text = line1 if not line2 else f"{line1}\n{line2}"

    existing = normalize_str(cell.value)
    if existing:
        cell.value = f"{existing}\n\n— — —\n\n{new_text}"
    else:
        cell.value = new_text

    fill_color = OVERVIEW_FILL_COLORS.get(normalize_str(content_type))
    if fill_color:
        cell.fill = PatternFill(fill_type="solid", fgColor=fill_color)

    font_color = "FF800080" if geofenced else "FF000000"
    cell.font = Font(bold=True, color=font_color)
    cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

def write_vertical_short_cell(ws, top_row: int, col: int, record: Dict):
    """
    Vertical Shorts use 2 stacked cells:
    - top cell: PROPERTY (uppercase) + Strategy, centered, black
    - bottom cell: Region/Tiers, centered, bold, red
      or yellow if SPAIN ONLY
    Ignores geofenced purple styling for this content type.
    """
    top_cell = get_writable_cell(ws, top_row, col)
    bottom_cell = get_writable_cell(ws, top_row + 1, col)

    prop_val = record.get(AT_FIELDS["property"])
    if isinstance(prop_val, list):
        prop = " / ".join(normalize_str(p) for p in prop_val if normalize_str(p))
    else:
        prop = normalize_str(prop_val)
    prop = prop.upper()

    strat_val = record.get(AT_FIELDS["strategy"])
    if isinstance(strat_val, list):
        strat = " / ".join(normalize_str(s) for s in strat_val if normalize_str(s))
    else:
        strat = normalize_str(strat_val)

    tier_text = normalize_overview_region(record.get(AT_FIELDS["region"]))

    # top cell text
    top_text = prop
    if strat:
        top_text = f"{prop}\n{strat}"

    top_cell.value = top_text
    top_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    top_cell.font = Font(bold=True, color="FF000000")
    top_cell.fill = PatternFill(fill_type="solid", fgColor=OVERVIEW_FILL_COLORS["Vertical Short"])

    # bottom cell text
    bottom_cell.value = tier_text
    bottom_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    if tier_text == "SPAIN ONLY":
        bottom_cell.font = Font(bold=True, color="FFFFFF00")  # yellow
    else:
        bottom_cell.font = Font(bold=True, color="FFFF0000")  # red

    bottom_cell.fill = PatternFill(fill_type="solid", fgColor=OVERVIEW_FILL_COLORS["Vertical Short"])

def auto_adjust_overview_row_height(ws, row: int, min_height: float = 24, per_line: float = 16, chars_per_line: int = 22):
    """
    Better approximation of row auto-height based on explicit line breaks and wrapped text.
    """
    max_visual_lines = 1

    for col in range(4, 11):  # D:J
        cell = get_writable_cell(ws, row, col)
        value = "" if cell.value is None else str(cell.value)

        if not value:
            continue

        visual_lines = 0
        for line in value.split("\n"):
            line = line.strip()
            if not line:
                visual_lines += 1
            else:
                visual_lines += max(1, math.ceil(len(line) / chars_per_line))

        max_visual_lines = max(max_visual_lines, visual_lines)

    ws.row_dimensions[row].height = max(min_height, max_visual_lines * per_line)


def clear_overview_week(ws, week_num: int):
    """
    Clear only content/event writable area in a week block.
    Do NOT clear header/date formula cells.
    """
    if week_num not in OVERVIEW_WEEK_LAYOUT:
        return

    layout = OVERVIEW_WEEK_LAYOUT[week_num]

    # Clear only event rows
    for row in layout["event_rows"]:
        for col in range(4, 11):
            cell = get_writable_cell(ws, row, col)
            cell.value = None

    # Clear only content area
    for row in range(layout["base_row"], layout["end_row"] + 1):
        for col in range(4, 11):
            cell = get_writable_cell(ws, row, col)
            cell.value = None


def write_overview_headers(ws, records: List[Dict]):
    month_name_upper = month_name_from_records(records)

    title_cell = get_writable_cell(ws, 2, 2)  # B2
    existing_title = normalize_str(title_cell.value)
    if existing_title:
        title_cell.value = replace_month_in_title(existing_title, month_name_upper)
    else:
        title_cell.value = f"DC EMEA YT {month_name_upper} 2026 CALENDAR"
    title_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for week_num, layout in OVERVIEW_WEEK_LAYOUT.items():
        for day_name, col in OVERVIEW_DAY_COLS.items():
            day_label_cell = get_writable_cell(ws, layout["header_row"], col)
            day_label_cell.value = OVERVIEW_DAY_LABELS[day_name]
            day_label_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    first_monday = first_visible_monday(records)
    if first_monday:
        monday_date_cell = get_writable_cell(ws, 5, OVERVIEW_DAY_COLS["Monday"])  # D5 only
        monday_date_cell.value = first_monday
        monday_date_cell.number_format = "dd-mmm"
        monday_date_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

def collect_overview_events(records: List[Dict]) -> Dict[int, List[Tuple[str, int, int]]]:
    """
    Returns:
    { week_num: [(event_text, start_col, end_col), ...] }

    Uses EVENT content type and merges across consecutive days if same text.
    """
    events_by_week = defaultdict(list)

    raw = []
    for rec in records:
        fields = rec.get("fields", rec)
        if normalize_str(fields.get(AT_FIELDS["content_type"])) != "EVENT":
            continue

        live = parse_date(fields.get(AT_FIELDS["live_date"]))
        if not live:
            continue

        wom = fields.get(AT_FIELDS["week_of_month"])
        try:
            week_num = int(str(wom).strip())
        except Exception:
            continue

        day_name = normalize_str(fields.get(AT_FIELDS["day_of_week"]))
        if day_name not in DAY_ORDER:
            day_name = DAY_ORDER[live.weekday()]

        event_text = normalize_str(fields.get(AT_FIELDS["activity"])) or normalize_str(fields.get(AT_FIELDS["title"]))
        if not event_text:
            continue

        raw.append((week_num, day_name, event_text))

    # group by week + text
    grouped = defaultdict(list)
    for week_num, day_name, event_text in raw:
        grouped[(week_num, event_text)].append(OVERVIEW_DAY_COLS[day_name])

    for (week_num, event_text), cols in grouped.items():
        cols = sorted(set(cols))
        # merge contiguous spans
        start = cols[0]
        prev = cols[0]
        for col in cols[1:]:
            if col == prev + 1:
                prev = col
            else:
                events_by_week[week_num].append((event_text, start, prev))
                start = col
                prev = col
        events_by_week[week_num].append((event_text, start, prev))

    return events_by_week


def write_overview_events(ws, records: List[Dict]):
    """
    Writes EVENT labels into grey rows.
    Uses 2 rows max per week.
    If event spans multiple days, merges across those day columns.
    """
    events_by_week = collect_overview_events(records)

    for week_num, event_spans in events_by_week.items():
        if week_num not in OVERVIEW_WEEK_LAYOUT:
            continue

        event_rows = OVERVIEW_WEEK_LAYOUT[week_num]["event_rows"]

        for idx, (event_text, start_col, end_col) in enumerate(event_spans[:2]):
            row = event_rows[idx]

            # merge if span > 1 col
            if end_col > start_col:
                ws.merge_cells(start_row=row, start_column=start_col, end_row=row, end_column=end_col)

            cell = get_writable_cell(ws, row, start_col)
            cell.value = event_text
            cell.font = Font(italic=True, bold=False, color="FF000000")
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for week_num, layout in OVERVIEW_WEEK_LAYOUT.items():
        for event_row in layout["event_rows"]:
            auto_adjust_overview_row_height(ws, event_row, min_height=18, per_line=14)

def populate_monthly_overview(
    wb,
    records: List[Dict],
    enable_mediaplanning: bool = True,
    channel_id_to_name_map: Optional[Dict[str, str]] = None,
):
    """
    Populate 'MONTHLY OVERVIEW YT DC EMEA'.

    Uses visual calendar positioning based on the first visible Monday,
    not Airtable Week of Month.
    """
    if channel_id_to_name_map is None:
        channel_id_to_name_map = {}    
    
    if MONTHLY_OVERVIEW_SHEET not in wb.sheetnames:
        logging.warning(f"Sheet '{MONTHLY_OVERVIEW_SHEET}' not found. Skipping MONTHLY OVERVIEW export.")
        return

    ws = wb[MONTHLY_OVERVIEW_SHEET]

    # Header/title and first Monday
    write_overview_headers(ws, records)

    # Events
    write_overview_events(ws, records)

    first_monday = first_visible_monday(records)
    if not first_monday:
        logging.warning("Could not determine first visible Monday for MONTHLY OVERVIEW.")
        return

    used_weeks = set()

    for rec in records:
        fields = rec.get("fields", rec)

        if asset_is_event(fields):
            continue

        owner = hub_role(fields.get(AT_FIELDS["owner_hub"]))
        if owner not in {"Programming", "Mediaplanning"}:
            continue

        if owner == "Mediaplanning" and not enable_mediaplanning:
            continue

        content_type = normalize_str(fields.get(AT_FIELDS["content_type"]))
        live_date = parse_date(fields.get(AT_FIELDS["live_date"]))
        if not live_date:
            continue

        pos = overview_visual_position(live_date, first_monday)
        if not pos:
            continue

        visual_week, day_name = pos
        if visual_week not in OVERVIEW_WEEK_LAYOUT:
            continue

        used_weeks.add(visual_week)

        row_offset = overview_row_offset(content_type, fields.get(AT_FIELDS["priority"]))
        if row_offset is None:
            continue

        base_row = OVERVIEW_WEEK_LAYOUT[visual_week]["base_row"]
        target_row = base_row + row_offset
        target_col = OVERVIEW_DAY_COLS[day_name]

        if content_type == "Vertical Short":
            write_vertical_short_cell(ws, target_row, target_col, fields)
            auto_adjust_overview_row_height(ws, target_row, min_height=26, per_line=16)
            auto_adjust_overview_row_height(ws, target_row + 1, min_height=22, per_line=16)
        else:
            line1, line2, geofenced = build_overview_lines(fields, channel_id_to_name_map=channel_id_to_name_map)
            write_overview_cell(ws, target_row, target_col, line1, line2, content_type, geofenced)
            auto_adjust_overview_row_height(ws, target_row)

    # If week 5 unused, clear it
    if 5 not in used_weeks:
        clear_overview_week(ws, 5)

    # Event row height approximation
    for week_num, layout in OVERVIEW_WEEK_LAYOUT.items():
        for event_row in layout["event_rows"]:
            auto_adjust_overview_row_height(ws, event_row, min_height=24, per_line=14)

    logging.info(f"Finished populating '{MONTHLY_OVERVIEW_SHEET}'.")

def infer_week_number_for_block(per_day_all_roles: Dict[str, List[Dict]], monday_date: date) -> str:
    """
    Infer 'Week Number (Year)' for the block from Airtable records.
    Use the most common non-empty value. Fallback to ISO week of monday_date.
    """
    values = []
    for recs in per_day_all_roles.values():
        for f in recs:
            wn = normalize_str(f.get(AT_FIELDS["week_number"]))
            if wn:
                values.append(wn)
    if values:
        most_common = Counter(values).most_common(1)[0][0]
        return most_common
    # fallback ISO week
    return str(monday_date.isocalendar()[1])

def populate_calendar(template_path: str, output_path: str, records: List[Dict],
                      week_anchor_cells: Optional[List[str]] = None,
                      enable_mediaplanning: bool = True,
                      anchor_dates_override: Optional[List[str]] = None,
                      channel_id_to_name_map: Optional[Dict[str, str]] = None):

    week_anchor_cells = week_anchor_cells or ["D6", "D59", "D112", "D165", "D218"]
    if channel_id_to_name_map is None:
        channel_id_to_name_map = {}

    wb = load_workbook(template_path, data_only=False)

    if TARGET_SHEET not in wb.sheetnames:
        raise RuntimeError(f"Target sheet '{TARGET_SHEET}' not found in workbook.")
    ws = wb[TARGET_SHEET]

    set_full_calendar_column_widths(ws)
    update_full_calendar_month_title(ws, records)

    anchor_dates_map = full_calendar_anchor_dates_from_records(records)

    if not anchor_dates_map:
        anchor_dates_map = get_anchor_dates_from_values_wb(template_path, week_anchor_cells)

    blocks = find_week_blocks(ws, anchor_dates_map)

    buckets = aggregate_by_week_day_and_role(records)

    start_monday = first_visible_monday_of_target_month(records)
    if start_monday:
        write_full_calendar_start_date(ws, start_monday)

    target_year, target_month = resolve_target_month_year(records)

    for mon_date, block in blocks.items():
        logging.info(f"Processing week block starting {block.monday_date} at row {block.anchor_row}...")

        per_day_role_records: Dict[Tuple[str, str], List[Dict]] = {}
        per_day_all_roles: Dict[str, List[Dict]] = {d: [] for d in DAY_ORDER}

        for day_name in DAY_ORDER:
            prog = buckets.get((mon_date, day_name, "Programming"), [])
            media = buckets.get((mon_date, day_name, "Mediaplanning"), [])

            per_day_role_records[(day_name, "Programming")] = prog
            per_day_role_records[(day_name, "Mediaplanning")] = media if enable_mediaplanning else []
            per_day_all_roles[day_name] = prog + (media if enable_mediaplanning else [])

        week_num_text = infer_week_number_for_block(per_day_all_roles, block.monday_date)
        write_week_label(ws, block, week_num_text)
        rewrite_full_calendar_owner_row(ws, block, DAY_LAYOUT)


        for day_name in DAY_ORDER:
            write_to_day_role_columns(ws, block, day_name, "Programming",
                                      per_day_role_records[(day_name, "Programming")], enable_mediaplanning)
            write_to_day_role_columns(ws, block, day_name, "Mediaplanning",
                                      per_day_role_records[(day_name, "Mediaplanning")], enable_mediaplanning)

        write_week_deadlines_fixed(ws, block, per_day_all_roles)
        write_countries_rollout_fixed(ws, block, per_day_role_records, channel_id_to_name_map)

        # reconstruir merges visuales al final de la semana
        rebuild_full_calendar_week_merges(ws, block)

    if target_year and target_month:
        shade_out_of_month_days_full_calendar(ws, blocks, target_year, target_month)

    auto_adjust_full_calendar_row_heights(ws, blocks)

    populate_monthly_overview(
        wb=wb,
        records=records,
        enable_mediaplanning=enable_mediaplanning,
        channel_id_to_name_map=channel_id_to_name_map,
    )

    logging.info(f"Saving populated workbook to: {output_path}")
    out_dir = os.path.dirname(output_path)
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)
    if output_path.lower().endswith(".xlsx"):
        wb.template = False
    elif output_path.lower().endswith(".xltx"):
        wb.template = True
    cleanup_stale_merged_cells(ws) 
    wb.save(output_path)
    logging.info("Done.")

def main():
    parser = argparse.ArgumentParser(description="Populate Disney Channel EMEA YouTube calendar into existing template.")
    parser.add_argument("--template", required=True, help="Path to existing Excel template (.xltx or .xlsx).")
    parser.add_argument("--output", required=True, help="Path to output .xlsx file.")
    parser.add_argument("--airtable-token", default=os.getenv("AIRTABLE_API_TOKEN"),
                        help="Airtable API token or set AIRTABLE_API_TOKEN.")
    parser.add_argument("--airtable-base", help="Airtable Base ID.")
    parser.add_argument("--airtable-table", help="Airtable Table ID or name.")
    parser.add_argument("--airtable-view", help="Optional Airtable view name.")
    parser.add_argument("--json", help="Optional path to local JSON export instead of API.")
    parser.add_argument("--anchors", nargs="*", help="Override weekly anchor cells, e.g., D6 D59 D112 D165 D218.")
    parser.add_argument("--anchor-dates", nargs="*", help="Optional ISO dates for anchors.")
    parser.add_argument("--disable-mediaplanning", action="store_true", help="If set, do not populate Mediaplanning Hub content.")
    parser.add_argument("--channel", default="Disney Channel", help='Filter by "Channel" (e.g., "Disney Channel", "Disney Jr.").')
    parser.add_argument("--month", help='Filter by "Month" (optional, case-insensitive).')
    parser.add_argument("--log-level", default="INFO", help="Logging level (DEBUG, INFO, WARNING, ERROR).")
    args = parser.parse_args()

    load_dotenv()
    token = args.airtable_token or os.getenv("AIRTABLE_API_TOKEN")
    base_id = args.airtable_base or os.getenv("AIRTABLE_BASE_ID")
    table_id = args.airtable_table or os.getenv("AIRTABLE_TABLE_ID")
    view = args.airtable_view or os.getenv("AIRTABLE_VIEW")

    if token: token = token.strip()
    if base_id: base_id = base_id.strip().rstrip(".")
    if table_id: table_id = table_id.strip().rstrip(".")
    if view: view = view.strip()
    logging.debug(f"Using Airtable base={base_id!r} table={table_id!r} view={view!r}")

    # Manejo de ENABLE_MEDIAPLANNING (env) si no usas el flag para desactivarlo
    enable_mediaplanning = not args.disable_mediaplanning
    env_mp = os.getenv("ENABLE_MEDIAPLANNING")
    if env_mp is not None and not args.disable_mediaplanning:
        enable_mediaplanning = env_mp.strip().lower() in ("1", "true", "yes", "y", "on")

    # (Opcional futuro) COUNTRY_SECTION está en .env. Tu script actualmente siempre escribe ambas
    # secciones. Si lo reactivas, podrías leerlo así:
    # country_section = (os.getenv("COUNTRY_SECTION") or "both").strip().lower()

    logging.basicConfig(stream=sys.stdout, level=getattr(logging, args.log_level.upper(), logging.INFO),
                        format="%(levelname)s: %(message)s")

    # Fuente de datos: JSON o Airtable
    channel_id_to_name_map = {}

    if args.json:
        logging.info(f"Loading records from JSON: {args.json}")
        records = load_local_json(args.json)

        # Even with local JSON, try to resolve linked YouTube Channel IDs if Airtable creds exist
        if token and base_id:
            logging.info("Loading channel mappings from Airtable...")
            channel_id_to_name_map = fetch_channel_id_to_name_map(token, base_id)
    else:
        if not token or not base_id or not table_id:
            missing = []
            if not token: missing.append("AIRTABLE_API_TOKEN")
            if not base_id: missing.append("AIRTABLE_BASE_ID")
            if not table_id: missing.append("AIRTABLE_TABLE_ID")
            parser.error(f"Missing Airtable config: {', '.join(missing)}. Set them in .env or pass flags.")

        logging.info("Fetching records from Airtable...")
        records = fetch_airtable_records(
            api_token=token,
            base_id=base_id,
            table_id=table_id,
            view=view,
        )

        logging.info("Loading channel mappings from Airtable...")
        channel_id_to_name_map = fetch_channel_id_to_name_map(token, base_id)

    logging.info(f"Fetched {len(records)} records from Airtable/JSON source")

    # Filtros (Channel/Month) permanecen igual
    records = filter_records_by_channel(records, args.channel)
    records = filter_records_by_month(records, args.month)
    logging.info(f"After month filter '{args.month}': {len(records)} records")

    week_anchors = args.anchors if args.anchors else ["D6", "D59", "D112", "D165", "D218"]
    populate_calendar(
        template_path=args.template,
        output_path=args.output,
        records=records,
        week_anchor_cells=week_anchors,
        enable_mediaplanning=enable_mediaplanning,
        anchor_dates_override=args.anchor_dates,
        channel_id_to_name_map=channel_id_to_name_map,
    )

if __name__ == "__main__":
    main()