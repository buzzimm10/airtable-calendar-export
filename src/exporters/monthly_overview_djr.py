import logging
from collections import defaultdict
from datetime import timedelta

from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

from config.fields import AT_FIELDS
from config.layouts import DAY_ORDER
from config.layouts_djr import DJR_MONTHLY_OVERVIEW_SHEET
from config.countries_djr import DJR_COUNTRY_SHORT_NAMES
from core.dates import (
    parse_date,
    first_visible_monday_of_target_month,
    month_name_from_records,
)
from core.text_formatting import normalize_str, lower_clean
from core.excel_utils import get_writable_cell, center_wrap
from core.countries import normalize_overview_region
from services.planner_service import hub_role, asset_is_event
from typing import Dict, List, Optional, Tuple

DJR_OVERVIEW_DAY_COLS = {
    "Monday": 4, "Tuesday": 5, "Wednesday": 6, "Thursday": 7,
    "Friday": 8, "Saturday": 9, "Sunday": 10,
}

DJR_OVERVIEW_WEEK_LAYOUT = {
    1: {"date_row": 9, "full_eps": 10, "exclusive": 11, "lifts": 12, "compilations": 13, "live": 14, "podcast_rows": (15, 16), "vertical_rows": (17, 18)},
    2: {"date_row": 20, "full_eps": 21, "exclusive": 22, "lifts": 23, "compilations": 24, "live": 25, "podcast_rows": (26, 27), "vertical_rows": (28, 29)},
    3: {"date_row": 31, "full_eps": 32, "exclusive": 33, "lifts": 34, "compilations": 35, "live": 36, "podcast_rows": (37, 38), "vertical_rows": (39, 40)},
    4: {"date_row": 42, "full_eps": 43, "exclusive": 44, "lifts": 45, "compilations": 46, "live": 47, "podcast_rows": (48, 49), "vertical_rows": (50, 51)},
    5: {"date_row": 53, "full_eps": 54, "exclusive": 55, "lifts": 56, "compilations": 57, "live": 58, "podcast_rows": (59, 60), "vertical_rows": (61, 62)},
}

DJR_OVERVIEW_FILL_COLORS = {
    "full_eps": "FFFF00", "exclusive": "B4C6E7", "lifts": "A9D08E",
    "compilations": "F4B084", "live": "00B0F0", "podcast": "CC99FF", "vertical": "FF99CC",
}

PROPERTY_ABBREVIATIONS = {
    "Marvel Iron Man and His Awesome Friends": "Marvel Iron Man AHAF",
    "Mickey Mouse Clubhouse": "MMCH+",
    "Mickey Mouse and the Roadster Racers": "MMRR",
    "Mickey Mouse and the Roadster Racers'": "MMRR",
}

def clean_property(value):
    if isinstance(value, list): value = " / ".join(normalize_str(x) for x in value if normalize_str(x))
    value = normalize_str(value).replace("[", "").replace("]", "").replace("'", "").replace('"', "")
    return PROPERTY_ABBREVIATIONS.get(value, value)

def get_episode_category(record):
    category = normalize_str(record.get(AT_FIELDS.get("episode_category", "Episode Category")))
    if not category: return "CURRENT"
    category_upper = category.upper()
    if "LIBRARY" in category_upper: return "LIBRARY"
    if "CURRENT" in category_upper: return "CURRENT"
    return category_upper

def get_djr_overview_region_text(record):
    region_text = normalize_overview_region(record.get(AT_FIELDS["region"]))
    return region_text.replace("TIER A, B & C", "TIER A, B, C")

def get_episode_number(record):
    title = normalize_str(record.get(AT_FIELDS["youtube_title"]))
    activity = normalize_str(record.get(AT_FIELDS["activity"]))
    source = f"{activity} {title}"
    import re
    match = re.search(r"\b(?:ep|episode|episodio)\.?\s*(\d+)", source, re.I)
    if match: return match.group(1)
    return ""

def build_djr_overview_text(record, target_country=None):
    content_type = normalize_str(record.get(AT_FIELDS["content_type"]))
    prop = clean_property(record.get(AT_FIELDS["property"]))
    title = normalize_str(record.get(AT_FIELDS["youtube_title"]))
    activity = normalize_str(record.get(AT_FIELDS["activity"]))
    
    region = "" if target_country else get_djr_overview_region_text(record)
    
    fields = record.get("fields", record)
    policy_val = fields.get("Rights Management Policy")
    policy_str = str(policy_val).strip() if policy_val else ""
    
    gf_suffix = ""
    if "Block AUS/NZ" in policy_str: gf_suffix = " [BLOCK AUS/NZ]"
    elif "Monetize Locally" in policy_str: gf_suffix = " GF"
    elif "Block Everywhere" in policy_str: gf_suffix = " [🚫 NO MONETIZE]"

    ep = get_episode_number(record)

    if content_type == "Full Episode":
        ep_text = f" ep. {ep}" if ep else ""
        line1 = f"{prop}{ep_text} {title}".strip()
        line2 = get_episode_category(record)
        return f"{line1}{gf_suffix}\n{line2}\n{region}".strip()

    if content_type in {"Exclusive", "Editorial Short"}:
        ep_text = f" ep. {ep}" if ep else ""
        return f"{prop} Short{ep_text}\n{region}{gf_suffix}".strip()

    if content_type in {"Lift", "Musical Lift", "Magical Moment"}:
        label = "Magical Moment" if content_type == "Magical Moment" else content_type
        return f"{label} {prop}{gf_suffix}".strip()

    if content_type == "Compilation":
        prefix = "MEGA Compilation" if "mega" in activity.lower() or "mega" in title.lower() else "Compilation"
        return f"{prefix} {prop}\n{region}{gf_suffix}".strip()

    if content_type == "Live Stream":
        if "full" in title.lower() or "episode" in title.lower(): return f"{prop} Compilation Full Eps\n{region}{gf_suffix}".strip()
        return f"MEGA {prop}\n{region}{gf_suffix}".strip()

    if content_type == "Podcast":
        return f"{prop} Podcast\n{region}{gf_suffix}".strip()

    if content_type == "Vertical Short":
        return prop

    return f"{prop}\n{region}{gf_suffix}".strip()

def get_djr_overview_bucket(content_type):
    if content_type == "Full Episode": return "full_eps"
    if content_type in {"Exclusive", "Editorial Short"}: return "exclusive"
    if content_type in {"Lift", "Musical Lift", "Magical Moment"}: return "lifts"
    if content_type == "Compilation": return "compilations"
    if content_type == "Live Stream": return "live"
    if content_type == "Podcast": return "podcast"
    if content_type == "Vertical Short": return "vertical"
    return None

def append_to_cell(cell, text):
    existing = normalize_str(cell.value)
    if existing: cell.value = f"{existing}\n\n— — —\n\n{text}"
    else: cell.value = text

def write_djr_overview_regular_cell(ws, row, col, text, bucket):
    cell = get_writable_cell(ws, row, col)
    append_to_cell(cell, text)
    cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    cell.font = Font(bold=True, color="FF000000")
    fill_color = DJR_OVERVIEW_FILL_COLORS.get(bucket)
    if fill_color: cell.fill = PatternFill(fill_type="solid", fgColor=fill_color)

def write_djr_overview_vertical(ws, layout, col, record, target_country=None):
    top_row, bottom_row = layout["vertical_rows"]
    top_cell = get_writable_cell(ws, top_row, col)
    bottom_cell = get_writable_cell(ws, bottom_row, col)
    top_text = build_djr_overview_text(record, target_country=target_country)
    bottom_text = get_djr_overview_region_text(record)
    append_to_cell(top_cell, top_text)

    fill = PatternFill(fill_type="solid", fgColor=DJR_OVERVIEW_FILL_COLORS["vertical"])
    top_cell.fill = fill
    bottom_cell.fill = fill
    top_cell.font = Font(bold=True, color="FF000000")
    top_cell.alignment = center_wrap()

    if target_country:
        bottom_cell.value = None
        try: ws.merge_cells(start_row=top_row, start_column=col, end_row=bottom_row, end_column=col)
        except Exception: pass
    else:
        bottom_cell.value = bottom_text
        bottom_cell.alignment = center_wrap()
        if bottom_text == "SPAIN ONLY": bottom_cell.font = Font(bold=True, color="FFFFFF00")
        else: bottom_cell.font = Font(bold=True, color="FFFF0000")

def write_djr_overview_header(ws, records):
    month_name = month_name_from_records(records)
    title_cell = get_writable_cell(ws, 5, 2)
    existing = normalize_str(title_cell.value)
    if existing: title_cell.value = existing.replace("MONTH", month_name)
    else: title_cell.value = f"DJR EMEA YT {month_name} 2026 CALENDAR"
    title_cell.alignment = center_wrap()

    first_monday = first_visible_monday_of_target_month(records)
    if first_monday:
        date_cell = get_writable_cell(ws, 9, 4)
        date_cell.value = first_monday
        date_cell.number_format = "dd-mmm"
        date_cell.alignment = center_wrap()
    return first_monday

def auto_adjust_djr_overview_heights(ws):
    for layout in DJR_OVERVIEW_WEEK_LAYOUT.values():
        rows = [
            layout["full_eps"], layout["exclusive"], layout["lifts"],
            layout["compilations"], layout["live"], *layout["podcast_rows"], *layout["vertical_rows"],
        ]
        for row in rows:
            ws.row_dimensions[row].height = max(ws.row_dimensions[row].height or 15, 42)

import re

def is_country_match(target_country: str, included_list: set) -> bool: 
    if not target_country: return True
    target_clean = re.sub(r'\s*\(.*?\)', '', target_country).strip().upper()
    for c in included_list:
        if target_clean == re.sub(r'\s*\(.*?\)', '', str(c)).strip().upper():
            return True
    return False

def populate_monthly_overview_djr(
    wb, records, enable_mediaplanning=True,
    channel_id_to_name_map=None, target_country=None 
):
    if DJR_MONTHLY_OVERVIEW_SHEET not in wb.sheetnames: return
    ws = wb[DJR_MONTHLY_OVERVIEW_SHEET]

    if target_country:
        from openpyxl.styles import PatternFill
        # Purga MASIVA celda por celda antes de escribir el local overview
        for col in range(4, 11):
            for row in range(10, 80):
                cell = ws.cell(row=row, column=col)
                cell.value = None
                cell.fill = PatternFill(fill_type=None)
                
            # EL FIX: El unmerge ahora está DENTRO del loop de las columnas
            for layout in DJR_OVERVIEW_WEEK_LAYOUT.values():
                try: ws.unmerge_cells(start_row=layout["vertical_rows"][0], start_column=col, end_row=layout["vertical_rows"][1], end_column=col)
                except: pass
                try: ws.unmerge_cells(start_row=layout["podcast_rows"][0], start_column=col, end_row=layout["podcast_rows"][1], end_column=col)
                except: pass

    first_monday = write_djr_overview_header(ws, records) 
    if not first_monday: return

    for rec in records:
        fields = rec.get("fields", rec)
        if asset_is_event(fields): continue
            
        if target_country:
            from exporters.full_calendar_djr import collect_included_countries_djr
            included = collect_included_countries_djr(record=rec, channel_id_to_name_map=channel_id_to_name_map)
            if not is_country_match(target_country, included): continue

        owner = hub_role(fields.get(AT_FIELDS["owner_hub"]))
        if owner == "Mediaplanning" and not enable_mediaplanning: continue

        live_date = parse_date(fields.get(AT_FIELDS["live_date"]))
        if not live_date: continue

        delta_days = (live_date - first_monday).days
        if delta_days < 0: continue

        week_index = (delta_days // 7) + 1
        if week_index not in DJR_OVERVIEW_WEEK_LAYOUT: continue

        day_name = DAY_ORDER[live_date.weekday()]
        col = DJR_OVERVIEW_DAY_COLS[day_name]
        content_type = normalize_str(fields.get(AT_FIELDS["content_type"]))
        bucket = get_djr_overview_bucket(content_type)
        if not bucket: continue

        layout = DJR_OVERVIEW_WEEK_LAYOUT[week_index]

        if bucket == "vertical":
            write_djr_overview_vertical(ws, layout, col, fields, target_country=target_country)
            continue

        if bucket == "podcast":
            row = layout["podcast_rows"][0]
            cell = get_writable_cell(ws, row, col)
            if normalize_str(cell.value): row = layout["podcast_rows"][1]
        else:
            row = layout[bucket]

        text = build_djr_overview_text(fields, target_country=target_country)
        write_djr_overview_regular_cell(ws, row, col, text, bucket)

    auto_adjust_djr_overview_heights(ws)
    logging.info(f"Finished populating '{DJR_MONTHLY_OVERVIEW_SHEET}'.")