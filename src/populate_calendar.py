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

from exporters.full_calendar import (
    rebuild_full_calendar_week_merges,
    write_to_day_role_columns,
    write_week_label,
    rewrite_full_calendar_owner_row,
    shade_out_of_month_days_full_calendar,
    set_full_calendar_column_widths,
    auto_adjust_full_calendar_row_heights,
    write_week_deadlines_fixed,
    write_countries_rollout_fixed,
    write_full_calendar_start_date,
    update_full_calendar_month_title,
    find_week_blocks,
    infer_week_number_for_block,
)

from exporters.monthly_overview import (
    overview_bucket,
    overview_row_offset,
    overview_visual_position,
    build_overview_lines,
    apply_overview_fill,
    write_overview_cell,
    write_vertical_short_cell,
    auto_adjust_overview_row_height,
    clear_overview_week,
    write_overview_headers,
    collect_overview_events,
    write_overview_events,
    populate_monthly_overview,
)

from config.brands import BRAND_CONFIG
from config.fields import AT_FIELDS
from config.layouts import (
    TARGET_SHEET,
    MONTHLY_OVERVIEW_SHEET,
    DAY_ORDER,
    DAY_LAYOUT,
    FULL_CALENDAR_COUNTRY_ROW_RANGES,
    FULL_CALENDAR_WEEKLY_COUNT_COL,
    OVERVIEW_DAY_COLS,
    OVERVIEW_WEEK_LAYOUT,
)
from config.styles import (
    OVERVIEW_FILL_COLORS,
    OUT_OF_MONTH_FILL,
    OVERVIEW_DAY_LABELS,
)

from config.countries import (
    SHEET_COUNTRIES,
    CHANNEL_COUNTRY_MAP,
    COUNTRY_ALIASES,
    COUNTRY_SHORT_NAMES,
    REGION_PRESETS,
    EXCLUDE_NOTE_PATTERNS,
)

from core.text_formatting import (
    normalize_str,
    lower_clean,
    bullet_join,
    tc_in_out_str,
    replace_month_in_title,
    join_comments,
)

from core.dates import (
    parse_date,
    monday_of_week,
    resolve_target_month_year,
    first_visible_monday,
    first_visible_monday_of_target_month,
    full_calendar_anchor_dates_from_records,
    month_name_from_records,
)

from core.filters import (
    filter_records_by_brand,
    filter_records_by_month,
)

from core.excel_utils import (
    get_writable_cell,
    ensure_writable_cell,
    cleanup_stale_merged_cells,
    safe_set_value,
    force_unmerged_writable_cell,
    cell_value,
    center_wrap,
    left_wrap,
    copy_cell_style,
    get_anchor_dates_from_values_wb,
    _to_indices,
)

from core.airtable_client import (
    fetch_airtable_records,
    load_local_json,
    fetch_channel_id_to_name_map,
)

from core.countries import (
    extract_excluded_countries_from_notes,
    normalize_excluded_market_name,
    format_excluded_markets_comment,
    collect_included_countries,
    resolve_region_countries,
    ordered_country_list,
    format_except_suffix,
    normalize_overview_region,
)

from services.planner_service import (
    WeekBlock,
    hub_role,
    asset_is_event,
    earliest_deadline_for_day,
    load_and_filter_records,
    build_comments_for_excel,
    assemble_text_for_single_asset,
    assemble_texts_for_multi_assets,
    get_priority_sort_value,
    aggregate_by_week_day_and_role,
)

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


BASE_DAY_TEMPLATE = {
    "Monday": {"programming": 1, "mediaplanning": 1, "uk": 1},
    "Tuesday": {"programming": 1, "mediaplanning": 1, "uk": 1},
    "Wednesday": {"programming": 1, "mediaplanning": 1, "uk": 1},
    "Thursday": {"programming": 1, "mediaplanning": 1, "uk": 1},
    "Friday": {"programming": 2, "mediaplanning": 1, "uk": 1},
    "Saturday": {"programming": 1, "mediaplanning": 1, "uk": 1},
    "Sunday": {"programming": 2, "mediaplanning": 1, "uk": 1},
}

MAX_ASSETS_PER_HUB_PER_DAY = 2


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
    parser.add_argument("--channel", default="Disney Channel", help='Filter by Brand, e.g. "Disney Channel", "Disney Jr.".')
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
    logging.info(f"Before filters: {len(records)} records")

    records = filter_records_by_brand(records, args.channel)
    logging.info(f"After brand filter '{args.channel}': {len(records)} records")

    records = filter_records_by_month(records, args.month)
    logging.info(f"After month filter '{args.month}': {len(records)} records")

    records, channel_id_to_name_map = load_and_filter_records(
        month=args.month,
        channel=args.channel,
        json_path=args.json,
        log_level=args.log_level,
)

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