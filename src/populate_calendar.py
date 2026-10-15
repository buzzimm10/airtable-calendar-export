#!/usr/bin/env python3
"""
Populate the 'FULL CALENDAR YT DC EMEA' sheet of an existing Excel template
from Airtable (or local JSON) using openpyxl.

PATCHES (2026-03-25):
- Week label: write "WEEK <Week Number (Year)>" into merged cells B6, B60, B114, B168, B222, B276 (per week block).
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
import logging
from typing import Dict, List, Optional, Tuple
from openpyxl import load_workbook
from dotenv import load_dotenv
load_dotenv()  # carga .env del cwd (raíz del proyecto)

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

from services.local_calendars_service import generate_local_calendars
from exporters.monthly_overview import populate_monthly_overview
from exporters.full_calendar_djr import populate_calendar_djr

from config.layouts import (
    TARGET_SHEET,
    DAY_ORDER,
    DAY_LAYOUT,
)
from config.template_config import get_template_config
from core.dates import (
    resolve_target_month_year,
    first_visible_monday_of_target_month,
    full_calendar_anchor_dates_from_records,
    month_name_from_records,
)
from core.filters import (
    filter_records_by_brand,
    filter_records_by_month,
)
from core.excel_utils import cleanup_stale_merged_cells, get_anchor_dates_from_values_wb, safe_set_value
from core.airtable_client import (
    fetch_airtable_records,
    load_local_json,
    fetch_channel_id_to_name_map,
)
from services.planner_service import load_and_filter_records, aggregate_by_week_day_and_role

# --- SLOTS INTELIGENTES PARA DISNEY CHANNEL ---
def get_target_columns(day_name, programming_count, mediaplanning_count, original_day_cols):
    all_cols = original_day_cols[day_name]
    
    p_count = min(programming_count, 4)
    m_count = min(mediaplanning_count, 4 - p_count)
    
    prog_slots = [0, 1, 2, 3][:p_count]
    prog_cols = [all_cols[i] for i in prog_slots]
    
    remaining_slots = [i for i in [2, 1, 0, 3] if i not in prog_slots]
    media_slots = remaining_slots[:m_count]
    media_cols = [all_cols[i] for i in media_slots]
    
    used_slots = set(prog_slots + media_slots)
    
    uk_cols = []
    if 3 not in used_slots:
        uk_cols = [all_cols[3]]
        used_slots.add(3)
        
    extra_cols = [all_cols[i] for i in [0, 1, 2, 3] if i not in used_slots]
    
    return {"Programming": prog_cols, "Mediaplanning": media_cols, "UK": uk_cols, "Extra": extra_cols}
# -----------------------------------------------

def populate_calendar(template_path: str, output_path: str, records: List[Dict],
                      channel: str = "Disney Channel", 
                      week_anchor_cells: Optional[List[str]] = None,
                      enable_mediaplanning: bool = True,
                      anchor_dates_override: Optional[List[str]] = None,
                      channel_id_to_name_map: Optional[Dict[str, str]] = None,
                      selected_countries: Optional[List[str]] = None):

    if channel == "Disney Junior":
        logging.info("Routing to Disney Junior dedicated exporter...")
        populate_calendar_djr(
            template_path=template_path,
            output_path=output_path,
            records=records,
            enable_mediaplanning=enable_mediaplanning,
            channel_id_to_name_map=channel_id_to_name_map
        )
        
        zip_path = None
        if selected_countries != []:
            target_year, target_month = resolve_target_month_year(records)
            m_str = f"{target_year}-{target_month:02d}" if target_year else "2026-05"
            zip_path = generate_local_calendars(
                master_path=output_path, 
                output_dir="country_calendars", 
                channel_type="DJR",
                selected_countries=selected_countries,
                month_str=m_str,
                records=records,
                channel_id_to_name_map=channel_id_to_name_map
            )
        return zip_path

    # === LÓGICA DE EXPORTACIÓN PARA DISNEY CHANNEL ===
    week_anchor_cells = week_anchor_cells or ["D6", "D60", "D114", "D168", "D222", "D276"]
    if channel_id_to_name_map is None:
        channel_id_to_name_map = {}

    wb = load_workbook(template_path, data_only=False)
    active_target_sheet = "FULL CALENDAR YT DC EMEA"

    if active_target_sheet not in wb.sheetnames:
        raise RuntimeError(f"Target sheet '{active_target_sheet}' not found in workbook.")
    
    ws = wb[active_target_sheet]

    set_full_calendar_column_widths(ws)
    update_full_calendar_month_title(ws, records)

    anchor_dates_map = full_calendar_anchor_dates_from_records(records)

    # PARCHE COORDENADAS: Añadida la Semana 6 (D276)
    if anchor_dates_map and channel == "Disney Channel":
        new_coords = ["D6", "D60", "D114", "D168", "D222", "D276"]
        updated_map = {}
        for idx, (old_addr, date_val) in enumerate(anchor_dates_map.items()):
            if idx < len(new_coords):
                updated_map[new_coords[idx]] = date_val
        anchor_dates_map = updated_map

    if not anchor_dates_map:
        anchor_dates_map = get_anchor_dates_from_values_wb(template_path, week_anchor_cells)

    blocks = find_week_blocks(ws, anchor_dates_map)
    buckets = aggregate_by_week_day_and_role(records)

    start_monday = first_visible_monday_of_target_month(records)
    if start_monday:
        write_full_calendar_start_date(ws, start_monday)

    target_year, target_month = resolve_target_month_year(records)

    # Guardamos una copia segura de las columnas base del Layout
    ORIGINAL_DAY_COLS = {
        day: sorted([col for _, col in DAY_LAYOUT[day]])
        for day in DAY_ORDER
    }

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

            # --- HACK DINÁMICO: Reescribimos el DAY_LAYOUT en memoria ---
            column_map = get_target_columns(day_name, len(prog), len(media if enable_mediaplanning else []), ORIGINAL_DAY_COLS)
            new_day_layout = []
            for col in column_map["Programming"]: new_day_layout.append(("Programming", col))
            for col in column_map["Mediaplanning"]: new_day_layout.append(("Mediaplanning", col))
            for col in column_map["UK"]: new_day_layout.append(("UK", col))
            for col in column_map["Extra"]: new_day_layout.append(("Extra", col))
            
            DAY_LAYOUT[day_name] = new_day_layout
            # -------------------------------------------------------------

        week_num_text = infer_week_number_for_block(per_day_all_roles, block.monday_date)
        write_week_label(ws, block, week_num_text)
        
        # Ahora que DAY_LAYOUT tiene los slots inteligentes de esta semana, escribimos cabeceras y contenidos
        rewrite_full_calendar_owner_row(ws, block, DAY_LAYOUT)

        for day_name in DAY_ORDER:
            write_to_day_role_columns(ws, block, day_name, "Programming",
                                      per_day_role_records[(day_name, "Programming")], enable_mediaplanning)
            write_to_day_role_columns(ws, block, day_name, "Mediaplanning",
                                      per_day_role_records[(day_name, "Mediaplanning")], enable_mediaplanning)

        write_week_deadlines_fixed(ws, block, per_day_all_roles)
        write_countries_rollout_fixed(ws, block, per_day_role_records, channel_id_to_name_map)
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
    
    # ---> CAMBIAR NOMBRE PESTAÑA HIGHLIGHTS Y FECHA E4 <---
    from datetime import date
    month_name_upper = month_name_from_records(records).capitalize()
    month_abbr = month_name_upper[:3] 
    
    for sn in wb.sheetnames:
        if "highlights" in sn.lower():
            new_title = f"YT DC EMEA {month_abbr} Highlights"
            wb[sn].title = new_title
            if target_year and target_month:
                ws_high = wb[new_title]
                safe_set_value(ws_high, 4, 5, date(target_year, target_month, 1), number_format="mmm-yy")

    cleanup_stale_merged_cells(ws) 
    wb.save(output_path)
    
    # ---> FASE 2: CALENDARIOS LOCALES (DC) <---
    zip_path = None
    if selected_countries != []:
        target_year, target_month = resolve_target_month_year(records)
        m_str = f"{target_year}-{target_month:02d}" if target_year else "2026-05"
        zip_path = generate_local_calendars(
            master_path=output_path, 
            output_dir="country_calendars", 
            channel_type="DC",
            selected_countries=selected_countries,
            month_str=m_str,
            dc_blocks=blocks,
            records=records,
            channel_id_to_name_map=channel_id_to_name_map 
        )
    
    logging.info("Done.")
    return zip_path

def main():
    parser = argparse.ArgumentParser(description="Populate Disney Channel EMEA YouTube calendar into existing template.")
    parser.add_argument("--template", help="Path to existing Excel template (.xltx or .xlsx). If omitted, uses brand default.")
    parser.add_argument("--output", required=True, help="Path to output .xlsx file.")
    parser.add_argument("--airtable-token", default=os.getenv("AIRTABLE_API_TOKEN"),
                        help="Airtable API token or set AIRTABLE_API_TOKEN.")
    parser.add_argument("--airtable-base", help="Airtable Base ID.")
    parser.add_argument("--airtable-table", help="Airtable Table ID or name.")
    parser.add_argument("--airtable-view", help="Optional Airtable view name.")
    parser.add_argument("--json", help="Optional path to local JSON export instead of API.")
    parser.add_argument("--anchors", nargs="*", help="Override weekly anchor cells, e.g., D6 D59 D112 D165 D218 D276.")
    parser.add_argument("--anchor-dates", nargs="*", help="Optional ISO dates for anchors.")
    parser.add_argument("--disable-mediaplanning", action="store_true", help="If set, do not populate Mediaplanning Hub content.")
    parser.add_argument("--countries", nargs="*", help="Lista específica de países a exportar para los calendarios locales.")
    parser.add_argument("--channel", default="Disney Channel", help='Filter by Brand, e.g. "Disney Channel", "Disney Junior".')
    parser.add_argument("--month", help='Filter by "Month" (optional, case-insensitive).')
    parser.add_argument("--log-level", default="INFO", help="Logging level (DEBUG, INFO, WARNING, ERROR).")
    args = parser.parse_args()

    template_config = get_template_config(args.channel)

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

    enable_mediaplanning = not args.disable_mediaplanning
    env_mp = os.getenv("ENABLE_MEDIAPLANNING")
    if env_mp is not None and not args.disable_mediaplanning:
        enable_mediaplanning = env_mp.strip().lower() in ("1", "true", "yes", "y", "on")

    logging.basicConfig(stream=sys.stdout, level=getattr(logging, args.log_level.upper(), logging.INFO),
                        format="%(levelname)s: %(message)s")

    channel_id_to_name_map = {}

    if args.json:
        logging.info(f"Loading records from JSON: {args.json}")
        records = load_local_json(args.json)
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

    # Añadida la Semana 6 (D276)
    week_anchors = args.anchors if args.anchors else ["D6", "D60", "D114", "D168", "D222", "D276"]
    template_path = args.template or str(template_config["template_path"])
    
    if args.anchors:
        week_anchors = args.anchors
    elif "week_anchor_cells" in template_config:
        config_anchors = template_config["week_anchor_cells"]
        week_anchors = config_anchors if len(config_anchors) >= 6 else ["D6", "D60", "D114", "D168", "D222", "D276"]

    populate_calendar(
        template_path=template_path,
        output_path=args.output,
        records=records,
        channel=args.channel,
        week_anchor_cells=week_anchors,
        enable_mediaplanning=enable_mediaplanning,
        anchor_dates_override=args.anchor_dates,
        channel_id_to_name_map=channel_id_to_name_map,
        selected_countries=args.countries,
    )

if __name__ == "__main__":
    main()