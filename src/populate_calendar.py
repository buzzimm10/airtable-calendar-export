#!/usr/bin/env python3
"""
Populate the 'FULL CALENDAR YT DC EMEA' sheet of an existing Excel template
from Airtable (or local JSON) using openpyxl.

PATCHES (2026-03-25):
- Week label: write "WEEK <Week Number (Year)>" into merged cells B6, B60, B114, B168, B222, B276 (per week block).
  Value inferred from Airtable records of that week; fallback to ISO week from Monday date.
- Countries by Hub: keep writing YES/NO in the correct hub column; uniform font.
- All prior behavior preserved: channel/month filters, merged-cell safe writing,
  dual Programming Fri/Sun, optional Mediaplanning, UK blank, real date headers.
"""
import argparse
import os
import re
import sys
import logging
from typing import Dict, List, Optional, Tuple
from openpyxl import load_workbook
from dotenv import load_dotenv
from datetime import timedelta, date

load_dotenv()  # carga .env del cwd (raíz del proyecto)

from exporters.full_calendar import (
    rebuild_full_calendar_week_merges,
    write_week_label,
    rewrite_full_calendar_owner_row,
    shade_out_of_month_days_full_calendar,
    set_full_calendar_column_widths,
    auto_adjust_full_calendar_row_heights,
    write_week_deadlines_fixed,
    write_full_calendar_start_date,
    update_full_calendar_month_title,
    find_week_blocks,
    infer_week_number_for_block,
)

from services.local_calendars_service import generate_local_calendars
from exporters.monthly_overview import populate_monthly_overview
from exporters.full_calendar_djr import populate_calendar_djr

from config.layouts import TARGET_SHEET, DAY_ORDER, DAY_LAYOUT
from config.template_config import get_template_config
from config.fields import AT_FIELDS
from core.dates import (
    resolve_target_month_year,
    first_visible_monday_of_target_month,
    full_calendar_anchor_dates_from_records,
    month_name_from_records,
)
from core.filters import filter_records_by_brand, filter_records_by_month
from core.excel_utils import cleanup_stale_merged_cells, get_anchor_dates_from_values_wb, safe_set_value, center_wrap
from core.text_formatting import normalize_str, lower_clean, process_multiline_branding, markdown_to_richtext
from core.countries import collect_included_countries
from core.airtable_client import fetch_airtable_records, load_local_json, fetch_channel_id_to_name_map
from services.planner_service import load_and_filter_records, aggregate_by_week_day_and_role, assemble_text_for_single_asset, hub_role

from openpyxl.styles import Alignment, Font, PatternFill


# --- NUEVO: SLOTS INTELIGENTES Y ESCRITURA SIN LÍMITES PARA DISNEY CHANNEL ---
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

def write_owner_row(ws, owner_row, day_name, column_map):
    for col in column_map["Programming"]: safe_set_value(ws, owner_row, col, "PROGRAMMING HUB", alignment=center_wrap())
    for col in column_map["Mediaplanning"]: safe_set_value(ws, owner_row, col, "MEDIAPLANNING HUB", alignment=center_wrap())
    for col in column_map["UK"]: safe_set_value(ws, owner_row, col, "YT TEAM UK", alignment=center_wrap())
    for col in column_map["Extra"]: safe_set_value(ws, owner_row, col, "", alignment=center_wrap())

def write_asset_to_col_dc(ws, block, col, record):
    single = assemble_text_for_single_asset(record)
    content_type = single.get("content_type", "")
    if normalize_str(record.get("fields", {}).get(AT_FIELDS["content_type"])) in {"Editorial Short", "Exclusive"}:
        content_type = "Exclusive Short"
    if normalize_str(record.get("fields", {}).get(AT_FIELDS["content_type"])) == "Vertical Short":
        content_type = "Vertical Short"

    safe_set_value(ws, block.property_row, col, single.get("property", ""), alignment=center_wrap())
    safe_set_value(ws, block.content_type_row, col, content_type, alignment=center_wrap())
    safe_set_value(ws, block.title_row, col, single.get("title", ""), alignment=center_wrap())
    safe_set_value(ws, block.synopsis_row, col, single.get("synopsis", ""), alignment=center_wrap())
    
    strat_val = record.get("fields", record).get(AT_FIELDS["strategy"])
    actual_strategy = " / ".join(normalize_str(s) for s in strat_val if normalize_str(s)) if isinstance(strat_val, list) else normalize_str(strat_val)
    safe_set_value(ws, block.strategy_row, col, actual_strategy, alignment=center_wrap())
    
    if hasattr(block, 'tx_row') and block.tx_row: safe_set_value(ws, block.tx_row, col, single.get("tx", ""), alignment=center_wrap())
    if hasattr(block, 'tc_row') and block.tc_row: safe_set_value(ws, block.tc_row, col, single.get("tc", ""), alignment=center_wrap())

    f = record.get("fields", record)
    if hasattr(block, 'branding_rows') and block.branding_rows:
        bl = block.branding_rows
        ac = Alignment(horizontal='center', vertical='center', wrap_text=True)
        al = Alignment(horizontal='left', vertical='center', wrap_text=True)

        def _set_rich_cell(row, column, text_val, alignment):
            if row: ws.cell(row=row, column=column, value=markdown_to_richtext(text_val)).alignment = alignment

        wt_lines = process_multiline_branding(f.get(AT_FIELDS["branding_watermark"]), max_lines=4)
        for i, r in enumerate(bl.get("watermark", [])): _set_rich_cell(r, col, wt_lines[i], ac)

        bug_lines = process_multiline_branding(f.get(AT_FIELDS["branding_bug"]), max_lines=3)
        for i, r in enumerate(bl.get("bug", [])): _set_rich_cell(r, col, bug_lines[i], ac)

        if bl.get("copyright"): _set_rich_cell(bl["copyright"], col, f.get(AT_FIELDS["branding_copyright"], ""), al)

        ec_lines = process_multiline_branding(f.get(AT_FIELDS["branding_endcard"]), max_lines=6)
        for i, r in enumerate(bl.get("endcard", [])): _set_rich_cell(r, col, ec_lines[i], ac)

    if hasattr(block, 'comments_rows') and block.comments_rows:
        role = hub_role(f.get(AT_FIELDS["owner_hub"]))
        crow = block.comments_rows.get(role)
        if crow: safe_set_value(ws, crow, col, markdown_to_richtext(single.get("comments", "")), alignment=center_wrap())

def write_country_rollout_dc(ws, block, record, col, channel_id_to_name_map):
    f = record.get("fields", record)
    included, _ = collect_included_countries(
        youtube_channel=f.get(AT_FIELDS["youtube_channel"]),
        region=f.get(AT_FIELDS["region"]),
        notes=f.get(AT_FIELDS["notes"]),
        excluded_markets=f.get(AT_FIELDS["excluded_markets"]),
        channel_id_to_name_map=channel_id_to_name_map,
    )
    FILL_GREEN = PatternFill(fill_type="solid", fgColor="00B050")
    for country, row_idx in getattr(block, 'rollout_country_rows', {}).items():
        c_clean = country.strip().lower()
        is_in = any(c_clean == inc.strip().lower() for inc in included)
        val = "YES" if is_in else "NO"
        font = Font(color="000000", bold=True)
        fill = FILL_GREEN if is_in else None
        if val == "NO": font = Font(color="FF0000", bold=True)
        if c_clean == "italy": font = Font(color="FF0000", bold=True)
        safe_set_value(ws, row_idx, col, val, alignment=center_wrap(), font=font, fill=fill)

    for country, row_idx in getattr(block, 'status_country_rows', {}).items():
        c_clean = country.strip().lower()
        is_in = any(c_clean == inc.strip().lower() for inc in included)
        val = "YES" if is_in else "NO"
        font = Font(color="000000", bold=True)
        if val == "NO": font = Font(color="FF0000", bold=True)
        safe_set_value(ws, row_idx, col, val, alignment=center_wrap(), font=font)
# -----------------------------------------------------------------------------

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
            template_path=template_path, output_path=output_path, records=records,
            enable_mediaplanning=enable_mediaplanning, channel_id_to_name_map=channel_id_to_name_map
        )
        zip_path = None
        if selected_countries != []:
            target_year, target_month = resolve_target_month_year(records)
            m_str = f"{target_year}-{target_month:02d}" if target_year else "2026-05"
            zip_path = generate_local_calendars(
                master_path=output_path, output_dir="country_calendars", channel_type="DJR",
                selected_countries=selected_countries, month_str=m_str,
                records=records, channel_id_to_name_map=channel_id_to_name_map
            )
        return zip_path

    # === LÓGICA DE EXPORTACIÓN PARA DISNEY CHANNEL ===
    if channel_id_to_name_map is None: channel_id_to_name_map = {}
    wb = load_workbook(template_path, data_only=False)
    active_target_sheet = "FULL CALENDAR YT DC EMEA"
    if active_target_sheet not in wb.sheetnames: raise RuntimeError(f"Target sheet '{active_target_sheet}' not found in workbook.")
    ws = wb[active_target_sheet]

    set_full_calendar_column_widths(ws)
    update_full_calendar_month_title(ws, records)

    # --- FORZAR 6 SEMANAS DINÁMICAMENTE ---
    start_monday = first_visible_monday_of_target_month(records)
    if start_monday: write_full_calendar_start_date(ws, start_monday)

    new_coords = ["D6", "D60", "D114", "D168", "D222", "D276"]
    anchor_dates_map = {}
    if start_monday:
        for i, coord in enumerate(new_coords): anchor_dates_map[coord] = start_monday + timedelta(days=i*7)

    blocks = find_week_blocks(ws, anchor_dates_map)
    buckets = aggregate_by_week_day_and_role(records)
    target_year, target_month = resolve_target_month_year(records)

    # PARCHE DINÁMICO: Añadir Semana 6 a Monthly Overview DC
    import exporters.monthly_overview
    if hasattr(exporters.monthly_overview, 'OVERVIEW_WEEK_LAYOUT'):
        mo_layout = exporters.monthly_overview.OVERVIEW_WEEK_LAYOUT
        if 5 in mo_layout and 6 not in mo_layout:
            diff = mo_layout[5]["date_row"] - mo_layout[4]["date_row"]
            w6 = {}
            for k, v in mo_layout[5].items():
                if isinstance(v, tuple): w6[k] = (v[0] + diff, v[1] + diff)
                else: w6[k] = v + diff
            mo_layout[6] = w6

    ORIGINAL_DAY_COLS = {day: sorted([col for _, col in DAY_LAYOUT[day]]) for day in DAY_ORDER}

    for mon_date, block in blocks.items():
        per_day_role_records: Dict[Tuple[str, str], List[Dict]] = {}
        per_day_all_roles: Dict[str, List[Dict]] = {d: [] for d in DAY_ORDER}

        for day_name in DAY_ORDER:
            prog = buckets.get((mon_date, day_name, "Programming"), [])
            media = buckets.get((mon_date, day_name, "Mediaplanning"), [])
            per_day_role_records[(day_name, "Programming")] = prog
            per_day_role_records[(day_name, "Mediaplanning")] = media if enable_mediaplanning else []
            per_day_all_roles[day_name] = prog + (media if enable_mediaplanning else [])

            # --- REESCRITURA DINÁMICA DEL DAY_LAYOUT ---
            column_map = get_target_columns(day_name, len(prog), len(media if enable_mediaplanning else []), ORIGINAL_DAY_COLS)
            new_day_layout = []
            for col in column_map["Programming"]: new_day_layout.append(("Programming", col))
            for col in column_map["Mediaplanning"]: new_day_layout.append(("Mediaplanning", col))
            for col in column_map["UK"]: new_day_layout.append(("UK", col))
            for col in column_map["Extra"]: new_day_layout.append(("Extra", col))
            DAY_LAYOUT[day_name] = new_day_layout

        week_num_text = infer_week_number_for_block(per_day_all_roles, block.monday_date)
        write_week_label(ws, block, week_num_text)
        rewrite_full_calendar_owner_row(ws, block, DAY_LAYOUT)

        # ESCRITURA DIRECTA SIN LÍMITES
        for day_name in DAY_ORDER:
            prog_records = per_day_role_records[(day_name, "Programming")]
            media_records = per_day_role_records[(day_name, "Mediaplanning")]
            column_map = get_target_columns(day_name, len(prog_records), len(media_records), ORIGINAL_DAY_COLS)

            for idx, record in enumerate(prog_records[:len(column_map["Programming"])]):
                col = column_map["Programming"][idx]
                write_asset_to_col_dc(ws, block, col, record)
                write_country_rollout_dc(ws, block, record, col, channel_id_to_name_map)

            for idx, record in enumerate(media_records[:len(column_map["Mediaplanning"])]):
                col = column_map["Mediaplanning"][idx]
                write_asset_to_col_dc(ws, block, col, record)
                write_country_rollout_dc(ws, block, record, col, channel_id_to_name_map)

        write_week_deadlines_fixed(ws, block, per_day_all_roles)
        rebuild_full_calendar_week_merges(ws, block)

    # SOMBREADO FUERA DE MES: Automáticamente abarca la semana 6
    if target_year and target_month:
        shade_out_of_month_days_full_calendar(ws, blocks, target_year, target_month)

    auto_adjust_full_calendar_row_heights(ws, blocks)
    populate_monthly_overview(wb=wb, records=records, enable_mediaplanning=enable_mediaplanning, channel_id_to_name_map=channel_id_to_name_map)

    out_dir = os.path.dirname(output_path)
    if out_dir: os.makedirs(out_dir, exist_ok=True)
    if output_path.lower().endswith(".xlsx"): wb.template = False
    elif output_path.lower().endswith(".xltx"): wb.template = True
    
    cleanup_stale_merged_cells(ws) 
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
            master_path=output_path, output_dir="country_calendars", channel_type="DC",
            selected_countries=selected_countries, month_str=m_str,
            dc_blocks=blocks, records=records, channel_id_to_name_map=channel_id_to_name_map 
        )
    return zip_path

def main():
    parser = argparse.ArgumentParser(description="Populate Disney Channel EMEA YouTube calendar into existing template.")
    parser.add_argument("--template", help="Path to existing Excel template (.xltx or .xlsx). If omitted, uses brand default.")
    parser.add_argument("--output", required=True, help="Path to output .xlsx file.")
    parser.add_argument("--airtable-token", default=os.getenv("AIRTABLE_API_TOKEN"), help="Airtable API token or set AIRTABLE_API_TOKEN.")
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

    enable_mediaplanning = not args.disable_mediaplanning
    env_mp = os.getenv("ENABLE_MEDIAPLANNING")
    if env_mp is not None and not args.disable_mediaplanning:
        enable_mediaplanning = env_mp.strip().lower() in ("1", "true", "yes", "y", "on")

    logging.basicConfig(stream=sys.stdout, level=getattr(logging, args.log_level.upper(), logging.INFO), format="%(levelname)s: %(message)s")
    channel_id_to_name_map = {}

    if args.json:
        records = load_local_json(args.json)
        if token and base_id: channel_id_to_name_map = fetch_channel_id_to_name_map(token, base_id)
    else:
        if not token or not base_id or not table_id:
            parser.error(f"Missing Airtable config.")
        records = fetch_airtable_records(api_token=token, base_id=base_id, table_id=table_id, view=view)
        channel_id_to_name_map = fetch_channel_id_to_name_map(token, base_id)

    records = filter_records_by_brand(records, args.channel)
    records = filter_records_by_month(records, args.month)
    records, channel_id_to_name_map = load_and_filter_records(month=args.month, channel=args.channel, json_path=args.json, log_level=args.log_level)

    week_anchors = ["D6", "D60", "D114", "D168", "D222", "D276"]
    template_path = args.template or str(template_config["template_path"])

    populate_calendar(
        template_path=template_path, output_path=args.output, records=records, channel=args.channel,
        week_anchor_cells=week_anchors, enable_mediaplanning=enable_mediaplanning, anchor_dates_override=args.anchor_dates,
        channel_id_to_name_map=channel_id_to_name_map, selected_countries=args.countries
    )

if __name__ == "__main__":
    main()