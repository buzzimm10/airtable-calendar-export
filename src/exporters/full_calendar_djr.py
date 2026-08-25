import os
import logging
from collections import defaultdict, Counter
from datetime import timedelta, date

from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter

from core.text_formatting import process_multiline_branding, markdown_to_richtext

from config.fields import AT_FIELDS
from config.layouts import DAY_ORDER
from config.layouts_djr import (
    DJR_TARGET_SHEET,
    DJR_DAY_LAYOUT,
    DJR_WEEK_LAYOUT,
)
from config.countries_djr import DJR_SHEET_COUNTRIES, DJR_REGION_PRESETS
from core.dates import (
    month_name_from_records, 
    resolve_target_month_year, 
    first_visible_monday_of_target_month,
    parse_date,
)

from core.text_formatting import normalize_str, lower_clean
from core.excel_utils import safe_set_value, center_wrap, cleanup_stale_merged_cells, get_writable_cell
from core.countries import collect_included_countries
from services.planner_service import (
    hub_role,
    asset_is_event,
    assemble_text_for_single_asset,
    earliest_deadline_for_day,
)
from exporters.monthly_overview_djr import populate_monthly_overview_djr
from exporters.shared.records import aggregate_by_visual_week_day_role  
from config.fields import AT_FIELDS
import warnings
import re

warnings.filterwarnings("ignore", message="wmf image format is not supported.*")
warnings.filterwarnings("ignore", message="Data Validation extension is not supported.*")

DC_TO_DJR_COUNTRY = {
    "France (French)": "France",
    "Germany": "GSA (German)",
    "Italy (Italian)": "Italy",
    "Middle East (Arabic - MSA)": "Middle East (Arabic)",
    "Poland (Polish)": "Poland",
    "Spain": "Spain",
    "Bulgaria (Bulgarian)": "Bulgaria",
    "Greece (Greek)": "Greece",
    "Israel (Hebrew)": "Israel (Hebrew)",
    "Netherlands (Dutch)": "Netherlands (Dutch)",
    "Portugal": "Portugal",
    "Romania (Romanian)": "Romania",
    "Turkey": "Türkiye",
    "Tukey": "Türkiye",
    "Disney Kids Turkey": "Türkiye",
    "Türkiye": "Türkiye",   
    "Sweden (Swedish)": "Sweden (Swedish)",
    "Denmark (Danish)": "Denmark (Danish)",
    "Norway (Norwegian)": "Norway",
}

def get_djr_target_columns(day_name, programming_count, mediaplanning_count):
    cols = DJR_DAY_LAYOUT[day_name]

    programming_base = [col for role, col in cols if role == "Programming"][0]
    extra_col = [col for role, col in cols if role == "Extra"][0]
    mediaplanning_base = [col for role, col in cols if role == "Mediaplanning"][0]
    uk_base = [col for role, col in cols if role == "UK"][0]

    programming_cols = [programming_base]
    mediaplanning_cols = [mediaplanning_base]
    uk_cols = [uk_base]
    extra_cols = [extra_col]

    if programming_count >= 2:
        programming_cols.append(extra_col)
        extra_cols = []

    if mediaplanning_count >= 2:
        mediaplanning_cols.append(uk_base)
        uk_cols = []

    return {
        "Programming": programming_cols,
        "Mediaplanning": mediaplanning_cols,
        "UK": uk_cols,
        "Extra": extra_cols,
    }


def write_djr_owner_row(ws, owner_row, day_name, column_map):
    for col in column_map["Programming"]:
        safe_set_value(ws, owner_row, col, "PROGRAMMING HUB", alignment=center_wrap())

    for col in column_map["Mediaplanning"]:
        safe_set_value(ws, owner_row, col, "MEDIAPLANNING HUB", alignment=center_wrap())

    for col in column_map["UK"]:
        safe_set_value(ws, owner_row, col, "YT TEAM UK", alignment=center_wrap())

    for col in column_map["Extra"]:
        safe_set_value(ws, owner_row, col, "", alignment=center_wrap())


def write_djr_start_date(ws, first_monday):
    cell = get_writable_cell(ws, 7, 4)  # D7, merged D:G
    cell.value = first_monday
    cell.alignment = center_wrap()
    cell.number_format="dd-mmm-yyyy"

def write_djr_month_title(ws, records):
    month_name = month_name_from_records(records)
    # Corrección solicitada: Solo escribir "DISNEY JUNIOR YOUTUBE EMEA CALENDAR" en B1
    cell_b1 = get_writable_cell(ws, 1, 2)
    cell_b1.value = "DISNEY JUNIOR YOUTUBE EMEA CALENDAR"
    cell_b1.alignment = center_wrap()

    # Y la fecha del mes en B2 (ej. 01/10/2026)
    target_year, target_month = resolve_target_month_year(records)
    if target_year and target_month:
        cell_b2 = get_writable_cell(ws, 2, 2)
        cell_b2.value = date(target_year, target_month, 1)
        cell_b2.alignment = center_wrap()
        cell_b2.number_format = "dd/mm/yyyy"

def write_djr_week_number(ws, layout, week_monday):
    week_num = week_monday.isocalendar()[1]

    safe_set_value(ws, layout["date_row"], 2, "WEEK", alignment=center_wrap())
    safe_set_value(ws, layout["date_row"], 3, week_num, alignment=center_wrap())

def write_djr_branding(ws, layout, col, record):
    branding_layout = layout.get("branding_rows")
    if not branding_layout:
        return
        
    f = record.get("fields", record)
    align_center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    align_left = Alignment(horizontal='left', vertical='center', wrap_text=True)

    def _set_rich_cell(row, column, text_val, alignment):
        if not row:
            return
        cell = ws.cell(row=row, column=column)
        cell.value = markdown_to_richtext(text_val)
        cell.alignment = alignment

    watermark_lines = process_multiline_branding(f.get(AT_FIELDS["branding_watermark"]), max_lines=4)
    for idx, row in enumerate(branding_layout["watermark"]):
        _set_rich_cell(row, col, watermark_lines[idx], align_center)

    bug_lines = process_multiline_branding(f.get(AT_FIELDS["branding_bug"]), max_lines=3)
    for idx, row in enumerate(branding_layout["bug"]):
        _set_rich_cell(row, col, bug_lines[idx], align_center)

    copyright_text = f.get(AT_FIELDS["branding_copyright"], "")
    _set_rich_cell(branding_layout["copyright"], col, copyright_text, align_left)

    endcard_lines = process_multiline_branding(f.get(AT_FIELDS["branding_endcard"]), max_lines=6)
    for idx, row in enumerate(branding_layout["endcard"]):
        _set_rich_cell(row, col, endcard_lines[idx], align_center)

def write_djr_asset_to_col(ws, layout, col, record):
    single = assemble_text_for_single_asset(record)
    content_type = single["content_type"]

    if normalize_str(record.get(AT_FIELDS["content_type"])) in {"Editorial Short", "Exclusive"}:
        content_type = "Exclusive Short"

    if normalize_str(record.get(AT_FIELDS["content_type"])) == "Vertical Short":
        content_type = "Vertical Short"

    safe_set_value(ws, layout["property_row"], col, single["property"], alignment=center_wrap())
    safe_set_value(ws, layout["content_type_row"], col, content_type, alignment=center_wrap())
    safe_set_value(ws, layout["title_row"], col, single["title"], alignment=center_wrap())
    safe_set_value(ws, layout["synopsis_row"], col, single["synopsis"], alignment=center_wrap())
    safe_set_value(ws, layout["strategy_row"], col, single["strategy"], alignment=center_wrap())
    safe_set_value(ws, layout["tx_row"], col, single["tx"], alignment=center_wrap())
    safe_set_value(ws, layout["tc_row"], col, single["tc"], alignment=center_wrap())

    write_djr_branding(ws, layout, col, record)

    role = hub_role(record.get(AT_FIELDS["owner_hub"]))
    comments_row = layout["comments_rows"].get(role)
    if comments_row:
        fields = record.get("fields", record)
        policy = fields.get("Rights Management Policy")
        policy_str = str(policy).strip() if policy else ""
        extra_note = ""
        
        if "Block AUS/NZ" in policy_str:
            extra_note = "\n\n⚠️ **Rights Policy:** Block AUS/NZ/TAIWAN/CHINA/HONG KONG - Monetize everywhere else"
        elif "Monetize Locally" in policy_str:
            extra_note = "\n\n⚠️ **Rights Policy:** Monetize Locally (Geo-block)"
        elif "Block Everywhere" in policy_str:
            extra_note = "\n\n🚫 **Rights Policy:** Block Everywhere / Do Not Monetize"
            
        comments_text = single["comments"] + extra_note
        safe_set_value(ws, comments_row, col, markdown_to_richtext(comments_text), alignment=center_wrap())


def write_djr_deadline(ws, layout, day_name, records):
    if not records:
        return

    deadline = earliest_deadline_for_day(records)
    if not deadline:
        return

    first_col = min(col for _, col in DJR_DAY_LAYOUT[day_name])

    cell = get_writable_cell(ws, layout["deadline_row"], first_col)
    cell.value = deadline
    cell.alignment = center_wrap()
    cell.number_format = "dd-mmm-yyyy"

def collect_included_countries_djr(record, channel_id_to_name_map=None):
    region = record.get(AT_FIELDS["region"])
    included = set()

    if region:
        if isinstance(region, list):
            region_items = [lower_clean(x) for x in region]
        else:
            region_items = [
                x.strip().lower()
                for x in str(region).replace("[", "").replace("]", "").replace("'", "").replace('"', "").split(",")
                if x.strip()
            ]

        for item in region_items:
            if item in DJR_REGION_PRESETS:
                included |= DJR_REGION_PRESETS[item]

    if not included:
        included, _ = collect_included_countries(
            youtube_channel=record.get(AT_FIELDS["youtube_channel"]),
            region=record.get(AT_FIELDS["region"]),
            notes=record.get(AT_FIELDS["notes"]),
            excluded_markets=record.get(AT_FIELDS["excluded_markets"]),
            channel_id_to_name_map=channel_id_to_name_map,
        )
        included = {DC_TO_DJR_COUNTRY.get(country, country) for country in included}

    return included

def write_djr_country_rollout(ws, layout, column_map, record, col, channel_id_to_name_map):
    included = collect_included_countries_djr(record, channel_id_to_name_map)

    country_start, country_end = layout["country_rows"]
    status_start, status_end = layout["status_rows"]

    TIER_B_COUNTRIES = {"Poland", "Romania", "Netherlands (Dutch)", "Portugal", "Sweden (Swedish)", "Denmark (Danish)", "Norway", "Bulgaria"}
    
    FILL_TIER_A_C_YES = PatternFill(fill_type="solid", fgColor="EDEDED")
    FILL_TIER_B_YES = PatternFill(fill_type="solid", fgColor="FFFFFF")
    FILL_GREEN_YES = PatternFill(fill_type="solid", fgColor="00B050")

    for idx, country in enumerate(DJR_SHEET_COUNTRIES):
        rollout_row = country_start + idx
        status_row = status_start + idx

        value = "YES" if country in included else "NO"

        if value == "YES":
            font_rollout = Font(color="FFFFFFFF", bold=True)
            fill_rollout = FILL_GREEN_YES
        else:
            font_rollout = Font(color="FFFF0000", bold=True)
            fill_rollout = None 

        safe_set_value(ws, rollout_row, col, value, alignment=center_wrap(), font=font_rollout, fill=fill_rollout)

        if value == "YES":
            font_status = Font(color="00000000", bold=True)
            if country in TIER_B_COUNTRIES:
                fill_status = FILL_TIER_B_YES
            else:
                fill_status = FILL_TIER_A_C_YES
        else:
            font_status = Font(color="FFFF0000", bold=True)
            fill_status = None 

        safe_set_value(ws, status_row, col, value, alignment=center_wrap(), font=font_status, fill=fill_status)


def rebuild_djr_week_merges(ws, layout):
    rows_to_rebuild = [
        layout["date_row"],
        layout["deadline_row"],
    ]

    for cr in list(ws.merged_cells.ranges):
        if cr.min_row == cr.max_row and cr.min_row in rows_to_rebuild:
            ws.unmerge_cells(str(cr))

    for day_name in DAY_ORDER:
        cols = [col for _, col in DJR_DAY_LAYOUT[day_name]]
        start_col = min(cols)
        end_col = max(cols)

        ws.merge_cells(
            start_row=layout["date_row"],
            start_column=start_col,
            end_row=layout["date_row"],
            end_column=end_col,
        )

        ws.merge_cells(
            start_row=layout["deadline_row"],
            start_column=start_col,
            end_row=layout["deadline_row"],
            end_column=end_col,
        )

def shade_djr_out_of_month_days(ws, first_monday, target_year, target_month):
    fill = PatternFill(fill_type="solid", fgColor="D6DCE4")

    for week_index, layout in DJR_WEEK_LAYOUT.items():
        week_monday = first_monday + timedelta(days=(week_index - 1) * 7)

        for idx, day_name in enumerate(DAY_ORDER):
            current_date = week_monday + timedelta(days=idx)

            if current_date.year == target_year and current_date.month == target_month:
                continue

            cols = [col for _, col in DJR_DAY_LAYOUT[day_name]]

            for row in range(layout["owner_row"], layout["status_comments_row"] + 1):
                for col in cols:
                    ws.cell(row=row, column=col).fill = fill

def set_djr_column_widths(ws):
    for day_name in DAY_ORDER:
        for role, col in DJR_DAY_LAYOUT[day_name]:
            width = 30 if role in {"Programming", "Extra", "Mediaplanning"} else 16
            ws.column_dimensions[get_column_letter(col)].width = width

def set_djr_row_heights(ws):
    for week_num, layout in DJR_WEEK_LAYOUT.items():
        for row in [
            layout["property_row"],
            layout["content_type_row"],
            layout["title_row"],
            layout["synopsis_row"],
            layout["strategy_row"],
            layout["tx_row"],
            layout["tc_row"],
            *layout["comments_rows"].values(),
        ]:
            ws.row_dimensions[row].height = max(ws.row_dimensions[row].height or 15, 38)


# --- NUEVO: Exportar ID's ---
def export_overon_ids(wb, records):
    sheet_name = "ID's"
    if sheet_name not in wb.sheetnames:
        logging.warning(f"Sheet {sheet_name} not found. Skipping ID export.")
        return

    ws_ids = wb[sheet_name]
    start_row = 4 # Asumiendo que los datos empiezan en la fila 4 (según captura)
    current_row = start_row

    align_center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    align_left = Alignment(horizontal='left', vertical='center', wrap_text=True)

    # Filtrar y ordenar los records (solo los que tienen TX ID y TC IN/OUT)
    valid_records = []
    for r in records:
        f = r.get("fields", r)
        tx = normalize_str(f.get(AT_FIELDS["tx_id"]))
        tcin = normalize_str(f.get(AT_FIELDS["tc_in"]))
        tcout = normalize_str(f.get(AT_FIELDS["tc_out"]))
        
        if tx and tcin and tcout:
            # Intentar parsear la fecha para ordenar
            live_date_str = f.get(AT_FIELDS["live_date"])
            parsed_date = parse_date(live_date_str) if live_date_str else date.min
            valid_records.append((parsed_date, f))
            
    # Ordenar por fecha
    valid_records.sort(key=lambda x: x[0])

    for _, f in valid_records:
        asset_label = normalize_str(f.get("Asset Label", f.get(AT_FIELDS["youtube_title"])))
        date_str = ""
        live_date = parse_date(f.get(AT_FIELDS["live_date"]))
        if live_date:
            date_str = f"{live_date.day}/{live_date.month}"

        tx_raw = normalize_str(f.get(AT_FIELDS["tx_id"]))
        tcin_raw = normalize_str(f.get(AT_FIELDS["tc_in"]))
        tcout_raw = normalize_str(f.get(AT_FIELDS["tc_out"]))

        tx_lines = [x.strip() for x in tx_raw.split('\n') if x.strip()]
        tcin_lines = [x.strip() for x in tcin_raw.split('\n') if x.strip()]
        tcout_lines = [x.strip() for x in tcout_raw.split('\n') if x.strip()]

        num_items = max(len(tx_lines), len(tcin_lines), len(tcout_lines))

        # Lógica para MEGA Compilations
        is_mega = "mega" in asset_label.lower() or "mega" in normalize_str(f.get("Content Subtype", "")).lower()

        if is_mega and num_items > 1:
            # Escribir el título principal
            safe_set_value(ws_ids, current_row, 1, f"{asset_label} {date_str}", alignment=align_center)
            
            # Parsear las notas para extraer los subtítulos
            notes = normalize_str(f.get(AT_FIELDS["notes"]))
            sub_comps = [line.strip() for line in notes.split('\n') if line.strip() and "compile" not in line.lower()]
            
            # Si no hay subtítulos claros, dividimos equitativamente por defecto
            if not sub_comps:
                sub_comps = [f"Part {i+1}" for i in range(num_items)]
            
            items_per_sub = num_items // len(sub_comps) if sub_comps else 1
            remainder = num_items % len(sub_comps) if sub_comps else 0

            start_merge_a = current_row
            item_idx = 0
            
            for sub_idx, sub_title in enumerate(sub_comps):
                count_for_this_sub = items_per_sub + (1 if sub_idx < remainder else 0)
                start_merge_b = current_row
                
                for _ in range(count_for_this_sub):
                    if item_idx >= num_items: break
                    
                    tx_val = tx_lines[item_idx] if item_idx < len(tx_lines) else ""
                    tcin_val = tcin_lines[item_idx] if item_idx < len(tcin_lines) else ""
                    tcout_val = tcout_lines[item_idx] if item_idx < len(tcout_lines) else ""

                    safe_set_value(ws_ids, current_row, 3, tx_val, alignment=align_left)
                    safe_set_value(ws_ids, current_row, 4, tcin_val, alignment=align_center)
                    safe_set_value(ws_ids, current_row, 5, tcout_val, alignment=align_center)
                    current_row += 1
                    item_idx += 1
                
                # Merge Column B para este sub-compilado
                if current_row - 1 > start_merge_b:
                    ws_ids.merge_cells(start_row=start_merge_b, start_column=2, end_row=current_row-1, end_column=2)
                safe_set_value(ws_ids, start_merge_b, 2, sub_title, alignment=align_center)

            # Merge Column A para el MEGA compilado
            if current_row - 1 > start_merge_a:
                ws_ids.merge_cells(start_row=start_merge_a, start_column=1, end_row=current_row-1, end_column=1)

        elif num_items > 0:
            # Compilado normal
            start_merge = current_row
            for i in range(num_items):
                tx_val = tx_lines[i] if i < len(tx_lines) else ""
                tcin_val = tcin_lines[i] if i < len(tcin_lines) else ""
                tcout_val = tcout_lines[i] if i < len(tcout_lines) else ""

                safe_set_value(ws_ids, current_row, 2, tx_val, alignment=align_left)
                safe_set_value(ws_ids, current_row, 3, tcin_val, alignment=align_center)
                safe_set_value(ws_ids, current_row, 4, tcout_val, alignment=align_center)
                current_row += 1

            if current_row - 1 > start_merge:
                ws_ids.merge_cells(start_row=start_merge, start_column=1, end_row=current_row-1, end_column=1)
            
            safe_set_value(ws_ids, start_merge, 1, f"{asset_label} {date_str}", alignment=align_center)
        
        # Espacio entre bloques
        current_row += 1


# --- NUEVO: Exportar Metadata Programming Spain ---
def export_metadata_spain(wb, records, first_monday):
    sheet_name = "Metadata Programming Spain"
    if sheet_name not in wb.sheetnames:
        logging.warning(f"Sheet {sheet_name} not found. Skipping Metadata Spain export.")
        return

    ws_meta = wb[sheet_name]
    
    # Mapeo de semanas a filas (basado en tus instrucciones)
    # Semana 1: 6-14 | Semana 2: 16-24 | Semana 3: 26-34 | Semana 4: 36-44 | Semana 5: 46-54
    META_WEEK_LAYOUT = {
        1: 6,
        2: 16,
        3: 26,
        4: 36,
        5: 46
    }
    
    # Filas relativas dentro de cada bloque semanal
    META_ROWS = {
        "day_of_week": 0,
        "date": 1,
        "deadline": 2,
        "owner": 3,
        "property": 4,
        "content_type": 5,
        "title": 6,
        "synopsis": 7,
        "strategy": 8
    }

    # Columnas por día (simplificadas para este tab, asumimos 1 columna por día C, D, E...)
    # Asumiendo C=3, D=4, E=5, F=6, G=7, H=8, I=9
    META_DAY_COLS = {
        "Monday": 3, "Tuesday": 4, "Wednesday": 5, "Thursday": 6, "Friday": 7, "Saturday": 8, "Sunday": 9
    }

    align_center = Alignment(horizontal='center', vertical='center', wrap_text=True)

    # Filtrar solo Programming
    prog_records = [r for r in records if hub_role(r.get("fields", r).get(AT_FIELDS["owner_hub"])) == "Programming"]
    buckets = aggregate_by_visual_week_day_role(prog_records, first_monday)

    for week_index, start_row in META_WEEK_LAYOUT.items():
        week_monday = first_monday + timedelta(days=(week_index - 1) * 7)
        
        for idx, day_name in enumerate(DAY_ORDER):
            col = META_DAY_COLS[day_name]
            current_date = week_monday + timedelta(days=idx)
            
            # Escribir cabeceras de día y fecha
            safe_set_value(ws_meta, start_row + META_ROWS["day_of_week"], col, day_name.upper(), alignment=align_center)
            safe_set_value(ws_meta, start_row + META_ROWS["date"], col, current_date, alignment=align_center, number_format="dd-mmm-yy")
            
            day_records = buckets.get((week_index, day_name, "Programming"), [])
            
            if not day_records:
                continue
                
            # Tomamos el primer record para llenar
            record = day_records[0]
            f = record.get("fields", record)
            single = assemble_text_for_single_asset(record)
            
            deadline = earliest_deadline_for_day(day_records)
            
            # Obtener ES Title y ES Description (Nuevos campos Airtable)
            es_title = normalize_str(f.get("ES Title", single["title"])) # Fallback al inglés si está vacío
            es_desc = normalize_str(f.get("ES Description", single["synopsis"]))

            safe_set_value(ws_meta, start_row + META_ROWS["deadline"], col, deadline, alignment=align_center, number_format="dd-mmm-yy")
            safe_set_value(ws_meta, start_row + META_ROWS["owner"], col, "PROGRAMMING HUB", alignment=align_center)
            safe_set_value(ws_meta, start_row + META_ROWS["property"], col, single["property"], alignment=align_center)
            safe_set_value(ws_meta, start_row + META_ROWS["content_type"], col, single["content_type"], alignment=align_center)
            safe_set_value(ws_meta, start_row + META_ROWS["title"], col, es_title, alignment=align_center)
            safe_set_value(ws_meta, start_row + META_ROWS["synopsis"], col, es_desc, alignment=align_center)
            safe_set_value(ws_meta, start_row + META_ROWS["strategy"], col, single["strategy"], alignment=align_center)



def populate_calendar_djr(
    template_path: str,
    output_path: str,
    records: list,
    enable_mediaplanning: bool = True,
    channel_id_to_name_map: dict | None = None,
):
    print("ENTERED populate_calendar_djr")
    print("DJR template_path:", template_path)
    print("DJR output_path:", output_path)
    print("DJR records:", len(records))
    
    if channel_id_to_name_map is None:
        channel_id_to_name_map = {}

    wb = load_workbook(template_path, data_only=False)

    if DJR_TARGET_SHEET not in wb.sheetnames:
        raise RuntimeError(f"Target sheet '{DJR_TARGET_SHEET}' not found in workbook.")

    ws = wb[DJR_TARGET_SHEET]

    logging.info(f"DJR records received: {len(records)}")

    for r in records[:5]:
        f = r.get("fields", r)
        logging.info(
            "DJR sample: live_date=%s | brand=%s | owner=%s | content_type=%s | property=%s",
            f.get(AT_FIELDS["live_date"]),
            f.get(AT_FIELDS["brand"]),
            f.get(AT_FIELDS["owner_hub"]),
            f.get(AT_FIELDS["content_type"]),
            f.get(AT_FIELDS["property"]),
        )

    set_djr_column_widths(ws)
    write_djr_month_title(ws, records)

    first_monday = first_visible_monday_of_target_month(records)
    
    if not first_monday:
        logging.warning("No se pudo determinar el primer lunes. Abortando generación.")
        return

    write_djr_start_date(ws, first_monday)
    print("DJR first_monday:", first_monday)

    buckets = aggregate_by_visual_week_day_role(records, first_monday)

    logging.info(f"DJR buckets created: {len(buckets)}")

    for key, items in list(buckets.items())[:10]:
        logging.info(f"DJR bucket {key}: {len(items)} records")

    for week_index, layout in DJR_WEEK_LAYOUT.items():
        if not first_monday:
            continue

        week_monday = first_monday + timedelta(days=(week_index - 1) * 7)

        logging.info(f"Processing DJR week {week_index}: {week_monday}")

        write_djr_week_number(ws, layout, week_monday)

        for day_name in DAY_ORDER:
            prog_records = buckets.get((week_index, day_name, "Programming"), [])
            media_records = buckets.get((week_index, day_name, "Mediaplanning"), [])

            if not enable_mediaplanning:
                media_records = []

            column_map = get_djr_target_columns(
                day_name=day_name,
                programming_count=len(prog_records),
                mediaplanning_count=len(media_records),
            )

            write_djr_owner_row(ws, layout["owner_row"], day_name, column_map)

            all_day_records = prog_records + media_records
            write_djr_deadline(ws, layout, day_name, all_day_records)

            for idx, record in enumerate(prog_records[:2]):
                col = column_map["Programming"][idx]
                write_djr_asset_to_col(ws, layout, col, record)
                write_djr_country_rollout(ws, layout, column_map, record, col, channel_id_to_name_map)

            for idx, record in enumerate(media_records[:2]):
                col = column_map["Mediaplanning"][idx]
                write_djr_asset_to_col(ws, layout, col, record)
                write_djr_country_rollout(ws, layout, column_map, record, col, channel_id_to_name_map)

        rebuild_djr_week_merges(ws, layout)

    target_year, target_month = resolve_target_month_year(records)

    if first_monday and target_year and target_month:
        shade_djr_out_of_month_days(ws, first_monday, target_year, target_month)

    set_djr_row_heights(ws)

    # Exportar las nuevas hojas de cálculo
    export_overon_ids(wb, records)
    export_metadata_spain(wb, records, first_monday)

    try:
        populate_monthly_overview_djr(
            wb=wb,
            records=records,
            enable_mediaplanning=enable_mediaplanning,
            channel_id_to_name_map=channel_id_to_name_map,
        )
    except Exception as e:
        logging.warning(f"Could not populate DJR monthly overview yet: {e}")

    out_dir = os.path.dirname(output_path)
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)

    cleanup_stale_merged_cells(ws)
    wb.template = False

    month_name_upper = month_name_from_records(records).capitalize()
    target_year, target_month = resolve_target_month_year(records)
    month_abbr = month_name_upper[:3] 

    for sn in wb.sheetnames:
        if "highlights" in sn.lower():
            new_title = f"YT DJR EMEA {month_abbr} Highlights"
            wb[sn].title = new_title
            
            if target_year and target_month:
                ws_high = wb[new_title]
                safe_set_value(ws_high, 3, 5, date(target_year, target_month, 1), number_format="mmm-yy")

    cleanup_stale_merged_cells(ws) 
    wb.save(output_path)

    logging.info(f"Saved DJR calendar to: {output_path}")