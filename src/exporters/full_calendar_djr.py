import os
import logging
from collections import defaultdict, Counter
from datetime import timedelta

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
    parse_date,
    monday_of_week,
    first_visible_monday_of_target_month,
    resolve_target_month_year,
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
from core.text_formatting import process_multiline_branding, markdown_to_richtext
from core.excel_utils import safe_set_value, center_wrap
from config.fields import AT_FIELDS
import warnings

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
    from core.dates import month_name_from_records

    month_name = month_name_from_records(records)

    cell = get_writable_cell(ws, 2, 2)  # B2 merged B:W
    cell.value = month_name
    cell.alignment = center_wrap()
    # cleanup_stale_merged_cells(ws)


def write_djr_week_number(ws, layout, week_monday):
    week_num = week_monday.isocalendar()[1]

    safe_set_value(ws, layout["date_row"], 2, "WEEK", alignment=center_wrap())
    safe_set_value(ws, layout["date_row"], 3, week_num, alignment=center_wrap())

def write_djr_branding(ws, layout, col, record):
    """
    Inyecta la sección BRANDING en DJR respetando las reglas de filas y Markdown.
    Pasa absolutamente todos los campos por markdown_to_richtext.
    """
    branding_layout = layout.get("branding_rows")
    if not branding_layout:
        return
        
    f = record.get("fields", record)
    align_center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    align_left = Alignment(horizontal='left', vertical='center', wrap_text=True)

    # Forzamos la escritura directa usando ws.cell para que openpyxl acepte el RichText puro
    def _set_rich_cell(row, column, text_val, alignment):
        if not row:
            return
        cell = ws.cell(row=row, column=column)
        # ---> AQUÍ OCURRE LA MAGIA DEL BOLD/ITALIC <---
        cell.value = markdown_to_richtext(text_val)
        cell.alignment = alignment

    # 1. WATERMARK: 4 filas independientes
    watermark_lines = process_multiline_branding(f.get(AT_FIELDS["branding_watermark"]), max_lines=4)
    for idx, row in enumerate(branding_layout["watermark"]):
        _set_rich_cell(row, col, watermark_lines[idx], align_center)

    # 2. BUG: 3 filas independientes
    bug_lines = process_multiline_branding(f.get(AT_FIELDS["branding_bug"]), max_lines=3)
    for idx, row in enumerate(branding_layout["bug"]):
        _set_rich_cell(row, col, bug_lines[idx], align_center)

    # 3. COPYRIGHT: 1 sola fila con todo el bloque de texto (\n interno), alineado a la IZQUIERDA
    copyright_text = f.get(AT_FIELDS["branding_copyright"], "")
    _set_rich_cell(branding_layout["copyright"], col, copyright_text, align_left)

    # 4. ENDCARD: 6 filas independientes
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
    safe_set_value(ws, layout["tx_row"], col, single["tx"], alignment=center_wrap())
    safe_set_value(ws, layout["tc_row"], col, single["tc"], alignment=center_wrap())

    write_djr_branding(ws, layout, col, record)

    role = hub_role(record.get(AT_FIELDS["owner_hub"]))
    comments_row = layout["comments_rows"].get(role)
    if comments_row:
        # APLICAMOS EL PARSEADOR AL COMENTARIO AQUÍ:
        safe_set_value(
            ws, 
            comments_row, 
            col, 
            markdown_to_richtext(single["comments"]), 
            alignment=center_wrap()
        )


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

    # fallback: if no region match, use existing generic logic
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
    included, _ = collect_included_countries(
        youtube_channel=record.get(AT_FIELDS["youtube_channel"]),
        region=record.get(AT_FIELDS["region"]),
        notes=record.get(AT_FIELDS["notes"]),
        excluded_markets=record.get(AT_FIELDS["excluded_markets"]),
        channel_id_to_name_map=channel_id_to_name_map,
    )
    included = collect_included_countries_djr(record, channel_id_to_name_map)

    country_start, country_end = layout["country_rows"]
    status_start, status_end = layout["status_rows"]

    # Definición de Tiers (Ajusta los nombres exactos si difieren en tu DJR_SHEET_COUNTRIES)
    TIER_B_COUNTRIES = {"Poland", "Romania", "Netherlands (Dutch)", "Portugal", "Sweden (Swedish)", "Denmark (Danish)", "Norway", "Bulgaria"}
    # Tier A y C usarán el mismo color (EDEDED), por lo que se agrupan en el bloque 'else' del YES.

    # Rellenos (Fills)
    FILL_TIER_A_C_YES = PatternFill(fill_type="solid", fgColor="EDEDED")
    FILL_TIER_B_YES = PatternFill(fill_type="solid", fgColor="FFFFFF")
    FILL_GREEN_YES = PatternFill(fill_type="solid", fgColor="00B050")

    for idx, country in enumerate(DJR_SHEET_COUNTRIES):
        rollout_row = country_start + idx
        status_row = status_start + idx

        value = "YES" if country in included else "NO"

        # 1. Lógica para Mediaplanning (Mantiene el original Verde / Letra Roja)
        if value == "YES":
            font_rollout = Font(color="FFFFFFFF", bold=True)
            fill_rollout = FILL_GREEN_YES
        else:
            font_rollout = Font(color="FFFF0000", bold=True)
            fill_rollout = None 

        safe_set_value(ws, rollout_row, col, value, alignment=center_wrap(), font=font_rollout, fill=fill_rollout)

        # 2. Lógica para Status Conforming (Gris/Blanco por Tiers)
        if value == "YES":
            font_status = Font(color="00000000", bold=True) # Letra oscura para contraste
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
    
    # 1. Validación estricta para tranquilizar al Type Checker
    if not first_monday:
        logging.warning("No se pudo determinar el primer lunes. Abortando generación.")
        return

    # 2. Como ya pasó el 'if', Python sabe que first_monday es 'date' 100% seguro
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
    wb.save(output_path)

    logging.info(f"Saved DJR calendar to: {output_path}")