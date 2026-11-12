import os
import glob
import zipfile
import re
from copy import copy
from typing import Dict, List, Optional, Set, Tuple
from datetime import datetime, date, timedelta

from openpyxl import load_workbook
from openpyxl.cell.cell import Cell
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from config.layouts import TARGET_SHEET as dc_sheet_name, DAY_ORDER
from config.countries import SHEET_COUNTRIES as dc_master_countries
from config.fields import AT_FIELDS
from exporters.full_calendar_djr import (
    DJR_TARGET_SHEET as djr_sheet_name,
    collect_included_countries_djr,
)
from config.countries_djr import DJR_SHEET_COUNTRIES as djr_master_countries
from config.layouts_djr import DJR_WEEK_LAYOUT as djr_layout, DJR_DAY_LAYOUT as djr_day
from core.excel_utils import safe_set_value, center_wrap
from core.dates import (
    first_visible_monday_of_target_month,
    month_name_from_records,
    resolve_target_month_year,
)
from core.text_formatting import normalize_str
from core.countries import collect_included_countries
from services.planner_service import (
    aggregate_by_week_day_and_role,
    assemble_text_for_single_asset,
    earliest_deadline_for_day,
)
from exporters.shared.records import aggregate_by_visual_week_day_role
from exporters.full_calendar import infer_week_number_for_block

LONG_DATE_FMT = "dddd, dd mmmm yyyy"

COUNTRY_SHEET_CODES = {
    "middle east": "MENA",
    "middle east (arabic - msa)": "MENA",
    "middle east (arabic)": "MENA",
    "mena": "MENA",
    "netherlands": "NLD",
    "netherlands (dutch)": "NLD",
    "belgium": "BEL",
    "belgium (french)": "BEL",
    "portugal": "PT",
    "spain": "ES",
    "españa": "ES",
    "italy": "IT",
    "italy (italian)": "IT",
    "france": "FR",
    "germany": "GSA",
    "gsa (german)": "GSA",
    "poland": "PL",
    "poland (polish)": "PL",
    "romania": "RO",
    "romania (romanian)": "RO",
    "greece": "GR",
    "greece (greek)": "GR",
    "israel": "IL",
    "israel (hebrew)": "IL",
    "bulgaria": "BG",
    "bulgaria (bulgarian)": "BG",
    "turkey": "TR",
    "türkiye": "TR",
    "hungary": "HU",
    "hungary (hungarian)": "HU",
    "czech r.": "CZ",
    "czech r. (czech)": "CZ",
    "sweden": "SE",
    "sweden (swedish)": "SE",
    "denmark": "DK",
    "denmark (danish)": "DK",
    "norway": "NO",
    "norway (norwegian)": "NO",
}

DC_TRANS_CONTENT_COLS = {
    "Monday":    [4, 6, 8, 10],    # D, F, H, J
    "Tuesday":   [12, 14, 16, 18], # L, N, P, R
    "Wednesday": [20, 22, 24, 26], # T, V, X, Z
    "Thursday":  [28, 30, 32, 34], # AB, AD, AF, AH
    "Friday":    [36, 38, 40, 42], # AJ, AL, AN, AP
    "Saturday":  [44, 46, 48, 50], # AR, AT, AV, AX
    "Sunday":    [52, 54, 56, 58], # AZ, BB, BD, BF
}
DC_TRANS_DAY_RANGES = {
    "Monday": (4, 11), "Tuesday": (12, 19), "Wednesday": (20, 27),
    "Thursday": (28, 35), "Friday": (36, 43), "Saturday": (44, 51), "Sunday": (52, 59),
}

DJR_TRANS_CONTENT_COLS = {
    "Monday":    [4, 6, 8, 10],    # D, F, H, J
    "Tuesday":   [11, 13, 15, 17], # K, M, O, Q
    "Wednesday": [18, 20, 21, 23], # R, T, U, W
    "Thursday":  [24, 26, 27, 29], # X, Z, AA, AC
    "Friday":    [30, 32, 33, 35], # AD, AF, AG, AI
    "Saturday":  [36, 38, 39, 41], # AJ, AL, AM, AO
    "Sunday":    [42, 44, 45, 47], # AP, AR, AS, AU
}
DJR_TRANS_DAY_RANGES = {
    "Monday": (4, 10), "Tuesday": (11, 17), "Wednesday": (18, 23),
    "Thursday": (24, 29), "Friday": (30, 35), "Saturday": (36, 41), "Sunday": (42, 47),
}


def get_clean_country_name(name: str) -> str:
    return re.sub(r'\s*\(.*?\)', '', name).strip()


def get_sheet_country_code(country_name: str) -> str:
    clean = country_name.lower().strip()
    if clean in COUNTRY_SHEET_CODES:
        return COUNTRY_SHEET_CODES[clean]
    base = get_clean_country_name(country_name).lower()
    if base in COUNTRY_SHEET_CODES:
        return COUNTRY_SHEET_CODES[base]
    for key, code in COUNTRY_SHEET_CODES.items():
        if key in clean or clean in key:
            return code
    return base.upper()[:4]


def find_translation_template(is_djr: bool) -> Optional[str]:
    brand_tag = "djr" if is_djr else "dc"
    candidates = [
        f"templates/template_translation_{brand_tag}.xltx",
        f"templates/template_translation_{brand_tag}.xlsx",
    ]
    for path in candidates:
        if os.path.exists(path):
            return path
    for ext in ("*.xltx", "*.xlsx"):
        for path in glob.glob(os.path.join("templates", ext)):
            fname = os.path.basename(path).lower()
            if ("translation" in fname or "mena" in fname) and brand_tag in fname:
                if brand_tag == "dc" and "djr" in fname:
                    continue
                return path
    return None


def _set_cell_keep_merge(
    ws,
    row: int,
    col: int,
    value,
    alignment=None,
    number_format: Optional[str] = None,
    font=None,
    fill=None,
):
    """Escribe en la celda superior izquierda sin destruir los merged_cells existentes."""
    cell = ws.cell(row=row, column=col)
    if isinstance(cell, Cell):
        cell.value = value
        if alignment is not None:
            cell.alignment = alignment
        if number_format is not None:
            cell.number_format = number_format
        if font is not None:
            cell.font = font
        if fill is not None:
            cell.fill = fill


def _ensure_merged_range(ws, min_row: int, min_col: int, max_row: int, max_col: int):
    """Descombina cualquier rango solapado antes de combinar, evitando XML corrupto en Excel."""
    target_str = f"{get_column_letter(min_col)}{min_row}:{get_column_letter(max_col)}{max_row}"
    for r in list(ws.merged_cells.ranges):
        if str(r) == target_str:
            return
        if not (r.max_col < min_col or r.min_col > max_col or r.max_row < min_row or r.min_row > max_row):
            try:
                ws.unmerge_cells(str(r))
            except Exception:
                pass
    try:
        ws.merge_cells(start_row=min_row, start_column=min_col, end_row=max_row, end_column=max_col)
    except Exception:
        pass


def _get_smart_slots(day_name: str, programming_count: int, mediaplanning_count: int, day_cols_map: Dict[str, List[int]]):
    all_cols = day_cols_map[day_name]
    p_count = min(programming_count, 4)
    m_count = min(mediaplanning_count, 4 - p_count)

    prog_slots = [0, 1, 2, 3][:p_count]
    prog_cols = [all_cols[i] for i in prog_slots]

    remaining_slots = [i for i in [2, 1, 0, 3] if i not in prog_slots]
    media_slots = remaining_slots[:m_count]
    media_cols = [all_cols[i] for i in media_slots]

    used_slots = set(prog_slots + media_slots)
    uk_cols = [all_cols[3]] if 3 not in used_slots else []
    if 3 not in used_slots:
        used_slots.add(3)
    extra_cols = [all_cols[i] for i in [0, 1, 2, 3] if i not in used_slots]
    return {"Programming": prog_cols, "Mediaplanning": media_cols, "UK": uk_cols, "Extra": extra_cols}


def _populate_translation_calendar(
    wb,
    is_djr: bool,
    country: str,
    clean_country: str,
    country_code: str,
    records: list,
    channel_id_to_name_map: Optional[dict],
    month_str: str,
    target_year: Optional[int],
    target_month: Optional[int],
    first_monday: Optional[date],
):
    is_italy = "ital" in country.lower()
    display_country_title = "ITALY" if is_italy else "MENA"
    lang_header = "Italian | Italiano" if is_italy else "Arabic | العربية"
    country_row_label = (
        "Italy (Italian)" if is_italy
        else ("Middle East (Arabic)" if is_djr else "Middle East (Arabic - MSA)")
    )

    if target_year and target_month:
        d_ref = date(target_year, target_month, 1)
        month_name_upper = d_ref.strftime("%B").upper()
        month_abbr = d_ref.strftime("%b").capitalize()
    else:
        month_name_upper = month_name_from_records(records or []).upper() if records else "MONTH"
        month_abbr = month_name_upper.capitalize()[:3]

    # 1. Actualizar Highlights sin romper merges
    for sn in wb.sheetnames:
        if "highlights" in sn.lower():
            ws_hi = wb[sn]
            if target_year and target_month:
                _set_cell_keep_merge(ws_hi, 3, 5, date(target_year, target_month, 1), number_format="[$-409]mmm-yy")
            for r in range(1, 15):
                for c in range(1, 10):
                    cell = ws_hi.cell(row=r, column=c)
                    if isinstance(cell, Cell) and isinstance(cell.value, str) and "EMEA" in cell.value:
                        cell.value = cell.value.replace("EMEA", country_code)

    # 2. Actualizar FULL CALENDAR
    fc_sheet = djr_sheet_name if is_djr else dc_sheet_name
    if fc_sheet not in wb.sheetnames:
        return
    ws = wb[fc_sheet]

    # Capturar los rellenos nativos de fecha y deadline antes de tocar celdas
    ref_date_row = 7 if is_djr else 222
    ref_dl_row = 8 if is_djr else 223
    in_month_date_fill = copy(ws.cell(row=ref_date_row, column=4).fill)
    in_month_dl_fill = copy(ws.cell(row=ref_dl_row, column=4).fill)

    brand_full = "DISNEY JUNIOR" if is_djr else "DISNEY CHANNEL"
    end_header_col = 47 if is_djr else 54  # AU=47 en DJR, BB=54 en DC
    _ensure_merged_range(ws, 1, 2, 2, end_header_col)
    _ensure_merged_range(ws, 3, 2, 4, end_header_col)

    _set_cell_keep_merge(ws, 1, 2, f"{brand_full} YOUTUBE {display_country_title} CALENDAR", alignment=center_wrap())
    if is_djr:
        if target_year and target_month:
            _set_cell_keep_merge(ws, 3, 2, date(target_year, target_month, 1), alignment=center_wrap(), number_format="mmm-yy")
    else:
        _set_cell_keep_merge(ws, 3, 2, month_name_upper, alignment=center_wrap())

    start_date_row = 7 if is_djr else 6
    if first_monday:
        _set_cell_keep_merge(ws, start_date_row, 4, first_monday, alignment=center_wrap(), number_format=LONG_DATE_FMT)

    step = 62 if is_djr else 54
    day_cols_map = DJR_TRANS_CONTENT_COLS if is_djr else DC_TRANS_CONTENT_COLS
    day_ranges_map = DJR_TRANS_DAY_RANGES if is_djr else DC_TRANS_DAY_RANGES

    translation_cols: Set[int] = set()
    for d_name in DAY_ORDER:
        c_start, c_end = day_ranges_map[d_name]
        c_list = set(day_cols_map[d_name])
        for c in range(c_start, c_end + 1):
            if c not in c_list:
                translation_cols.add(c)

    fill_teal_owner = PatternFill(fill_type="solid", fgColor="147D83")
    fill_mint_trans = PatternFill(fill_type="solid", fgColor="EAF6F5")
    fill_djr_owner = PatternFill(fill_type="solid", fgColor="D6DCE4")
    fill_out_month = PatternFill(fill_type="solid", fgColor="D6DCE4")
    fill_green_yes = PatternFill(fill_type="solid", fgColor="00B050")
    fill_grey_used = PatternFill(fill_type="solid", fgColor="BFBFBF")
    fill_white_link = PatternFill(fill_type="solid", fgColor="FFFFFF")
    fill_none = PatternFill(fill_type=None)

    font_white_bold = Font(bold=True, color="FFFFFF")
    font_black_bold = Font(bold=True, color="000000")

    djr_buckets: Dict[Tuple[int, str, str], List[dict]] = {}
    dc_buckets: Dict[Tuple[date, str, str], List[dict]] = {}
    if records and first_monday:
        if is_djr:
            djr_buckets = aggregate_by_visual_week_day_role(records, first_monday)
        else:
            dc_buckets = aggregate_by_week_day_and_role(records)

    rows_to_hide_and_clear: Set[int] = set()

    for w_idx in range(6):
        offset = w_idx * step
        if is_djr:
            r_date = 7 + offset
            r_dl = 8 + offset
            r_owner = 9 + offset
            r_prop = 10 + offset
            r_ctype = 11 + offset
            r_title = 12 + offset
            r_syn = 13 + offset
            r_strat = 14 + offset
            r_country = 18 + offset
            r_used = 19 + offset
            r_link = 20 + offset
            block_end = r_link + 1
        else:
            r_date = 6 + offset
            r_dl = 7 + offset
            r_owner = 8 + offset
            r_prop = 9 + offset
            r_ctype = 10 + offset
            r_title = 11 + offset
            r_syn = 12 + offset
            r_strat = 13 + offset
            r_country = 21 + offset
            r_used = 22 + offset
            r_link = 23 + offset
            block_end = r_link + 1

        week_monday = (first_monday + timedelta(days=w_idx * 7)) if first_monday else None
        week_sunday = (week_monday + timedelta(days=6)) if week_monday else None

        if (
            week_monday is not None
            and week_sunday is not None
            and target_month is not None
            and week_monday.month != target_month
            and week_sunday.month != target_month
        ):
            for r_h in range(r_date, block_end + 1):
                rows_to_hide_and_clear.add(r_h)
            continue

        if not is_djr:
            _ensure_merged_range(ws, r_date, 2, r_date, 3)
        _ensure_merged_range(ws, r_country, 2, r_country, 3)
        for d_name in DAY_ORDER:
            c_start, c_end = day_ranges_map[d_name]
            _ensure_merged_range(ws, r_date, c_start, r_date, c_end)
            _ensure_merged_range(ws, r_dl, c_start, r_dl, c_end)

        _set_cell_keep_merge(ws, r_country, 2, country_row_label, alignment=center_wrap())

        per_day_prog: Dict[str, List[dict]] = {}
        per_day_media: Dict[str, List[dict]] = {}
        per_day_all: Dict[str, List[dict]] = {}
        for d_name in DAY_ORDER:
            if is_djr:
                p_recs = djr_buckets.get((w_idx + 1, d_name, "Programming"), [])
                m_recs = djr_buckets.get((w_idx + 1, d_name, "Mediaplanning"), [])
            elif week_monday is not None:
                p_recs = dc_buckets.get((week_monday, d_name, "Programming"), [])
                m_recs = dc_buckets.get((week_monday, d_name, "Mediaplanning"), [])
            else:
                p_recs = []
                m_recs = []
            per_day_prog[d_name] = p_recs
            per_day_media[d_name] = m_recs
            per_day_all[d_name] = p_recs + m_recs

        if week_monday is not None:
            if is_djr:
                _set_cell_keep_merge(ws, r_date, 2, "WEEK", alignment=center_wrap())
                _set_cell_keep_merge(ws, r_date, 3, week_monday.isocalendar()[1], alignment=center_wrap())
            else:
                wk_label = infer_week_number_for_block(per_day_all, week_monday)
                if not wk_label.upper().startswith("WEEK"):
                    wk_label = f"WEEK {wk_label}"
                _set_cell_keep_merge(ws, r_date, 2, wk_label, alignment=center_wrap())

        for d_idx, d_name in enumerate(DAY_ORDER):
            c_start, c_end = day_ranges_map[d_name]
            current_date = (week_monday + timedelta(days=d_idx)) if week_monday else None
            in_month = False
            if current_date is not None and target_year is not None and target_month is not None:
                in_month = (current_date.year == target_year and current_date.month == target_month)

            date_cell = ws.cell(row=r_date, column=c_start)
            if isinstance(date_cell, Cell):
                date_cell.number_format = LONG_DATE_FMT
                date_cell.alignment = center_wrap()

            if current_date is not None:
                dl_val = earliest_deadline_for_day(per_day_all[d_name]) or (current_date - timedelta(days=7))
                _set_cell_keep_merge(ws, r_dl, c_start, dl_val, alignment=center_wrap(), number_format=LONG_DATE_FMT)

            for c in range(c_start, c_end + 1):
                is_trans = c in translation_cols
                for r_clear in (r_prop, r_ctype, r_title, r_syn, r_strat, r_country, r_used, r_link):
                    cell = ws.cell(row=r_clear, column=c)
                    if isinstance(cell, Cell):
                        cell.value = None

                if not in_month:
                    # Días fuera de mes: gris en TODAS las filas incluyendo fecha (r_date) y deadline (r_dl)
                    for r_shade in (r_date, r_dl, r_owner, r_prop, r_ctype, r_title, r_syn, r_strat, r_country):
                        ws.cell(row=r_shade, column=c).fill = copy(fill_out_month)
                    if is_trans:
                        c_own_out = ws.cell(row=r_owner, column=c)
                        if isinstance(c_own_out, Cell):
                            c_own_out.value = lang_header
                        c_own_out.alignment = center_wrap()
                else:
                    # Días dentro del mes: restaurar color nativo en fecha (r_date) y deadline (r_dl)
                    ws.cell(row=r_date, column=c).fill = copy(in_month_date_fill)
                    ws.cell(row=r_dl, column=c).fill = copy(in_month_dl_fill)
                    if is_trans:
                        c_own = ws.cell(row=r_owner, column=c)
                        if isinstance(c_own, Cell):
                            c_own.value = lang_header
                        c_own.fill = fill_teal_owner
                        c_own.font = font_white_bold
                        c_own.alignment = center_wrap()
                        for r_m in (r_prop, r_ctype, r_title, r_syn, r_strat, r_country, r_used, r_link):
                            ws.cell(row=r_m, column=c).fill = fill_mint_trans
                    else:
                        ws.cell(row=r_owner, column=c).fill = fill_djr_owner if is_djr else fill_none
                        for r_w in (r_prop, r_ctype, r_title, r_syn, r_strat, r_country):
                            ws.cell(row=r_w, column=c).fill = fill_none
                        ws.cell(row=r_used, column=c).fill = fill_grey_used
                        ws.cell(row=r_link, column=c).fill = fill_white_link

            col_map = _get_smart_slots(d_name, len(per_day_prog[d_name]), len(per_day_media[d_name]), day_cols_map)
            for col in col_map["Programming"]:
                _set_cell_keep_merge(ws, r_owner, col, "PROGRAMMING HUB", alignment=center_wrap())
            for col in col_map["Mediaplanning"]:
                _set_cell_keep_merge(ws, r_owner, col, "MEDIAPLANNING HUB", alignment=center_wrap())
            for col in col_map["UK"]:
                _set_cell_keep_merge(ws, r_owner, col, "YT TEAM UK", alignment=center_wrap())
            for col in col_map["Extra"]:
                _set_cell_keep_merge(ws, r_owner, col, "", alignment=center_wrap())

            def _write_slot(col: int, rec: dict):
                f = rec.get("fields", rec)
                if is_djr:
                    inc = collect_included_countries_djr(rec, channel_id_to_name_map)
                else:
                    inc, _ = collect_included_countries(
                        youtube_channel=f.get(AT_FIELDS["youtube_channel"]),
                        region=f.get(AT_FIELDS["region"]),
                        notes=f.get(AT_FIELDS["notes"]),
                        excluded_markets=f.get(AT_FIELDS["excluded_markets"]),
                        channel_id_to_name_map=channel_id_to_name_map,
                    )
                target_norm = clean_country.lower()
                is_included = any(
                    target_norm in str(x).lower() or str(x).lower().split("(")[0].strip() in target_norm
                    for x in inc
                )
                if not is_included:
                    return

                single = assemble_text_for_single_asset(rec)
                raw_ct = normalize_str(f.get(AT_FIELDS["content_type"]))
                ctype = single.get("content_type", "")
                if raw_ct in {"Editorial Short", "Exclusive"}:
                    ctype = "Exclusive Short"
                elif raw_ct == "Vertical Short":
                    ctype = "Vertical Short"

                strat_val = f.get(AT_FIELDS["strategy"])
                strat_str = (
                    " / ".join(normalize_str(s) for s in strat_val if normalize_str(s))
                    if isinstance(strat_val, list)
                    else normalize_str(strat_val)
                )

                _set_cell_keep_merge(ws, r_prop, col, single.get("property", ""), alignment=center_wrap())
                _set_cell_keep_merge(ws, r_ctype, col, ctype, alignment=center_wrap())
                _set_cell_keep_merge(ws, r_title, col, single.get("title", ""), alignment=center_wrap())
                _set_cell_keep_merge(ws, r_syn, col, single.get("synopsis", ""), alignment=center_wrap())
                _set_cell_keep_merge(ws, r_strat, col, strat_str, alignment=center_wrap())
                _set_cell_keep_merge(
                    ws, r_country, col, "YES",
                    alignment=center_wrap(),
                    font=font_white_bold if is_djr else font_black_bold,
                    fill=fill_green_yes,
                )

            for idx, rec in enumerate(per_day_prog[d_name][:len(col_map["Programming"])]):
                _write_slot(col_map["Programming"][idx], rec)
            for idx, rec in enumerate(per_day_media[d_name][:len(col_map["Mediaplanning"])]):
                _write_slot(col_map["Mediaplanning"][idx], rec)

    for m in list(ws.merged_cells.ranges):
        if m.min_row in rows_to_hide_and_clear or m.max_row in rows_to_hide_and_clear:
            try:
                ws.unmerge_cells(str(m))
            except Exception:
                pass
    for r_hide in sorted(rows_to_hide_and_clear):
        ws.row_dimensions[r_hide].hidden = True
        ws.row_dimensions[r_hide].height = 0
        for c in range(1, end_header_col + 15):
            cell = ws.cell(row=r_hide, column=c)
            if isinstance(cell, Cell):
                cell.value = None
            cell.fill = PatternFill(fill_type=None)
            cell.border = Border()

    # 3. Actualizar MONTHLY OVERVIEW local y su cabecera
    if records:
        if is_djr:
            from exporters.monthly_overview_djr import populate_monthly_overview_djr
            populate_monthly_overview_djr(
                wb, records=records, enable_mediaplanning=True,
                channel_id_to_name_map=channel_id_to_name_map, target_country=clean_country,
            )
        else:
            from exporters.monthly_overview import populate_monthly_overview
            populate_monthly_overview(
                wb, records=records, enable_mediaplanning=True,
                channel_id_to_name_map=channel_id_to_name_map, target_country=clean_country,
            )

    for sn in wb.sheetnames:
        if "overview" in sn.lower() or "monthly" in sn.lower():
            ws_ov = wb[sn]
            title_row = 5 if is_djr else 2
            brand_short = "DJR" if is_djr else "DC"
            yr_str = str(target_year) if target_year else "2026"
            _set_cell_keep_merge(
                ws_ov, title_row, 2,
                f"{brand_short} {display_country_title} YT {month_name_upper} {yr_str} CALENDAR",
                alignment=center_wrap(),
            )

    # 4. Renombrar todas las pestañas usando códigos cortos (<= 31 chars)
    for sn in list(wb.sheetnames):
        if "highlights" in sn.lower():
            wb[sn].title = f"YT {'DJR' if is_djr else 'DC'} {country_code} {month_abbr} Highlights"[:31]
        elif "EMEA" in sn:
            wb[sn].title = sn.replace("EMEA", country_code)[:31]


def generate_local_calendars(
    master_path: str,
    output_dir: str,
    channel_type: str,
    selected_countries: Optional[List[str]] = None,
    month_str: str = "2026-05",
    dc_blocks: Optional[dict] = None,
    records: Optional[list] = None,
    channel_id_to_name_map: Optional[dict] = None,
):
    if not os.path.exists(master_path) or not selected_countries:
        return None

    os.makedirs(output_dir, exist_ok=True)
    for f in os.listdir(output_dir):
        if f.endswith(".xlsx") or f.endswith(".zip"):
            os.remove(os.path.join(output_dir, f))

    is_djr = channel_type.upper() == "DJR"
    sheet_name = djr_sheet_name if is_djr else dc_sheet_name
    master_countries = djr_master_countries if is_djr else dc_master_countries

    target_year: Optional[int] = None
    target_month: Optional[int] = None
    try:
        dt = datetime.strptime(month_str, "%Y-%m")
        target_year = dt.year
        target_month = dt.month
        month_num = dt.strftime("%m")
        month_name_year = dt.strftime("%B %Y")
    except Exception:
        target_year, target_month = resolve_target_month_year(records or [])
        month_num = f"{target_month:02d}" if target_month else "00"
        month_name_year = month_str

    first_monday: Optional[date] = first_visible_monday_of_target_month(records) if records else None
    if first_monday is None and target_year is not None and target_month is not None:
        d1 = date(target_year, target_month, 1)
        first_monday = d1 - timedelta(days=d1.weekday())

    brand_label = "DJR" if is_djr else "DC"
    generated_files = []

    fill_grey = PatternFill(fill_type="solid", fgColor="BFBFBF")
    fill_white = PatternFill(fill_type="solid", fgColor="FFFFFF")
    font_bold = Font(bold=True, color="000000")

    thin_side = Side(border_style="thin", color="000000")
    thick_side = Side(border_style="medium", color="000000")
    border_normal = Border(top=thin_side, bottom=thin_side, left=thin_side, right=thin_side)
    border_thick_bottom = Border(top=thin_side, bottom=thick_side, left=thin_side, right=thin_side)

    def style_recycled_row(sheet, row_idx, label, fill_color, is_last=False):
        if not isinstance(row_idx, int):
            return
        applied_border = border_thick_bottom if is_last else border_normal
        cell_a = sheet.cell(row=row_idx, column=1)
        if isinstance(cell_a, Cell):
            cell_a.value = None
        cell_c = sheet.cell(row=row_idx, column=3)
        if isinstance(cell_c, Cell):
            cell_c.value = None
        safe_set_value(sheet, row_idx, 2, label)
        sheet.row_dimensions[row_idx].hidden = False
        sheet.row_dimensions[row_idx].height = 30
        for c in range(2, 32):
            cell = sheet.cell(row=row_idx, column=c)
            if c >= 3 and isinstance(cell, Cell):
                cell.value = None
            cell.fill = fill_color
            cell.border = applied_border
        cell2 = sheet.cell(row=row_idx, column=2)
        cell2.font = font_bold
        cell2.alignment = Alignment(horizontal="right", vertical="center")

    def clear_cell(sheet, row_idx, col_idx):
        if isinstance(row_idx, int) and isinstance(col_idx, int):
            cell_obj = sheet.cell(row=row_idx, column=col_idx)
            if isinstance(cell_obj, Cell):
                cell_obj.value = None

    def get_val(sheet, r, c):
        if not isinstance(r, int) or not isinstance(c, int):
            return ""
        v = sheet.cell(row=r, column=c).value
        return str(v).strip().lower() if v is not None else ""

    def format_israel_tx_tc(sheet, row_idx, is_tx=False):
        if not isinstance(row_idx, int):
            return
        sheet.row_dimensions[row_idx].hidden = False
        sheet.row_dimensions[row_idx].height = 60 if is_tx else 35
        for c in range(4, 32):
            cell = sheet.cell(row=row_idx, column=c)
            if isinstance(cell, Cell):
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    translation_tpl_path = find_translation_template(is_djr)

    for country in selected_countries:
        clean_country = get_clean_country_name(country)
        country_code = get_sheet_country_code(country)
        c_lower = country.lower()

        is_translation_market = any(k in c_lower for k in ["middle east", "mena", "italy", "italia"])

        # === RAMA A: MERCADOS CON COLUMNAS DE TRADUCCIÓN (MENA E ITALIA) ===
        if is_translation_market and translation_tpl_path:
            country_wb = load_workbook(translation_tpl_path, data_only=False)
            _populate_translation_calendar(
                wb=country_wb,
                is_djr=is_djr,
                country=country,
                clean_country=clean_country,
                country_code=country_code,
                records=records or [],
                channel_id_to_name_map=channel_id_to_name_map,
                month_str=month_str,
                target_year=target_year,
                target_month=target_month,
                first_monday=first_monday,
            )
            filename = f"{month_num}_Social Media Calendar {brand_label} YouTube {clean_country} - {month_name_year}.xlsx"
            output_path = os.path.join(output_dir, filename)
            country_wb.template = False
            country_wb.save(output_path)
            country_wb.close()
            generated_files.append(output_path)
            continue

        # === RAMA B: RESTO DE PAÍSES (INCLUYENDO ESPAÑA / TIER S & TIER A) ===
        country_wb = load_workbook(master_path, data_only=False)

        for sn in list(country_wb.sheetnames):
            sn_lower = sn.lower()
            if any(x in sn_lower for x in ["compilation", "sheet 1", "sheet1", "id's", "metadata"]):
                del country_wb[sn]

        if sheet_name not in country_wb.sheetnames:
            country_wb.close()
            continue

        ws = country_wb[sheet_name]
        is_israel = clean_country.lower() == "israel"
        max_content_row: int = 0
        safe_rows: Set[int] = set()

        ws.column_dimensions["AG"].hidden = True
        ws.column_dimensions["AF"].hidden = True

        if is_djr:
            for day in DAY_ORDER:
                day_cols = sorted([col for _, col in djr_day[day]])
                if len(day_cols) >= 4:
                    uk_col_letter = get_column_letter(day_cols[3])
                    ws.column_dimensions[uk_col_letter].width = 22
            for layout in djr_layout.values():
                for r_fmt in (layout.get("date_row"), layout.get("deadline_row")):
                    if isinstance(r_fmt, int):
                        for day in DAY_ORDER:
                            first_c = min(col for _, col in djr_day[day])
                            c_obj = ws.cell(row=r_fmt, column=first_c)
                            if isinstance(c_obj, Cell):
                                c_obj.number_format = LONG_DATE_FMT

        if records:
            if is_djr:
                from exporters.monthly_overview_djr import populate_monthly_overview_djr
                populate_monthly_overview_djr(
                    country_wb, records=records, enable_mediaplanning=True,
                    channel_id_to_name_map=channel_id_to_name_map, target_country=clean_country,
                )
            else:
                from exporters.monthly_overview import populate_monthly_overview
                populate_monthly_overview(
                    country_wb, records=records, enable_mediaplanning=True,
                    channel_id_to_name_map=channel_id_to_name_map, target_country=clean_country,
                )

        # ---------------- LÓGICA DISNEY JUNIOR ----------------
        if is_djr:
            country_idx = next((i for i, mc in enumerate(master_countries) if mc.lower() == country.lower()), -1)
            if country_idx == -1:
                country_wb.close()
                continue

            for week_num, layout in djr_layout.items():
                rollout_start = layout["country_rows"][0]
                my_rollout_row = rollout_start + country_idx

                if my_rollout_row > 0:
                    safe_rows.add(my_rollout_row)
                    cell_b = ws.cell(row=my_rollout_row, column=2)
                    cell_c = ws.cell(row=my_rollout_row, column=3)
                    country_name_val = cell_c.value if isinstance(cell_c, Cell) else None
                    if isinstance(cell_b, Cell):
                        cell_b.value = country_name_val
                        cell_b.alignment = Alignment(horizontal="center", vertical="center")
                    if isinstance(cell_c, Cell):
                        cell_c.value = None
                    try:
                        ws.merge_cells(start_row=my_rollout_row, start_column=2, end_row=my_rollout_row, end_column=3)
                    except Exception:
                        pass

                djr_tc = layout.get("tc_row")
                if isinstance(djr_tc, int) and djr_tc > max_content_row:
                    max_content_row = djr_tc

                candidate_rows = [r for r in range(rollout_start, rollout_start + len(master_countries)) if r > my_rollout_row]
                if len(candidate_rows) < 3:
                    status_start = layout.get("status_rows", [0])[0]
                    candidate_rows.extend([r for r in range(status_start, status_start + 30) if r > my_rollout_row])

                candidate_rows = sorted(list(set(candidate_rows)))
                if len(candidate_rows) >= 3:
                    used_row = candidate_rows[0]
                    link_row = candidate_rows[1]
                    blank_row = candidate_rows[2]

                    style_recycled_row(ws, used_row, "USED OR NOT", fill_grey, is_last=False)
                    style_recycled_row(ws, link_row, "LINK YOUTUBE", fill_white, is_last=True)
                    ws.row_dimensions[blank_row].hidden = False
                    ws.row_dimensions[blank_row].height = 15
                    for c in range(1, 32):
                        cell = ws.cell(row=blank_row, column=c)
                        if isinstance(cell, Cell):
                            cell.value = None
                        cell.fill = PatternFill(fill_type=None)
                        cell.border = Border()
                    safe_rows.update([used_row, link_row, blank_row])

                for day in DAY_ORDER:
                    for role, col in djr_day[day]:
                        cell = ws.cell(row=my_rollout_row, column=col)
                        cell_val = str(cell.value).strip().upper() if hasattr(cell, "value") else ""
                        if "YES" not in cell_val:
                            if isinstance(cell, Cell):
                                cell.value = None
                            rows_to_clear = [
                                layout["property_row"], layout["content_type_row"],
                                layout["title_row"], layout["synopsis_row"], layout["strategy_row"],
                            ]
                            if is_israel:
                                rows_to_clear.extend([layout.get("tx_row"), layout.get("tc_row")])
                            for r in rows_to_clear:
                                clear_cell(ws, r, col)

                if is_israel:
                    israel_tx = layout.get("tx_row")
                    if isinstance(israel_tx, int):
                        format_israel_tx_tc(ws, israel_tx, is_tx=True)
                        safe_rows.add(israel_tx)
                    israel_tc = layout.get("tc_row")
                    if isinstance(israel_tc, int):
                        format_israel_tx_tc(ws, israel_tc, is_tx=False)
                        safe_rows.add(israel_tc)

        # ---------------- LÓGICA DISNEY CHANNEL ----------------
        else:
            if not dc_blocks:
                country_wb.close()
                continue
            for block in dc_blocks.values():
                target_key = next((k for k in block.rollout_country_rows.keys() if k.lower() == country.lower()), None)
                country_row_rollout = block.rollout_country_rows.get(target_key) if target_key else None

                if country_row_rollout:
                    safe_rows.add(country_row_rollout)
                    cell_b = ws.cell(row=country_row_rollout, column=2)
                    cell_c = ws.cell(row=country_row_rollout, column=3)
                    country_name_val = cell_c.value if isinstance(cell_c, Cell) else None
                    if isinstance(cell_b, Cell):
                        cell_b.value = country_name_val
                        cell_b.alignment = Alignment(horizontal="center", vertical="center")
                    if isinstance(cell_c, Cell):
                        cell_c.value = None
                    try:
                        ws.merge_cells(start_row=country_row_rollout, start_column=2, end_row=country_row_rollout, end_column=3)
                    except Exception:
                        pass

                    candidate_rows = [r for r in block.rollout_country_rows.values() if r > country_row_rollout]
                    if len(candidate_rows) < 3:
                        candidate_rows.extend([r for r in block.status_country_rows.values() if r > country_row_rollout])

                    candidate_rows = sorted(list(set(candidate_rows)))
                    if len(candidate_rows) >= 3:
                        used_row = candidate_rows[0]
                        link_row = candidate_rows[1]
                        blank_row = candidate_rows[2]
                        style_recycled_row(ws, used_row, "USED OR NOT", fill_grey, is_last=False)
                        style_recycled_row(ws, link_row, "LINK YOUTUBE", fill_white, is_last=True)
                        ws.row_dimensions[blank_row].hidden = False
                        ws.row_dimensions[blank_row].height = 15
                        for c in range(1, 32):
                            cell = ws.cell(row=blank_row, column=c)
                            if isinstance(cell, Cell):
                                cell.value = None
                            cell.fill = PatternFill(fill_type=None)
                            cell.border = Border()
                        safe_rows.update([used_row, link_row, blank_row])

                dc_tc = block.tc_row
                if isinstance(dc_tc, int) and dc_tc > max_content_row:
                    max_content_row = dc_tc

                if country_row_rollout:
                    for col in range(4, 32):
                        cell = ws.cell(row=country_row_rollout, column=col)
                        cell_val = str(cell.value).strip().upper() if hasattr(cell, "value") else ""
                        if "YES" not in cell_val:
                            if isinstance(cell, Cell):
                                cell.value = None
                            rows_to_clear = [
                                block.content_type_row, block.property_row,
                                block.title_row, block.synopsis_row, block.strategy_row,
                            ]
                            if is_israel:
                                rows_to_clear.extend([block.tx_row, block.tc_row])
                            rows_to_clear = [r for r in rows_to_clear if r is not None]
                            for r in rows_to_clear:
                                clear_cell(ws, r, col)

                if is_israel:
                    israel_dc_tx = block.tx_row
                    if isinstance(israel_dc_tx, int):
                        format_israel_tx_tc(ws, israel_dc_tx, is_tx=True)
                        safe_rows.add(israel_dc_tx)
                    israel_dc_tc = block.tc_row
                    if isinstance(israel_dc_tc, int):
                        format_israel_tx_tc(ws, israel_dc_tc, is_tx=False)
                        safe_rows.add(israel_dc_tc)

        for r in range(5, 500):
            if get_val(ws, r, 2).startswith("week"):
                prev_r = r - 1
                safe_rows.add(r)
                safe_rows.add(prev_r)
                ws.row_dimensions[prev_r].hidden = False
                ws.row_dimensions[prev_r].height = 15
                for c in range(1, 32):
                    cell = ws.cell(row=prev_r, column=c)
                    if isinstance(cell, Cell):
                        cell.value = None
                    cell.fill = PatternFill(fill_type=None)
                    cell.border = Border()

        # ---------------- LA GUILLOTINA LÁSER ----------------
        rows_to_hide_and_clear = set()
        in_junk_section = False
        country_names_lower = [c.lower() for c in master_countries]

        standard_headers = [
            "deadline", "owner", "property", "content type", "asset",
            "title", "synopsis", "planning strategy", "used or not", "link youtube",
        ]
        junk_keywords = [
            "watermark", "bug", "copyright", "endcard", "branding",
            "status conforming", "comments", "total vod", "burned",
            "tx id", "tc in", "tz id", "tc im",
        ]

        weeks_out_of_month: Set[int] = set()
        if target_year is not None and target_month is not None and first_monday is not None:
            for w in range(6):
                w_monday = first_monday + timedelta(days=w * 7)
                w_sunday = w_monday + timedelta(days=6)
                if w_monday.month != target_month and w_sunday.month != target_month:
                    weeks_out_of_month.add(w)

        current_week_idx = -1
        sheet_max_row = ws.max_row if isinstance(ws.max_row, int) else 500

        for r in range(1, sheet_max_row + 1):
            val_a = get_val(ws, r, 1)
            val_b = get_val(ws, r, 2)
            val_c = get_val(ws, r, 3)
            header_col_text = f"{val_a} {val_b} {val_c}".lower()

            if val_b.startswith("week"):
                current_week_idx += 1

            if current_week_idx in weeks_out_of_month:
                rows_to_hide_and_clear.add(r)
                continue

            if "disney" in header_col_text and "channel" in header_col_text and "country" in header_col_text and r > 200:
                for row_to_end in range(r - 2, sheet_max_row + 1):
                    rows_to_hide_and_clear.add(row_to_end)
                break

            if any(h in header_col_text for h in standard_headers) or "week" in header_col_text:
                safe_rows.add(r)
                in_junk_section = False
                if "week" not in header_col_text:
                    if "title" in header_col_text:
                        ws.row_dimensions[r].height = 50
                    elif "synopsis" in header_col_text:
                        ws.row_dimensions[r].height = 85
                    else:
                        ws.row_dimensions[r].height = 35
                continue

            if r in safe_rows:
                in_junk_section = False
                continue

            if any(k in header_col_text for k in junk_keywords):
                in_junk_section = True

            if in_junk_section:
                rows_to_hide_and_clear.add(r)
                continue

            if val_b.startswith("tier ") or val_a.startswith("tier "):
                rows_to_hide_and_clear.add(r)
                continue

            if any(c in header_col_text for c in country_names_lower) or "total" in header_col_text:
                rows_to_hide_and_clear.add(r)
                continue

        ranges_to_unmerge = []
        for m in list(ws.merged_cells.ranges):
            if m.min_row in rows_to_hide_and_clear or m.max_row in rows_to_hide_and_clear:
                ranges_to_unmerge.append(m)

        for m in ranges_to_unmerge:
            try:
                ws.unmerge_cells(str(m))
            except Exception:
                pass

        for r in sorted(list(rows_to_hide_and_clear)):
            ws.row_dimensions[r].hidden = True
            ws.row_dimensions[r].height = 0
            for c in range(1, 40):
                try:
                    cell = ws.cell(row=r, column=c)
                    if isinstance(cell, Cell):
                        cell.value = None
                    cell.fill = PatternFill(fill_type=None)
                    cell.border = Border()
                except Exception:
                    pass

        sheet_max_col = ws.max_column if isinstance(ws.max_column, int) else 40
        for c in range(1, sheet_max_col + 1):
            c_val = str(ws.cell(row=1, column=c).value).lower()
            if "total vod" in c_val:
                ws.column_dimensions[get_column_letter(c)].hidden = True

        # --- REEMPLAZO DE NOMBRES EMEA -> COUNTRY CODE (<= 31 CHARS) ---
        for sn in country_wb.sheetnames:
            ws_tmp = country_wb[sn]
            for r in range(1, 15):
                for c in range(1, 10):
                    c_obj = ws_tmp.cell(row=r, column=c)
                    if isinstance(c_obj, Cell):
                        v = c_obj.value
                        if isinstance(v, str) and "EMEA" in v:
                            c_obj.value = v.replace("EMEA", country_code)

        for sn in list(country_wb.sheetnames):
            if "EMEA" in sn:
                new_title = sn.replace("EMEA", country_code)
                country_wb[sn].title = new_title[:31]

        filename = f"{month_num}_Social Media Calendar {brand_label} YouTube {clean_country} - {month_name_year}.xlsx"
        output_path = os.path.join(output_dir, filename)
        country_wb.template = False
        country_wb.save(output_path)
        country_wb.close()
        generated_files.append(output_path)

    zip_filename = f"Local_Calendars_{brand_label}_{month_num}.zip"
    zip_path = os.path.join(output_dir, zip_filename)
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zipf:
        for f in generated_files:
            zipf.write(f, os.path.basename(f))

    return zip_path