import os
import logging
import zipfile
import re
from typing import List, Optional
from openpyxl import load_workbook
from datetime import datetime
from openpyxl.cell.cell import MergedCell

from config.layouts import TARGET_SHEET as dc_sheet_name
from config.countries import SHEET_COUNTRIES as dc_master_countries
from config.layouts import DAY_ORDER
from exporters.full_calendar_djr import DJR_TARGET_SHEET as djr_sheet_name
from config.countries_djr import DJR_SHEET_COUNTRIES as djr_master_countries
from config.layouts_djr import DJR_WEEK_LAYOUT as djr_layout
from config.layouts_djr import DJR_DAY_LAYOUT as djr_day
from core.excel_utils import safe_set_value
from openpyxl.styles import PatternFill, Font, Alignment

def get_clean_country_name(name: str) -> str:
    """Elimina el idioma entre paréntesis, ej: 'Israel (Hebrew)' -> 'Israel'"""
    return re.sub(r'\s*\(.*?\)', '', name).strip()

def generate_local_calendars(master_path: str, output_dir: str, channel_type: str, 
                             selected_countries: Optional[List[str]] = None,
                             month_str: str = "2026-05", 
                             dc_blocks: Optional[dict] = None,
                             records: Optional[list] = None, 
                             channel_id_to_name_map: Optional[dict] = None):
    
    if not os.path.exists(master_path) or not selected_countries: return None

    os.makedirs(output_dir, exist_ok=True)
    for f in os.listdir(output_dir):
        if f.endswith(".xlsx") or f.endswith(".zip"): os.remove(os.path.join(output_dir, f))

    is_djr = channel_type.upper() == "DJR"
    sheet_name = djr_sheet_name if is_djr else dc_sheet_name
    master_countries = djr_master_countries if is_djr else dc_master_countries
    
    try:
        dt = datetime.strptime(month_str, "%Y-%m")
        month_num = dt.strftime("%m")
        month_name_year = dt.strftime("%B %Y")
    except:
        month_num = "00"; month_name_year = month_str

    brand_label = "DJR" if is_djr else "DC"
    generated_files = []

    fill_grey = PatternFill(fill_type="solid", fgColor="D0CECE")
    fill_white = PatternFill(fill_type="solid", fgColor="FFFFFF")
    font_bold = Font(bold=True, color="000000")
    align_right = Alignment(horizontal="right", vertical="center")

    def style_recycled_row(sheet, row_idx, label, fill_color):
        if not row_idx: return
        safe_set_value(sheet, row_idx, 2, label)
        sheet.row_dimensions[row_idx].hidden = False
        sheet.row_dimensions[row_idx].height = 15  
        for c in range(2, 34): 
            cell = sheet.cell(row=row_idx, column=c)
            cell.fill = fill_color
            if c == 2:
                cell.font = font_bold
                cell.alignment = align_right

    def clear_cell(sheet, row_idx, col_idx):
        if row_idx and col_idx:
            cell_obj = sheet.cell(row=row_idx, column=col_idx)
            if not isinstance(cell_obj, MergedCell):
                cell_obj.value = None

    def get_val(sheet, r, c):
        v = sheet.cell(row=r, column=c).value
        return str(v).strip().lower() if v is not None else ""

    for country in selected_countries:
        if country.lower() in ["spain", "españa"]: continue
        country_wb = load_workbook(master_path, data_only=False)
        clean_country = get_clean_country_name(country)
        
        for sn in country_wb.sheetnames:
            if "compilation" in sn.lower(): del country_wb[sn]
            
        if sheet_name not in country_wb.sheetnames:
            country_wb.close()
            continue
            
        ws = country_wb[sheet_name]
        is_israel = clean_country.lower() == "israel"
        last_processed_row = 0

        # Ocultar columna AG (Total VOD)
        ws.column_dimensions['AG'].hidden = True

        for r in range(1, 400):
            if get_val(ws, r, 2).startswith("week "):
                ws.row_dimensions[r - 1].hidden = False
                ws.row_dimensions[r - 1].height = 25

        # ---------------- LÓGICA DISNEY JUNIOR ----------------
        if is_djr:
            country_idx = next((i for i, mc in enumerate(master_countries) if mc.lower() == country.lower()), -1)
            if country_idx == -1: country_wb.close(); continue
                
            for week_num, layout in djr_layout.items():
                rollout_start = layout["country_rows"][0]
                my_rollout_row = rollout_start + country_idx
                
                tc_r = layout.get("tc_row", 0)
                if tc_r > last_processed_row: last_processed_row = tc_r
                
                for day in DAY_ORDER:
                    for role, col in djr_day[day]:
                        cell_val = str(ws.cell(row=my_rollout_row, column=col).value).strip().upper()
                        if "YES" not in cell_val:
                            rows_to_clear = [layout["property_row"], layout["content_type_row"], layout["title_row"], layout["synopsis_row"], layout["strategy_row"], layout.get("tx_row"), layout.get("tc_row")]
                            for r in rows_to_clear: clear_cell(ws, r, col)
                        elif not is_israel:
                            clear_cell(ws, layout.get("tx_row"), col)
                            clear_cell(ws, layout.get("tc_row"), col)

                if not is_israel:
                    style_recycled_row(ws, layout.get("tx_row"), "USED OR NOT", fill_grey)
                    style_recycled_row(ws, layout.get("tc_row"), "LINK YOUTUBE", fill_white)
                else:
                    # FIX TYPING PYLANCE
                    israel_tx = layout.get("tx_row")
                    if israel_tx: 
                        ws.row_dimensions[israel_tx].hidden = False
                        ws.row_dimensions[israel_tx].height = 20
                    israel_tc = layout.get("tc_row")
                    if israel_tc: 
                        ws.row_dimensions[israel_tc].hidden = False
                        ws.row_dimensions[israel_tc].height = 20

        # ---------------- LÓGICA DISNEY CHANNEL ----------------
        else:
            if not dc_blocks: country_wb.close(); continue
            for block in dc_blocks.values():
                target_key = next((k for k in block.rollout_country_rows.keys() if k.lower() == country.lower()), None)
                country_row_rollout = block.rollout_country_rows.get(target_key) if target_key else None
                
                if block.end_row > last_processed_row: last_processed_row = block.end_row
                
                if country_row_rollout:
                    for col in range(4, 34): 
                        cell_val = str(ws.cell(row=country_row_rollout, column=col).value).strip().upper()
                        if "YES" not in cell_val:
                            rows_to_clear = [block.content_type_row, block.property_row, block.title_row, block.synopsis_row, block.strategy_row, block.tx_row, block.tc_row, block.comments_row]
                            for r in rows_to_clear: clear_cell(ws, r, col)
                        elif not is_israel:
                            clear_cell(ws, block.tx_row, col)
                            clear_cell(ws, block.tc_row, col)

                if not is_israel:
                    style_recycled_row(ws, block.tx_row, "USED OR NOT", fill_grey)
                    style_recycled_row(ws, block.tc_row, "LINK YOUTUBE", fill_white)
                else:
                    # FIX TYPING PYLANCE
                    israel_dc_tx = block.tx_row
                    if israel_dc_tx: 
                        ws.row_dimensions[israel_dc_tx].hidden = False
                        ws.row_dimensions[israel_dc_tx].height = 20
                    israel_dc_tc = block.tc_row
                    if israel_dc_tc: 
                        ws.row_dimensions[israel_dc_tc].hidden = False
                        ws.row_dimensions[israel_dc_tc].height = 20

        # ---------------- OCULTAR EXCEDENTES (PURGA RADICAL) ----------------
        country_names_lower = [c.lower() for c in master_countries]
        
        for r in range(1, 400):
            val_a = get_val(ws, r, 1)
            val_b = get_val(ws, r, 2)
            val_c = get_val(ws, r, 3)
            val_d = get_val(ws, r, 4) 
            
            if any(x in val_a + val_b + val_c + val_d for x in ["watermark", "bug", "copyright", "endcard", "branding", "comments", "total vod", "status conforming"]):
                ws.row_dimensions[r].hidden = True
                for offset in range(1, 8):
                    if r + offset <= ws.max_row and get_val(ws, r+offset, 1) == "" and get_val(ws, r+offset, 2) == "":
                        ws.row_dimensions[r+offset].hidden = True
                    else: break
            
            if any(c in val_a + val_b for c in country_names_lower) or val_b == "total" or val_c == "total":
                ws.row_dimensions[r].hidden = True

        if last_processed_row > 0:
            for r in range(last_processed_row + 5, 500): ws.row_dimensions[r].hidden = True

        # --- REEMPLAZO DE NOMBRES EMEA -> COUNTRY (FIX PYLANCE MERGEDCELL) ---
        for sn in country_wb.sheetnames:
            ws_tmp = country_wb[sn]
            for r in range(1, 15):
                for c in range(1, 10):
                    c_obj = ws_tmp.cell(row=r, column=c)
                    # Solo reemplazamos si es una celda real, esquivando las combinadas
                    if not isinstance(c_obj, MergedCell):
                        v = c_obj.value
                        if isinstance(v, str) and "EMEA" in v:
                            c_obj.value = v.replace("EMEA", clean_country.upper())
        
        for sn in country_wb.sheetnames:
            if "EMEA" in sn:
                new_title = sn.replace("EMEA", clean_country.upper())
                country_wb[sn].title = new_title[:31] 

        # --- GENERAR OVERVIEW LOCAL ---
        if records:
            if is_djr:
                from exporters.monthly_overview_djr import populate_monthly_overview_djr
                populate_monthly_overview_djr(country_wb, records=records, enable_mediaplanning=True, channel_id_to_name_map=channel_id_to_name_map, target_country=country)
            else:
                from exporters.monthly_overview import populate_monthly_overview
                populate_monthly_overview(country_wb, records=records, enable_mediaplanning=True, channel_id_to_name_map=channel_id_to_name_map, target_country=country)

        filename = f"{month_num}_Social Media Calendar {brand_label} YouTube {clean_country} - {month_name_year}.xlsx"
        output_path = os.path.join(output_dir, filename)
        country_wb.template = False
        country_wb.save(output_path)
        country_wb.close()
        generated_files.append(output_path)

    zip_filename = f"Local_Calendars_{brand_label}_{month_num}.zip"
    zip_path = os.path.join(output_dir, zip_filename)
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for f in generated_files: zipf.write(f, os.path.basename(f))
            
    return zip_path