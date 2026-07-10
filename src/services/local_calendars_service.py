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
from openpyxl.styles import PatternFill, Font, Alignment, Border

def get_clean_country_name(name: str) -> str:
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

    def style_recycled_row(sheet, row_idx, label, fill_color):
        if not isinstance(row_idx, int): return
        safe_set_value(sheet, row_idx, 2, label)
        sheet.row_dimensions[row_idx].hidden = False
        sheet.row_dimensions[row_idx].height = 15  
        
        # Limpiar y pintar toda la fila (Borra los YES/NO viejos)
        for c in range(3, 34): 
            cell = sheet.cell(row=row_idx, column=c)
            if not isinstance(cell, MergedCell):
                cell.value = None
            cell.fill = fill_color
            
        cell2 = sheet.cell(row=row_idx, column=2)
        cell2.fill = fill_color
        cell2.font = font_bold
        cell2.alignment = Alignment(horizontal="right", vertical="center")

    def clear_cell(sheet, row_idx, col_idx):
        if isinstance(row_idx, int) and isinstance(col_idx, int):
            cell_obj = sheet.cell(row=row_idx, column=col_idx)
            if not isinstance(cell_obj, MergedCell):
                cell_obj.value = None

    def get_val(sheet, r, c):
        if not isinstance(r, int) or not isinstance(c, int): return ""
        v = sheet.cell(row=r, column=c).value
        return str(v).strip().lower() if v is not None else ""

    def format_israel_tx_tc(sheet, row_idx, is_tx=False):
        if not isinstance(row_idx, int): return
        sheet.row_dimensions[row_idx].hidden = False
        sheet.row_dimensions[row_idx].height = 45 if is_tx else 25 
        for c in range(4, 34):
            cell = sheet.cell(row=row_idx, column=c)
            if not isinstance(cell, MergedCell):
                cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

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
        
        max_content_row: int = 0
        current_max_row: int = ws.max_row if isinstance(ws.max_row, int) else 500

        ws.column_dimensions['AG'].hidden = True

        if records:
            if is_djr:
                from exporters.monthly_overview_djr import populate_monthly_overview_djr
                populate_monthly_overview_djr(country_wb, records=records, enable_mediaplanning=True, channel_id_to_name_map=channel_id_to_name_map, target_country=clean_country)
            else:
                from exporters.monthly_overview import populate_monthly_overview
                populate_monthly_overview(country_wb, records=records, enable_mediaplanning=True, channel_id_to_name_map=channel_id_to_name_map, target_country=clean_country)

        # ---------------- LÓGICA DISNEY JUNIOR ----------------
        if is_djr:
            country_idx = next((i for i, mc in enumerate(master_countries) if mc.lower() == country.lower()), -1)
            if country_idx == -1: country_wb.close(); continue
                
            for week_num, layout in djr_layout.items():
                rollout_start = layout["country_rows"][0]
                my_rollout_row = rollout_start + country_idx
                
                djr_tc = layout.get("tc_row")
                if isinstance(djr_tc, int) and djr_tc > max_content_row:
                    max_content_row = djr_tc
                
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
                    format_israel_tx_tc(ws, layout.get("tx_row"), is_tx=True)
                    format_israel_tx_tc(ws, layout.get("tc_row"), is_tx=False)
                    
                    # Buscar filas FÍSICAMENTE DEBAJO de TC IN & OUT
                    available_rows = [r for r in range(rollout_start, rollout_start + len(master_countries)) if isinstance(djr_tc, int) and r > djr_tc and r != my_rollout_row]
                    available_rows.sort()
                    if len(available_rows) >= 2:
                        style_recycled_row(ws, available_rows[0], "USED OR NOT", fill_grey)
                        style_recycled_row(ws, available_rows[1], "LINK YOUTUBE", fill_white)

        # ---------------- LÓGICA DISNEY CHANNEL ----------------
        else:
            if not dc_blocks: country_wb.close(); continue
            for block in dc_blocks.values():
                target_key = next((k for k in block.rollout_country_rows.keys() if k.lower() == country.lower()), None)
                country_row_rollout = block.rollout_country_rows.get(target_key) if target_key else None
                
                dc_tc = block.tc_row
                if isinstance(dc_tc, int) and dc_tc > max_content_row:
                    max_content_row = dc_tc
                
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
                    format_israel_tx_tc(ws, block.tx_row, is_tx=True)
                    format_israel_tx_tc(ws, block.tc_row, is_tx=False)
                    
                    # Buscar filas FÍSICAMENTE DEBAJO de TC IN & OUT
                    available_rows = [r for r in block.rollout_country_rows.values() if isinstance(dc_tc, int) and r > dc_tc and r != country_row_rollout]
                    available_rows.sort()
                    if len(available_rows) >= 2:
                        style_recycled_row(ws, available_rows[0], "USED OR NOT", fill_grey)
                        style_recycled_row(ws, available_rows[1], "LINK YOUTUBE", fill_white)

        # ---------------- LA PURGA INVENCIBLE (LAS 11 FILAS) ----------------
        in_junk_section = False
        country_names_lower = [c.lower() for c in master_countries]
        sniper_limit: int = max_content_row + 25 if max_content_row > 0 else 400
        
        for r in range(1, sniper_limit):
            val_a = get_val(ws, r, 1)
            val_b = get_val(ws, r, 2)
            val_c = get_val(ws, r, 3)
            val_d = get_val(ws, r, 4)
            combined_val = val_a + " " + val_b + " " + val_c + " " + val_d
            
            # Resetear la alarma cuando empieza una nueva semana
            if val_b.startswith("week "):
                in_junk_section = False
            
            # Activar alarma al detectar basura (Esto aniquila las 11 filas de Status Conforming seguidas)
            if any(x in combined_val for x in ["watermark", "bug", "copyright", "endcard", "branding", "status conforming", "comments", "total vod"]):
                in_junk_section = True

            if in_junk_section:
                ws.row_dimensions[r].hidden = True
                
            # Ocultar Tiers extra
            if val_b.startswith("tier ") or val_a.startswith("tier "):
                ws.row_dimensions[r].hidden = True

            # Ocultar países no usados
            if any(c in val_a + val_b + val_c for c in country_names_lower) or val_b == "total" or val_c == "total":
                ws.row_dimensions[r].hidden = True

        # ---------------- LA FILA BLANCA PERFECTA ENTRE SEMANAS ----------------
        for r in range(5, 400):
            if get_val(ws, r, 2).startswith("week "):
                prev_r = r - 1
                ws.row_dimensions[prev_r].hidden = False
                ws.row_dimensions[prev_r].height = 15
                for c in range(1, 35):
                    cell = ws.cell(row=prev_r, column=c)
                    if not isinstance(cell, MergedCell):
                        cell.value = None
                        cell.fill = PatternFill(fill_type=None)
                        cell.border = Border() # Borde invisible

        # ---------------- LA GUILLOTINA FINAL (Fila 273+) ----------------
        if max_content_row > 0:
            for r in range(max_content_row + 10, current_max_row + 50): 
                ws.row_dimensions[r].hidden = True

        # --- REEMPLAZO DE NOMBRES EMEA -> COUNTRY ---
        for sn in country_wb.sheetnames:
            ws_tmp = country_wb[sn]
            for r in range(1, 15):
                for c in range(1, 10):
                    c_obj = ws_tmp.cell(row=r, column=c)
                    if not isinstance(c_obj, MergedCell):
                        v = c_obj.value
                        if isinstance(v, str) and "EMEA" in v:
                            c_obj.value = v.replace("EMEA", clean_country.upper())
        
        for sn in country_wb.sheetnames:
            if "EMEA" in sn:
                new_title = sn.replace("EMEA", clean_country.upper())
                country_wb[sn].title = new_title[:31] 

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