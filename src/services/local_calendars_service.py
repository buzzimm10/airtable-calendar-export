import os
import logging
import zipfile
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

def generate_local_calendars(master_path: str, output_dir: str, channel_type: str, 
                             selected_countries: Optional[List[str]] = None,
                             month_str: str = "2026-05", 
                             dc_blocks: Optional[dict] = None,
                             records: Optional[list] = None, 
                             channel_id_to_name_map: Optional[dict] = None):
    
    if not os.path.exists(master_path): return None

    os.makedirs(output_dir, exist_ok=True)
    for f in os.listdir(output_dir):
        if f.endswith(".xlsx") or f.endswith(".zip"): os.remove(os.path.join(output_dir, f))

    is_djr = channel_type.upper() == "DJR"
    sheet_name = djr_sheet_name if is_djr else dc_sheet_name
    master_countries = djr_master_countries if is_djr else dc_master_countries
    countries_to_process = selected_countries if selected_countries is not None else master_countries
    
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

    for country in countries_to_process:
        if country.lower() in ["spain", "españa"]: continue
        country_wb = load_workbook(master_path, data_only=False)
        
        for sn in country_wb.sheetnames:
            if "compilation" in sn.lower(): del country_wb[sn]
        if sheet_name not in country_wb.sheetnames:
            country_wb.close()
            continue
            
        ws = country_wb[sheet_name]
        is_israel = country.lower() == "israel"
        last_processed_row = 0

        # ---------------- LÓGICA DISNEY JUNIOR ----------------
        if is_djr:
            country_idx = next((i for i, mc in enumerate(master_countries) if mc.lower() == country.lower()), -1)
            if country_idx == -1: country_wb.close(); continue
                
            for week_num, layout in djr_layout.items():
                rollout_start = layout["country_rows"][0]
                my_rollout_row = rollout_start + country_idx
                
                for i in range(len(master_countries)):
                    ws.row_dimensions[rollout_start + i].hidden = True
                if "status_rows" in layout:
                    for i in range(len(master_countries)):
                        ws.row_dimensions[layout["status_rows"][0] + i].hidden = True

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
                    # FORZAR A ISRAEL A MOSTRAR LAS FILAS (Unhide)
                    tx_r = layout.get("tx_row")
                    tc_r = layout.get("tc_row")
                    if tx_r is not None: ws.row_dimensions[tx_r].hidden = False
                    if tc_r is not None: ws.row_dimensions[tc_r].hidden = False

        # ---------------- LÓGICA DISNEY CHANNEL ----------------
        else:
            if not dc_blocks: country_wb.close(); continue
            for block in dc_blocks.values():
                target_key = next((k for k in block.rollout_country_rows.keys() if k.lower() == country.lower()), None)
                country_row_rollout = block.rollout_country_rows.get(target_key) if target_key else None
                
                if block.end_row > last_processed_row: last_processed_row = block.end_row

                for r in block.rollout_country_rows.values(): ws.row_dimensions[r].hidden = True
                for r in block.status_country_rows.values(): ws.row_dimensions[r].hidden = True
                
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
                    # FORZAR A ISRAEL A MOSTRAR LAS FILAS (Unhide)
                    if block.tx_row: ws.row_dimensions[block.tx_row].hidden = False
                    if block.tc_row: ws.row_dimensions[block.tc_row].hidden = False

        # ---------------- OCULTAR EXCEDENTES (PURGA RADICAL) ----------------
        country_names_lower = [c.lower() for c in master_countries]
        
        for r in range(1, 400):
            val_a = get_val(ws, r, 1)
            val_b = get_val(ws, r, 2)
            val_c = get_val(ws, r, 3)
            
            # Buscar Comments en A, B o C aniquilará la fila roja de Status Conforming
            if any(x in val_a for x in ["watermark", "bug", "copyright", "endcard", "branding", "comments"]) or \
               any(x in val_b for x in ["watermark", "bug", "copyright", "endcard", "branding", "comments"]) or \
               any(x in val_c for x in ["comments"]):
                ws.row_dimensions[r].hidden = True
                for offset in range(1, 8):
                    if r + offset <= ws.max_row and get_val(ws, r+offset, 1) == "" and get_val(ws, r+offset, 2) == "":
                        ws.row_dimensions[r+offset].hidden = True
                    else: break
            
            if "status conforming" in val_a or "status conforming" in val_b or \
               val_b == "total" or val_c == "total" or \
               val_b in country_names_lower or val_a in country_names_lower:
                ws.row_dimensions[r].hidden = True

        if last_processed_row > 0:
            for r in range(last_processed_row + 5, 500): ws.row_dimensions[r].hidden = True

        if records:
            if is_djr:
                from exporters.monthly_overview_djr import populate_monthly_overview_djr
                populate_monthly_overview_djr(country_wb, records=records, enable_mediaplanning=True, channel_id_to_name_map=channel_id_to_name_map, target_country=country)
            else:
                from exporters.monthly_overview import populate_monthly_overview
                populate_monthly_overview(country_wb, records=records, enable_mediaplanning=True, channel_id_to_name_map=channel_id_to_name_map, target_country=country)

        clean_country = country.replace("/", "_").replace(" ", "_")
        filename = f"{month_num}_Social Media Calendar {brand_label} YouTube EMEA - {month_name_year} - {clean_country}.xlsx"
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