import os
import logging
import zipfile
from typing import List, Optional
from openpyxl import load_workbook
from datetime import datetime

# Movemos las importaciones arriba para que Pylance sepa que siempre existen
from config.layouts import TARGET_SHEET as dc_sheet_name
from config.countries import SHEET_COUNTRIES as dc_master_countries
from config.layouts import DAY_ORDER
from exporters.full_calendar_djr import DJR_TARGET_SHEET as djr_sheet_name
from config.countries_djr import DJR_SHEET_COUNTRIES as djr_master_countries
from config.layouts_djr import DJR_WEEK_LAYOUT as djr_layout
from config.layouts_djr import DJR_DAY_LAYOUT as djr_day

def generate_local_calendars(master_path: str, output_dir: str, channel_type: str, 
                             selected_countries: Optional[List[str]] = None,
                             month_str: str = "2026-05", 
                             dc_blocks: Optional[dict] = None): # <-- FIX: Optional[dict]
    """
    Fase 2: Genera un archivo Excel por país, oculta filas técnicas, 
    elimina sheets innecesarios y comprime el resultado en un ZIP.
    """
    if not os.path.exists(master_path):
        logging.error(f"Master Calendar no encontrado: {master_path}")
        return None

    os.makedirs(output_dir, exist_ok=True)
    for f in os.listdir(output_dir):
        if f.endswith(".xlsx") or f.endswith(".zip"):
            os.remove(os.path.join(output_dir, f))

    is_djr = channel_type.upper() == "DJR"
    
    # Asignamos las variables según el canal
    sheet_name = djr_sheet_name if is_djr else dc_sheet_name
    master_countries = djr_master_countries if is_djr else dc_master_countries

    countries_to_process = selected_countries if selected_countries is not None else master_countries
    
    # Manejo robusto del nombre del mes
    try:
        dt = datetime.strptime(month_str, "%Y-%m")
        month_num = dt.strftime("%m")
        month_name_year = dt.strftime("%B %Y")
    except:
        month_num = "00"
        month_name_year = month_str

    brand_label = "DJR" if is_djr else "DC"
    generated_files = []

    for country in countries_to_process:
        if country.lower() in ["spain", "españa"]:
            continue

        # Usamos data_only=False para NO romper los estilos ni fórmulas del Master
        country_wb = load_workbook(master_path, data_only=False)
        
        # Eliminar pestaña de Compilations
        for sn in country_wb.sheetnames:
            if "compilation" in sn.lower():
                del country_wb[sn]

        if sheet_name not in country_wb.sheetnames:
            country_wb.close()
            continue
            
        ws = country_wb[sheet_name]

        # --- LÓGICA DISNEY JUNIOR ---
        if is_djr:
            country_idx = -1
            for i, mc in enumerate(master_countries):
                if mc.lower() == country.lower():
                    country_idx = i
                    break
                    
            if country_idx == -1:
                country_wb.close()
                continue
                
            for week_num, layout in djr_layout.items():
                rollout_start = layout["country_rows"][0]
                status_start = layout["status_rows"][0]
                my_rollout_row = rollout_start + country_idx

                for day in DAY_ORDER:
                    for role, col in djr_day[day]:
                        cell_val = str(ws.cell(row=my_rollout_row, column=col).value).strip().upper()
                        if "YES" not in cell_val:
                            rows_to_clear = [
                                layout["property_row"], layout["content_type_row"],
                                layout["title_row"], layout["synopsis_row"],
                                layout["strategy_row"], layout["tx_row"], layout["tc_row"]
                            ]
                            rows_to_clear.extend(layout["comments_rows"].values())
                            
                            for r in rows_to_clear: 
                                if r: ws.cell(row=r, column=col).value = None

                for i in range(len(master_countries)):
                    ws.row_dimensions[rollout_start + i].hidden = True
                    ws.row_dimensions[status_start + i].hidden = True

                for r in layout["comments_rows"].values(): ws.row_dimensions[r].hidden = True
                
                if country.lower() != "israel":
                    ws.row_dimensions[layout["tx_row"]].hidden = True
                    ws.row_dimensions[layout["tc_row"]].hidden = True

        # --- LÓGICA DISNEY CHANNEL ---
        else:
            if not dc_blocks:
                country_wb.close()
                continue
                
            for block in dc_blocks.values():
                target_key = next((k for k in block.rollout_country_rows.keys() if k.lower() == country.lower()), None)
                country_row_rollout = block.rollout_country_rows.get(target_key) if target_key else None
                
                if country_row_rollout:
                    for col in range(4, 34): 
                        cell_val = str(ws.cell(row=country_row_rollout, column=col).value).strip().upper()
                        if "YES" not in cell_val:
                            rows_to_clear = [
                                block.content_type_row, block.property_row, block.title_row,
                                block.synopsis_row, block.strategy_row, block.tx_row,
                                block.tc_row, block.comments_row
                            ]
                            for r in rows_to_clear: 
                                if r: ws.cell(row=r, column=col).value = None

                for r in block.rollout_country_rows.values(): ws.row_dimensions[r].hidden = True
                for r in block.status_country_rows.values(): ws.row_dimensions[r].hidden = True
                if block.comments_row: ws.row_dimensions[block.comments_row].hidden = True

                if country.lower() != "israel":
                    if block.tx_row: ws.row_dimensions[block.tx_row].hidden = True
                    if block.tc_row: ws.row_dimensions[block.tc_row].hidden = True

        # Ocultar Branding Dinámicamente
        for r in range(1, ws.max_row + 1):
            label = str(ws.cell(row=r, column=1).value).lower()
            if any(x in label for x in ["watermark", "bug", "copyright", "endcard", "branding"]):
                ws.row_dimensions[r].hidden = True

        # Guardar archivo local
        clean_country = country.replace("/", "_").replace(" ", "_")
        filename = f"{month_num}_Social Media Calendar {brand_label} YouTube EMEA - {month_name_year} - {clean_country}.xlsx"
        output_path = os.path.join(output_dir, filename)
        
        country_wb.template = False
        country_wb.save(output_path)
        country_wb.close()
        generated_files.append(output_path)
        logging.info(f"Creado: {filename}")

    # Empaquetar en ZIP
    zip_filename = f"Local_Calendars_{brand_label}_{month_num}.zip"
    zip_path = os.path.join(output_dir, zip_filename)
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for f in generated_files:
            zipf.write(f, os.path.basename(f))
            
    return zip_path