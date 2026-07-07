import os
import logging
from typing import List, Optional
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

def generate_local_calendars(master_path: str, output_dir: str, channel_type: str, selected_countries: Optional[List[str]] = None):
    """
    Fase 2: Genera un archivo Excel independiente por cada país filtrando 
    los contenidos que no aplican y ocultando las filas técnicas.
    """
    if not os.path.exists(master_path):
        logging.error(f"No se encontró el Master Calendar en: {master_path}")
        return

    os.makedirs(output_dir, exist_ok=True)
    
    # 1. Definimos las variables base según el canal
    if channel_type.upper() == "DJR":
        from exporters.full_calendar_djr import DJR_TARGET_SHEET as sheet_name
        from config.countries_djr import DJR_SHEET_COUNTRIES as master_countries
        anchors = {"B7": None, "B60": None, "B122": None, "B184": None, "B246": None}
    else:
        from config.layouts import TARGET_SHEET as sheet_name
        from config.countries import SHEET_COUNTRIES as master_countries
        anchors = {"B6": None, "B59": None, "B112": None, "B165": None, "B218": None}

    # Definimos qué países vamos a procesar
    countries_to_process = selected_countries if selected_countries else master_countries

    wb_meta = load_workbook(master_path, data_only=True)
    ws_meta = wb_meta[sheet_name]
    
    # Rellenamos el mock con fechas ficticias solo para que find_week_blocks las detecte
    from datetime import date
    from exporters.full_calendar import find_week_blocks
    anchor_dates = {addr: date.today() for addr in anchors}
    blocks = find_week_blocks(ws_meta, anchor_dates)
    wb_meta.close()

    logging.info(f"Iniciando Fase 2 para {channel_type}. Países seleccionados: {len(countries_to_process)}")

    # 2. Iteración principal por cada País
    for country in countries_to_process:
        if country.lower() in ["spain", "españa"]:
            logging.info(f"Saltando {country} (usa el Master Calendar).")
            continue

        logging.info(f"Generando calendario local para: {country}")
        
        country_wb = load_workbook(master_path)
        ws = country_wb[sheet_name]
        
        # 3. Procesamos cada bloque semanal independientemente
        for block in blocks.values():
            for col in range(4, 33):
                country_row_rollout = block.rollout_country_rows.get(country)
                
                if country_row_rollout:
                    cell_value = str(ws.cell(row=country_row_rollout, column=col).value).strip().upper()
                    
                    if "YES" not in cell_value:
                        rows_to_clear = [
                            block.content_type_row, block.property_row, block.title_row,
                            block.synopsis_row, block.strategy_row, block.tx_row,
                            block.tc_row, block.comments_row
                        ]
                        if block.tc_row:
                            for offset in range(1, 7):
                                rows_to_clear.append(block.tc_row + offset)
                                
                        for r in rows_to_clear:
                            if r: ws.cell(row=r, column=col).value = None

            # 4. LIMPIEZA VISUAL: Ocultar Filas Técnicas
            for r in block.rollout_country_rows.values(): ws.row_dimensions[r].hidden = True
            for r in block.status_country_rows.values(): ws.row_dimensions[r].hidden = True
            if block.comments_row: ws.row_dimensions[block.comments_row].hidden = True
                
            for r in range(block.anchor_row, block.end_row + 1):
                label = str(ws.cell(row=r, column=1).value).lower()
                if any(x in label for x in ["watermark", "bug", "copyright", "endcard", "branding"]):
                    ws.row_dimensions[r].hidden = True

            # EXCEPCIÓN ISRAEL
            if country.lower() != "israel":
                if block.tx_row: ws.row_dimensions[block.tx_row].hidden = True
                if block.tc_row: ws.row_dimensions[block.tc_row].hidden = True

        # 5. Guardamos el archivo final del país
        clean_country_name = country.replace("/", "_").replace(" ", "_")
        filename = f"Calendario_Local_{clean_country_name}.xlsx"
        output_path = os.path.join(output_dir, filename)
        
        country_wb.template = False
        country_wb.save(output_path)
        country_wb.close()
        logging.info(f"Guardado exitosamente: {filename}")