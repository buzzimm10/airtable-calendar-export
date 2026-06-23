import os
import sys
from datetime import datetime
import pandas as pd
import streamlit as st

# Allow imports from src/
CURRENT_DIR = os.path.dirname(__file__)
SRC_DIR = os.path.abspath(os.path.join(CURRENT_DIR, ".."))
if SRC_DIR not in sys.path:
    sys.path.append(SRC_DIR)

from services.planner_service import load_and_filter_records
from config.fields import AT_FIELDS
from populate_calendar import populate_calendar 

# --- PAGE CONFIG ---
st.set_page_config(
    page_title="Disney EMEA Calendar",
    page_icon="✨",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- STATE MANAGEMENT ---
if "records" not in st.session_state:
    st.session_state.records = None
if "channel_map" not in st.session_state:
    st.session_state.channel_map = {}

# --- HEADER PRINCIPAL ---
st.title("✨ Disney YouTube Calendar Tool")
st.markdown("Genera el **Full Calendar** y **Monthly Overview** en un par de clics.")
st.divider()

# --- SIDEBAR (CONTROLES) ---
with st.sidebar:
    st.header("1️⃣ Configuración")
    channel = st.selectbox("📺 Brand", ["Disney Channel", "Disney Jr."], index=0)
    
    col_yr, col_mo = st.columns(2)
    with col_yr:
        year = st.selectbox("📅 Year", ["2026", "2027", "2028"], index=0)
    with col_mo:
        month_number = st.selectbox(
            "📆 Month",
            [
                ("01", "Jan"), ("02", "Feb"), ("03", "Mar"), ("04", "Apr"), 
                ("05", "May"), ("06", "Jun"), ("07", "Jul"), ("08", "Aug"), 
                ("09", "Sep"), ("10", "Oct"), ("11", "Nov"), ("12", "Dec")
            ],
            index=4, format_func=lambda x: x[1]
        )
    month_str = f"{year}-{month_number[0]}"
    
    st.divider()
    
    # Ocultamos la complejidad técnica
    with st.expander("⚙️ Advanced Settings"):
        include_mediaplanning = st.checkbox("Include Mediaplanning", value=True)
        use_json = st.checkbox("Use local JSON", value=False)
        json_path = st.text_input("JSON path", value="data/sample_airtable.json") if use_json else None
        template_path = st.text_input("Template path", value="templates/X_Social Media Calendar DC YouTube EMEA - month 2026.xltx")
        output_folder = st.text_input("Output folder", value="output")

    st.divider()
    
    # Botón primario en el sidebar
    if st.button("🔍 Cargar Datos", type="primary", use_container_width=True):
        with st.spinner("Conectando con Airtable..."):
            try:
                records, channel_map = load_and_filter_records(
                    month=month_str,
                    channel=channel,
                    json_path=json_path,
                    log_level="INFO"
                )
                st.session_state.records = records
                st.session_state.channel_map = channel_map
            except Exception as e:
                st.error(f"Error cargando datos: {str(e)}")

# --- CONTENIDO PRINCIPAL ---
if st.session_state.records is None:
    # Estado inicial (Vacío)
    st.info("👈 Por favor, selecciona la marca y el mes en el menú lateral izquierdo y haz clic en **Cargar Datos**.")
else:
    records = st.session_state.records
    
    if len(records) == 0:
        st.warning(f"No se encontraron assets en Airtable para {channel} en el mes seleccionado.")
    else:
        # --- 2️⃣ QA SUMMARY ---
        st.subheader("2️⃣ Resumen de QA")
        qa_col1, qa_col2, qa_col3, qa_col4 = st.columns(4)
        qa_col1.metric("Total Assets", len(records))
        qa_col2.metric("🛑 Blockers", "0") # Mock por ahora
        qa_col3.metric("⚠️ Warnings", "0") # Mock por ahora
        qa_col4.metric("ℹ️ Info", "0")      # Mock por ahora
        
        # PREVIEW DATA (Escondido por defecto para que sea más clean)
        with st.expander("👀 Auditar datos extraídos (Preview)", expanded=False):
            preview_data = []
            for r in records:
                f = r.get("fields", r)
                preview_data.append({
                    "Live Date": f.get(AT_FIELDS["live_date"], "MISSING"),
                    "Content Type": f.get(AT_FIELDS["content_type"], ""),
                    "Title": f.get(AT_FIELDS["youtube_title"], f.get(AT_FIELDS["title"], "")),
                    "Owner": f.get(AT_FIELDS["owner_hub"], "")
                })
            # Ocultamos el index para que la tabla se vea más profesional
            st.dataframe(pd.DataFrame(preview_data), use_container_width=True, hide_index=True)
        
        st.divider()
        
        # --- 3️⃣ GENERAR EXCEL ---
        st.subheader("3️⃣ Exportar Reporte")
        output_filename = f"calendar_{channel.lower().replace(' ', '_')}_{month_str.lower()}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        output_path = os.path.join(output_folder, output_filename)
        
        # Usamos columnas para que el botón no sea absurdamente ancho
        btn_col1, btn_col2, btn_col3 = st.columns([1, 1, 2])
        
        with btn_col1:
            if st.button("🚀 Generar Excel", type="primary", use_container_width=True):
                with st.spinner("Escribiendo Excel... No cierres esta ventana."):
                    try:
                        os.makedirs(output_folder, exist_ok=True)
                        populate_calendar(
                            records=records,
                            template_path=template_path,
                            output_path=output_path,
                            enable_mediaplanning=include_mediaplanning,
                            channel_id_to_name_map=st.session_state.channel_map
                        )
                        st.session_state.export_success = True
                        st.session_state.output_path = output_path
                        st.session_state.output_filename = output_filename
                    except Exception as e:
                        st.error(f"Error generando Excel: {str(e)}")
        
        # Si el excel se generó bien, mostramos el botón de descarga al lado
        if st.session_state.get("export_success", False):
            with btn_col2:
                with open(st.session_state.output_path, "rb") as file:
                    st.download_button(
                        label="⬇️ Descargar Archivo",
                        data=file,
                        file_name=st.session_state.output_filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="secondary",
                        use_container_width=True
                    )
            st.success("✅ ¡Reporte listo! Haz clic en Descargar Archivo.")