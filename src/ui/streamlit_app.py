import os
import sys
from datetime import datetime
from collections import Counter
import pandas as pd
import streamlit as st

# Allow imports from src/
CURRENT_DIR = os.path.dirname(__file__)
SRC_DIR = os.path.abspath(os.path.join(CURRENT_DIR, ".."))
if SRC_DIR not in sys.path:
    sys.path.append(SRC_DIR)

from services.planner_service import load_and_filter_records, hub_role
from config.fields import AT_FIELDS
from config.styles import OVERVIEW_FILL_COLORS # <-- NUEVO
from config.content_types import MONTHLY_OVERVIEW_MAPPING # <-- NUEVO
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
    
    # --- Lógica de default: Próximo Mes ---
    hoy = datetime.now()
    next_m = hoy.month + 1 if hoy.month < 12 else 1
    next_y = hoy.year if hoy.month < 12 else hoy.year + 1
    
    year_options = ["2026", "2027", "2028", "2029"]
    try:
        default_year_idx = year_options.index(str(next_y))
    except ValueError:
        default_year_idx = 0  # Fallback por si acaso
        
    default_month_idx = next_m - 1  # Los índices van de 0 a 11
    # --------------------------------------

    with col_yr:
        year = st.selectbox("📅 Year", year_options, index=default_year_idx)
    with col_mo:
        month_number = st.selectbox(
            "📆 Month",
            [
                ("01", "Jan"), ("02", "Feb"), ("03", "Mar"), ("04", "Apr"), 
                ("05", "May"), ("06", "Jun"), ("07", "Jul"), ("08", "Aug"), 
                ("09", "Sep"), ("10", "Oct"), ("11", "Nov"), ("12", "Dec")
            ],
            index=default_month_idx, format_func=lambda x: x[1]
        )
    month_str = f"{year}-{month_number[0]}"
    
    st.divider()
    
    # Lógica del nombre del Template (Fijo a YYYY)
    if channel == "Disney Channel":
        template_name = "X_Social Media Calendar DC YouTube EMEA - month YYYY.xltx"
    else:
        template_name = "X_Social Media Calendar DJR YouTube EMEA - month YYYY.xltx"
        
    default_template_path = f"templates/{template_name}"
    
    # Ocultamos la complejidad técnica
    with st.expander("⚙️ Advanced Settings"):
        include_mediaplanning = st.checkbox("Include Mediaplanning", value=True)
        use_json = st.checkbox("Use local JSON", value=False)
        json_path = st.text_input("JSON path", value="data/sample_airtable.json") if use_json else None
        
        # El text_input ahora se alimenta de la ruta dinámica
        template_path = st.text_input("Template path", value=default_template_path)
        output_folder = st.text_input("Output folder", value="output")

    st.divider()
    
    if st.button("📡 Fetch Data", type="primary", use_container_width=True):
        with st.spinner("Conectando con Airtable API..."):
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
    # --- FIRMA / CRÉDITO ---
    st.markdown("<br>" * 5, unsafe_allow_html=True) # Empujamos la firma hacia abajo
    st.caption("⚙️ Built by Augusto Zimmermann")

# --- CONTENIDO PRINCIPAL ---
if st.session_state.records is None:
    # Estado inicial: Dashboard Técnico (Geek Mode)
    st.markdown("### ⚡ Sistema de Automatización de Calendarios")
    st.info("👈 Esperando parámetros de ejecución. Selecciona el entorno en el panel izquierdo.")
    
    st.markdown("<br>", unsafe_allow_html=True)
    
    col_dash1, col_dash2 = st.columns([1, 1.5])
    
    with col_dash1:
        st.markdown("#### 🔧 Pipeline Config")
        # Mostramos la configuración en formato JSON para ese look "developer"
        pipeline_config = {
            "source": "Airtable API v1" if not st.session_state.get("use_json") else "Local JSON",
            "target_brand": channel,
            "target_period": month_str,
            "engine": "openpyxl (Python 3)",
            "template_file": template_name,
            "status": "AWAITING_EXECUTION"
        }
        st.json(pipeline_config)
        
    with col_dash2:
        st.markdown("#### 🏗️ Arquitectura de Exportación")
        # Diagrama ASCII estilo consola
        st.code('''
[📡 Airtable] 
     │ (Fetch JSON Payload)
     ▼
[🧠 Planner Service] ──(Sorting & Rules)──> [🚦 QA Engine]
     │                                            │ (Blockers / Warnings)
     ▼                                            ▼
[📝 Excel Exporters] <──(openpyxl)── [📁 Templates (.xltx)]
     │
     ▼
[📊 Final .xlsx Output]
        ''', language="text")

else:
    records = st.session_state.records
    
    if len(records) == 0:
        st.warning(f"No se encontraron assets en Airtable para {channel} en el mes seleccionado.")
    else:
    #--- 2️⃣ QA SUMMARY ---
       # --- NUEVO: Procesamiento de Métricas (Insights) ---
        hub_counts = Counter()
        ct_counts = Counter()

        for r in records:
            f = r.get("fields", r)
            
            # 1. Contar Hubs
            raw_hub = f.get(AT_FIELDS["owner_hub"])
            # Usamos tu función hub_role para saber si es Programming o Mediaplanning
            role = hub_role(raw_hub)
            if role:
                hub_counts[role] += 1
            else:
                hub_counts["Other/Ignored"] += 1

            # 2. Contar Content Types (Excluyendo EVENT)
            ct = f.get(AT_FIELDS["content_type"], "Missing Type")
            # Ignoramos mayúsculas/minúsculas para asegurar que atrape todos los eventos
            if str(ct).strip().upper() != "EVENT":
                ct_counts[ct] += 1

        # --- 2️⃣ INSIGHTS & QA SUMMARY ---
        st.subheader("2️⃣ Insights & QA")
        
        # Fila 1: Distribución por Hub
        st.markdown("##### 🎬 Distribución de Assets")
        hub_col1, hub_col2, hub_col3 = st.columns(3)
        hub_col1.metric("Total Assets", len(records))
        hub_col2.metric("📺 Programming Hub", hub_counts.get("Programming", 0))
        hub_col3.metric("📱 Mediaplanning Hub", hub_counts.get("Mediaplanning", 0))
        
        # Expander para ver el detalle de Content Types CON COLORES
        with st.expander("📊 Ver desglose por Content Type", expanded=False):
            ct_df = pd.DataFrame(ct_counts.items(), columns=["Content Type", "Count"])
            ct_df = ct_df.sort_values(by="Count", ascending=False)
            
            # Función para pintar las celdas usando la misma lógica de Excel
            def style_content_type(val):
                ct_norm = str(val).strip()
                mapped_category = ct_norm
                
                # Buscar en la configuración
                for key, value in MONTHLY_OVERVIEW_MAPPING.items():
                    if key.lower() == ct_norm.lower():
                        mapped_category = value
                        break
                
                # Obtener color base
                fill_color = OVERVIEW_FILL_COLORS.get(ct_norm)
                
                if not fill_color:
                    # Lógica para Lifts/Sneak Peeks (Verde claro)
                    if mapped_category.lower() == "scene lifts / sneak peeks / videoclips" or ct_norm.lower() in {"lift", "musical lift", "sneak peek", "videoclip"}:
                        fill_color = "FFA9D08E"
                    else:
                        fill_color = OVERVIEW_FILL_COLORS.get(mapped_category, "FFFFFFFF")
                
                # --- FIX: Convertir ARGB a RGB válido para CSS ---
                # Si el color tiene 8 caracteres (ej. FFA9D08E), ignoramos los dos primeros (FF)
                if isinstance(fill_color, str) and len(fill_color) == 8:
                    css_color = f"#{fill_color[2:]}"
                elif isinstance(fill_color, str) and len(fill_color) == 6:
                    css_color = f"#{fill_color}"
                else:
                    css_color = "#FFFFFF" # Fallback a blanco
                    
                return f"background-color: {css_color}; color: #000000; font-weight: bold;"

            # Aplicamos el estilo de pandas (Styler) a la columna
            styled_df = ct_df.style.map(style_content_type, subset=["Content Type"])
            st.dataframe(styled_df, hide_index=True, use_container_width=True)

        st.markdown("<br>", unsafe_allow_html=True)

        # Fila 2: QA Engine (Blockers & Warnings - Mock por ahora)
        st.markdown("##### 🚦 Estado de Validación (QA)")
        qa_col1, qa_col2, qa_col3 = st.columns(3)
        qa_col1.metric("🛑 Blockers", "0")
        qa_col2.metric("⚠️ Warnings", "0")
        qa_col3.metric("ℹ️ Info", "0")
        
        # PREVIEW DATA
        with st.expander("👀 Auditar datos extraídos (Data Preview)", expanded=False):
            preview_data = []
            for r in records:
                f = r.get("fields", r)
                preview_data.append({
                    "Live Date": f.get(AT_FIELDS["live_date"], "MISSING"),
                    "Content Type": f.get(AT_FIELDS["content_type"], ""),
                    "Title": f.get(AT_FIELDS["youtube_title"], f.get(AT_FIELDS["title"], "")),
                    "Owner": f.get(AT_FIELDS["owner_hub"], "")
                })
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
                            channel=channel, # <-- ESTO ES VITAL
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