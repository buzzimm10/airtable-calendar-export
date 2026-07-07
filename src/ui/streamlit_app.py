import os
import sys
from datetime import datetime
from collections import Counter
import pandas as pd
import streamlit as st

# 🚨 1. PRIMERO le enseñamos a Python dónde buscar 🚨
CURRENT_DIR = os.path.dirname(__file__)
SRC_DIR = os.path.abspath(os.path.join(CURRENT_DIR, ".."))
if SRC_DIR not in sys.path:
    sys.path.append(SRC_DIR)

# 🚨 2. LUEGO importamos nuestras carpetas locales 🚨
from config.countries import SHEET_COUNTRIES as dc_countries
from config.countries_djr import DJR_SHEET_COUNTRIES as djr_countries
from services.planner_service import load_and_filter_records, hub_role
from services.qa_service import run_qa_on_records
from config.fields import AT_FIELDS
from config.styles import OVERVIEW_FILL_COLORS
from config.content_types import MONTHLY_OVERVIEW_MAPPING
from populate_calendar import populate_calendar

# --- PAGE CONFIG ---
# ... tu código sigue normal a partir de aquí ...

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
if "qa_results" not in st.session_state:
    st.session_state.qa_results = None

# --- HEADER PRINCIPAL ---
st.title("✨ Disney YouTube Calendar Tool")
st.markdown("Genera el **Full Calendar** y **Monthly Overview** en un par de clics.")
st.divider()

# --- SIDEBAR (CONTROLES) ---
with st.sidebar:
    st.header("1️⃣ Configuración")
    channel = st.selectbox("📺 Brand", ["Disney Channel", "Disney Junior"], index=0)
    
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
        
        template_path = st.text_input("Template path", value=default_template_path)
        output_folder = st.text_input("Output folder", value="output")

    st.divider()

    # ---> NUEVO: Selector visual de Países Fase 2 <---
    st.markdown("### 🌍 Calendarios Locales (Fase 2)")
    country_options = dc_countries if channel == "Disney Channel" else djr_countries
    country_options = [c for c in country_options if c.lower() not in ["spain", "españa"]]
    
    selected_countries = st.multiselect(
        "Selecciona para qué países generar calendarios limpios individuales:",
        options=country_options,
        default=country_options,
        help="Los archivos se guardarán automáticamente en la carpeta 'country_calendars'."
    )
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
                
                # --- NUEVO: Ejecutar motor de QA inmediatamente ---
                st.session_state.qa_results = run_qa_on_records(records)
                # -------------------------------------------------
                
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
        # Recuperamos los resultados reales de QA
        qa = st.session_state.get("qa_results", {"blockers": [], "warnings": [], "info": []})
        num_blockers = len(qa["blockers"])
        num_warnings = len(qa["warnings"])
        num_info = len(qa["info"])

        # --- PROCESAMIENTO DE MÉTRICAS (Filtrando EVENT radicalmente) ---
        hub_counts = Counter()
        ct_counts = Counter()
        total_content_assets = 0  # Contador exclusivo de contenidos reales

        for r in records:
            f = r.get("fields", r)
            ct = f.get(AT_FIELDS["content_type"], "")
            
            # SI ES UN EVENTO, SE IGNORA PARA LAS TARJETAS Y DESGLOSE VISUAL
            if str(ct).strip().upper() == "EVENT":
                continue
                
            total_content_assets += 1
            
            raw_hub = f.get(AT_FIELDS["owner_hub"])
            role = hub_role(raw_hub)
            if role:
                hub_counts[role] += 1
            else:
                hub_counts["Other/Ignored"] += 1

            if ct:
                ct_counts[ct] += 1

        # --- 2️⃣ INSIGHTS & QA SUMMARY ---
        st.subheader("2️⃣ Insights & QA")
        
        # Fila 1: Distribución por Hub (Muestra datos reales sin EVENT)
        st.markdown("##### 🎬 Distribución de Assets")
        hub_col1, hub_col2, hub_col3 = st.columns(3)
        hub_col1.metric("Total Assets", total_content_assets)
        hub_col2.metric("📺 Programming Hub", hub_counts.get("Programming", 0))
        hub_col3.metric("📱 Mediaplanning Hub", hub_counts.get("Mediaplanning", 0))
        
        with st.expander("📊 Ver desglose por Content Type", expanded=False):
            ct_df = pd.DataFrame(ct_counts.items(), columns=["Content Type", "Count"])
            ct_df = ct_df.sort_values(by="Count", ascending=False)
            
            # FIX: Función de pintado Case-Insensitive
            def style_content_type(val):
                ct_norm = str(val).strip()
                mapped_category = ct_norm
                
                # 1. Resolver alias de forma insensible a mayúsculas
                for key, value in MONTHLY_OVERVIEW_MAPPING.items():
                    if key.lower() == ct_norm.lower():
                        mapped_category = value
                        break
                
                # 2. Búsqueda insensible a mayúsculas en OVERVIEW_FILL_COLORS
                fill_color = None
                for key, color in OVERVIEW_FILL_COLORS.items():
                    if key.lower() == ct_norm.lower():
                        fill_color = color
                        break
                
                if not fill_color:
                    for key, color in OVERVIEW_FILL_COLORS.items():
                        if key.lower() == mapped_category.lower():
                            fill_color = color
                            break
                
                # 3. Fallback histórico para celdas verdes (Lifts/Sneak peeks)
                if not fill_color:
                    if mapped_category.lower() == "scene lifts / sneak peeks / videoclips" or ct_norm.lower() in {"lift", "musical lift", "sneak peek", "videoclip"}:
                        fill_color = "FFA9D08E"
                    else:
                        fill_color = "FFFFFFFF"
                
                # Convertir ARGB/RGB a CSS hex estándar
                if isinstance(fill_color, str) and len(fill_color) == 8:
                    css_color = f"#{fill_color[2:]}"
                elif isinstance(fill_color, str) and len(fill_color) == 6:
                    css_color = f"#{fill_color}"
                else:
                    css_color = "#FFFFFF"
                    
                return f"background-color: {css_color}; color: #000000; font-weight: bold;"

            styled_df = ct_df.style.map(style_content_type, subset=["Content Type"])
            st.dataframe(styled_df, hide_index=True, use_container_width=True)

        st.markdown("<br>", unsafe_allow_html=True)

        # Fila 2: QA Engine (DATOS REALES)
        st.markdown("##### 🚦 Estado de Validación (QA)")
        qa_col1, qa_col2, qa_col3 = st.columns(3)
        
        # Usamos delta_color para resaltar si hay blockers
        qa_col1.metric("🛑 Blockers", num_blockers, delta=f"-{num_blockers}" if num_blockers > 0 else None, delta_color="inverse")
        qa_col2.metric("⚠️ Warnings", num_warnings)
        qa_col3.metric("ℹ️ Info", num_info)
        
        # DESGLOSE DE ALERTAS DE QA (Solo se muestran los expanders si hay mensajes)
        if num_blockers > 0:
            with st.container(): # <-- CORREGIDO AQUÍ
                st.markdown("###### ❌ Corrige los siguientes Blockers para poder exportar:")
                for b in qa["blockers"]:
                    st.error(b)

        if num_warnings > 0:
            with st.expander(f"⚠️ Ver {num_warnings} Warnings", expanded=False):
                for w in qa["warnings"]:
                    st.warning(w)

        if num_info > 0:
            with st.expander(f"ℹ️ Ver {num_info} Notas de Información", expanded=False):
                for i in qa["info"]:
                    st.info(i)

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
        
        btn_col1, btn_col2, _ = st.columns([1, 1, 2])
        
        with btn_col1:
            # INTERCEPTOR CRÍTICO: Si num_blockers > 0, el botón se deshabilita automáticamente
            button_disabled = num_blockers > 0
            
            if button_disabled:
                st.markdown("⚠️ **Exportación bloqueada.** Revisa las alertas rojas de arriba.")
            
            if st.button("🚀 Generar Excel", type="primary", use_container_width=True, disabled=button_disabled):
                with st.spinner("Escribiendo Excel... No cierres esta ventana."):
                    try:
                        os.makedirs(output_folder, exist_ok=True)
                        populate_calendar(
                            records=records,
                            template_path=template_path,
                            output_path=output_path,
                            channel=channel,
                            enable_mediaplanning=include_mediaplanning,
                            channel_id_to_name_map=st.session_state.channel_map,
                            selected_countries=selected_countries # <-- NUEVO: Pasamos la selección a tu backend
                        )
                        st.session_state.export_success = True
                        st.session_state.output_path = output_path
                        st.session_state.output_filename = output_filename
                    except Exception as e:
                        st.error(f"Error generando Excel: {str(e)}")
        
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