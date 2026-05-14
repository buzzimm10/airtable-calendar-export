import os
import sys
from datetime import datetime

import streamlit as st
from config.brands import get_brand_config
from services.planner_service import load_and_filter_records, export_calendar
from exporters.shared.validation import run_basic_qa

# Allow imports from src/
CURRENT_DIR = os.path.dirname(__file__)
SRC_DIR = os.path.abspath(os.path.join(CURRENT_DIR, ".."))
if SRC_DIR not in sys.path:
    sys.path.append(SRC_DIR)

from services.planner_service import load_and_filter_records


st.set_page_config(
    page_title="Disney YouTube Calendar Tool",
    page_icon="📅",
    layout="centered",
)

st.title("Disney YouTube Calendar Tool")

st.write("Generate FULL CALENDAR and MONTHLY OVERVIEW from Airtable or local JSON.")

brand = st.selectbox(
    "Brand",
    ["Disney Channel", "Disney Jr."],
    index=0,
)

brand_config = get_brand_config(brand)

template_path = st.text_input(
    "Template path",
    value=brand_config["template"],
)

output_folder = st.text_input(
    "Output folder",
    value="output",
)

year = st.selectbox("Year", ["2026", "2027", "2028"], index=0)

month_number = st.selectbox(
    "Month",
    [
        ("01", "January"),
        ("02", "February"),
        ("03", "March"),
        ("04", "April"),
        ("05", "May"),
        ("06", "June"),
        ("07", "July"),
        ("08", "August"),
        ("09", "September"),
        ("10", "October"),
        ("11", "November"),
        ("12", "December"),
    ],
    index=4,
    format_func=lambda x: x[1],
)

month = f"{year}-{month_number[0]}"



include_mediaplanning = st.checkbox(
    "Include Mediaplanning / Creative Services",
    value=True,
)

run_qa = st.checkbox("Run QA before export", value=True)

use_json = st.checkbox("Use local JSON instead of Airtable", value=False)

json_path = None
if use_json:
    json_path = st.text_input("JSON path", value="data/sample_airtable.json")

output_filename = st.text_input(
    "Output filename",
    value=f"calendar_{brand.lower().replace(' ', '_')}_{month.lower()}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
)

if st.button("Generate calendar"):
    try:
        os.makedirs(output_folder, exist_ok=True)
        output_path = os.path.join(output_folder, output_filename)

        result = load_and_filter_records(
            template_path=template_path,
            output_path=output_path,
            month=month,
            channel=brand,
            json_path=json_path,
            include_mediaplanning=include_mediaplanning,
            log_level="INFO",
            run_qa=run_qa,
        )
        records, channel_id_to_name_map = load_and_filter_records(
            month=month,
            channel=brand,
            json_path=json_path,
            log_level="INFO",
        )

        if run_qa:
            issues = run_basic_qa(records)
            if issues:
                st.warning(f"QA found {len(issues)} issues.")
                with st.expander("View QA issues"):
                    st.write(issues)

        export_calendar(
            template_path=template_path,
            output_path=output_path,
            records=records,
            channel_id_to_name_map=channel_id_to_name_map,
            brand=brand,
            include_mediaplanning=include_mediaplanning,
        )

        st.success(f"Calendar generated successfully. Records used: {len(records)}")    

        with open(output_path, "rb") as f:
            st.download_button(
                label="Download Excel",
                data=f,
                file_name=os.path.basename(output_path),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    except Exception as e:
        st.error(str(e))