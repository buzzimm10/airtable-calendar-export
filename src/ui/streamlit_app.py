import os
import sys
from datetime import datetime

import streamlit as st

# Allow imports from src/
CURRENT_DIR = os.path.dirname(__file__)
SRC_DIR = os.path.abspath(os.path.join(CURRENT_DIR, ".."))
if SRC_DIR not in sys.path:
    sys.path.append(SRC_DIR)

from services.planner_service import run_planner


st.set_page_config(
    page_title="Disney YouTube Calendar Tool",
    page_icon="📅",
    layout="centered",
)

st.title("Disney YouTube Calendar Tool")

st.write("Generate FULL CALENDAR and MONTHLY OVERVIEW from Airtable or local JSON.")

template_path = st.text_input(
    "Template path",
    value="templates/X_Social Media Calendar DC YouTube EMEA - month 2026.xltx",
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

channel = st.selectbox(
    "Channel",
    ["Disney Channel", "Disney Jr."],
    index=0,
)

include_mediaplanning = st.checkbox(
    "Include Mediaplanning / Creative Services",
    value=True,
)

use_json = st.checkbox("Use local JSON instead of Airtable", value=False)

json_path = None
if use_json:
    json_path = st.text_input("JSON path", value="data/sample_airtable.json")

output_filename = st.text_input(
    "Output filename",
    value=f"calendar_{channel.lower().replace(' ', '_')}_{month.lower()}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
)

if st.button("Generate calendar"):
    try:
        os.makedirs(output_folder, exist_ok=True)
        output_path = os.path.join(output_folder, output_filename)

        result = load_and_filter_records(
            template_path=template_path,
            output_path=output_path,
            month=month,
            channel=channel,
            json_path=json_path,
            include_mediaplanning=include_mediaplanning,
            log_level="INFO",
        )

        st.success(f"Calendar generated successfully. Records used: {result['records_count']}")

        with open(output_path, "rb") as f:
            st.download_button(
                label="Download Excel",
                data=f,
                file_name=os.path.basename(output_path),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    except Exception as e:
        st.error(str(e))