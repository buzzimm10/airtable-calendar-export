import os
import sys
from datetime import datetime

import streamlit as st

# Allow imports from src/
CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
SRC_DIR = os.path.abspath(os.path.join(CURRENT_DIR, ".."))

if SRC_DIR not in sys.path:
    sys.path.insert(0, SRC_DIR)

from config.brands import get_brand_config
from services.planner_service import (
    load_and_filter_records,
    export_calendar,
)
from exporters.shared.validation import run_basic_qa


st.set_page_config(
    page_title="Disney YouTube Calendar Exporter",
    page_icon="📅",
    layout="centered",
)

st.title("Disney YouTube Calendar Exporter")
st.caption(
    "Create Excel calendars from Airtable data for Disney Channel and Disney Jr."
)

# ---------------------------------------------------
# SETUP
# ---------------------------------------------------

st.subheader("1. Calendar setup")

brand = st.selectbox(
    "Brand",
    ["Disney Channel", "Disney Jr."],
    index=0,
)

brand_config = get_brand_config(brand)

template_option = st.selectbox(
    "Template",
    ["DC", "DJR"],
    index=0 if brand == "Disney Channel" else 1,
)

template_paths = {
    "DC": "templates/X_Social Media Calendar DC YouTube EMEA - month YYYY.xltx",
    "DJR": "templates/X_Social Media Calendar DJR YouTube EMEA - month YYYY.xltx",
}

template_path = template_paths[template_option]

st.caption(f"Using template: {template_path}")

year = st.selectbox(
    "Year",
    ["2026", "2027", "2028"],
    index=0,
)

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
    index=5,
    format_func=lambda x: x[1],
)

month = f"{year}-{month_number[0]}"

# ---------------------------------------------------
# EXPORT OPTIONS
# ---------------------------------------------------

st.subheader("2. Export options")

include_mediaplanning = st.checkbox(
    "Include Mediaplanning / Creative Services",
    value=True,
)

run_qa = st.checkbox(
    "Run QA before export",
    value=True,
)

use_json = st.checkbox(
    "Use local JSON instead of Airtable",
    value=False,
)

json_path = None

if use_json:
    json_path = st.text_input(
        "JSON path",
        value="data/sample_airtable.json",
    )

# ---------------------------------------------------
# OUTPUT
# ---------------------------------------------------

st.subheader("3. Output")

output_folder = st.selectbox(
    "Output folder",
    ["output"],
    index=0,
)

with st.expander("Advanced output settings"):
    output_folder = st.text_input(
        "Custom output folder",
        value=output_folder,
    )

output_filename = st.text_input(
    "Output filename",
    value=(
        f"calendar_{brand.lower().replace(' ', '_')}_"
        f"{month.lower()}_"
        f"{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    ),
)

# ---------------------------------------------------
# GENERATE
# ---------------------------------------------------

if st.button("Generate calendar"):
    try:
        os.makedirs(output_folder, exist_ok=True)

        output_path = os.path.join(
            output_folder,
            output_filename,
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
                    for issue in issues:
                        st.write(f"- {issue}")

        export_calendar(
            template_path=template_path,
            output_path=output_path,
            records=records,
            channel_id_to_name_map=channel_id_to_name_map,
            brand=brand,
            include_mediaplanning=include_mediaplanning,
        )

        st.success(
            f"Calendar generated successfully. "
            f"Records used: {len(records)}"
        )

        with open(output_path, "rb") as f:
            st.download_button(
                label="Download Excel",
                data=f,
                file_name=os.path.basename(output_path),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    except Exception as e:
        st.error(str(e))