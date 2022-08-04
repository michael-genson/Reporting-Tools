from datetime import datetime, timedelta

import streamlit as st

from scripts.calculate_first_call_resolution import (
    main as calculate_first_call_resolution,
)
from scripts.utils import format_percent

st.title("Reporting Tools")
st.header("First Call Resolution Calculator")

calculations: dict[str, float | int] = {}
with st.form("fcr_calculator"):
    fcr_reopened_file = st.file_uploader(
        "1. FCR Re-opened Report", type=["xlsx", "xls"]
    )
    fcr_closed_file = st.file_uploader("2. FCR Closed Report", type=["xlsx", "xls"])
    fcr_parent_file = st.file_uploader(
        "3. FCR Parent Cases Report", type=["xlsx", "xls"]
    )

    st.markdown("---")

    report_date = st.date_input(
        "Report Date", value=datetime.today() - timedelta(days=15)
    )
    child_case_threshold = st.number_input("Child Case Threshold", value=4, min_value=0)
    fcr_submitted = st.form_submit_button("Calculate First Call Resolution")

    if fcr_submitted:
        if not all([fcr_reopened_file, fcr_closed_file, fcr_parent_file]):
            st.markdown(
                '<span style="color:red">**All three files are required**</span>',
                unsafe_allow_html=True,
            )

        else:
            try:
                calculations = calculate_first_call_resolution(
                    report_date,
                    fcr_reopened_file,
                    fcr_closed_file,
                    fcr_parent_file,
                    child_case_threshold,
                )

            except ValueError as e:
                st.markdown(
                    f'<span style="color:red">**Error: _{e}_**</span>',
                    unsafe_allow_html=True,
                )

if calculations:
    col1, col2, col3 = st.columns(3)
    col2.metric("First Call Resolution", format_percent(calculations["fcr"]))

    with st.expander("See calculated data"):
        subcol1, subcol2, subcol3 = st.columns(3)
        subcol1.metric("Closed Case Count", calculations["closed_case_count"])
        subcol2.metric("Escalated Case Count", calculations["escalated_case_count"])
        subcol3.metric("Child Case Count", calculations["child_case_count"])
