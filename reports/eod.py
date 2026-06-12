"""EOD Report — render function for the multi-report app."""
import sys
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

def render():
    import streamlit as st
    import io
    import pandas as pd
    import numpy as np
    from datetime import datetime

    from reports.eod_full import compute_eod_data, build_eod_ppt

    st.title("📋 EOD Weekly Report")
    st.caption("Upload EOD Excel to generate a 2-slide PPT report.")

    uploaded_file = st.file_uploader("Upload EOD Excel (.xlsx)", type=["xlsx", "xls"], key="eod_upload")

    if uploaded_file:
        try:
            all_sheets = pd.read_excel(uploaded_file, sheet_name=None)
            df = None
            for name, sheet_df in all_sheets.items():
                if "Host" in sheet_df.columns:
                    df = sheet_df
                    break
            if df is None:
                df = list(all_sheets.values())[0]
            if "Host" not in df.columns:
                st.error("Could not find 'Host' column.")
                return

            df = df.dropna(subset=["Host"]).reset_index(drop=True)
            st.success(f"Loaded **{uploaded_file.name}** — {len(df)} rows found.")

            summary, unable_fix_breakdown, host_counts, rep_chart, new_incidents = compute_eod_data(df)

            col1, col2 = st.columns(2)
            with col1:
                st.subheader("EOD Summary")
                st.dataframe(pd.DataFrame([summary]), use_container_width=True)
            with col2:
                rep_filtered = host_counts[host_counts["Count"] > 1] if not host_counts.empty else pd.DataFrame()
                st.subheader(f"Repetition (Count > 1): {len(rep_filtered)} lanes")
                if not rep_filtered.empty:
                    st.dataframe(rep_filtered, use_container_width=True)

            if not new_incidents.empty:
                st.subheader(f"New Incidents: {len(new_incidents)}")
                st.dataframe(new_incidents, use_container_width=True)

            if st.button("🚀 Generate EOD Report", type="primary", use_container_width=True, key="eod_gen"):
                with st.spinner("Generating..."):
                    today_str = datetime.now().strftime("%m/%d/%Y")
                    prs = build_eod_ppt(summary, unable_fix_breakdown, host_counts, rep_chart, new_incidents, today_str)
                    buffer = io.BytesIO()
                    prs.save(buffer)
                    buffer.seek(0)
                    filename = f"EOD_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pptx"
                    st.download_button("📥 Download EOD Report", data=buffer, file_name=filename, mime="application/vnd.openxmlformats-officedocument.presentationml.presentation", use_container_width=True)
                    st.success(f"Report generated: **{filename}**")
        except Exception as e:
            st.error(f"Error: {e}")
            import traceback
            st.code(traceback.format_exc())
