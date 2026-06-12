"""POS Readiness Report — render function for the multi-report app."""
import sys
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

def render(report_type="REG Readiness"):
    import streamlit as st
    import io
    import pandas as pd
    import numpy as np
    from datetime import datetime

    from reports.pos_full import compute_pos_data, build_pos_ppt

    st.subheader(f"🖥️ POS - {report_type} Report")
    st.caption(f"Upload the POS-{report_type.split()[0]} Excel to generate a 2-slide PPT report.")

    uploaded_file = st.file_uploader(f"Upload POS-{report_type.split()[0]} Excel (.xlsx)", type=["xlsx", "xls"], key=f"pos_{report_type}_upload")

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

            summary, unable_fix_chart, host_counts, rep_chart = compute_pos_data(df, report_type)

            col1, col2 = st.columns(2)
            with col1:
                st.subheader("Summary")
                st.dataframe(pd.DataFrame([summary]), use_container_width=True)
            with col2:
                if not unable_fix_chart.empty:
                    st.subheader("Unable to Fix Breakdown")
                    st.dataframe(unable_fix_chart, use_container_width=True)

            rep_filtered = host_counts[host_counts["Count"] > 1] if not host_counts.empty else pd.DataFrame()
            if not rep_filtered.empty:
                st.subheader(f"Repetition (Count > 1): {len(rep_filtered)} lanes")
                st.dataframe(rep_filtered, use_container_width=True)

            if st.button("🚀 Generate POS Report", type="primary", use_container_width=True, key=f"pos_{report_type}_gen"):
                with st.spinner("Generating..."):
                    today_str = datetime.now().strftime("%m/%d/%Y")
                    prs = build_pos_ppt(summary, unable_fix_chart, host_counts, rep_chart, report_type, today_str)
                    buffer = io.BytesIO()
                    prs.save(buffer)
                    buffer.seek(0)
                    filename = f"POS_{report_type.replace(' ', '_')}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pptx"
                    st.download_button("📥 Download POS Report", data=buffer, file_name=filename, mime="application/vnd.openxmlformats-officedocument.presentationml.presentation", use_container_width=True)
                    st.success(f"Report generated: **{filename}**")
        except Exception as e:
            st.error(f"Error: {e}")
            import traceback
            st.code(traceback.format_exc())
