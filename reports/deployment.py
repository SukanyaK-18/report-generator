"""Deployment Report — render function for the multi-report app."""
import sys
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

def render():
    # Import and run the deployment report logic
    # We exec the file content but skip the st.set_page_config line
    import streamlit as st
    import io
    import pandas as pd
    import numpy as np
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.dml.color import RGBColor
    from pptx.chart.data import CategoryChartData
    from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION
    from pptx.enum.text import PP_ALIGN
    from pptx.oxml.ns import qn
    from lxml import etree
    from datetime import datetime

    # Import the build function from the full file
    from reports.deployment_full import build_ppt, set_category_axis_horizontal, HEADER_BLUE, CHART_COLORS

    st.title("📊 Deployment Report")
    st.caption("Upload deployment Excel to generate a single-slide PPT with status table and chart.")

    uploaded_file = st.file_uploader("Upload Deployment Excel (.xlsx)", type=["xlsx", "xls"], key="deploy_upload")

    if uploaded_file:
        try:
            all_sheets = pd.read_excel(uploaded_file, sheet_name=None)
            df = pd.concat(all_sheets.values(), ignore_index=True)

            required = ["Deployment Type", "Current Status", "Count"]
            missing = [c for c in required if c not in df.columns]
            if missing:
                st.error(f"Missing required columns: {', '.join(missing)}")
                return

            df = df.dropna(subset=["Deployment Type", "Current Status", "Count"])
            df = df[df["Count"] != 0]
            df = df[df["Deployment Type"].astype(str).str.strip() != ""]

            pivot = pd.pivot_table(
                df, index="Deployment Type", columns="Current Status",
                values="Count", aggfunc="sum", fill_value=0
            ).reset_index()
            numeric_cols = pivot.select_dtypes(include=["number"]).columns
            pivot[numeric_cols] = pivot[numeric_cols].astype(int)
            pivot["Total Lanes"] = pivot.drop(columns=["Deployment Type"]).sum(axis=1)
            pivot = pivot[pivot["Total Lanes"] > 0]

            if "Date" in df.columns:
                df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
                start_dates = df.dropna(subset=["Date"]).groupby("Deployment Type")["Date"].min().reset_index().rename(columns={"Date": "Start Date"})
                pivot = pivot.merge(start_dates, on="Deployment Type", how="left")
            else:
                pivot["Start Date"] = ""
            pivot["End Date"] = "In Progress"

            st.success(f"Loaded **{uploaded_file.name}** — {len(pivot)} deployment types found.")
            st.dataframe(pivot.sort_values(by="Total Lanes", ascending=False), use_container_width=True)

            if st.button("🚀 Generate PPT Report", type="primary", use_container_width=True, key="deploy_gen"):
                with st.spinner("Generating..."):
                    today_str = datetime.now().strftime("%m/%d/%Y")
                    prs = build_ppt(pivot.sort_values(by="Total Lanes", ascending=False).reset_index(drop=True), today_str)
                    buffer = io.BytesIO()
                    prs.save(buffer)
                    buffer.seek(0)
                    filename = f"Deployment_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pptx"
                    st.download_button("📥 Download PPT", data=buffer, file_name=filename, mime="application/vnd.openxmlformats-officedocument.presentationml.presentation", use_container_width=True)
                    st.success(f"Report generated: **{filename}**")
        except Exception as e:
            st.error(f"Error: {e}")
