"""
Multi-Report Generator — Streamlit Web App
Combines Deployment, EOD, and POS Readiness reports into one tool.
"""

import streamlit as st

st.set_page_config(page_title="Report Generator", page_icon="📊", layout="wide")

# Sidebar navigation
st.sidebar.title("📊 Report Generator")
report_type = st.sidebar.radio(
    "Select Report Type",
    ["Deployment", "EOD", "POS - REG Readiness", "POS - SCO Readiness"],
    index=0
)

st.sidebar.divider()
st.sidebar.caption("Upload Excel → Generate PPT → Download")

# Route to the selected report
if report_type == "Deployment":
    from reports import deployment
    deployment.render()
elif report_type == "EOD":
    from reports import eod
    eod.render()
elif report_type == "POS - REG Readiness":
    from reports import pos
    pos.render("REG Readiness")
elif report_type == "POS - SCO Readiness":
    from reports import pos
    pos.render("SCO Readiness")
