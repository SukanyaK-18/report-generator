"""
Multi-Report Generator — Streamlit Web App
Combines Deployment, EOD, and POS Readiness reports into one tool.
"""

import streamlit as st

st.set_page_config(page_title="Report Generator", page_icon="📊", layout="wide")

# Custom CSS
st.markdown("""
<style>
@keyframes fadeIn {
    from { opacity: 0; transform: translateY(-5px); }
    to { opacity: 1; transform: translateY(0); }
}
.app-header {
    display: flex;
    align-items: center;
    gap: 14px;
    padding: 16px 0 8px 0;
    animation: fadeIn 0.4s ease-out;
}
.app-icon {
    font-size: 2.5rem;
}
.app-title {
    font-size: 1.8rem;
    font-weight: 700;
    background: linear-gradient(135deg, #4472C4, #ED7D31);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    margin: 0;
}
.app-subtitle {
    font-size: 0.85rem;
    color: #888;
    margin-top: 2px;
}
section[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #1a1a2e 0%, #16213e 100%);
}
section[data-testid="stSidebar"] .stRadio label {
    font-size: 0.95rem;
}
div[data-testid="stDownloadButton"] > button {
    background: linear-gradient(135deg, #4472C4, #1ABC9C);
    color: white;
    border: none;
    border-radius: 8px;
    font-weight: 600;
}
.stButton > button[kind="primary"] {
    border-radius: 8px;
}
</style>
""", unsafe_allow_html=True)

# Header
st.markdown("""
<div class="app-header">
    <div class="app-icon">📊</div>
    <div>
        <p class="app-title">Report Generator</p>
        <p class="app-subtitle">Upload Excel → Select Report Type → Generate PPT</p>
    </div>
</div>
""", unsafe_allow_html=True)

# Sidebar navigation
st.sidebar.title("📊 Report Type")
report_type = st.sidebar.radio(
    "Select",
    ["Deployment", "EOD", "POS - REG Readiness", "POS - SCO Readiness"],
    index=0,
    label_visibility="collapsed"
)

st.sidebar.divider()
st.sidebar.markdown("**How to use:**")
st.sidebar.markdown("1. Select report type above")
st.sidebar.markdown("2. Upload the Excel file")
st.sidebar.markdown("3. Click Generate")
st.sidebar.markdown("4. Download the PPT")

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
