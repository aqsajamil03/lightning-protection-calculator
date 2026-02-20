import streamlit as st
import pandas as pd
import datetime
import base64

st.set_page_config(page_title="Lightning Protection", page_icon="⚡")
st.title("⚡ Lightning Protection Calculator")
st.markdown("---")

# Sidebar
with st.sidebar:
    length = st.number_input("Length (m)", 70.0)
    width = st.number_input("Width (m)", 38.0)
    height = st.number_input("Height (m)", 20.0)
    lightning_density = st.number_input("Lightning Density", 1.0)
    calc_btn = st.button("Calculate", type="primary")

# Main area
if calc_btn:
    # Calculations
    collection_area = length * width + 2 * height * (length + width)
    annual_risk = collection_area * lightning_density / 1000000
    
    # Risk level
    if annual_risk < 0.001:
        risk = "LOW"
        color = "green"
    elif annual_risk < 0.01:
        risk = "MEDIUM"
        color = "orange"
    else:
        risk = "HIGH"
        color = "red"
    
    # Display
    col1, col2 = st.columns(2)
    with col1:
        st.metric("Collection Area", f"{collection_area:.0f} m²")
        st.metric("Annual Risk", f"{annual_risk:.6f}")
    with col2:
        st.markdown(f"## :{color}[{risk} RISK]")
    
    # Simple HTML Report
    report_html = f"""
    <html>
    <head><title>Lightning Report</title></head>
    <body>
        <h1>Lightning Protection Report</h1>
        <p>Date: {datetime.datetime.now().strftime('%Y-%m-%d')}</p>
        <hr>
        <h2>Building Details</h2>
        <p>Length: {length}m | Width: {width}m | Height: {height}m</p>
        <h2>Results</h2>
        <p>Collection Area: {collection_area:.0f} m²</p>
        <p>Annual Risk: {annual_risk:.6f}</p>
        <p>Risk Level: {risk}</p>
    </body>
    </html>
    """
    
    # Download button
    st.markdown("---")
    st.subheader("Download Report")
    
    b64 = base64.b64encode(report_html.encode()).decode()
    href = f'<a href="data:text/html;base64,{b64}" download="report.html">📥 Download HTML Report</a>'
    st.markdown(href, unsafe_allow_html=True)

else:
    st.info("👈 Enter values and click Calculate")

st.caption("Version 1.0")