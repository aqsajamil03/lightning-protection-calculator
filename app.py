import streamlit as st
import datetime

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

if calc_btn:
    # Simple calculations
    collection_area = length * width + 2 * height * (length + width)
    annual_risk = collection_area * lightning_density / 1000000
    
    if annual_risk < 0.001:
        risk = "LOW"
    elif annual_risk < 0.01:
        risk = "MEDIUM"
    else:
        risk = "HIGH"
    
    # Show results
    st.metric("Collection Area", f"{collection_area:.0f} m²")
    st.metric("Annual Risk", f"{annual_risk:.6f}")
    st.metric("Risk Level", risk)
    
    # Create report text
    report = f"""
    LIGHTNING PROTECTION REPORT
    ===========================
    Date: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M')}
    
    BUILDING DETAILS:
    Length: {length} m
    Width: {width} m
    Height: {height} m
    Lightning Density: {lightning_density} flashes/km²/year
    
    RESULTS:
    Collection Area: {collection_area:.0f} m²
    Annual Risk: {annual_risk:.6f}
    Risk Level: {risk}
    """
    
    # Download button
    st.download_button(
        label="📥 Download Report (TXT)",
        data=report,
        file_name=f"lightning_report_{datetime.datetime.now().strftime('%Y%m%d')}.txt",
        mime="text/plain"
    )

else:
    st.info("👈 Enter values and click Calculate")

st.markdown("---")
st.caption("Simple Version")