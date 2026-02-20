import streamlit as st
import math
import datetime

st.set_page_config(page_title="Lightning Protection", page_icon="⚡")
st.title("⚡ Lightning Protection Calculator")
st.markdown("---")

# Sidebar inputs
with st.sidebar:
    st.header("Building Parameters")
    length = st.number_input("Length (m)", 1.0, 200.0, 70.0)
    width = st.number_input("Width (m)", 1.0, 200.0, 38.0)
    height = st.number_input("Height (m)", 1.0, 100.0, 20.0)
    
    st.header("Location")
    lightning_density = st.number_input("Lightning flashes/km²/year", 0.1, 30.0, 1.0)
    
    calculate = st.button("Calculate", type="primary")

# Main calculation
if calculate:
    # Simple calculations
    collection_area = length * width + 2 * height * (length + width)
    annual_risk = collection_area * lightning_density / 1000000
    
    # Risk level
    if annual_risk < 0.001:
        risk_level = "LOW"
        color = "green"
        air_terminals = 2
    elif annual_risk < 0.01:
        risk_level = "MEDIUM"
        color = "orange"
        air_terminals = 4
    else:
        risk_level = "HIGH"
        color = "red"
        air_terminals = 6
    
    # Display results
    col1, col2 = st.columns(2)
    with col1:
        st.metric("Collection Area", f"{collection_area:.0f} m²")
        st.metric("Annual Risk", f"{annual_risk:.6f}")
    
    with col2:
        st.markdown(f"### :{color}[{risk_level} Risk]")
        st.metric("Air Terminals Required", air_terminals)
    
    st.info("PDF report feature coming soon!")

else:
    st.info("👈 Enter values in sidebar and click Calculate")

st.markdown("---")
st.caption(f"Version 1.0 | {datetime.datetime.now().strftime('%Y-%m-%d')}")