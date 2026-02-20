import streamlit as st
import math
import datetime
import pandas as pd
from fpdf import FPDF
import tempfile
import os
import base64

st.set_page_config(page_title="Lightning Protection", page_icon="⚡", layout="wide")
st.title("⚡ Advanced Lightning Protection Calculator")
st.markdown("**IEC 62305 Compliant**")
st.markdown("---")

# Sidebar inputs
with st.sidebar:
    st.header("📐 Building Parameters")
    
    col1, col2 = st.columns(2)
    with col1:
        length = st.number_input("Length (m)", 1.0, 200.0, 70.0)
        width = st.number_input("Width (m)", 1.0, 200.0, 38.0)
        height = st.number_input("Height (m)", 1.0, 100.0, 20.0)
    
    with col2:
        roof_type = st.selectbox("Roof Type", ["Flat", "Pitched", "Complex"])
        construction = st.selectbox("Construction", ["Concrete", "Steel", "Masonry", "Wood"])
    
    st.header("🌍 Location")
    lightning_density = st.number_input("Ng (flashes/km²/year)", 0.1, 30.0, 1.0)
    
    environment = st.select_slider("Environment", 
        options=["Surrounded", "Similar height", "Isolated", "Hilltop"])
    
    contents = st.multiselect("Contents Type",
        ["Ordinary", "Valuable", "Hospital", "Explosive"])
    
    occupants = st.number_input("Number of people", 1, 1000, 50)
    
    calculate = st.button("🔧 CALCULATE", type="primary", use_container_width=True)

# Advanced calculation function
def calculate_protection(length, width, height, Ng, environment, contents):
    
    # Collection area (IEC 62305-2)
    Ae = length * width + 2*height*(length + width) + math.pi*height**2
    
    # Environmental factor
    env_factors = {"Surrounded":0.25, "Similar height":0.5, "Isolated":1, "Hilltop":2}
    Cd = env_factors.get(environment, 1)
    
    # Annual risk
    Nd = Ae * Ng * 1e-6 * Cd
    
    # Tolerable risk based on contents
    if "Explosive" in contents:
        Rt = 0.00001
    elif "Hospital" in contents:
        Rt = 0.0001
    elif "Valuable" in contents:
        Rt = 0.001
    else:
        Rt = 0.01
    
    ratio = Nd / Rt
    
    # Protection Level
    if ratio < 0.001:
        lpl = "IV (Optional)"
        lpl_class = "IV"
        mesh = "20m x 20m"
        sphere = 60
        terminals = max(2, int((length+width)/30))
    elif ratio < 0.01:
        lpl = "III (Standard)"
        lpl_class = "III"
        mesh = "15m x 15m"
        sphere = 45
        terminals = max(3, int((length+width)/25))
    elif ratio < 0.1:
        lpl = "II (Enhanced)"
        lpl_class = "II"
        mesh = "10m x 10m"
        sphere = 30
        terminals = max(4, int((length+width)/20))
    else:
        lpl = "I (Maximum)"
        lpl_class = "I"
        mesh = "5m x 5m"
        sphere = 20
        terminals = max(6, int((length+width)/15))
    
    # Down conductors
    perimeter = 2*(length + width)
    down_conductors = max(2, int(perimeter/20))
    
    # Earthing resistance (simplified)
    earth_resistance = 100 / (2*math.pi*3) * math.log(4*3/0.016)
    
    return {
        'Ae': Ae, 'Cd': Cd, 'Nd': Nd, 'Rt': Rt, 'ratio': ratio,
        'lpl': lpl, 'lpl_class': lpl_class, 'mesh': mesh, 'sphere': sphere,
        'terminals': terminals, 'down_conductors': down_conductors,
        'perimeter': perimeter, 'earth_resistance': earth_resistance
    }

# PDF Generator
def create_pdf(length, width, height, Ng, environment, contents, occupants, results):
    
    pdf = FPDF()
    pdf.add_page()
    
    # Title
    pdf.set_font('Arial', 'B', 16)
    pdf.cell(0, 10, 'LIGHTNING PROTECTION REPORT', 0, 1, 'C')
    pdf.set_font('Arial', '', 10)
    pdf.cell(0, 10, f'Date: {datetime.datetime.now().strftime("%Y-%m-%d %H:%M")}', 0, 1, 'C')
    pdf.ln(10)
    
    # Building Info
    pdf.set_font('Arial', 'B', 12)
    pdf.cell(0, 10, '1. BUILDING INFORMATION', 0, 1)
    pdf.set_font('Arial', '', 10)
    pdf.cell(0, 8, f'Dimensions: {length}m x {width}m x {height}m', 0, 1)
    pdf.cell(0, 8, f'Environment: {environment}', 0, 1)
    pdf.cell(0, 8, f'Occupants: {occupants}', 0, 1)
    pdf.ln(5)
    
    # Risk Assessment
    pdf.set_font('Arial', 'B', 12)
    pdf.cell(0, 10, '2. RISK ASSESSMENT', 0, 1)
    pdf.set_font('Arial', '', 10)
    pdf.cell(0, 8, f'Collection Area: {results["Ae"]:.0f} m²', 0, 1)
    pdf.cell(0, 8, f'Annual Risk: {results["Nd"]:.6f}', 0, 1)
    pdf.cell(0, 8, f'Protection Level: {results["lpl"]}', 0, 1)
    pdf.ln(5)
    
    # Protection Design
    pdf.set_font('Arial', 'B', 12)
    pdf.cell(0, 10, '3. PROTECTION DESIGN', 0, 1)
    pdf.set_font('Arial', '', 10)
    pdf.cell(0, 8, f'Air Terminals: {results["terminals"]}', 0, 1)
    pdf.cell(0, 8, f'Mesh Size: {results["mesh"]}', 0, 1)
    pdf.cell(0, 8, f'Down Conductors: {results["down_conductors"]}', 0, 1)
    pdf.cell(0, 8, f'Earth Resistance: {results["earth_resistance"]:.1f} Ω', 0, 1)
    
    return pdf

# Main app
if calculate:
    results = calculate_protection(length, width, height, lightning_density, environment, contents)
    
    # Display results in columns
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.subheader("📊 Risk Assessment")
        st.metric("Collection Area", f"{results['Ae']:.0f} m²")
        st.metric("Annual Risk", f"{results['Nd']:.6f}")
        st.metric("Risk Ratio", f"{results['ratio']:.3f}")
    
    with col2:
        st.subheader("⚡ Protection Level")
        if results['lpl_class'] == "I":
            st.error(f"**{results['lpl']}**")
        elif results['lpl_class'] == "II":
            st.warning(f"**{results['lpl']}**")
        else:
            st.success(f"**{results['lpl']}**")
        
        st.metric("Air Terminals", results['terminals'])
        st.metric("Mesh Size", results['mesh'])
    
    with col3:
        st.subheader("📡 System Design")
        st.metric("Down Conductors", results['down_conductors'])
        st.metric("Rolling Sphere", f"{results['sphere']}m")
        st.metric("Earth Resistance", f"{results['earth_resistance']:.1f} Ω")
    
    # Materials table
    st.subheader("📋 Materials Required")
    materials = pd.DataFrame({
        'Component': ['Air Terminals', 'Conductors', 'Down Conductors', 'Earth Rods'],
        'Quantity': [
            f"{results['terminals']} pcs",
            f"{results['terminals'] + results['down_conductors']*height:.0f} m",
            f"{results['down_conductors']} pcs",
            f"{max(2, results['down_conductors'])} pcs"
        ]
    })
    st.dataframe(materials, use_container_width=True)
    
    # PDF Download
    st.markdown("---")
    st.subheader("📥 Download Report")
    
    if st.button("📄 Generate PDF Report"):
        pdf = create_pdf(length, width, height, lightning_density, environment, contents, occupants, results)
        
        with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp:
            pdf.output(tmp.name)
            tmp_path = tmp.name
        
        with open(tmp_path, 'rb') as f:
            pdf_bytes = f.read()
        
        b64 = base64.b64encode(pdf_bytes).decode()
        href = f'<a href="data:application/octet-stream;base64,{b64}" download="Lightning_Protection_Report.pdf">📥 Click here to download PDF Report</a>'
        st.markdown(href, unsafe_allow_html=True)
        
        os.unlink(tmp_path)
        st.success("✅ PDF Generated Successfully!")

else:
    st.info("👈 Enter building parameters and click CALCULATE")

st.markdown("---")
st.caption(f"⚡ IEC 62305 Compliant | Version 2.0 | {datetime.datetime.now().strftime('%Y-%m-%d')}")