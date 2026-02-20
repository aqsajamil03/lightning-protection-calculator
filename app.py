import streamlit as st
import math
import datetime
import pandas as pd
from fpdf import FPDF
import tempfile
import os
import base64

# Page setup
st.set_page_config(
    page_title="Lightning Protection Design",
    page_icon="⚡",
    layout="wide"
)

st.title("⚡ Lightning Protection Design System")
st.markdown("**IEC 62305 Compliant Calculator**")
st.markdown("---")

# Sidebar inputs
with st.sidebar:
    st.header("📐 Building Parameters")
    
    length = st.number_input("Length (m)", 1.0, 200.0, 70.0)
    width = st.number_input("Width (m)", 1.0, 200.0, 38.0)
    height = st.number_input("Height (m)", 1.0, 100.0, 20.0)
    
    roof_type = st.selectbox("Roof Type", ["Flat", "Pitched", "Complex"])
    construction = st.selectbox("Construction", ["Concrete", "Steel", "Masonry", "Wood"])
    
    st.header("🌍 Location")
    lightning_density = st.number_input("Lightning flashes/km²/year", 0.1, 30.0, 1.0)
    
    environment = st.selectbox("Environment", ["Surrounded", "Similar height", "Isolated", "Hilltop"])
    
    contents = st.multiselect("Contents", ["Ordinary", "Valuable", "Hospital", "Explosive"])
    
    occupants = st.number_input("People", 1, 1000, 50)
    
    calculate = st.button("🔧 CALCULATE", type="primary", use_container_width=True)

# Simple calculation function
def calculate(length, width, height, Ng, environment, contents):
    
    # Collection area
    Ae = length * width + 2*height*(length + width) + math.pi*height**2
    
    # Environment factor
    env_factors = {"Surrounded":0.25, "Similar height":0.5, "Isolated":1, "Hilltop":2}
    Cd = env_factors.get(environment, 1)
    
    Nd = Ae * Ng * 1e-6 * Cd
    
    # Risk tolerance
    if "Explosive" in contents:
        Rt = 0.00001
    elif "Hospital" in contents:
        Rt = 0.0001
    elif "Valuable" in contents:
        Rt = 0.001
    else:
        Rt = 0.01
    
    ratio = Nd / Rt
    
    # Protection level
    if ratio < 0.001:
        lpl = "IV (Optional)"
        mesh = "20m x 20m"
        sphere = 60
        terminals = max(2, int((length+width)/30))
    elif ratio < 0.01:
        lpl = "III (Standard)"
        mesh = "15m x 15m"
        sphere = 45
        terminals = max(3, int((length+width)/25))
    elif ratio < 0.1:
        lpl = "II (Enhanced)"
        mesh = "10m x 10m"
        sphere = 30
        terminals = max(4, int((length+width)/20))
    else:
        lpl = "I (Maximum)"
        mesh = "5m x 5m"
        sphere = 20
        terminals = max(6, int((length+width)/15))
    
    down_conductors = max(2, int(2*(length+width)/25))
    
    return {
        'Ae': Ae,
        'Cd': Cd,
        'Nd': Nd,
        'Rt': Rt,
        'ratio': ratio,
        'lpl': lpl,
        'mesh': mesh,
        'sphere': sphere,
        'terminals': terminals,
        'down_conductors': down_conductors,
        'perimeter': 2*(length+width)
    }

# PDF Generator
def create_pdf(params, length, width, height, roof_type, construction, occupants, contents, results):
    
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
    pdf.cell(0, 8, f'Roof: {roof_type}, Construction: {construction}', 0, 1)
    pdf.cell(0, 8, f'Occupants: {occupants}, Contents: {", ".join(contents)}', 0, 1)
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
    pdf.cell(0, 8, f'Rolling Sphere: {results["sphere"]}m', 0, 1)
    pdf.ln(5)
    
    # Materials
    pdf.set_font('Arial', 'B', 12)
    pdf.cell(0, 10, '4. MATERIALS REQUIRED', 0, 1)
    pdf.set_font('Arial', '', 10)
    pdf.cell(0, 8, f'- Air Terminals: {results["terminals"]} pcs', 0, 1)
    pdf.cell(0, 8, f'- Conductors: {results["terminals"] + results["down_conductors"]*height:.0f} m', 0, 1)
    pdf.cell(0, 8, f'- Earth Rods: {max(2, results["down_conductors"])} pcs', 0, 1)
    pdf.cell(0, 8, f'- Ring Conductor: {2*(length+width):.0f} m', 0, 1)
    
    return pdf

# Main area
if calculate:
    results = calculate(length, width, height, lightning_density, environment, contents)
    
    # Show results in columns
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.subheader("📊 Risk Assessment")
        st.metric("Collection Area", f"{results['Ae']:.0f} m²")
        st.metric("Annual Risk", f"{results['Nd']:.6f}")
        st.metric("Risk Ratio", f"{results['ratio']:.3f}")
    
    with col2:
        st.subheader("⚡ Protection Level")
        st.info(f"**{results['lpl']}**")
        st.metric("Air Terminals", results['terminals'])
        st.metric("Mesh Size", results['mesh'])
    
    with col3:
        st.subheader("📡 System Design")
        st.metric("Down Conductors", results['down_conductors'])
        st.metric("Rolling Sphere", f"{results['sphere']}m")
    
    # Materials table
    st.subheader("📋 Materials List")
    materials_df = pd.DataFrame({
        'Component': ['Air Terminals', 'Conductors', 'Down Conductors', 'Earth Rods', 'Ring Conductor'],
        'Quantity': [
            f"{results['terminals']} pcs",
            f"{results['terminals'] + results['down_conductors']*height:.0f} m",
            f"{results['down_conductors']} pcs",
            f"{max(2, results['down_conductors'])} pcs",
            f"{2*(length+width):.0f} m"
        ]
    })
    st.dataframe(materials_df, use_container_width=True)
    
    # Download PDF section - SEPARATE FROM BUTTON
    st.markdown("---")
    st.subheader("📥 Download Report")
    
    if st.button("📄 Generate PDF"):
        pdf = create_pdf(locals(), length, width, height, roof_type, construction, occupants, contents, results)
        
        # Save to temp file
        with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp:
            pdf.output(tmp.name)
            tmp_path = tmp.name
        
        # Read file
        with open(tmp_path, 'rb') as f:
            pdf_bytes = f.read()
        
        # Create download link
        b64 = base64.b64encode(pdf_bytes).decode()
        href = f'<a href="data:application/octet-stream;base64,{b64}" download="Lightning_Report.pdf">📥 Click here to download PDF</a>'
        st.markdown(href, unsafe_allow_html=True)
        
        # Clean up
        os.unlink(tmp_path)

else:
    st.info("👈 Enter values and click CALCULATE")

st.markdown("---")
st.caption(f"⚡ Version 1.0 | {datetime.datetime.now().strftime('%Y-%m-%d')}")