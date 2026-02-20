import streamlit as st
import math
import datetime
import pandas as pd

st.set_page_config(page_title="Complete Lightning Protection", page_icon="⚡", layout="wide")
st.title("⚡ Complete Lightning Protection Design System")
st.markdown("**IEC 62305 & NFPA 780 Compliant**")
st.markdown("---")

# Sidebar mein checkboxes
with st.sidebar:
    st.header("🔘 Select Sections to Display")
    
    show_risk = st.checkbox("📊 Risk Assessment", value=True)
    show_protection = st.checkbox("⚡ Protection Level", value=True)
    show_air = st.checkbox("🏗️ Air Termination", value=True)
    show_down = st.checkbox("📡 Down Conductors", value=True)
    show_earthing = st.checkbox("🔧 Earthing System", value=True)
    show_materials = st.checkbox("📋 Bill of Materials", value=True)
    show_report = st.checkbox("📥 Download Report", value=True)
    
    st.markdown("---")
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

if calculate:
    # ========== IEC 62305-2 CALCULATIONS ==========
    
    # 1. COLLECTION AREAS
    Ad = length * width + 2*length*height + 2*width*height + math.pi * height**2
    
    if environment == "Surrounded":
        Am = 500 * length * width
    elif environment == "Similar height":
        Am = 1000 * length * width
    elif environment == "Isolated":
        Am = 2000 * length * width
    else:
        Am = 3000 * length * width
    
    # 2. ENVIRONMENTAL FACTOR
    env_factors = {"Surrounded":0.25, "Similar height":0.5, "Isolated":1, "Hilltop":2}
    Cd = env_factors.get(environment, 1)
    
    # 3. ANNUAL RISK
    Nd = Ad * lightning_density * 1e-6 * Cd
    Nm = Am * lightning_density * 1e-6
    N_total = Nd + Nm
    
    # 4. TOLERABLE RISK
    if "Explosive" in contents:
        Rt = 0.00001
        risk_type = "Very High (Explosive)"
    elif "Hospital" in contents:
        Rt = 0.0001
        risk_type = "High (Hospital)"
    elif "Valuable" in contents:
        Rt = 0.001
        risk_type = "Medium (Valuable)"
    else:
        Rt = 0.01
        risk_type = "Low (Ordinary)"
    
    risk_ratio = N_total / Rt
    
    # 5. PROTECTION LEVEL
    if risk_ratio < 0.001:
        lpl = "IV (Optional)"
        lpl_class = "IV"
        lpl_desc = "Basic protection"
        mesh = "20m x 20m"
        sphere = 60
        protection_angle = 60
    elif risk_ratio < 0.01:
        lpl = "III (Standard)"
        lpl_class = "III"
        lpl_desc = "Standard commercial"
        mesh = "15m x 15m"
        sphere = 45
        protection_angle = 45
    elif risk_ratio < 0.1:
        lpl = "II (Enhanced)"
        lpl_class = "II"
        lpl_desc = "Enhanced protection"
        mesh = "10m x 10m"
        sphere = 30
        protection_angle = 35
    else:
        lpl = "I (Maximum)"
        lpl_class = "I"
        lpl_desc = "Maximum protection"
        mesh = "5m x 5m"
        sphere = 20
        protection_angle = 25
    
    # 6. AIR TERMINALS
    perimeter = 2*(length + width)
    
    if height <= sphere:
        protection_width = 2 * math.sqrt(sphere**2 - (sphere - height)**2)
        terminals_length = math.ceil(length / protection_width) + 1
        terminals_width = math.ceil(width / protection_width) + 1
        air_terminals = terminals_length * terminals_width
    else:
        air_terminals = math.ceil(perimeter / 10) + math.ceil((length * width) / 100)
    
    # 7. DOWN CONDUCTORS
    down_conductors = max(2, math.ceil(perimeter / 20))
    
    # 8. EARTHING
    soil_resistivity = 100
    
    if perimeter < 40:
        earth_resistance = soil_resistivity / (2 * math.pi * 3) * math.log(4 * 3 / 0.016)
        earthing_type = "Type A (Vertical rods)"
    else:
        ring_radius = math.sqrt(length * width / math.pi)
        earth_resistance = soil_resistivity / (2 * math.pi * ring_radius)
        earthing_type = "Type B (Ring electrode)"
    
    # 9. SEPARATION DISTANCE
    ki = {"I": 0.1, "II": 0.075, "III": 0.05, "IV": 0.05}.get(lpl_class, 0.05)
    separation = ki * height
    
    # ========== DISPLAY SECTIONS BASED ON CHECKBOXES ==========
    
    st.markdown("## 📋 Calculation Results")
    st.markdown("---")
    
    # 1. Risk Assessment Section
    if show_risk:
        with st.container():
            st.subheader("📊 Risk Assessment")
            col_r1, col_r2, col_r3, col_r4 = st.columns(4)
            with col_r1:
                st.metric("Direct Area (Ad)", f"{Ad:.0f} m²")
                st.metric("Direct Risk (Nd)", f"{Nd:.6f}")
            with col_r2:
                st.metric("Near Area (Am)", f"{Am:.0f} m²")
                st.metric("Near Risk (Nm)", f"{Nm:.6f}")
            with col_r3:
                st.metric("Total Risk", f"{N_total:.6f}")
                st.metric("Tolerable Risk", f"{Rt:.6f}")
            with col_r4:
                st.metric("Risk Ratio", f"{risk_ratio:.3f}")
                st.info(f"**{risk_type}**")
            st.markdown("---")
    
    # 2. Protection Level Section
    if show_protection:
        with st.container():
            st.subheader("⚡ Protection Level")
            col_p1, col_p2, col_p3 = st.columns(3)
            with col_p1:
                if lpl_class == "I":
                    st.error(f"**{lpl}**")
                elif lpl_class == "II":
                    st.warning(f"**{lpl}**")
                else:
                    st.success(f"**{lpl}**")
                st.caption(lpl_desc)
            with col_p2:
                st.metric("Mesh Size", mesh)
                st.metric("Rolling Sphere", f"{sphere}m")
            with col_p3:
                st.metric("Protection Angle", f"{protection_angle}°")
                st.metric("Risk Classification", risk_type)
            st.markdown("---")
    
    # 3. Air Termination Section
    if show_air:
        with st.container():
            st.subheader("🏗️ Air Termination Design")
            col_a1, col_a2 = st.columns(2)
            with col_a1:
                st.metric("Air Terminals Required", air_terminals)
                st.metric("Mesh Size", mesh)
            with col_a2:
                st.metric("Rolling Sphere Radius", f"{sphere}m")
                st.metric("Protection Angle", f"{protection_angle}°")
                
                # Air terminal placement guide
                st.info("**Placement Guide:**")
                st.markdown("- Install on roof edges")
                st.markdown(f"- Max spacing: {sphere:.0f}m")
            st.markdown("---")
    
    # 4. Down Conductors Section
    if show_down:
        with st.container():
            st.subheader("📡 Down Conductor System")
            col_d1, col_d2 = st.columns(2)
            with col_d1:
                st.metric("Number of Down Conductors", down_conductors)
                st.metric("Spacing", f"{perimeter/down_conductors:.1f}m")
            with col_d2:
                st.metric("Separation Distance", f"{separation:.2f}m")
                st.info("**Requirements:**")
                st.markdown("- 50mm² Cu minimum")
                st.markdown("- Test joints at each")
            st.markdown("---")
    
    # 5. Earthing System Section
    if show_earthing:
        with st.container():
            st.subheader("🔧 Earthing System Design")
            col_e1, col_e2 = st.columns(2)
            with col_e1:
                st.metric("Earth Resistance", f"{earth_resistance:.1f} Ω")
                st.info(f"**Type:** {earthing_type}")
            with col_e2:
                if earth_resistance < 10:
                    st.success("✅ Excellent - Meets IEC requirements")
                elif earth_resistance < 30:
                    st.warning("⚠️ Acceptable - Consider improvement")
                else:
                    st.error("❌ Too high - Redesign required")
                
                st.markdown("**Recommendations:**")
                st.markdown("- Use multiple rods in parallel")
                st.markdown("- Connect to building foundation")
            st.markdown("---")
    
    # 6. Materials Section
    if show_materials:
        with st.container():
            st.subheader("📋 Bill of Materials")
            
            materials_data = {
                'Component': [
                    'Air Termination Rods',
                    'Conductors (Air/Down)',
                    'Test Joints',
                    'Earth Rods',
                    'Ring Conductor',
                    'Connectors'
                ],
                'Quantity': [
                    f"{air_terminals} pcs",
                    f"{air_terminals + down_conductors*height:.0f} m",
                    f"{down_conductors} pcs",
                    f"{max(2, down_conductors)} pcs",
                    f"{perimeter:.0f} m",
                    f"{air_terminals*2 + down_conductors*2} pcs"
                ],
                'Specification': [
                    '10mm Cu, 1.5m length',
                    '50mm² Cu',
                    'Stainless steel',
                    '16mm Cu, 3m length',
                    '95mm² Cu',
                    'Stainless steel'
                ]
            }
            
            st.dataframe(pd.DataFrame(materials_data), use_container_width=True, hide_index=True)
            st.markdown("---")
    
    # 7. Download Report Section
    if show_report:
        with st.container():
            st.subheader("📥 Download Complete Report")
            
            # Create report text
            report = f"""
================================================
    COMPLETE LIGHTNING PROTECTION REPORT
    IEC 62305-2 & NFPA 780 COMPLIANT
================================================
Date: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M')}

BUILDING DETAILS:
-----------------------------------------------
Dimensions: {length}m × {width}m × {height}m
Roof Type: {roof_type}
Construction: {construction}
Environment: {environment}
Contents: {', '.join(contents) if contents else 'Standard'}
Occupants: {occupants}

RISK ASSESSMENT:
-----------------------------------------------
Collection Area (Ad): {Ad:.0f} m²
Near Strike Area (Am): {Am:.0f} m²
Direct Risk (Nd): {Nd:.6f}
Near Risk (Nm): {Nm:.6f}
Total Risk: {N_total:.6f}
Tolerable Risk: {Rt:.6f}
Risk Ratio: {risk_ratio:.3f}
Risk Type: {risk_type}

PROTECTION LEVEL:
-----------------------------------------------
Level: {lpl}
Description: {lpl_desc}
Mesh Size: {mesh}
Rolling Sphere: {sphere}m
Protection Angle: {protection_angle}°

PROTECTION DESIGN:
-----------------------------------------------
Air Terminals: {air_terminals}
Down Conductors: {down_conductors}
Spacing: {perimeter/down_conductors:.1f}m
Separation Distance: {separation:.2f}m

EARTHING SYSTEM:
-----------------------------------------------
Type: {earthing_type}
Resistance: {earth_resistance:.1f} Ω
Target: <10 Ω

MATERIALS REQUIRED:
-----------------------------------------------
Air Terminals: {air_terminals} pcs (10mm Cu, 1.5m)
Conductors: {air_terminals + down_conductors*height:.0f} m (50mm² Cu)
Test Joints: {down_conductors} pcs
Earth Rods: {max(2, down_conductors)} pcs (16mm Cu, 3m)
Ring Conductor: {perimeter:.0f} m (95mm² Cu)
Connectors: {air_terminals*2 + down_conductors*2} pcs

================================================
Generated by: Advanced Lightning Protection System
================================================
"""
            
            st.download_button(
                label="📥 Download Report (TXT)",
                data=report,
                file_name=f"Lightning_Report_{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.txt",
                mime="text/plain"
            )

else:
    st.info("👈 Select sections from sidebar and click CALCULATE")

st.markdown("---")
st.caption(f"⚡ Complete Lightning Protection System | Version 7.0 | {datetime.datetime.now().strftime('%Y%m%d')}")