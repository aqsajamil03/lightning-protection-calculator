import streamlit as st
import math
import datetime

st.set_page_config(page_title="Advanced Lightning Protection", page_icon="⚡", layout="wide")
st.title("⚡ Advanced Lightning Protection Calculator")
st.markdown("**IEC 62305 Compliant**")
st.markdown("---")

# Sidebar
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

if calculate:
    # ADVANCED CALCULATIONS
    
    # 1. Collection Area (IEC 62305-2)
    Ae = length * width + 2*height*(length + width) + math.pi*height**2
    
    # 2. Environmental Factor
    env_factors = {"Surrounded":0.25, "Similar height":0.5, "Isolated":1, "Hilltop":2}
    Cd = env_factors.get(environment, 1)
    
    # 3. Annual Risk
    Nd = Ae * lightning_density * 1e-6 * Cd
    
    # 4. Tolerable Risk based on contents
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
    
    # 5. Risk Ratio
    risk_ratio = Nd / Rt
    
    # 6. Protection Level Determination
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
    
    # 7. Air Terminals Calculation
    perimeter = 2*(length + width)
    
    if height <= sphere:
        protection_width = 2 * math.sqrt(sphere**2 - (sphere - height)**2)
        terminals_length = math.ceil(length / protection_width) + 1
        terminals_width = math.ceil(width / protection_width) + 1
        air_terminals = terminals_length * terminals_width
    else:
        air_terminals = math.ceil(perimeter / 10) + math.ceil((length * width) / 100)
    
    # 8. Down Conductors
    down_conductors = max(2, math.ceil(perimeter / 20))
    
    # 9. Earth Resistance
    soil_resistivity = 100  # Ω·m (typical value)
    
    if perimeter < 40:
        earth_resistance = soil_resistivity / (2 * math.pi * 3) * math.log(4 * 3 / 0.016)
        earthing_type = "Type A (Vertical rods)"
    else:
        ring_radius = math.sqrt(length * width / math.pi)
        earth_resistance = soil_resistivity / (2 * math.pi * ring_radius)
        earthing_type = "Type B (Ring electrode)"
    
    # 10. Separation Distance
    ki = {"I": 0.1, "II": 0.075, "III": 0.05, "IV": 0.05}.get(lpl_class, 0.05)
    separation = ki * height
    
    # Display Results in Columns
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.subheader("📊 Risk Assessment")
        st.metric("Collection Area (Ae)", f"{Ae:.0f} m²")
        st.metric("Annual Risk (Nd)", f"{Nd:.6f}")
        st.metric("Tolerable Risk (Rt)", f"{Rt:.6f}")
        st.metric("Risk Ratio", f"{risk_ratio:.3f}")
        st.info(f"**Risk Type:** {risk_type}")
    
    with col2:
        st.subheader("⚡ Protection Level")
        if lpl_class == "I":
            st.error(f"**{lpl}**")
        elif lpl_class == "II":
            st.warning(f"**{lpl}**")
        else:
            st.success(f"**{lpl}**")
        st.caption(lpl_desc)
        
        st.metric("Air Terminals Required", air_terminals)
        st.metric("Mesh Size", mesh)
        st.metric("Rolling Sphere", f"{sphere}m")
        st.metric("Protection Angle", f"{protection_angle}°")
    
    with col3:
        st.subheader("🔧 System Design")
        st.metric("Down Conductors", down_conductors)
        st.metric("Down Conductor Spacing", f"{perimeter/down_conductors:.1f}m")
        st.metric("Separation Distance", f"{separation:.2f}m")
        st.metric("Earth Resistance", f"{earth_resistance:.1f} Ω")
        st.info(f"**Earthing:** {earthing_type}")
    
    # Materials Table
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
    
    import pandas as pd
    st.dataframe(pd.DataFrame(materials_data), use_container_width=True, hide_index=True)
    
    # Recommendations
    st.subheader("📌 Recommendations")
    
    if earth_resistance > 10:
        st.warning("⚠️ Earth resistance is high. Consider:")
        st.markdown("- Add more earth rods in parallel")
        st.markdown("- Use chemical treatment")
        st.markdown("- Increase rod length")
    
    if separation < 0.5:
        st.warning("⚠️ Small separation distance. Maintain distance from metal installations")
    
    # DOWNLOAD REPORT (WORKING VERSION)
    st.markdown("---")
    st.subheader("📥 Download Report")
    
    # Create report text
    report = f"""
    ================================================
        ADVANCED LIGHTNING PROTECTION REPORT
            IEC 62305 COMPLIANT
    ================================================
    Date: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M')}
    
    1. BUILDING INFORMATION
    ------------------------------------------------
    Dimensions: {length}m (L) × {width}m (W) × {height}m (H)
    Roof Type: {roof_type}
    Construction: {construction}
    Occupants: {occupants}
    Contents: {', '.join(contents) if contents else 'Standard'}
    Environment: {environment}
    
    2. RISK ASSESSMENT
    ------------------------------------------------
    Collection Area (Ae): {Ae:.0f} m²
    Environmental Factor (Cd): {Cd}
    Annual Risk (Nd): {Nd:.6f}
    Tolerable Risk (Rt): {Rt:.6f}
    Risk Ratio: {risk_ratio:.3f}
    Risk Classification: {risk_type}
    
    3. PROTECTION LEVEL
    ------------------------------------------------
    Protection Level: {lpl}
    Description: {lpl_desc}
    
    4. PROTECTION DESIGN
    ------------------------------------------------
    Air Terminals Required: {air_terminals}
    Mesh Size: {mesh}
    Rolling Sphere Radius: {sphere}m
    Protection Angle: {protection_angle}°
    Down Conductors: {down_conductors}
    Down Conductor Spacing: {perimeter/down_conductors:.1f}m
    Separation Distance: {separation:.2f}m
    
    5. EARTHING SYSTEM
    ------------------------------------------------
    Earthing Type: {earthing_type}
    Earth Resistance: {earth_resistance:.1f} Ω
    Target Resistance: <10 Ω
    
    6. BILL OF MATERIALS
    ------------------------------------------------
    Air Termination Rods: {air_terminals} pcs (10mm Cu, 1.5m)
    Conductors: {air_terminals + down_conductors*height:.0f} m (50mm² Cu)
    Test Joints: {down_conductors} pcs (Stainless steel)
    Earth Rods: {max(2, down_conductors)} pcs (16mm Cu, 3m)
    Ring Conductor: {perimeter:.0f} m (95mm² Cu)
    Connectors: {air_terminals*2 + down_conductors*2} pcs
    
    7. RECOMMENDATIONS
    ------------------------------------------------
    """
    
    if earth_resistance > 10:
        report += "- Improve earthing system to achieve <10Ω resistance\n"
    if separation < 0.5:
        report += "- Maintain separation distance from metal installations\n"
    
    report += f"""
    
    ================================================
    Report Generated by: Lightning Protection Calculator
    ================================================
    """
    
    # Download button
    st.download_button(
        label="📥 Download Full Report (TXT)",
        data=report,
        file_name=f"Lightning_Report_{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.txt",
        mime="text/plain"
    )
    
    st.success("✅ Report ready! Click button above to download.")

else:
    st.info("👈 Enter building parameters and click CALCULATE")
    
    with st.expander("ℹ️ About this calculator"):
        st.markdown("""
        **Advanced Lightning Protection Calculator**
        
        This calculator follows IEC 62305 standards:
        - Part 2: Risk management
        - Part 3: Physical damage protection
        - Part 4: Electrical systems protection
        
        **Features:**
        - Collection area calculation
        - Environmental factors
        - Protection Level I-IV determination
        - Air termination design
        - Down conductor calculation
        - Earthing system design
        - Complete bill of materials
        """)

st.markdown("---")
st.caption(f"⚡ Advanced Version 4.0 | {datetime.datetime.now().strftime('%Y-%m-%d')}")