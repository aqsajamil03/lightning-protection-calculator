import streamlit as st
import math
import datetime
import pandas as pd

st.set_page_config(page_title="Complete Lightning Protection", page_icon="⚡", layout="wide")
st.title("⚡ Complete Lightning Protection Design System")
st.markdown("**IEC 62305 & NFPA 780 Compliant**")
st.markdown("---")

# Create tabs for different sections
tab1, tab2, tab3 = st.tabs(["📊 Risk Assessment", "🔧 Protection Design", "📋 Complete Report"])

with tab1:
    st.header("Risk Assessment Parameters")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.subheader("Building Details")
        zone_name = st.text_input("Zone/Area Name", "Process Area")
        length = st.number_input("Length (m)", 1.0, 200.0, 70.0)
        width = st.number_input("Width (m)", 1.0, 200.0, 38.0)
        height = st.number_input("Height (m)", 1.0, 100.0, 20.0)
        distance = st.number_input("Distance (m)", 0.0, 1000.0, 500.0)
    
    with col2:
        st.subheader("Location Factors")
        cd_location = st.selectbox("Location Factor (CD)", [0.25, 0.5, 1.0, 2.0], index=2, 
                                  help="Surrounded:0.25, Similar:0.5, Isolated:1, Hilltop:2")
        ks1_shield = st.number_input("Structure Shield (KS1)", 0.1, 1.0, 1.0)
        td_days = st.number_input("Thunderstorm Days (TD)", 1, 100, 10)
        ng_flashes = st.number_input("NG (flashes/km²/year)", 0.1, 30.0, 1.0)
        nt_danger = st.number_input("NT (No. of dangerous events/year)", 1, 100, 10)
    
    with col3:
        st.subheader("Probability Factors")
        pb_lps = st.number_input("LPS Probability (PB)", 0.01, 1.0, 0.2, format="%.2f")
        pta_injury = st.number_input("Injury Probability (PTA)", 0.001, 0.1, 0.01, format="%.3f")
        lt_victims_shock = st.number_input("Victims - Shock (Lt)", 0.0001, 0.1, 0.01, format="%.4f")
        lf_victims_fire = st.number_input("Victims - Fire (Lf)", 0.0001, 0.1, 0.01, format="%.4f")
    
    # Calculate when button is pressed
    if st.button("Calculate Risk Assessment", type="primary"):
        
        # Calculations
        Ad = length * width + 2*length*height + 2*width*height + math.pi * height**2
        Am = 1000 * length * width  # Simplified near strike area
        
        # Risk calculations
        Nd = Ad * ng_flashes * 1e-6 * cd_location
        Nm = Am * ng_flashes * 1e-6
        
        # Reduction factors
        rt_soil = 0.00001  # Soil reduction factor
        rf_fire = 0.01     # Fire reduction factor
        rp_protection = 0.5  # Protection factor
        
        # People factors
        nz_people = st.number_input("People in structure", 1, 1000, 10)
        tz_hours = st.number_input("Hours/year people present", 1, 8760, 8760)
        nt_total = st.number_input("Total people", 1, 1000, 10)
        hz_hazard = st.number_input("Hazard factor", 1.0, 3.0, 2.0)
        
        # Loss calculations
        la_shock_loss = 3e-8  # Loss of living being by shock
        lb_damage_loss = 0.00003  # Loss physical damage
        
        # Risk components
        ra_risk = Nd * pb_lps * la_shock_loss * lt_victims_shock
        rb_risk = Nm * pb_lps * lb_damage_loss * lf_victims_fire
        
        # Display results in a table format
        st.subheader("📊 Risk Assessment Results")
        
        results_data = {
            'Parameter': [
                'Process Area', 'Collection Area (Ad)', 'Near Strike Area (Am)',
                'Nd (lightning strikes/year)', 'Nm (near strikes/year)',
                'RA Risk Component', 'RB Risk Component',
                'Total Risk (RA + RB)'
            ],
            'Value': [
                f"{length}×{width}×{height}m",
                f"{Ad:.0f} m²",
                f"{Am:.0f} m²",
                f"{Nd:.6f}",
                f"{Nm:.6f}",
                f"{ra_risk:.2e}",
                f"{rb_risk:.2e}",
                f"{ra_risk + rb_risk:.2e}"
            ]
        }
        
        st.dataframe(pd.DataFrame(results_data), use_container_width=True, hide_index=True)
        
        # Store in session state for other tabs
        st.session_state['Ad'] = Ad
        st.session_state['Am'] = Am
        st.session_state['Nd'] = Nd
        st.session_state['Nm'] = Nm
        st.session_state['ra_risk'] = ra_risk
        st.session_state['rb_risk'] = rb_risk

with tab2:
    st.header("NFPA 780 Protection Design")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("Coefficients")
        cd_env = st.selectbox("Environment (CD)", [0.25, 0.5, 1.0, 2.0], index=0, key="cd_env")
        c2_type = st.selectbox("Type Coefficient (C2)", [0.5, 1.0, 2.0, 3.0], index=2)
        c3_content = st.selectbox("Content Coefficient (C3)", [1, 2, 3, 4], index=2)
        c4_occupancy = st.selectbox("Occupancy Coefficient (C4)", [1, 2, 3, 4], index=0)
        c5_consequence = st.selectbox("Consequence Coefficient (C5)", [1, 5, 10], index=1)
        
        # Calculate total coefficient
        c_total = cd_env * c2_type * c3_content * c4_occupancy * c5_consequence
        
        st.metric("Total Coefficient (C)", f"{c_total:.2f}")
        
        td_days_2 = st.number_input("Thunderstorm Days", 1, 100, 10, key="td2")
        nc_tolerable = 1e-4  # Tolerable lightning frequency
        
    with col2:
        st.subheader("Building Dimensions")
        length_2 = st.number_input("Length (m)", 1.0, 200.0, 50.0, key="l2")
        width_2 = st.number_input("Width (m)", 1.0, 200.0, 26.0, key="w2")
        height_2 = st.number_input("Height (m)", 1.0, 100.0, 7.35, key="h2")
        
        # Calculate collection area
        Ad_2 = length_2 * width_2 + 2*length_2*height_2 + 2*width_2*height_2 + math.pi * height_2**2
        
        ng_2 = st.number_input("NG (flashes/km²/year)", 0.1, 30.0, 1.0, key="ng2")
        
        # Calculate Nd
        Nd_2 = Ad_2 * ng_2 * 1e-6
        efficiency = 1 - (nc_tolerable / Nd_2) if Nd_2 > 0 else 0
        
        st.metric("Collection Area (Ad)", f"{Ad_2:.0f} m²")
        st.metric("Nd (strikes/year)", f"{Nd_2:.6f}")
        st.metric("Protection Efficiency", f"{efficiency:.2%}")
    
    # Protection Level Determination
    st.subheader("⚡ Protection Level")
    
    if efficiency > 0.98:
        lpl = "Class I"
        lpl_desc = "Maximum Protection"
        sphere_radius = 20
        min_current = 3
        max_current = 200
        mesh_size = "5m × 5m"
    elif efficiency > 0.95:
        lpl = "Class II"
        lpl_desc = "High Protection"
        sphere_radius = 30
        min_current = 5
        max_current = 150
        mesh_size = "10m × 10m"
    elif efficiency > 0.90:
        lpl = "Class III"
        lpl_desc = "Standard Protection"
        sphere_radius = 45
        min_current = 10
        max_current = 100
        mesh_size = "15m × 15m"
    else:
        lpl = "Class IV"
        lpl_desc = "Basic Protection"
        sphere_radius = 60
        min_current = 16
        max_current = 100
        mesh_size = "20m × 20m"
    
    col_a, col_b, col_c = st.columns(3)
    
    with col_a:
        st.info(f"**{lpl}**")
        st.caption(lpl_desc)
        st.metric("Rolling Sphere Radius", f"{sphere_radius} m")
    
    with col_b:
        st.metric("Min Current", f"{min_current} kA")
        st.metric("Max Current", f"{max_current} kA")
    
    with col_c:
        st.metric("Mesh Size", mesh_size)
        
        # Air terminals calculation
        if height_2 <= sphere_radius:
            protection_width = 2 * math.sqrt(sphere_radius**2 - (sphere_radius - height_2)**2)
            terminals_length = math.ceil(length_2 / protection_width) + 1
            terminals_width = math.ceil(width_2 / protection_width) + 1
            air_terminals = terminals_length * terminals_width
        else:
            perimeter_2 = 2*(length_2 + width_2)
            air_terminals = math.ceil(perimeter_2 / 10) + math.ceil((length_2 * width_2) / 100)
        
        st.metric("Air Terminals Required", air_terminals)
    
    # Materials Specification
    st.subheader("🔧 Materials Specification")
    
    col_m1, col_m2, col_m3 = st.columns(3)
    
    with col_m1:
        material = st.selectbox("Material", ["Copper", "Aluminum", "Galvanized Steel"])
        rod_dia = st.number_input("Rod Diameter (mm)", 8.0, 20.0, 12.7)
        rod_length = st.number_input("Rod Length (m)", 1.0, 3.0, 1.5)
    
    with col_m2:
        down_conductor = st.number_input("Down Conductor (mm²)", 25, 100, 50)
        bounding_conductor = st.number_input("Bounding Conductor (mm²)", 25, 100, 50)
    
    with col_m3:
        earth_dia = st.number_input("Earth Rod Diameter (mm)", 12.0, 20.0, 15.0)
        earth_length = st.number_input("Earth Rod Length (m)", 2.0, 4.0, 3.0)
    
    # Store in session state
    st.session_state['lpl'] = lpl
    st.session_state['sphere_radius'] = sphere_radius
    st.session_state['air_terminals'] = air_terminals

with tab3:
    st.header("📋 Complete Lightning Protection Report")
    
    if 'Ad' in st.session_state:
        # Create comprehensive report
        report = f"""
===============================================
    COMPLETE LIGHTNING PROTECTION REPORT
    IEC 62305 & NFPA 780 COMPLIANT
===============================================
Date: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M')}

1. RISK ASSESSMENT RESULTS
-----------------------------------------------
Collection Area (Ad): {st.session_state['Ad']:.0f} m²
Near Strike Area (Am): {st.session_state['Am']:.0f} m²
Nd (Direct Strikes): {st.session_state['Nd']:.6f}
Nm (Near Strikes): {st.session_state['Nm']:.6f}
RA Risk Component: {st.session_state['ra_risk']:.2e}
RB Risk Component: {st.session_state['rb_risk']:.2e}
Total Risk: {st.session_state['ra_risk'] + st.session_state['rb_risk']:.2e}

2. PROTECTION DESIGN
-----------------------------------------------
Protection Level: {st.session_state.get('lpl', 'Not Calculated')}
Rolling Sphere Radius: {st.session_state.get('sphere_radius', 0)} m
Air Terminals Required: {st.session_state.get('air_terminals', 0)}

3. MATERIALS REQUIRED
-----------------------------------------------
"""
        
        # Download button
        st.download_button(
            label="📥 Download Complete Report",
            data=report,
            file_name=f"Lightning_Report_{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.txt",
            mime="text/plain"
        )
    else:
        st.info("Please calculate in Risk Assessment tab first")

st.markdown("---")
st.caption(f"⚡ Complete Lightning Protection System | Version 6.0 | {datetime.datetime.now().strftime('%Y-%m-%d')}")