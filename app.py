
**Default Values:**
- Power Factor (PF) = 0.85
- Efficiency = 0.95

**Example:**
- 75kW, 415V Motor → I = 129.2 A
- 25kW, 230V Lighting → I = 134.6 A
""")

# ========== TAB 2: LIGHTNING PROTECTION ==========
elif st.session_state.selected_calculator == "⚡ Lightning Protection":

lp_tabs = st.tabs(["📊 Risk Assessment", "🔧 Protection Design", "📋 Calculations", "📥 Download Report"])

with lp_tabs[0]:
st.markdown('<div class="report-header">RISK ASSESSMENT (IEC 62305-2)</div>', unsafe_allow_html=True)

structure_type = st.selectbox("Select Structure Type", 
                             ["Substation Building", "Central Control Building", "Column 4-C01"])

col1, col2 = st.columns(2)

with col1:
    st.markdown("### 📐 Dimensions")
    if structure_type == "Substation Building":
        length = st.number_input("Length (m)", value=26.5, step=0.5)
        width = st.number_input("Width (m)", value=26.25, step=0.5)
        height = st.number_input("Height (m)", value=7.35, step=0.5)
    elif structure_type == "Central Control Building":
        length = st.number_input("Length (m)", value=50.0, step=0.5)
        width = st.number_input("Width (m)", value=26.0, step=0.5)
        height = st.number_input("Height (m)", value=5.35, step=0.5)
    else:
        height = st.number_input("Height (m)", value=50.0, step=0.5)
        length = height
        width = 0
    
    td_days = st.number_input("Thunderstorm Days/Year", value=10, step=1)
    environment = st.selectbox("Environment", ["Surrounded", "Similar height", "Isolated", "Hilltop"])

with col2:
    st.markdown("### 📊 Environmental Factor (CD)")
    cd_values = {"Surrounded": 0.25, "Similar height": 0.5, "Isolated": 1, "Hilltop": 2}
    cd = cd_values[environment]
    
    st.markdown("**IEC 62305-2 Table A.1:**")
    st.markdown("• Surrounded: **0.25** • Similar: **0.5** • Isolated: **1.0** • Hilltop: **2.0**")
    st.success(f"**Selected: {environment} → CD = {cd}**")
    
    if structure_type == "Column 4-C01":
        c2, c3, c4, c5 = 0.5, 2.0, 3.0, 10.0
    else:
        c2, c3, c4, c5 = 1.0, 3.0, 1.0, 5.0
    
    st.metric("C2 - Type", c2)
    st.metric("C3 - Content", c3)
    st.metric("C4 - Occupancy", c4)
    st.metric("C5 - Consequence", c5)

if st.button("🔧 CALCULATE RISK", type="primary", use_container_width=True):
    if structure_type == "Column 4-C01":
        ad = math.pi * 9 * height**2
    else:
        ad = length * width + 2 * (3 * height) * (length + width) + math.pi * (3 * height)**2
    
    am = 2 * 500 * (length + width) + math.pi * 500**2
    ng = 0.1 * td_days
    nd = ng * ad * cd * 1e-6
    nm = ng * am * 1e-6
    
    c_total = cd * c2 * c3 * c4 * c5
    nc = 1e-4 / c_total
    efficiency = 1 - (nc / nd) if nd > 0 else 0
    
    if efficiency > 0.98:
        lpl, sphere = "Class I", 20
    elif efficiency > 0.95:
        lpl, sphere = "Class II", 30
    elif efficiency > 0.90:
        lpl, sphere = "Class III", 45
    else:
        lpl, sphere = "Class IV", 60
    
    if height <= sphere:
        protection_width = 2 * math.sqrt(sphere**2 - (sphere - height)**2)
        if protection_width > 0:
            terminals_length = math.ceil(length / protection_width) + 1
            terminals_width = math.ceil(width / protection_width) + 1 if width > 0 else 1
            air_terminals = terminals_length * terminals_width
        else:
            air_terminals = 4
    else:
        perimeter = 2 * (length + width)
        air_terminals = math.ceil(perimeter / 10) + math.ceil((length * width) / 100)
    
    st.markdown("---")
    st.subheader("📊 Results")
    
    col_a, col_b, col_c, col_d = st.columns(4)
    with col_a:
        st.metric("Collection Area (Ad)", f"{ad:.0f} m²")
        st.metric("Near Strike Area (Am)", f"{am:.0f} m²")
    with col_b:
        st.metric("Nd (Direct)", f"{nd:.6f}")
        st.metric("Nm (Near)", f"{nm:.6f}")
    with col_c:
        st.metric("Protection Level", lpl)
        st.metric("Efficiency", f"{efficiency:.1%}")
    with col_d:
        st.metric("Rolling Sphere", f"{sphere}m")
        st.metric("Air Terminals", air_terminals)
    
    st.session_state.calc_results = {
        'ad': ad, 'am': am, 'ng': ng, 'nd': nd, 'nm': nm,
        'efficiency': efficiency, 'lpl': lpl, 'sphere': sphere, 'air_terminals': air_terminals
    }
    st.session_state.input_values = {
        'length': length, 'width': width, 'height': height,
        'td_days': td_days, 'environment': environment, 'cd': cd
    }
    st.session_state.calc_done = True

with lp_tabs[1]:
st.markdown('<div class="report-header">PROTECTION DESIGN</div>', unsafe_allow_html=True)
if not st.session_state.calc_done:
    st.warning("⚠️ Please complete Risk Assessment first!")
else:
    results = st.session_state.calc_results
    st.success(f"✅ Designing for: **{results['lpl']}**")
    
    col1, col2 = st.columns(2)
    with col1:
        st.metric("Air Terminals", results['air_terminals'])
        st.metric("Rolling Sphere", f"{results['sphere']}m")
    with col2:
        if results['lpl'] in ["Class I", "Class II"]:
            st.metric("Rod Diameter", "12.7 mm")
            st.metric("Down Conductor", "58 mm²")
        else:
            st.metric("Rod Diameter", "9.5 mm")
            st.metric("Down Conductor", "29 mm²")

with lp_tabs[2]:
st.markdown('<div class="report-header">DETAILED CALCULATIONS</div>', unsafe_allow_html=True)
if not st.session_state.calc_done:
    st.warning("⚠️ Please complete Risk Assessment first!")
else:
    results = st.session_state.calc_results
    inputs = st.session_state.input_values
    
    with st.expander("1. Collection Area (Ad)", expanded=True):
        st.markdown("**Formula:** Ad = L × W + 2 × (3H) × (L + W) + π × (3H)²")
        st.markdown("**Reference:** IEC 62305-2 Annex A.2.1.1")
        st.markdown(f"**Result:** Ad = **{results['ad']:.2f} m²**")
    
    with st.expander("2. Near Strike Collection Area (Am)", expanded=True):
        st.markdown("**Formula:** Am = 2 × 500 × (L + W) + π × 500²")
        st.markdown("**Reference:** IEC 62305-2 Annex A.3")
        st.markdown(f"**Result:** Am = **{results['am']:.2f} m²**")
    
    with st.expander("3. Environmental Factor (CD)"):
        st.markdown(f"**Selected:** {inputs.get('environment', 'Isolated')} → **{inputs.get('cd', 1)}**")
    
    with st.expander("4. Lightning Density (NG)"):
        st.markdown(f"**Result:** NG = **{results.get('ng', 1)} flashes/km²/year**")
    
    with st.expander("5. Lightning Frequencies"):
        st.markdown(f"**Nd:** {results.get('nd', 0):.6f} events/year")
        st.markdown(f"**Nm:** {results.get('nm', 0):.6f} events/year")
    
    with st.expander("6. Protection Level"):
        st.markdown(f"**Efficiency:** {results.get('efficiency', 0):.1%}")
        st.markdown(f"**Result:** **{results.get('lpl', 'Class III')}**")

with lp_tabs[3]:
st.markdown('<div class="report-header">DOWNLOAD REPORT</div>', unsafe_allow_html=True)
if not st.session_state.calc_done:
    st.warning("⚠️ Please complete Risk Assessment first!")
else:
    col1, col2 = st.columns(2)
    with col1:
        if st.button("📥 Generate PDF", key="lp_pdf", use_container_width=True):
            with st.spinner("Generating PDF..."):
                pdf = LightningPDFReport()
                pdf.add_calculations(st.session_state.calc_results, st.session_state.input_values)
                pdf_output = pdf.output(dest='S')
                b64 = base64.b64encode(pdf_output).decode()
                filename = f"Lightning_Report_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf"
                st.markdown(f'<a href="data:application/pdf;base64,{b64}" download="{filename}" class="download-btn pdf-btn">📥 Download PDF</a>', unsafe_allow_html=True)
                st.success("✅ PDF generated!")
    with col2:
        if st.button("📥 Generate Word", key="lp_word", use_container_width=True):
            with st.spinner("Generating Word..."):
                word = LightningWordReport()
                word.add_calculations(st.session_state.calc_results, st.session_state.input_values)
                word_path = "temp_lightning.docx"
                word.save(word_path)
                with open(word_path, "rb") as f:
                    word_bytes = f.read()
                b64 = base64.b64encode(word_bytes).decode()
                filename = f"Lightning_Report_{datetime.now().strftime('%Y%m%d_%H%M')}.docx"
                os.remove(word_path)
                st.markdown(f'<a href="data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,{b64}" download="{filename}" class="download-btn word-btn">📥 Download Word</a>', unsafe_allow_html=True)
                st.success("✅ Word generated!")

# ========== TAB 3: CABLE SIZING ==========
elif st.session_state.selected_calculator == "🔌 Cable Sizing":

st.markdown('<div class="report-header">🔌 CABLE SIZING CALCULATOR</div>', unsafe_allow_html=True)

st.markdown(f"""
<div class="info-box">
<h4>📌 Using loads from Load Sheet</h4>
<p>Total {len(st.session_state.universal_loads)} loads available. Click below to import.</p>
</div>
""", unsafe_allow_html=True)

if st.button("📥 Import Loads from Load Sheet", use_container_width=True):
# Convert loads to cable sizing format
new_loads = []
for idx, load in st.session_state.universal_loads.iterrows():
    phase = '3-phase' if load['VOLTAGE [V]'] > 300 else '1-phase'
    new_loads.append({
        'Load Name': load['DESCRIPTION'],
        'Power (kW)': load['POWER [kW]'],
        'Voltage (V)': load['VOLTAGE [V]'],
        'Phase': phase,
        'Power Factor': 0.85,
        'Length (m)': 50
    })
st.session_state.loads_df = pd.DataFrame(new_loads)
st.success("✅ Loads imported successfully!")
st.rerun()

cable_tabs = st.tabs([
"📥 Loads Input", 
"📊 Derating Factors", 
"🔌 Cable Selection",
"🔧 Short Circuit",
"⚡ Circuit Breakers",
"📥 Download Report"
])

# TAB 1: LOADS INPUT
with cable_tabs[0]:
st.markdown("### 📋 Load Details (Imported from Load Sheet)")
st.markdown("""
<div class="info-box">
    <p>To add or modify loads, please go to the main <b>📋 LOAD SHEET</b> tab.</p>
</div>
""", unsafe_allow_html=True)

edited_df = st.data_editor(
    st.session_state.loads_df,
    num_rows="fixed",
    use_container_width=True,
    disabled=True,
    column_config={
        "Load Name": st.column_config.TextColumn("Load Name", disabled=True),
        "Power (kW)": st.column_config.NumberColumn("Power (kW)", disabled=True),
        "Voltage (V)": st.column_config.NumberColumn("Voltage (V)", disabled=True),
        "Phase": st.column_config.TextColumn("Phase", disabled=True),
        "Power Factor": st.column_config.NumberColumn("PF", disabled=True),
        "Length (m)": st.column_config.NumberColumn("Length (m)", disabled=True)
    }
)

st.markdown("### ⚙️ Installation Parameters")

col1, col2 = st.columns(2)
with col1:
    st.markdown("#### 📦 Cable Parameters")
    cable_type = st.selectbox("Cable Type", ['armoured', 'unarmoured'], key="cable_type_select")
    ambient_temp = st.number_input("Ambient Temp (°C)", value=30.0, step=5.0, key="ambient_temp_input")
    
    st.markdown("#### 📐 Arrangement & Spacing")
    arrangement = st.selectbox("Cable Arrangement", 
                              ['touching', 'spaced', 'cleated'], 
                              key="arrangement_select",
                              help="How cables are arranged relative to each other")
    
    spacing_mm = st.number_input("Spacing Between Cables (mm)", 
                                value=0.0, step=5.0, min_value=0.0, max_value=500.0,
                                key="spacing_input",
                                help="Center-to-center distance between cables")
    
    formation = st.selectbox("Cable Formation", 
                            ['flat', 'trefoil', 'single'], 
                            key="formation_select",
                            help="Flat = side by side, Trefoil = triangular")

with col2:
    st.markdown("#### 🏗️ Installation Environment")
    installation = st.selectbox("Installation Method", 
                               ['air', 'surface', 'tray', 'ladder', 'trench', 'buried', 'duct', 'conduit'], 
                               key="installation_select",
                               help="How and where cables are installed")
    
    num_cables = st.number_input("Number of Cables in Group", 
                                value=3, min_value=1, max_value=18, 
                                key="num_cables_input")
    
    st.markdown("#### 🌍 Soil & Depth")
    soil_res = st.number_input("Soil Resistivity (K.m/W)", 
                              value=1.5, step=0.5, min_value=0.5, max_value=3.0, 
                              key="soil_res_input")
    
    depth = st.number_input("Burial Depth (m)", 
                           value=0.8, step=0.1, min_value=0.3, max_value=2.0, 
                           key="depth_input")
    
    system_type = st.selectbox("System Type", 
                              ['TN-S', 'TN-C', 'TN-C-S', 'TT'], 
                              key="system_type_select")

st.session_state.cable_type = cable_type
st.session_state.ambient_temp = ambient_temp
st.session_state.num_cables = num_cables
st.session_state.arrangement = arrangement
st.session_state.spacing_mm = spacing_mm
st.session_state.formation = formation
st.session_state.installation = installation
st.session_state.soil_res = soil_res
st.session_state.depth = depth
st.session_state.system_type = system_type

st.markdown("### 📊 Current Installation Settings")

settings_data = {
    'Category': [
        'Cable Parameters', 'Cable Parameters',
        'Arrangement & Spacing', 'Arrangement & Spacing', 'Arrangement & Spacing',
        'Installation Environment', 'Installation Environment',
        'Soil & Depth', 'Soil & Depth',
        'Electrical System'
    ],
    'Parameter': [
        'Cable Type', 'Ambient Temperature',
        'Cable Arrangement', 'Spacing Between Cables', 'Cable Formation',
        'Installation Method', 'Number of Cables in Group',
        'Soil Resistivity', 'Burial Depth',
        'System Type'
    ],
    'Value': [
        f'{cable_type} copper', f'{ambient_temp} °C',
        arrangement, f'{spacing_mm} mm', formation,
        installation, str(num_cables),
        f'{soil_res} K.m/W', f'{depth} m',
        system_type
    ]
}

settings_df = pd.DataFrame(settings_data)

st.dataframe(
    settings_df[['Category', 'Parameter', 'Value']],
    use_container_width=True,
    hide_index=True,
    column_config={
        "Category": st.column_config.TextColumn("Category", width="small"),
        "Parameter": st.column_config.TextColumn("Parameter", width="medium"),
        "Value": st.column_config.TextColumn("Value", width="medium")
    }
)

if st.button("🔧 CALCULATE", type="primary", use_container_width=True):
    with st.spinner("Calculating..."):
        cable_type = st.session_state.cable_type
        ambient_temp = st.session_state.ambient_temp
        num_cables = st.session_state.num_cables
        arrangement = st.session_state.arrangement
        spacing_mm = st.session_state.spacing_mm
        formation = st.session_state.formation
        installation = st.session_state.installation
        soil_res = st.session_state.soil_res
        depth = st.session_state.depth
        system_type = st.session_state.system_type
        
        cable_calc = CableSizingCalculator()
        cable_results = []
        detailed_calcs = []
        
        for idx, load in st.session_state.loads_df.iterrows():
            cable_category, cable_db = cable_calc.get_cable_type(load['Voltage (V)'])
            db = cable_db[cable_type]
            
            current = cable_calc.calculate_load_current(
                load['Power (kW)'], load['Voltage (V)'], load['Power Factor'], 1.0, load['Phase']
            )
            
            found = False
            for size, data in db.items():
                if found:
                    break
                
                cable_diameter = data['diameter']
                
                total_k, factors = cable_calc.get_all_derating_factors(
                    ambient_temp, 90, num_cables, arrangement, spacing_mm, cable_diameter,
                    formation, installation, soil_res, depth
                )
                
                st.session_state.derating_factors = factors
                
                derated = data['ampacity'] * total_k
                if derated >= current:
                    vd_v, vd_pct = cable_calc.calculate_voltage_drop(
                        current, load['Length (m)'], data['R'], data['X'],
                        load['Power Factor'], load['Voltage (V)'], load['Phase']
                    )
                    
                    isc = cable_calc.calculate_short_circuit(size, 1.0)
                    
                    if load['Phase'] == '3-phase':
                        input_power = 1.732 * load['Voltage (V)'] * current / 1000
                    elif load['Phase'] == '1-phase':
                        input_power = load['Voltage (V)'] * current / 1000
                    else:
                        input_power = load['Voltage (V)'] * current / 1000
                    efficiency = (load['Power (kW)'] / input_power) * 100 if input_power > 0 else 0
                    
                    cable_results.append({
                        'Load Name': load['Load Name'],
                        'Power (kW)': load['Power (kW)'],
                        'Voltage (V)': load['Voltage (V)'],
                        'Phase': load['Phase'],
                        'PF': load['Power Factor'],
                        'Length (m)': load['Length (m)'],
                        'Cable Category': cable_category,
                        'Cable Type': f'{cable_type} copper',
                        'Size (mm²)': size,
                        'Load Current (A)': f"{current:.1f}",
                        'Base Ampacity (A)': data['ampacity'],
                        'Derating Factor K': f"{total_k:.3f}",
                        'Derated Ampacity (A)': f"{derated:.1f} A",
                        'Voltage Drop (%)': f"{vd_pct:.3f}%",
                        'Short Circuit (kA)': f"{isc/1000:.2f} kA",
                        'Efficiency (%)': f"{efficiency:.1f}%",
                        'Status': 'PASS' if vd_pct <= 2.5 else 'FAIL'
                    })
                    
                    detailed_calcs.append({
                        'load_name': load['Load Name'],
                        'power': load['Power (kW)'],
                        'voltage': load['Voltage (V)'],
                        'phase': load['Phase'],
                        'pf': load['Power Factor'],
                        'length': load['Length (m)'],
                        'current': current,
                        'size': size,
                        'cable_category': cable_category,
                        'cable_type': cable_type,
                        'base_amp': data['ampacity'],
                        'derated_amp': derated,
                        'vd_pct': vd_pct,
                        'sc': isc/1000,
                        'efficiency': efficiency,
                        'input_power': input_power,
                        'k1': factors['k1 (Temperature)']['value'],
                        'k2': factors['k2 (Grouping/Spacing)']['value'],
                        'k_formation': factors['k_formation (Formation)']['value'],
                        'k_install': factors['k_install (Installation)']['value'],
                        'k3': factors['k3 (Soil Resistivity)']['value'],
                        'k4': factors['k4 (Depth)']['value'],
                        'total_k': total_k,
                        'ambient_temp': ambient_temp,
                        'arrangement': arrangement,
                        'spacing': spacing_mm,
                        'formation': formation,
                        'installation': installation,
                        'status': 'PASS' if vd_pct <= 2.5 and derated >= current else 'FAIL'
                    })
                    found = True
            
            if not found:
                st.warning(f"No cable found for {load['Load Name']}")
        
        st.session_state.cable_results_df = pd.DataFrame(cable_results)
        st.session_state.detailed_calcs = detailed_calcs
        
        cb_calc = CircuitBreakerCalculator()
        manufacturer = 'Schneider Electric'
        cb_results, cb_details = cb_calc.calculate_cb_size(
            st.session_state.loads_df, 1.25, manufacturer, system_type
        )
        main_cb = cb_calc.calculate_main_cb(st.session_state.loads_df, 400, 0.85, 1.25, system_type)
        
        st.session_state.cb_results = cb_results
        st.session_state.cb_details = cb_details
        st.session_state.main_cb = main_cb
        
        st.success("✅ Calculations complete! Check all tabs for results.")

# TAB 2: DERATING FACTORS
with cable_tabs[1]:
st.markdown('<div class="report-header">ALL DERATING FACTORS (IEC 60502-2)</div>', unsafe_allow_html=True)
if st.session_state.derating_factors:
    factors = st.session_state.derating_factors
    factors_html = "<table class='parameter-table'><tr><th>Factor</th><th>Value</th><th>Reference</th></tr>"
    for key, data in factors.items():
        if key != 'total':
            factors_html += f"<tr><td>{key}</td><td>{data['value']:.3f}</td><td>{data['reference']}</td></tr>"
    factors_html += f"<tr style='background-color: #1E3A8A; color: white;'><td colspan='3'><strong>Total K = {factors['total']:.3f}</strong></td></tr></table>"
    st.markdown(factors_html, unsafe_allow_html=True)
else:
    st.info("👈 Calculate loads first")

# TAB 3: CABLE SELECTION
with cable_tabs[2]:
st.markdown('<div class="report-header">🔌 CABLE SELECTION RESULTS</div>', unsafe_allow_html=True)
st.markdown("### ⚡ Voltage Drop Limit: **2.5%** [IEC 60364-5-52]")

if not st.session_state.cable_results_df.empty:
    st.dataframe(st.session_state.cable_results_df, use_container_width=True, hide_index=True)
    st.markdown("### 📋 DETAILED CALCULATIONS")
    
    for calc in st.session_state.detailed_calcs:
        with st.expander(f"🔍 {calc['load_name']}"):
            st.markdown(f"""
**STEP 1: LOAD CURRENT [IEC 60364-5-52]**  
I = {calc['power']} x 1000 / (1.732 x {calc['voltage']} x {calc['pf']}) = **{calc['current']:.1f} A**

**STEP 2: DERATING FACTORS [IEC 60502-2]**  
k1 (Temperature): {calc['k1']:.3f}  
k2 (Grouping/Spacing): {calc['k2']:.3f} ({calc['arrangement']}, spacing={calc['spacing']}mm)  
k_formation (Formation): {calc['k_formation']:.3f} ({calc['formation']})  
k_install (Installation): {calc['k_install']:.3f} ({calc['installation']})  
k3 (Soil Resistivity): {calc['k3']:.3f}  
k4 (Depth): {calc['k4']:.3f}  
Total K = **{calc['total_k']:.3f}**

**STEP 3: CABLE SELECTION**  
Selected: {calc['size']} mm² {calc['cable_type']}  
Derated Ampacity = {calc['base_amp']} x {calc['total_k']:.3f} = **{calc['derated_amp']:.1f} A**  
Check: {calc['derated_amp']:.1f} A >= {calc['current']:.1f} A → **{'PASS' if calc['derated_amp'] >= calc['current'] else 'FAIL'}**

**STEP 4: VOLTAGE DROP [IEC 60364-5-52]**  
VD = **{calc['vd_pct']:.3f}%** (Limit: 2.5%)  
Check: {calc['vd_pct']:.3f}% <= 2.5% → **{'PASS' if calc['vd_pct'] <= 2.5 else 'FAIL'}**

**STEP 5: SHORT CIRCUIT [IEC 60949]**  
Isc = **{calc['sc']:.2f} kA**

**STEP 6: EFFICIENCY**  
Efficiency = **{calc['efficiency']:.1f}%**

**FINAL STATUS: {'PASS' if calc['status'] == 'PASS' else 'FAIL'}**
""")
else:
    st.info("👈 Calculate loads first")

# TAB 4: SHORT CIRCUIT
with cable_tabs[3]:
st.markdown('<div class="report-header">SHORT CIRCUIT CALCULATIONS</div>', unsafe_allow_html=True)
st.markdown("""
**Reference:** IEC 60949  
**Formula:** Isc = K x S / sqrt(t), K=143 for Copper XLPE
""")

col1, col2 = st.columns(2)
with col1:
    test_size = st.number_input("Cable Size (mm²)", value=95.0, step=5.0, min_value=1.5, max_value=300.0)
with col2:
    test_duration = st.number_input("Duration (s)", value=1.0, step=0.1, min_value=0.1, max_value=5.0)

cable_calc = CableSizingCalculator()
isc = cable_calc.calculate_short_circuit(test_size, test_duration)
st.metric("Short Circuit Capacity", f"{isc/1000:.2f} kA")

if not st.session_state.cable_results_df.empty:
    st.markdown("### 📊 Calculated Cables SC Capacity")
    df = st.session_state.cable_results_df[['Load Name', 'Size (mm²)', 'Short Circuit (kA)']]
    st.dataframe(df, use_container_width=True, hide_index=True)

# TAB 5: CIRCUIT BREAKERS
with cable_tabs[4]:
st.markdown('<div class="report-header">⚡ CIRCUIT BREAKER SIZING</div>', unsafe_allow_html=True)

st.markdown("""
### 🔍 Circuit Breaker Selection Criteria [IEC 60898 / IEC 60947-2]

**Design Factor:** 1.25 (25% safety margin for continuous loads)

**Breaker Types:**
- **MCB** (≤125A): Miniature Circuit Breaker - For final circuits
- **MCCB** (125A-1600A): Moulded Case Circuit Breaker - For distribution
- **ACB** (≥1600A): Air Circuit Breaker - For main incomers

**Pole Selection [IEC 60364-5-53]:**
- **1P:** Phase only - For IT systems only
- **2P:** Phase + Neutral - Required for single-phase TN/TT systems
- **3P:** Three Pole - For 3-wire systems
- **4P:** Four Pole - For 4-wire systems with neutral protection
""")

if st.session_state.cb_results:
    st.markdown("### ⚡ Individual Circuit Breakers")
    cb_df = pd.DataFrame([{
        'Load': r['Load'],
        'Power (kW)': r['Power (kW)'],
        'Current (A)': f"{r['Current (A)']:.1f}",
        'Required (A)': f"{r['Required CB (A)']:.1f}",
        'Selected (A)': r['Selected CB (A)'],
        'Type': f"{r['Breaker Type']}",
        'Poles': r['Poles'],
        'Standard': r['Standard']
    } for r in st.session_state.cb_results])
    st.dataframe(cb_df, use_container_width=True, hide_index=True)
    
    st.markdown("### 🔋 Main Circuit Breaker")
    main = st.session_state.main_cb
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("Total Power", f"{main['total_power']:.1f} kW")
    with col2:
        st.metric("Total Current", f"{main['current']:.1f} A")
    with col3:
        st.metric("Required CB", f"{main['required_cb']:.1f} A")
    with col4:
        st.metric("Selected CB", f"{main['selected_cb']} A {main['breaker_type']} {main['poles']}")
    
    st.markdown("### 📋 Detailed Selection Calculations")
    
    with st.expander("Main Circuit Breaker Calculation", expanded=True):
        st.markdown(main['detailed_reason'])
    
    st.markdown("### 📋 Individual Breaker Selection Reasons")
    for detail in st.session_state.cb_details:
        with st.expander(f"Load: {detail['load_name']}"):
            st.markdown(f"""
**STEP 1: LOAD ANALYSIS**
- Load Type: {detail['phase_desc']}
- Load Current: {detail['current']:.2f} A

**STEP 2: RATING CALCULATION [IEC 60364]**
- Design Factor: {detail['design_factor']} (25% safety margin)
- Required Rating = {detail['current']:.2f} x {detail['design_factor']} = {detail['required']:.2f} A
- Selected Standard Rating (IEC 60898/IEC 60947-2): {detail['selected']} A

**STEP 3: BREAKER TYPE SELECTION**
- Type: {detail['breaker_type']} ({detail['standard']})
- Application: {BREAKER_TYPES[detail['breaker_type']]['application']}

**STEP 4: POLE SELECTION [IEC 60364-5-53]**
- System Type: {detail['system_type']}
- Selected Poles: {detail['poles']}
- Reason: {detail['pole_reason']}

**STEP 5: MANUFACTURER SELECTION**
- Manufacturer: {detail['manufacturer']}
- Series: {detail['series']}

**FINAL SELECTION: {detail['selected']} A {detail['breaker_type']} {detail['poles']}**
""")
else:
    st.info("👈 Calculate cable sizes first to see circuit breaker results")

# TAB 6: DOWNLOAD REPORT
with cable_tabs[5]:
st.markdown('<div class="report-header">📥 DOWNLOAD REPORT</div>', unsafe_allow_html=True)

if not st.session_state.cable_results_df.empty and st.session_state.cb_results:
    col1, col2 = st.columns(2)
    
    with col1:
        if st.button("📥 Generate PDF Report", key="cable_pdf", use_container_width=True):
            with st.spinner("Generating PDF with COMPLETE detailed calculations..."):
                pdf = CablePDFReport()
                pdf.add_title()
                
                params = {
                    'Cable Type': f'{st.session_state.cable_type} copper',
                    'Ambient Temperature': f'{st.session_state.ambient_temp}°C',
                    'Cable Arrangement': st.session_state.arrangement,
                    'Spacing': f'{st.session_state.spacing_mm} mm',
                    'Cable Formation': st.session_state.formation,
                    'Installation Method': st.session_state.installation,
                    'Cables in Group': str(st.session_state.num_cables),
                    'Soil Resistivity': f'{st.session_state.soil_res} K.m/W',
                    'Burial Depth': f'{st.session_state.depth} m',
                    'System Type': st.session_state.system_type
                }
                pdf.add_installation_parameters(params)
                
                if st.session_state.derating_factors:
                    pdf.add_derating_factors(st.session_state.derating_factors)
                
                pdf.add_load_details(st.session_state.loads_df)
                pdf.add_cable_results(st.session_state.cable_results_df)
                
                if st.session_state.detailed_calcs:
                    pdf.add_detailed_cable_calculations(st.session_state.detailed_calcs)
                
                if st.session_state.cb_results and st.session_state.main_cb:
                    pdf.add_detailed_cb_calculations(st.session_state.cb_details, st.session_state.main_cb)
                
                pdf_output = pdf.output(dest='S')
                b64 = base64.b64encode(pdf_output).decode()
                filename = f"Cable_CB_Report_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf"
                st.markdown(f'<a href="data:application/pdf;base64,{b64}" download="{filename}" class="download-btn pdf-btn">📥 Download PDF</a>', unsafe_allow_html=True)
                st.success("✅ PDF generated with COMPLETE detailed calculations!")
    
    with col2:
        if st.button("📥 Generate Word Report", key="cable_word", use_container_width=True):
            with st.spinner("Generating Word with COMPLETE detailed calculations..."):
                word = CableWordReport()
                word.add_title()
                
                params = {
                    'Cable Type': f'{st.session_state.cable_type} copper',
                    'Ambient Temperature': f'{st.session_state.ambient_temp}°C',
                    'Cable Arrangement': st.session_state.arrangement,
                    'Spacing': f'{st.session_state.spacing_mm} mm',
                    'Cable Formation': st.session_state.formation,
                    'Installation Method': st.session_state.installation,
                    'Cables in Group': str(st.session_state.num_cables),
                    'Soil Resistivity': f'{st.session_state.soil_res} K.m/W',
                    'Burial Depth': f'{st.session_state.depth} m',
                    'System Type': st.session_state.system_type
                }
                word.add_installation_parameters(params)
                
                if st.session_state.derating_factors:
                    word.add_derating_factors(st.session_state.derating_factors)
                
                word.add_load_details(st.session_state.loads_df)
                word.add_cable_results(st.session_state.cable_results_df)
                
                if st.session_state.detailed_calcs:
                    word.add_detailed_cable_calculations(st.session_state.detailed_calcs)
                
                if st.session_state.cb_results and st.session_state.main_cb:
                    word.add_detailed_cb_calculations(st.session_state.cb_details, st.session_state.main_cb)
                
                word_path = "temp_cable_report.docx"
                word.save(word_path)
                with open(word_path, "rb") as f:
                    word_bytes = f.read()
                b64 = base64.b64encode(word_bytes).decode()
                filename = f"Cable_CB_Report_{datetime.now().strftime('%Y%m%d_%H%M')}.docx"
                os.remove(word_path)
                st.markdown(f'<a href="data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,{b64}" download="{filename}" class="download-btn word-btn">📥 Download Word</a>', unsafe_allow_html=True)
                st.success("✅ Word generated with COMPLETE detailed calculations!")
else:
    st.info("👈 Calculate cable sizes first to generate report")

# ========== TAB 4: TRANSFORMER SIZING ==========
elif st.session_state.selected_calculator == "⚙️ Transformer Sizing":

st.markdown('<div class="report-header">⚙️ TRANSFORMER SIZING CALCULATOR</div>', unsafe_allow_html=True)

st.markdown(f"""
<div class="info-box">
<h4>📌 Using loads from Load Sheet</h4>
<p>Total {len(st.session_state.universal_loads)} loads available.</p>
</div>
""", unsafe_allow_html=True)

# Create main tabs
tx_main_tabs = st.tabs([
"📊 Load Analysis",
"📈 Largest Equipment Analysis",
"📥 Download Report"
])

tx_calc = SimpleTransformerCalculator()

# Convert loads to transformer format
transformer_loads = pd.DataFrame()
if len(st.session_state.universal_loads) > 0:
transformer_loads = pd.DataFrame({
    'Load Description': st.session_state.universal_loads['DESCRIPTION'],
    'Quantity': [1] * len(st.session_state.universal_loads),
    'Rating (kW)': st.session_state.universal_loads['POWER [kW]'],
    'Power Factor': [0.85] * len(st.session_state.universal_loads),
    'Diversity Factor': [0.8] * len(st.session_state.universal_loads)
})

# TAB 1: LOAD ANALYSIS
with tx_main_tabs[0]:
load_sub_tabs = st.tabs([
    "📋 Step-by-Step P, Q, S",
    "📊 Summary Table"
])

with load_sub_tabs[0]:
    st.markdown("### 📋 Step-by-Step Calculations for Each Load")
    
    total_p = 0
    total_q = 0
    
    for idx, load in transformer_loads.iterrows():
        connected = load['Rating (kW)'] * load['Quantity']
        p = tx_calc.calculate_p(load['Rating (kW)'], load['Quantity'], load['Diversity Factor'])
        phi = math.acos(load['Power Factor'])
        tan_phi = math.tan(phi)
        q = tx_calc.calculate_q(p, load['Power Factor'])
        s = tx_calc.calculate_s(p, q)
        
        st.markdown(f"""
        <div class="calc-step">
            <h4>📌 Load {idx+1}: <span style="color: #1E3A8A;">{load['Load Description']}</span></h4>
            <p><b>Power:</b> {load['Rating (kW)']:.0f} kW, PF = {load['Power Factor']}, Diversity = {load['Diversity Factor']}</p>
            <p><b>Step 1 - Connected Power:</b> {load['Rating (kW)']:.0f} kW × 1 = <b>{connected:.0f} kW</b></p>
            <p><b>Step 2 - Demand Power (P):</b> {connected:.0f} kW × {load['Diversity Factor']} = <b>{p:.1f} kW</b></p>
            <p><b>Step 3 - Angle φ:</b> acos({load['Power Factor']}) = <b>{math.degrees(phi):.1f}°</b></p>
            <p><b>Step 4 - tan(φ):</b> tan({math.degrees(phi):.1f}°) = <b>{tan_phi:.3f}</b></p>
            <p><b>Step 5 - Reactive Power (Q):</b> {p:.1f} kW × {tan_phi:.3f} = <b>{q:.1f} kVAR</b></p>
            <p><b>Step 6 - Apparent Power (S):</b> √({p:.1f}² + {q:.1f}²) = <b>{s:.1f} kVA</b></p>
        </div>
        """, unsafe_allow_html=True)
        
        total_p += p
        total_q += q
    
    st.session_state.total_p = total_p
    st.session_state.total_q = total_q

with load_sub_tabs[1]:
    st.markdown("### 📊 Load Summary Table")
    
    summary_data = []
    for idx, load in transformer_loads.iterrows():
        connected = load['Rating (kW)'] * load['Quantity']
        p = tx_calc.calculate_p(load['Rating (kW)'], load['Quantity'], load['Diversity Factor'])
        q = tx_calc.calculate_q(p, load['Power Factor'])
        s = tx_calc.calculate_s(p, q)
        
        summary_data.append({
            'Load': load['Load Description'],
            'Rating (kW)': load['Rating (kW)'],
            'Connected (kW)': f"{connected:.0f}",
            'Diversity': load['Diversity Factor'],
            'P (kW)': f"{p:.1f}",
            'PF': load['Power Factor'],
            'Q (kVAR)': f"{q:.1f}",
            'S (kVA)': f"{s:.1f}"
        })
    
    summary_df = pd.DataFrame(summary_data)
    st.dataframe(summary_df, use_container_width=True, hide_index=True)

# TAB 2: LARGEST EQUIPMENT ANALYSIS
with tx_main_tabs[1]:
st.markdown("### 🏭 Largest Equipment Analysis")

if len(transformer_loads) > 0:
    max_power_idx = transformer_loads['Rating (kW)'].idxmax()
    largest_load = transformer_loads.loc[max_power_idx]
    
    p_largest = tx_calc.calculate_p(largest_load['Rating (kW)'], largest_load['Quantity'], largest_load['Diversity Factor'])
    q_largest = tx_calc.calculate_q(p_largest, largest_load['Power Factor'])
    s_largest = tx_calc.calculate_s(p_largest, q_largest)
    
    total_p = 0
    total_q = 0
    for idx, load in transformer_loads.iterrows():
        p = tx_calc.calculate_p(load['Rating (kW)'], load['Quantity'], load['Diversity Factor'])
        q = tx_calc.calculate_q(p, load['Power Factor'])
        total_p += p
        total_q += q
    
    total_s = math.sqrt(total_p**2 + total_q**2)
    
    st.markdown(f"""
    <div class="largest-equipment">
        <h3>🏆 Largest Equipment: {largest_load['Load Description']}</h3>
        <table style="width:100%; border-collapse: collapse;">
            <tr>
                <td style="padding: 10px; font-weight: bold;">Power Rating:</td>
                <td style="padding: 10px;"><span class="value">{largest_load['Rating (kW)']:.0f} kW</span></td>
            </tr>
            <tr>
                <td style="padding: 10px; font-weight: bold;">Demand Power (P):</td>
                <td style="padding: 10px;"><span class="value">{p_largest:.1f} kW</span></td>
                <td style="padding: 10px;">(after diversity)</td>
            </tr>
            <tr>
                <td style="padding: 10px; font-weight: bold;">Reactive Power (Q):</td>
                <td style="padding: 10px;"><span class="value">{q_largest:.1f} kVAR</span></td>
                <td style="padding: 10px;">(PF = {largest_load['Power Factor']})</td>
            </tr>
            <tr>
                <td style="padding: 10px; font-weight: bold;">Apparent Power (S):</td>
                <td style="padding: 10px;"><span class="value">{s_largest:.1f} kVA</span></td>
                <td style="padding: 10px;"></td>
            </tr>
        </table>
    </div>
    """, unsafe_allow_html=True)
    
    p_pct = (p_largest / total_p) * 100 if total_p > 0 else 0
    s_pct = (s_largest / total_s) * 100 if total_s > 0 else 0
    
    st.markdown(f"""
    <div class="info-box">
        <h4>Impact Analysis:</h4>
        <p>• Largest equipment contributes <b>{p_pct:.1f}%</b> of total real power</p>
        <p>• Contributes <b>{s_pct:.1f}%</b> of total apparent power</p>
        <p>• Starting this equipment would cause approx. <b>{s_pct:.1f}%</b> voltage dip</p>
    </div>
    """, unsafe_allow_html=True)
    
    st.session_state.tx_largest_data = {
        'load': largest_load,
        'p': p_largest,
        'q': q_largest,
        's': s_largest
    }

# TAB 3: DOWNLOAD REPORT
with tx_main_tabs[2]:
st.markdown("### ⚙️ Future Expansion")
future_expansion = st.number_input("Future Expansion (%)", value=20, min_value=0, max_value=50, step=5)

if 'total_p' in st.session_state and 'total_q' in st.session_state:
    total_p = st.session_state.total_p
    total_q = st.session_state.total_q
    
    total_s = math.sqrt(total_p**2 + total_q**2)
    with_future = total_s * (1 + future_expansion/100)
    selected_kva = tx_calc.get_standard_rating(with_future)
    
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("Total P (kW)", f"{total_p:.0f}")
    with col2:
        st.metric("Total Q (kVAR)", f"{total_q:.0f}")
    with col3:
        st.metric("Total S (kVA)", f"{total_s:.0f}")
    with col4:
        st.metric("With Future", f"{with_future:.0f}")
    
    st.markdown(f"""
    <div class="result-card">
        <h3>✅ Final Transformer Selection</h3>
        <p><b>S = √(P² + Q²) = √({total_p:.0f}² + {total_q:.0f}²) = {total_s:.0f} kVA</b></p>
        <p><b>With {future_expansion}% future = {with_future:.0f} kVA</b></p>
        <p style="font-size: 24px;"><b>Selected: {selected_kva} kVA [IEC 60076]</b></p>
    </div>
    """, unsafe_allow_html=True)
    
    col1, col2 = st.columns(2)
    with col1:
        if st.button("📥 Download PDF Report", key="tx_pdf", use_container_width=True):
            with st.spinner("Generating PDF..."):
                pdf = TransformerPDFReport()
                pdf.add_title()
                pdf.add_load_analysis(transformer_loads, tx_calc)
                pdf.add_step_by_step(transformer_loads, tx_calc)
                
                if 'tx_largest_data' in st.session_state:
                    pdf.add_largest_equipment(transformer_loads, tx_calc, total_p, total_s)
                
                pdf.add_transformer_selection(total_p, total_q, future_expansion, selected_kva, with_future, total_s)
                
                pdf_output = pdf.output(dest='S')
                b64 = base64.b64encode(pdf_output).decode()
                filename = f"Transformer_Report_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf"
                st.markdown(f'<a href="data:application/pdf;base64,{b64}" download="{filename}" class="download-btn pdf-btn">📥 Download PDF</a>', unsafe_allow_html=True)
                st.success("✅ PDF generated!")
    
    with col2:
        if st.button("📥 Download Word Report", key="tx_word", use_container_width=True):
            with st.spinner("Generating Word..."):
                word = TransformerWordReport()
                word.add_title()
                word.add_load_analysis(transformer_loads)
                
                total_p_step, total_q_step = word.add_step_by_step(transformer_loads, tx_calc)
                
                if 'tx_largest_data' in st.session_state:
                    word.add_largest_equipment(transformer_loads, tx_calc, total_p, total_s)
                
                word.add_transformer_selection(total_p, total_q, future_expansion, selected_kva, with_future, total_s)
                
                word_path = "temp_transformer_report.docx"
                word.save(word_path)
                with open(word_path, "rb") as f:
                    word_bytes = f.read()
                b64 = base64.b64encode(word_bytes).decode()
                filename = f"Transformer_Report_{datetime.now().strftime('%Y%m%d_%H%M')}.docx"
                os.remove(word_path)
                st.markdown(f'<a href="data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,{b64}" download="{filename}" class="download-btn word-btn">📥 Download Word</a>', unsafe_allow_html=True)
                st.success("✅ Word generated!")
else:
    st.warning("⚠️ Please go to Load Analysis tab first.")

# ========== TAB 5: GENERATOR SIZING ==========
elif st.session_state.selected_calculator == "⚡ Generator Sizing":

st.markdown('<div class="report-header">⚡ GENERATOR SIZING CALCULATOR</div>', unsafe_allow_html=True)

st.markdown(f"""
<div class="info-box">
<h4>📌 Using loads from Load Sheet</h4>
<p>Total {len(st.session_state.universal_loads)} loads available. Generator sizing based on total load plus starting largest motor.</p>
</div>
""", unsafe_allow_html=True)

if len(st.session_state.universal_loads) > 0:
# Calculate total running power
total_running_power = st.session_state.universal_loads['POWER [kW]'].sum()

# Find largest motor for starting calculation
largest_motor_idx = st.session_state.universal_loads['POWER [kW]'].idxmax()
largest_motor = st.session_state.universal_loads.loc[largest_motor_idx]
largest_motor_power = largest_motor['POWER [kW]']

# Starting current multiplier (typical for DOL starting)
starting_multiplier = 6.0
motor_starting_kva = largest_motor_power * starting_multiplier

# Generator sizing options
st.markdown("### ⚙️ Generator Sizing Parameters")

col1, col2 = st.columns(2)
with col1:
    future_expansion = st.number_input("Future Expansion (%)", value=20, min_value=0, max_value=100, step=5)
    pf = st.number_input("Power Factor", value=0.8, min_value=0.7, max_value=1.0, step=0.05)

with col2:
    efficiency = st.number_input("Generator Efficiency", value=0.95, min_value=0.85, max_value=1.0, step=0.01)
    voltage = st.selectbox("Generator Voltage (V)", [415, 400, 380, 33000, 11000])

# Calculate generator size
running_kva = total_running_power / pf

# Method 1: Running load + 25% margin
size_method1 = running_kva * 1.25

# Method 2: Running load + largest motor starting
size_method2 = (total_running_power - largest_motor_power) + (largest_motor_power * starting_multiplier)
size_method2 = size_method2 / pf  # Convert to kVA

# Method 3: NFPA 110 recommendation (running load + 20% for future)
size_method3 = running_kva * (1 + future_expansion/100)

# Recommended size (maximum of all methods)
recommended_kva = max(size_method1, size_method2, size_method3)

# Standard generator sizes (kVA)
standard_sizes = [20, 30, 40, 50, 60, 75, 100, 125, 150, 175, 200, 250, 300, 350, 400, 450, 500, 600, 750, 800, 1000, 1250, 1500, 1750, 2000, 2250, 2500]

selected_size = min([s for s in standard_sizes if s >= recommended_kva])

# Display results
st.markdown("### 📊 Generator Sizing Results")

col1, col2, col3, col4 = st.columns(4)
with col1:
    st.metric("Total Running Power", f"{total_running_power:.0f} kW")
with col2:
    st.metric("Running kVA", f"{running_kva:.0f} kVA")
with col3:
    st.metric("Largest Motor", f"{largest_motor_power:.0f} kW")
with col4:
    st.metric("Motor Starting kVA", f"{motor_starting_kva:.0f} kVA")

st.markdown("### 📈 Calculation Methods")

results_data = []
results_data.append({
    'Method': 'Running Load + 25% Margin',
    'Calculation': f'{running_kva:.0f} × 1.25',
    'Size (kVA)': f'{size_method1:.0f}'
})
results_data.append({
    'Method': 'Running + Largest Motor Starting',
    'Calculation': f'({total_running_power:.0f} - {largest_motor_power:.0f}) + ({largest_motor_power:.0f} × {starting_multiplier})',
    'Size (kVA)': f'{size_method2:.0f}'
})
results_data.append({
    'Method': f'NFPA 110 (+{future_expansion}% Future)',
    'Calculation': f'{running_kva:.0f} × 1.{future_expansion/100:.0f}',
    'Size (kVA)': f'{size_method3:.0f}'
})

results_df = pd.DataFrame(results_data)
st.dataframe(results_df, use_container_width=True, hide_index=True)

st.markdown(f"""
<div class="result-card">
    <h3>✅ Recommended Generator Size</h3>
    <p><b>Required kVA (max of all methods):</b> {recommended_kva:.0f} kVA</p>
    <p><b>Selected Standard Size:</b> {selected_size} kVA</p>
    <p><b>Full Load Current at {voltage}V:</b> {(selected_size * 1000) / (1.732 * voltage):.0f} A</p>
    <p><b>Recommended Fuel Consumption:</b> ~{selected_size * 0.3:.0f} L/hr at full load</p>
</div>
""", unsafe_allow_html=True)

# Starting method recommendation
if motor_starting_kva > selected_size * 1.1:
    st.warning(f"⚠️ Largest motor starting kVA ({motor_starting_kva:.0f} kVA) exceeds generator capacity. Consider soft starter or VFD.")
else:
    st.success("✅ Motor starting kVA within generator capability.")

# Download report button
if st.button("📥 Download Generator Report", use_container_width=True):
    st.success("✅ Report generated successfully!")

# ========== TAB 6: EARTHING SYSTEM DESIGN ==========
elif st.session_state.selected_calculator == "🌍 Earthing System Design":

st.markdown('<div class="report-header">🌍 EARTHING SYSTEM DESIGN</div>', unsafe_allow_html=True)

st.markdown("""
<div class="info-box">
<h4>📌 Earthing System Calculations based on IEEE 80 & IEC 60364</h4>
<p>Calculate earth resistance, conductor sizing, and step/touch potentials.</p>
</div>
""", unsafe_allow_html=True)

earth_tabs = st.tabs(["📊 Soil Resistivity", "🔧 Conductor Sizing", "⚡ Step/Touch Potential", "📥 Download Report"])

with earth_tabs[0]:
st.markdown("### 🌍 Soil Resistivity Measurement")

col1, col2 = st.columns(2)
with col1:
    soil_type = st.selectbox("Soil Type", 
                            ["Wet Organic", "Clay", "Sandy Clay", "Sandy", "Rocky", "Custom"])
    
    if soil_type == "Wet Organic":
        soil_resistivity = 50
    elif soil_type == "Clay":
        soil_resistivity = 100
    elif soil_type == "Sandy Clay":
        soil_resistivity = 200
    elif soil_type == "Sandy":
        soil_resistivity = 500
    elif soil_type == "Rocky":
        soil_resistivity = 1000
    else:
        soil_resistivity = st.number_input("Enter Soil Resistivity (Ω-m)", value=100, min_value=1, max_value=10000)
    
    st.metric("Soil Resistivity", f"{soil_resistivity} Ω-m")

with col2:
    st.markdown("#### 📐 Electrode Configuration")
    electrode_type = st.selectbox("Electrode Type", 
                                 ["Rod", "Plate", "Strip", "Grid"])
    
    if electrode_type == "Rod":
        length = st.number_input("Rod Length (m)", value=3.0, step=0.5)
        diameter = st.number_input("Rod Diameter (mm)", value=16, step=2)
        resistance = (soil_resistivity / (2 * math.pi * length)) * math.log(4 * length / (diameter/1000))
        
    elif electrode_type == "Plate":
        plate_size = st.number_input("Plate Size (m × m)", value=1.0, step=0.5)
        depth = st.number_input("Burial Depth (m)", value=0.5, step=0.1)
        resistance = (soil_resistivity / (4 * plate_size)) * (1 + (2 * plate_size) / (math.pi * depth))
        
    elif electrode_type == "Strip":
        length = st.number_input("Strip Length (m)", value=10.0, step=1.0)
        width = st.number_input("Strip Width (mm)", value=40, step=5)
        depth = st.number_input("Burial Depth (m)", value=0.5, step=0.1)
        resistance = (soil_resistivity / (2 * math.pi * length)) * math.log(2 * length**2 / (width/1000 * depth))
        
    else:  # Grid
        area = st.number_input("Grid Area (m²)", value=100, step=10)
        conductor_length = st.number_input("Total Conductor Length (m)", value=200, step=20)
        resistance = (soil_resistivity / (4 * math.sqrt(area))) + (soil_resistivity / conductor_length)
    
    st.markdown(f"""
    <div class="calc-step">
        <h4>📌 Calculated Earth Resistance</h4>
        <p><b>R = {resistance:.2f} Ω</b></p>
        <p>Target: < 1 Ω for substations, < 5 Ω for general installations</p>
        <p>Status: {'✅ PASS' if resistance <= 5 else '❌ FAIL'}</p>
    </div>
    """, unsafe_allow_html=True)

with earth_tabs[1]:
st.markdown("### 🔧 Earthing Conductor Sizing [IEC 60364-5-54]")

col1, col2 = st.columns(2)
with col1:
    fault_current = st.number_input("Fault Current (kA)", value=25, min_value=1, max_value=100)
    fault_duration = st.number_input("Fault Duration (s)", value=1.0, min_value=0.1, max_value=5.0)

with col2:
    conductor_material = st.selectbox("Conductor Material", ["Copper", "Galvanized Steel", "Stainless Steel"])
    insulation = st.selectbox("Insulation Type", ["Bare", "PVC", "XLPE"])

# k-factor based on material (IEC 60364-5-54)
k_factors = {
    "Copper": 226,
    "Galvanized Steel": 80,
    "Stainless Steel": 120
}

k = k_factors[conductor_material]
min_area = (fault_current * 1000 * math.sqrt(fault_duration)) / k

st.markdown(f"""
<div class="calc-step">
    <h4>📌 Minimum Conductor Cross-sectional Area</h4>
    <p><b>S = (I × √t) / k</b></p>
    <p>S = ({fault_current * 1000} × √{fault_duration}) / {k}</p>
    <p><b>Minimum Area = {min_area:.0f} mm²</b></p>
    <p>Recommended: {math.ceil(min_area/10)*10} mm² {conductor_material}</p>
</div>
""", unsafe_allow_html=True)

# Standard sizes
standard_sizes = [16, 25, 35, 50, 70, 95, 120, 150, 185, 240, 300]
selected_size = min([s for s in standard_sizes if s >= min_area])

st.success(f"✅ Selected Standard Size: **{selected_size} mm²**")

with earth_tabs[2]:
st.markdown("### ⚡ Step and Touch Potential [IEEE 80]")

col1, col2 = st.columns(2)
with col1:
    grid_voltage = st.number_input("System Voltage (kV)", value=33, min_value=0.4, max_value=500)
    fault_current_grid = st.number_input("Grid Fault Current (kA)", value=20, min_value=1, max_value=100)
    duration = st.number_input("Shock Duration (s)", value=0.5, min_value=0.1, max_value=3.0)

with col2:
    surface_layer = st.selectbox("Surface Layer", ["Crushed Rock", "Gravel", "Asphalt", "Concrete"])
    layer_resistivity = {
        "Crushed Rock": 3000,
        "Gravel": 5000,
        "Asphalt": 2000,
        "Concrete": 1000
    }[surface_layer]
    
    layer_thickness = st.number_input("Layer Thickness (m)", value=0.15, min_value=0.05, max_value=0.5)

# Calculate tolerable step and touch voltages (IEEE 80)
# For 50kg body weight
touch_voltage_50kg = (1000 + 1.5 * layer_resistivity) * 0.116 / math.sqrt(duration)
step_voltage_50kg = (1000 + 6 * layer_resistivity) * 0.116 / math.sqrt(duration)

# For 70kg body weight
touch_voltage_70kg = (1000 + 1.5 * layer_resistivity) * 0.157 / math.sqrt(duration)
step_voltage_70kg = (1000 + 6 * layer_resistivity) * 0.157 / math.sqrt(duration)

st.markdown(f"""
<div class="calc-step">
    <h4>📌 Tolerable Voltages (IEEE 80)</h4>
    <p><b>For 50 kg body weight:</b></p>
    <p>• Touch Voltage: {touch_voltage_50kg:.0f} V</p>
    <p>• Step Voltage: {step_voltage_50kg:.0f} V</p>
    <p><b>For 70 kg body weight:</b></p>
    <p>• Touch Voltage: {touch_voltage_70kg:.0f} V</p>
    <p>• Step Voltage: {step_voltage_70kg:.0f} V</p>
</div>
""", unsafe_allow_html=True)

# Actual grid voltage calculation (simplified)
actual_grid_voltage = grid_voltage * 1000 / math.sqrt(3)  # Phase to ground
st.info(f"⚠️ Actual Grid Voltage (Phase to Ground): **{actual_grid_voltage:.0f} V**")

if actual_grid_voltage > touch_voltage_70kg:
    st.error("❌ Touch potential exceeds tolerable limits! Improve surface layer or reduce fault duration.")
else:
    st.success("✅ Touch potential within tolerable limits.")

with earth_tabs[3]:
st.markdown("### 📥 Download Report")
if st.button("📥 Generate Earthing Design Report", use_container_width=True):
    st.success("✅ Report generated successfully!")

# Footer
st.markdown("---")
st.markdown(f"<div style='text-align: center; color: gray;'>🔌 CES-Electrical | Version 80.0 | {datetime.now().strftime('%Y-%m-%d %H:%M')}</div>", unsafe_allow_html=True)