import streamlit as st
import math
import datetime
import pandas as pd
import base64
from PIL import Image
import io
from datetime import datetime
from fpdf import FPDF
import os

st.set_page_config(page_title="Professional Engineering Tools", page_icon="⚡", layout="wide")

# Custom CSS
st.markdown("""
<style>
    .report-header {
        background-color: #1E3A8A;
        color: white;
        padding: 20px;
        border-radius: 10px;
        text-align: center;
        margin-bottom: 20px;
    }
    .formula-box {
        background-color: #F3F4F6;
        padding: 15px;
        border-radius: 8px;
        border-left: 5px solid #1E3A8A;
        margin: 10px 0;
        font-family: 'Courier New', monospace;
    }
    .sidebar-nav {
        padding: 10px;
        margin-bottom: 20px;
    }
    .nav-item {
        padding: 10px;
        margin: 5px 0;
        border-radius: 5px;
        cursor: pointer;
        transition: all 0.3s;
    }
    .nav-item:hover {
        background-color: #f0f2f6;
    }
    .nav-item-active {
        background-color: #1E3A8A;
        color: white;
        font-weight: bold;
    }
    .nav-item-active:hover {
        background-color: #1E3A8A;
    }
</style>
""", unsafe_allow_html=True)

# Initialize session state
if 'company_logo' not in st.session_state:
    st.session_state.company_logo = None
if 'contractor_logo' not in st.session_state:
    st.session_state.contractor_logo = None
if 'calc_results' not in st.session_state:
    st.session_state.calc_results = {}
if 'calc_done' not in st.session_state:
    st.session_state.calc_done = False
if 'selected_calculator' not in st.session_state:
    st.session_state.selected_calculator = " Lightning Protection"
if 'cover_details' not in st.session_state:
    st.session_state.cover_details = {
        'title': 'PLANT LIGHTNING CALCULATION',
        'revision': 'A',
        'date': '02 Sep 2025',
        'purpose': 'ISSUED FOR APPROVAL',
        'prepared_by': '',
        'reviewed_by': '',
        'approved_by': ''
    }
if 'project_info' not in st.session_state:
    st.session_state.project_info = {
        'company': 'COMPANY',
        'contractor': 'CONTRACTOR',
        'project_title': 'BASIC AND DETAIL ENGINEERING DESIGN SERVICES FOR\n70,000 BPD CDU & LPG UNIT FOR MAYSAN REFINERY',
        'document_number': 'XXXX-XXX-XXXX-XX-XXX-XXXX',
        'project_number': 'B049'
    }
if 'revision_history' not in st.session_state:
    st.session_state.revision_history = [
        {'rev': 'A', 'date': '02-Sep-2025', 'purpose': 'ISSUED FOR APPROVAL', 'prpd': '', 'revd': '', 'appd': ''}
    ]

# ========== SIDEBAR WITH NAVIGATION ==========
with st.sidebar:
    st.markdown("### Lightning Protection Systems")
    st.markdown("---")
    
    # Calculator Navigation
    st.markdown("### 📌 Select Calculator")
    
    calculators = [
        "⚡ Lightning Protection",
        "🔌 Cable Sizing",
        "⚙️ Transformer Sizing",
        "📊 Load Flow Analysis",
        "🔧 Short Circuit",
        "📈 Voltage Drop"
    ]
    
    for calc in calculators:
        if st.button(calc, key=f"nav_{calc}", use_container_width=True):
            st.session_state.selected_calculator = calc
            st.rerun()
    
    st.markdown("---")
    
    # Common Project Information (shared across all calculators)
    st.markdown("### 📋 Project Information")
    
    st.session_state.project_info['company'] = st.text_input("Company Name", st.session_state.project_info['company'])
    st.session_state.project_info['contractor'] = st.text_input("Contractor Name", st.session_state.project_info['contractor'])
    st.session_state.project_info['project_title'] = st.text_area("Project Title", st.session_state.project_info['project_title'], height=60)
    st.session_state.project_info['document_number'] = st.text_input("Document Number", st.session_state.project_info['document_number'])
    st.session_state.project_info['project_number'] = st.text_input("Project Number", st.session_state.project_info['project_number'])
    
    st.markdown("---")
    
    # Logo Uploads (shared across all calculators)
    st.markdown("### 🏢 Company Logo")
    company_logo = st.file_uploader("Upload Company Logo", type=['png', 'jpg', 'jpeg'], key="company")
    if company_logo is not None:
        st.session_state.company_logo = Image.open(io.BytesIO(company_logo.getvalue()))
        st.image(st.session_state.company_logo, width=100)
    
    st.markdown("### 🏭 Contractor Logo")
    contractor_logo = st.file_uploader("Upload Contractor Logo", type=['png', 'jpg', 'jpeg'], key="contractor")
    if contractor_logo is not None:
        st.session_state.contractor_logo = Image.open(io.BytesIO(contractor_logo.getvalue()))
        st.image(st.session_state.contractor_logo, width=100)

# ========== MAIN CONTENT AREA ==========
st.title(f"⚡ {st.session_state.selected_calculator} Calculator")

# Show different calculators based on selection
if st.session_state.selected_calculator == " Lightning Protection":
    # ========== LIGHTNING PROTECTION CALCULATOR (EXISTING) ==========
    
    # Create tabs for lightning protection
    lp_tabs = st.tabs([
        "🏢 Title Page", 
        "📊 Risk Assessment", 
        "🔧 Protection Design", 
        "📋 Calculations",
        "📝 Revision History",
        "📥 PDF Report"
    ])
    
    # ========== TAB 1: TITLE PAGE ==========
    with lp_tabs[0]:
        st.markdown('<div class="report-header">', unsafe_allow_html=True)
        st.markdown("## TITLE PAGE DESIGN")
        st.markdown('</div>', unsafe_allow_html=True)
        
        col1, col2 = st.columns(2)
        with col1:
            st.markdown("### 📝 Edit Title Page")
            st.session_state.cover_details['title'] = st.text_input("Document Title", st.session_state.cover_details['title'])
            st.session_state.cover_details['revision'] = st.text_input("Revision", st.session_state.cover_details['revision'])
            st.session_state.cover_details['date'] = st.text_input("Date", st.session_state.cover_details['date'])
        
        with col2:
            st.markdown("### 📄 Logos Preview")
            if st.session_state.company_logo:
                st.image(st.session_state.company_logo, width=100, caption="Company Logo")
            if st.session_state.contractor_logo:
                st.image(st.session_state.contractor_logo, width=100, caption="Contractor Logo")
    
    # ========== TAB 2: RISK ASSESSMENT ==========
    with lp_tabs[1]:
        st.markdown('<div class="report-header">', unsafe_allow_html=True)
        st.markdown("## RISK ASSESSMENT (IEC 62305-2)")
        st.markdown('</div>', unsafe_allow_html=True)
        
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
            st.markdown("### 📊 Coefficients")
            cd_values = {"Surrounded": 0.25, "Similar height": 0.5, "Isolated": 1, "Hilltop": 2}
            cd = cd_values[environment]
            st.info(f"**CD = {cd}** (IEC 62305-2 Table A.1)")
            
            if structure_type == "Column 4-C01":
                c2, c3, c4, c5 = 0.5, 2.0, 3.0, 10.0
                st.caption("Column coefficients applied")
            else:
                c2, c3, c4, c5 = 1.0, 3.0, 1.0, 5.0
                st.caption("Building coefficients applied")
            
            st.metric("C2 - Type Coefficient", c2)
            st.metric("C3 - Content Coefficient", c3)
            st.metric("C4 - Occupancy Coefficient", c4)
            st.metric("C5 - Consequence Coefficient", c5)
        
        if st.button("🔧 CALCULATE RISK", type="primary", use_container_width=True):
            
            if structure_type == "Column 4-C01":
                ad = math.pi * 9 * height**2
            else:
                ad = length * width + 2 * (3 * height) * (length + width) + math.pi * (3 * height)**2
            
            ng = 0.1 * td_days
            nd = ng * ad * cd * 1e-6
            
            c_total = cd * c2 * c3 * c4 * c5
            nc = 1e-4 / c_total
            efficiency = 1 - (nc / nd) if nd > 0 else 0
            
            if efficiency > 0.98:
                lpl, lpl_desc, sphere = "Class I", "Maximum Protection", 20
            elif efficiency > 0.95:
                lpl, lpl_desc, sphere = "Class II", "High Protection", 30
            elif efficiency > 0.90:
                lpl, lpl_desc, sphere = "Class III", "Standard Protection", 45
            else:
                lpl, lpl_desc, sphere = "Class IV", "Basic Protection", 60
            
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
            st.subheader("📊 Risk Assessment Results")
            
            col_a, col_b, col_c = st.columns(3)
            with col_a:
                st.metric("Collection Area (Ad)", f"{ad:.0f} m²")
                st.metric("Expected Frequency (Nd)", f"{nd:.6f}")
            with col_b:
                st.metric("Protection Level", lpl)
                st.metric("Efficiency", f"{efficiency:.1%}")
            with col_c:
                st.metric("Rolling Sphere", f"{sphere}m")
                st.metric("Air Terminals", air_terminals)
            
            st.session_state.calc_results = {
                'ad': ad, 'ng': ng, 'nd': nd, 'efficiency': efficiency,
                'lpl': lpl, 'lpl_desc': lpl_desc, 'sphere': sphere,
                'air_terminals': air_terminals
            }
            st.session_state.input_values = {
                'length': length, 'width': width, 'height': height,
                'td_days': td_days, 'environment': environment, 'cd': cd
            }
            st.session_state.calc_done = True
            st.success("✅ Risk Assessment Complete!")
    
    # ========== TAB 3: PROTECTION DESIGN ==========
    with lp_tabs[2]:
        st.markdown('<div class="report-header">', unsafe_allow_html=True)
        st.markdown("## PROTECTION DESIGN (IEC 62305-3 & NFPA 780)")
        st.markdown('</div>', unsafe_allow_html=True)
        
        if not st.session_state.calc_done:
            st.warning("⚠️ Please complete Risk Assessment first!")
        else:
            results = st.session_state.calc_results
            
            st.success(f"✅ Designing for: **{results['lpl']}**")
            
            col1, col2 = st.columns(2)
            
            with col1:
                st.markdown("### 🔧 Air Termination System")
                st.metric("Air Terminals Required", results['air_terminals'])
                st.metric("Rolling Sphere Radius", f"{results['sphere']}m")
                
                if results['lpl'] in ["Class I", "Class II"]:
                    rod_dia, down_size = 12.7, 58
                else:
                    rod_dia, down_size = 9.5, 29
                
                st.metric("Rod Diameter", f"{rod_dia} mm")
                st.metric("Down Conductor", f"{down_size} mm²")
            
            with col2:
                st.markdown("### 🌍 Earthing System")
                st.metric("Earthing Type", "Type A - Vertical Rods")
                st.metric("Earth Rod Length", "3.0 m")
                st.metric("Earth Rod Diameter", "15 mm")
                st.metric("Target Resistance", "<10 Ω")
    
    # ========== TAB 4: CALCULATIONS ==========
    with lp_tabs[3]:
        st.markdown('<div class="report-header">', unsafe_allow_html=True)
        st.markdown("## DETAILED CALCULATIONS")
        st.markdown('</div>', unsafe_allow_html=True)
        
        if not st.session_state.calc_done:
            st.warning("⚠️ Please complete Risk Assessment first!")
        else:
            with st.expander("Collection Area (Ad)", expanded=True):
                st.markdown("**Formula:** Ad = L×W + 2×(3H)×(L+W) + π×(3H)²")
                st.markdown(f"**Result:** Ad = **{st.session_state.calc_results['ad']:.2f} m²**")
    
    # ========== TAB 5: REVISION HISTORY ==========
    with lp_tabs[4]:
        st.markdown('<div class="report-header">', unsafe_allow_html=True)
        st.markdown("## REVISION HISTORY")
        st.markdown('</div>', unsafe_allow_html=True)
        
        col1, col2, col3, col4, col5, col6 = st.columns(6)
        with col1:
            st.session_state.revision_history[0]['rev'] = st.text_input("Rev", st.session_state.revision_history[0]['rev'])
        with col2:
            st.session_state.revision_history[0]['date'] = st.text_input("Date", st.session_state.revision_history[0]['date'])
        with col3:
            st.session_state.revision_history[0]['purpose'] = st.text_input("Purpose", st.session_state.revision_history[0]['purpose'])
        with col4:
            st.session_state.revision_history[0]['prpd'] = st.text_input("PRPD", st.session_state.revision_history[0]['prpd'])
        with col5:
            st.session_state.revision_history[0]['revd'] = st.text_input("REVD", st.session_state.revision_history[0]['revd'])
        with col6:
            st.session_state.revision_history[0]['appd'] = st.text_input("APPD", st.session_state.revision_history[0]['appd'])
    
    # ========== TAB 6: PDF REPORT ==========
    with lp_tabs[5]:
        st.markdown('<div class="report-header">', unsafe_allow_html=True)
        st.markdown("## GENERATE PDF REPORT")
        st.markdown('</div>', unsafe_allow_html=True)
        
        if not st.session_state.calc_done:
            st.warning("⚠️ Please complete Risk Assessment first!")
        else:
            if st.button("📥 GENERATE PDF REPORT", type="primary", use_container_width=True):
                st.info("PDF generation will be available in next update")
                # PDF generation code here

# ========== OTHER CALCULATORS (PLACEHOLDERS) ==========
elif st.session_state.selected_calculator == "🔌 Cable Sizing":
    st.markdown('<div class="report-header">', unsafe_allow_html=True)
    st.markdown("## CABLE SIZING CALCULATOR")
    st.markdown('</div>', unsafe_allow_html=True)
    
    st.info("🔧 Cable sizing calculator will be implemented here")
    
    # Placeholder tabs for cable sizing
    cable_tabs = st.tabs(["📐 Input Parameters", "📊 Calculations", "📋 Results", "📥 Report"])
    
    with cable_tabs[0]:
        st.markdown("### Cable Parameters")
        col1, col2 = st.columns(2)
        with col1:
            st.number_input("Current (A)", value=100)
            st.number_input("Length (m)", value=50)
        with col2:
            st.selectbox("Cable Type", ["Copper", "Aluminum"])
            st.selectbox("Installation", ["Air", "Ground", "Conduit"])

elif st.session_state.selected_calculator == "⚙️ Transformer Sizing":
    st.markdown('<div class="report-header">', unsafe_allow_html=True)
    st.markdown("## TRANSFORMER SIZING CALCULATOR")
    st.markdown('</div>', unsafe_allow_html=True)
    
    st.info("⚙️ Transformer sizing calculator will be implemented here")
    
    transformer_tabs = st.tabs(["📐 Input", "📊 Calculations", "📋 Results"])
    
    with transformer_tabs[0]:
        st.markdown("### Load Parameters")
        st.number_input("Total Load (kVA)", value=500)
        st.number_input("Future Expansion (%)", value=20)

elif st.session_state.selected_calculator == "📊 Load Flow Analysis":
    st.markdown('<div class="report-header">', unsafe_allow_html=True)
    st.markdown("## LOAD FLOW ANALYSIS")
    st.markdown('</div>', unsafe_allow_html=True)
    st.info("📊 Load flow analysis calculator will be implemented here")

elif st.session_state.selected_calculator == "🔧 Short Circuit":
    st.markdown('<div class="report-header">', unsafe_allow_html=True)
    st.markdown("## SHORT CIRCUIT CALCULATION")
    st.markdown('</div>', unsafe_allow_html=True)
    st.info("🔧 Short circuit calculation will be implemented here")

elif st.session_state.selected_calculator == "📈 Voltage Drop":
    st.markdown('<div class="report-header">', unsafe_allow_html=True)
    st.markdown("## VOLTAGE DROP CALCULATOR")
    st.markdown('</div>', unsafe_allow_html=True)
    st.info("📈 Voltage drop calculator will be implemented here")

# Footer
st.markdown("---")
st.markdown(f"<div style='text-align: center; color: gray;'>⚡ Professional Engineering Tools | Version 1.0 | {datetime.now().strftime('%Y-%m-%d')}</div>", unsafe_allow_html=True)