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

st.set_page_config(page_title="Professional Lightning Protection", page_icon="⚡", layout="wide")

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
    .title-page-preview {
        background-color: white;
        padding: 30px;
        border: 2px solid #1E3A8A;
        border-radius: 10px;
        margin-bottom: 20px;
        font-family: 'Arial', sans-serif;
    }
</style>
""", unsafe_allow_html=True)

# Initialize session state
if 'company_logo' not in st.session_state:
    st.session_state.company_logo = None
if 'calc_results' not in st.session_state:
    st.session_state.calc_results = {}
if 'calc_done' not in st.session_state:
    st.session_state.calc_done = False
if 'cover_details' not in st.session_state:
    st.session_state.cover_details = {
        'title': 'PLANT LIGHTING CALCULATION',
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

# ========== PDF Report Generator Class ==========
class PDF_Report(FPDF):
    def __init__(self):
        super().__init__()
        self.set_auto_page_break(auto=True, margin=15)
        
    def header(self):
        if self.page_no() > 1:
            self.set_font('Arial', 'I', 8)
            self.cell(0, 10, 'Lightning Protection Calculation', 0, 0, 'L')
            self.cell(0, 10, f'Page {self.page_no()}', 0, 0, 'R')
            self.ln(15)
    
    def footer(self):
        self.set_y(-15)
        self.set_font('Arial', 'I', 8)
        self.cell(0, 10, f'Document: {st.session_state.project_info["document_number"]} | Rev: {st.session_state.cover_details["revision"]}', 0, 0, 'C')
    
    def add_title_page(self, logo_path=None):
        self.add_page()
        
        # Company and Contractor
        self.set_y(15)
        self.set_font('Arial', 'B', 14)
        self.set_text_color(0, 0, 0)
        self.cell(0, 6, st.session_state.project_info['company'], 0, 1, 'C')
        
        self.set_font('Arial', 'B', 12)
        self.cell(0, 6, st.session_state.project_info['contractor'], 0, 1, 'C')
        
        # Project Title
        self.set_y(35)
        self.set_font('Arial', 'B', 10)
        self.multi_cell(0, 4, st.session_state.project_info['project_title'], 0, 'C')
        
        # Revision
        self.set_y(55)
        self.set_font('Arial', 'B', 10)
        self.cell(0, 5, f"Rev: {st.session_state.cover_details['revision']}", 0, 1, 'C')
        
        # Document Title
        self.set_y(65)
        self.set_font('Arial', 'B', 14)
        self.set_text_color(0, 51, 102)
        self.cell(0, 6, st.session_state.cover_details['title'], 0, 1, 'C')
        
        # Revision Legend
        self.set_y(80)
        self.set_font('Arial', '', 8)
        legend = [
            "ACCEPTED FOR INFORMATION ONLY.",
            "APPROVED - NO COMMENTS",
            "APPROVED WITH COMMENTS - WORK MAY PROCEED",
            "REJECTED. TO BE RESUBMITTED - WORK SHALL NOT PROCEED"
        ]
        for line in legend:
            self.cell(0, 3.5, line, 0, 1, 'L')
        
        # Revision Table
        self.set_y(105)
        self.set_font('Arial', 'B', 8)
        self.set_fill_color(240, 240, 240)
        
        # Table Header
        self.cell(15, 5, 'REV', 1, 0, 'C', 1)
        self.cell(25, 5, 'DATE', 1, 0, 'C', 1)
        self.cell(70, 5, 'ISSUE PURPOSE', 1, 0, 'C', 1)
        self.cell(25, 5, 'PRPD BY', 1, 0, 'C', 1)
        self.cell(25, 5, 'REVD BY', 1, 0, 'C', 1)
        self.cell(25, 5, 'APPD BY', 1, 1, 'C', 1)
        
        # Table Data
        self.set_font('Arial', '', 8)
        rev = st.session_state.revision_history[0]
        self.cell(15, 5, rev['rev'], 1, 0, 'C')
        self.cell(25, 5, rev['date'], 1, 0, 'C')
        self.cell(70, 5, rev['purpose'], 1, 0, 'C')
        self.cell(25, 5, rev['prpd'], 1, 0, 'C')
        self.cell(25, 5, rev['revd'], 1, 0, 'C')
        self.cell(25, 5, rev['appd'], 1, 1, 'C')
        
        # Document Numbers
        self.set_y(135)
        self.set_font('Arial', 'B', 9)
        self.cell(60, 5, 'DOCUMENT NUMBER', 0, 0)
        self.cell(60, 5, '', 0, 0)
        self.cell(60, 5, 'PROJECT NUMBER', 0, 1)
        
        self.set_font('Arial', '', 9)
        self.cell(60, 5, st.session_state.project_info['document_number'], 0, 0)
        self.cell(60, 5, '', 0, 0)
        self.cell(60, 5, st.session_state.project_info['project_number'], 0, 1)
        
        # SEAL
        self.set_y(160)
        
        # Outer circle
        self.set_draw_color(0, 51, 102)
        self.set_line_width(1.5)
        self.ellipse(150, 160, 45, 32)
        
        # Inner circle
        self.set_draw_color(0, 51, 102)
        self.set_line_width(0.8)
        self.ellipse(152, 162, 41, 28)
        
        # Seal text
        self.set_font('Arial', 'B', 12)
        self.set_text_color(0, 51, 102)
        self.set_xy(155, 168)
        self.cell(35, 6, 'APPROVED', 0, 1, 'C')
        
        self.set_font('Arial', 'B', 10)
        self.set_xy(155, 174)
        self.cell(35, 5, 'SIGNATURE', 0, 1, 'C')
        
        self.set_font('Arial', '', 9)
        self.set_xy(155, 179)
        self.cell(35, 4, 'SHD', 0, 1, 'C')
        
        # Date in seal
        self.set_font('Arial', '', 7)
        self.set_xy(155, 183)
        self.cell(35, 4, datetime.now().strftime('%d-%b-%Y'), 0, 1, 'C')
        
        # Reset colors
        self.set_text_color(0, 0, 0)
        self.set_draw_color(0, 0, 0)
        
        # Page Number
        self.set_y(200)
        self.set_font('Arial', 'I', 8)
        self.set_text_color(100, 100, 100)
        self.cell(0, 4, 'Page 1 of 9', 0, 1, 'C')
    
    def add_calculations(self, results, inputs):
        self.add_page()
        
        self.set_font('Arial', 'B', 14)
        self.set_text_color(0, 51, 102)
        self.cell(0, 10, '1. RISK ASSESSMENT CALCULATIONS', 0, 1)
        self.ln(5)
        
        # Collection Area
        self.set_font('Arial', 'B', 11)
        self.cell(0, 7, '1.1 Collection Area (Ad)', 0, 1)
        self.set_font('Arial', '', 9)
        self.multi_cell(0, 5, 'Formula: Ad = L x W + 2 x (3H) x (L + W) + 3.14 x (3H)^2')
        self.cell(0, 5, 'Reference: IEC 62305-2 Annex A.2.1.1, Equation A.2', 0, 1)
        self.set_font('Arial', 'B', 9)
        self.cell(0, 6, f'Calculation: Ad = {inputs["length"]} x {inputs["width"]} + 2 x (3 x {inputs["height"]}) x ({inputs["length"]} + {inputs["width"]}) + 3.14 x (3 x {inputs["height"]})^2', 0, 1)
        self.cell(0, 6, f'Result: Ad = {results["ad"]:.2f} m^2', 0, 1)
        self.ln(3)
        
        # Environmental Factor
        self.set_font('Arial', 'B', 11)
        self.cell(0, 7, '1.2 Environmental Factor (CD)', 0, 1)
        self.set_font('Arial', '', 9)
        self.cell(0, 5, 'Reference: IEC 62305-2 Table A.1 - Location Factor', 0, 1)
        self.cell(0, 5, f'Environment: {inputs["environment"]}', 0, 1)
        self.set_font('Arial', 'B', 9)
        self.cell(0, 6, f'Result: CD = {inputs["cd"]}', 0, 1)
        self.ln(3)
        
        # Lightning Density
        self.set_font('Arial', 'B', 11)
        self.cell(0, 7, '1.3 Lightning Ground Flash Density (NG)', 0, 1)
        self.set_font('Arial', '', 9)
        self.cell(0, 5, 'Formula: NG = 0.1 x Td', 0, 1)
        self.cell(0, 5, 'Reference: IEC 62305-2 Annex A.1, Equation A.1', 0, 1)
        self.cell(0, 5, f'Calculation: NG = 0.1 x {inputs["td_days"]}', 0, 1)
        self.set_font('Arial', 'B', 9)
        self.cell(0, 6, f'Result: NG = {results["ng"]} flashes/km^2/year', 0, 1)
        self.ln(3)
        
        # Expected Frequency
        self.set_font('Arial', 'B', 11)
        self.cell(0, 7, '1.4 Expected Annual Frequency (Nd)', 0, 1)
        self.set_font('Arial', '', 9)
        self.cell(0, 5, 'Formula: Nd = NG x Ad x CD x 10^-6', 0, 1)
        self.cell(0, 5, 'Reference: IEC 62305-2 Annex A.2.4, Equation A.4', 0, 1)
        self.cell(0, 5, f'Calculation: Nd = {results["ng"]} x {results["ad"]:.0f} x {inputs["cd"]} x 10^-6', 0, 1)
        self.set_font('Arial', 'B', 9)
        self.cell(0, 6, f'Result: Nd = {results["nd"]:.6f} events/year', 0, 1)
        self.ln(3)
        
        # Protection Level
        self.set_font('Arial', 'B', 11)
        self.cell(0, 7, '1.5 Protection Level Determination', 0, 1)
        self.set_font('Arial', '', 9)
        self.cell(0, 5, 'Reference: IEC 62305-1 Table 1 & Figure 1', 0, 1)
        self.cell(0, 5, f'Protection Efficiency: {results["efficiency"]:.1%}', 0, 1)
        self.set_font('Arial', 'B', 9)
        self.cell(0, 6, f'Result: {results["lpl"]} - {results["lpl_desc"]}', 0, 1)
        self.cell(0, 6, f'Rolling Sphere Radius: {results["sphere"]}m (IEC 62305-3 Table 2)', 0, 1)
        self.ln(3)
        
        # Air Terminals
        self.set_font('Arial', 'B', 11)
        self.cell(0, 7, '1.6 Air Terminals Required', 0, 1)
        self.set_font('Arial', '', 9)
        self.cell(0, 5, 'Method: Rolling Sphere Method', 0, 1)
        self.cell(0, 5, 'Reference: IEC 62305-3 Clause 5.2.2, Table 2', 0, 1)
        self.set_font('Arial', 'B', 9)
        self.cell(0, 6, f'Result: {results["air_terminals"]} air terminals required', 0, 1)

# ========== SIDEBAR ==========
with st.sidebar:
    st.markdown("### ⚡ Lightning Protection Systems")
    
    # Logo Upload
    st.markdown("---")
    st.markdown("### 🖼️ Company Logo")
    
    uploaded_logo = st.file_uploader("Upload Logo", type=['png', 'jpg', 'jpeg'])
    
    if uploaded_logo is not None:
        st.session_state.logo_bytes = uploaded_logo.getvalue()
        logo_image = Image.open(io.BytesIO(st.session_state.logo_bytes))
        st.session_state.company_logo = logo_image
        st.image(logo_image, width=150, caption="Your Logo")
        st.success("✅ Logo uploaded!")
    
    st.markdown("---")
    
    # Project Management
    st.header("📁 Project Management")
    if st.button("➕ New Project", use_container_width=True):
        st.session_state.calc_done = False
        st.session_state.calc_results = {}
        st.rerun()
    
    st.markdown("---")
    
    # Project Information
    st.header("📋 Project Information")
    
    st.session_state.project_info['company'] = st.text_input("Company Name", st.session_state.project_info['company'])
    st.session_state.project_info['contractor'] = st.text_input("Contractor Name", st.session_state.project_info['contractor'])
    st.session_state.project_info['project_title'] = st.text_area("Project Title", st.session_state.project_info['project_title'], height=80)
    st.session_state.project_info['document_number'] = st.text_input("Document Number", st.session_state.project_info['document_number'])
    st.session_state.project_info['project_number'] = st.text_input("Project Number", st.session_state.project_info['project_number'])

# ========== MAIN TABS ==========
tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "🏢 Title Page", 
    "📊 Risk Assessment", 
    "🔧 Protection Design", 
    "📋 Calculations",
    "📝 Revision History",
    "📥 PDF Report"
])

# ========== TAB 1: TITLE PAGE ==========
with tab1:
    st.markdown('<div class="report-header">', unsafe_allow_html=True)
    st.markdown("## TITLE PAGE DESIGN")
    st.markdown('</div>', unsafe_allow_html=True)
    
    col1, col2 = st.columns([1, 1])
    
    with col1:
        st.markdown("### 📄 Title Page Preview")
        
        # Title Page Preview HTML
        title_html = f"""
        <div class="title-page-preview">
            <h3 style="text-align:center;">{st.session_state.project_info['company']}</h3>
            <h4 style="text-align:center;">{st.session_state.project_info['contractor']}</h4>
            <p style="text-align:center; font-size:12px; margin:20px 0;">{st.session_state.project_info['project_title']}</p>
            <p style="text-align:center;"><strong>Rev: {st.session_state.cover_details['revision']}</strong></p>
            <h2 style="text-align:center; color:#1E3A8A;">{st.session_state.cover_details['title']}</h2>
            <div style="margin:20px 0; font-size:11px;">
                <div>ACCEPTED FOR INFORMATION ONLY.</div>
                <div>APPROVED - NO COMMENTS</div>
                <div>APPROVED WITH COMMENTS - WORK MAY PROCEED</div>
                <div>REJECTED. TO BE RESUBMITTED - WORK SHALL NOT PROCEED</div>
            </div>
            <table style="width:100%; border-collapse:collapse; margin:20px 0;">
                <tr style="background-color:#f2f2f2;">
                    <th style="border:1px solid #ddd; padding:5px;">REV</th>
                    <th style="border:1px solid #ddd; padding:5px;">DATE</th>
                    <th style="border:1px solid #ddd; padding:5px;">ISSUE PURPOSE</th>
                    <th style="border:1px solid #ddd; padding:5px;">PRPD BY</th>
                    <th style="border:1px solid #ddd; padding:5px;">REVD BY</th>
                    <th style="border:1px solid #ddd; padding:5px;">APPD BY</th>
                </tr>
        """
        
        for rev in st.session_state.revision_history:
            title_html += f"""
                <tr>
                    <td style="border:1px solid #ddd; padding:5px;">{rev['rev']}</td>
                    <td style="border:1px solid #ddd; padding:5px;">{rev['date']}</td>
                    <td style="border:1px solid #ddd; padding:5px;">{rev['purpose']}</td>
                    <td style="border:1px solid #ddd; padding:5px;">{rev['prpd']}</td>
                    <td style="border:1px solid #ddd; padding:5px;">{rev['revd']}</td>
                    <td style="border:1px solid #ddd; padding:5px;">{rev['appd']}</td>
                </tr>
            """
        
        title_html += f"""
            </table>
            <div style="margin-top:20px;">
                <strong>DOCUMENT NUMBER</strong> {st.session_state.project_info['document_number']} &nbsp;&nbsp;&nbsp;&nbsp; 
                <strong>PROJECT NUMBER</strong> {st.session_state.project_info['project_number']}
            </div>
            <div style="margin-top:30px; text-align:center; font-style:italic;">Page 1 of 9</div>
        </div>
        """
        
        st.markdown(title_html, unsafe_allow_html=True)
    
    with col2:
        st.markdown("### ✏️ Edit Title Page")
        st.session_state.cover_details['title'] = st.text_input("Document Title", st.session_state.cover_details['title'])
        st.session_state.cover_details['revision'] = st.text_input("Revision", st.session_state.cover_details['revision'])

# ========== TAB 2: RISK ASSESSMENT ==========
with tab2:
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
            'structure': structure_type,
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
with tab3:
    st.markdown('<div class="report-header">', unsafe_allow_html=True)
    st.markdown("## PROTECTION DESIGN (IEC 62305-3 & NFPA 780)")
    st.markdown('</div>', unsafe_allow_html=True)
    
    if not st.session_state.calc_done:
        st.warning("⚠️ Please complete Risk Assessment first in Tab 2!")
    else:
        results = st.session_state.calc_results
        inputs = st.session_state.input_values
        
        st.success(f"✅ Designing for: **{results['structure']}** with **{results['lpl']}**")
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("### 🔧 Air Termination System")
            st.metric("Air Terminals Required", results['air_terminals'])
            st.metric("Rolling Sphere Radius", f"{results['sphere']}m")
            
            if results['lpl'] in ["Class I", "Class II"]:
                rod_dia, down_size = 12.7, 58
                mesh = "5m x 5m" if results['lpl'] == "Class I" else "10m x 10m"
            else:
                rod_dia, down_size = 9.5, 29
                mesh = "15m x 15m" if results['lpl'] == "Class III" else "20m x 20m"
            
            st.metric("Rod Diameter", f"{rod_dia} mm")
            st.metric("Down Conductor", f"{down_size} mm²")
            st.metric("Mesh Size", mesh)
        
        with col2:
            st.markdown("### 🌍 Earthing System")
            if inputs['width'] > 0:
                earth_type = "Type B - Ring Electrode"
                earth_len = 3.0
                st.info("**Ring electrode around building perimeter**")
            else:
                earth_type = "Type A - Vertical Rods"
                earth_len = 2.4
                st.info("**Vertical rods at base of column**")
            
            st.metric("Earthing Type", earth_type)
            st.metric("Earth Rod Length", f"{earth_len}m")
            st.metric("Earth Rod Diameter", "15 mm (min)")
            st.metric("Target Resistance", "<10 Ω")
        
        st.markdown("---")
        st.markdown("### 📋 Bill of Materials")
        
        materials = pd.DataFrame({
            'Component': ['Air Termination Rods', 'Down Conductors', 'Earth Rods', 'Test Joints', 'Connectors'],
            'Quantity': [f"{results['air_terminals']} pcs", f"{down_size} mm²", f"{earth_len}m x 4 nos", f"{max(2, results['air_terminals']//2)} pcs", "As required"],
            'Material': ['Copper (10mm dia)', 'Copper (50mm²)', 'Copper Coated Steel', 'Stainless Steel', 'Copper/Stainless'],
            'Reference': ['IEC 62305-3 Table 6', 'IEC 62305-3 Table 6', 'IEC 62305-3 Table 7', 'Clause 5.3.6', 'Clause 5.5.3']
        })
        st.dataframe(materials, use_container_width=True, hide_index=True)

# ========== TAB 4: CALCULATIONS ==========
with tab4:
    st.markdown('<div class="report-header">', unsafe_allow_html=True)
    st.markdown("## DETAILED CALCULATIONS WITH FORMULAS")
    st.markdown('</div>', unsafe_allow_html=True)
    
    if not st.session_state.calc_done:
        st.warning("⚠️ Please complete Risk Assessment first in Tab 2!")
    else:
        results = st.session_state.calc_results
        inputs = st.session_state.input_values
        
        with st.expander("1. Collection Area (Ad)", expanded=True):
            st.markdown('<div class="formula-box">', unsafe_allow_html=True)
            st.markdown("**Formula:** Ad = L × W + 2 × (3H) × (L + W) + π × (3H)²")
            st.markdown("**Reference:** IEC 62305-2 Annex A.2.1.1, Equation A.2")
            st.markdown(f"**Calculation:** Ad = {inputs['length']} × {inputs['width']} + 2 × (3 × {inputs['height']}) × ({inputs['length']} + {inputs['width']}) + 3.14 × (3 × {inputs['height']})²")
            st.markdown(f"**Result:** Ad = **{results['ad']:.2f} m²**")
            st.markdown('</div>', unsafe_allow_html=True)
        
        with st.expander("2. Environmental Factor (CD)"):
            st.markdown('<div class="formula-box">', unsafe_allow_html=True)
            st.markdown("**Reference:** IEC 62305-2 Table A.1 - Location Factor")
            st.markdown(f"**Selected Environment:** {inputs['environment']}")
            st.markdown(f"**Result:** CD = **{inputs['cd']}**")
            st.markdown('</div>', unsafe_allow_html=True)
        
        with st.expander("3. Lightning Ground Flash Density (NG)"):
            st.markdown('<div class="formula-box">', unsafe_allow_html=True)
            st.markdown("**Formula:** NG = 0.1 × Td")
            st.markdown("**Reference:** IEC 62305-2 Annex A.1, Equation A.1")
            st.markdown(f"**Calculation:** NG = 0.1 × {inputs['td_days']}")
            st.markdown(f"**Result:** NG = **{results['ng']} flashes/km²/year**")
            st.markdown('</div>', unsafe_allow_html=True)
        
        with st.expander("4. Expected Annual Frequency (Nd)"):
            st.markdown('<div class="formula-box">', unsafe_allow_html=True)
            st.markdown("**Formula:** Nd = NG × Ad × CD × 10⁻⁶")
            st.markdown("**Reference:** IEC 62305-2 Annex A.2.4, Equation A.4")
            st.markdown(f"**Calculation:** Nd = {results['ng']} × {results['ad']:.0f} × {inputs['cd']} × 10⁻⁶")
            st.markdown(f"**Result:** Nd = **{results['nd']:.6f} events/year**")
            st.markdown('</div>', unsafe_allow_html=True)
        
        with st.expander("5. Protection Level Determination"):
            st.markdown('<div class="formula-box">', unsafe_allow_html=True)
            st.markdown("**Reference:** IEC 62305-1 Table 1 & Figure 1")
            st.markdown(f"**Efficiency:** {results['efficiency']:.1%}")
            st.markdown(f"**Result:** **{results['lpl']}** - {results['lpl_desc']}")
            st.markdown('</div>', unsafe_allow_html=True)

# ========== TAB 5: REVISION HISTORY ==========
with tab5:
    st.markdown('<div class="report-header">', unsafe_allow_html=True)
    st.markdown("## REVISION HISTORY")
    st.markdown('</div>', unsafe_allow_html=True)
    
    st.markdown("### 📝 Edit Revision Entries")
    
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
with tab6:
    st.markdown('<div class="report-header">', unsafe_allow_html=True)
    st.markdown("## GENERATE PDF REPORT")
    st.markdown('</div>', unsafe_allow_html=True)
    
    if not st.session_state.calc_done:
        st.warning("⚠️ Please complete Risk Assessment first in Tab 2!")
    else:
        st.success("✅ Calculations completed! Generate PDF report with title page.")
        
        if st.button("📥 GENERATE PDF REPORT", type="primary", use_container_width=True):
            with st.spinner("Generating PDF report... Please wait"):
                
                temp_logo = ""
                if st.session_state.company_logo is not None:
                    temp_logo = "temp_logo.png"
                    st.session_state.company_logo.save(temp_logo)
                
                pdf = PDF_Report()
                pdf.add_title_page(temp_logo)
                pdf.add_calculations(st.session_state.calc_results, st.session_state.input_values)
                
                if temp_logo and os.path.exists(temp_logo):
                    os.remove(temp_logo)
                
                pdf_output = pdf.output(dest='S')
                b64 = base64.b64encode(pdf_output).decode()
                
                filename = f"LPS_Report_{st.session_state.cover_details['revision']}_{datetime.now().strftime('%Y%m%d')}.pdf"
                
                st.markdown(f'<a href="data:application/pdf;base64,{b64}" download="{filename}" style="display: inline-block; padding: 15px 30px; background-color: #4CAF50; color: white; text-decoration: none; border-radius: 5px; font-size: 18px; margin: 20px 0;">📥 CLICK HERE TO DOWNLOAD PDF</a>', unsafe_allow_html=True)
                st.success("✅ PDF Generated Successfully!")

# Footer
st.markdown("---")
st.markdown(f"<div style='text-align: center; color: gray;'>⚡ Professional Lightning Protection System | Version 14.0 | {datetime.now().strftime('%Y-%m-%d')}</div>", unsafe_allow_html=True)