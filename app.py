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
    st.session_state.selected_calculator = "⚡ Lightning Protection"
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
if 'input_values' not in st.session_state:
    st.session_state.input_values = {}

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
    
    def add_title_page(self, company_logo_path=None, contractor_logo_path=None):
        self.add_page()
        
        # Full page border
        self.set_draw_color(0, 0, 0)
        self.set_line_width(0.5)
        self.rect(10, 10, 190, 277)
        
        # ROW 1: 3 Columns - Company Logo, Title, Contractor Logo
        # Column 1 - Company Logo
        self.set_xy(15, 15)
        self.rect(15, 15, 60, 25)
        if company_logo_path and os.path.exists(company_logo_path):
            try:
                self.image(company_logo_path, 18, 17, 54, 21)
            except:
                self.set_xy(20, 22)
                self.set_font('Arial', 'B', 8)
                self.cell(50, 10, 'COMPANY', 0, 1, 'C')
        else:
            self.set_xy(20, 22)
            self.set_font('Arial', 'B', 8)
            self.cell(50, 10, 'COMPANY', 0, 1, 'C')
        
        # Column 2 - Project Title
        self.set_xy(75, 15)
        self.rect(75, 15, 60, 25)
        self.set_xy(77, 20)
        self.set_font('Arial', 'B', 7)
        self.multi_cell(56, 3.5, st.session_state.project_info['project_title'], 0, 'C')
        
        # Column 3 - Contractor Logo
        self.set_xy(135, 15)
        self.rect(135, 15, 60, 25)
        if contractor_logo_path and os.path.exists(contractor_logo_path):
            try:
                self.image(contractor_logo_path, 138, 17, 54, 21)
            except:
                self.set_xy(150, 22)
                self.set_font('Arial', 'B', 8)
                self.cell(40, 10, 'CONTRACTOR', 0, 1, 'C')
        else:
            self.set_xy(150, 22)
            self.set_font('Arial', 'B', 8)
            self.cell(40, 10, 'CONTRACTOR', 0, 1, 'C')
        
        # ROW 2: 3 Columns - Rev, Title, Date
        # Column 1 - Revision
        self.set_xy(15, 45)
        self.rect(15, 45, 60, 15)
        self.set_xy(20, 48)
        self.set_font('Arial', 'B', 10)
        self.cell(50, 8, f"Rev: {st.session_state.cover_details['revision']}", 0, 1, 'C')
        
        # Column 2 - Document Title
        self.set_xy(75, 45)
        self.rect(75, 45, 60, 15)
        self.set_xy(77, 48)
        self.set_font('Arial', 'B', 9)
        self.cell(56, 8, st.session_state.cover_details['title'], 0, 1, 'C')
        
        # Column 3 - Date
        self.set_xy(135, 45)
        self.rect(135, 45, 60, 15)
        self.set_xy(150, 48)
        self.set_font('Arial', 'B', 9)
        self.cell(40, 8, st.session_state.cover_details['date'], 0, 1, 'C')
        
        # Large empty rectangle
        self.set_xy(15, 65)
        self.rect(15, 65, 180, 30)
        
        # Main Title
        self.set_xy(15, 100)
        self.set_font('Arial', 'B', 14)
        self.set_text_color(0, 51, 102)
        self.cell(180, 10, 'LIGHTNING PROTECTION CALCULATION', 0, 1, 'C')
        
        # Revision Legend Table
        self.set_xy(30, 120)
        self.rect(30, 120, 150, 40)
        self.line(70, 120, 70, 160)
        
        legend_items = [
            ('I', 'ACCEPTED FOR INFORMATION ONLY.'),
            ('A', 'APPROVED - NO COMMENTS'),
            ('B', 'APPROVED WITH COMMENTS - WORK MAY PROCEED'),
            ('C', 'REJECTED. TO BE RESUBMITTED - WORK SHALL NOT PROCEED')
        ]
        
        self.set_font('Arial', '', 9)
        y_pos = 125
        for code, desc in legend_items:
            self.set_xy(35, y_pos)
            self.cell(30, 6, code, 0, 0, 'C')
            self.set_xy(75, y_pos)
            self.cell(100, 6, desc, 0, 1, 'L')
            y_pos += 8
        
        # Revision History Table
        self.set_xy(15, 175)
        self.rect(15, 175, 180, 35)
        
        # Table Header
        self.set_xy(17, 178)
        self.set_font('Arial', 'B', 8)
        self.set_fill_color(240, 240, 240)
        
        self.cell(15, 6, 'REV', 1, 0, 'C', 1)
        self.cell(25, 6, 'DATE', 1, 0, 'C', 1)
        self.cell(60, 6, 'ISSUE PURPOSE', 1, 0, 'C', 1)
        self.cell(25, 6, 'PRPD BY', 1, 0, 'C', 1)
        self.cell(25, 6, 'REVD BY', 1, 0, 'C', 1)
        self.cell(25, 6, 'APPD BY', 1, 1, 'C', 1)
        
        # Table Data
        self.set_font('Arial', '', 8)
        self.set_xy(17, 184)
        rev = st.session_state.revision_history[0]
        self.cell(15, 6, rev['rev'], 1, 0, 'C')
        self.cell(25, 6, rev['date'], 1, 0, 'C')
        self.cell(60, 6, rev['purpose'], 1, 0, 'C')
        self.cell(25, 6, rev['prpd'], 1, 0, 'C')
        self.cell(25, 6, rev['revd'], 1, 0, 'C')
        self.cell(25, 6, rev['appd'], 1, 1, 'C')
        
        # Document Numbers
        self.set_y(220)
        self.set_font('Arial', 'B', 9)
        self.cell(60, 6, 'DOCUMENT NUMBER', 0, 0)
        self.cell(60, 6, '', 0, 0)
        self.cell(60, 6, 'PROJECT NUMBER', 0, 1)
        
        self.set_font('Arial', '', 9)
        self.cell(60, 6, st.session_state.project_info['document_number'], 0, 0)
        self.cell(60, 6, '', 0, 0)
        self.cell(60, 6, st.session_state.project_info['project_number'], 0, 1)
        
        # Page number
        self.set_y(250)
        self.set_font('Arial', 'I', 8)
        self.cell(0, 5, 'Page 1 of 9', 0, 1, 'C')
    
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
        self.cell(0, 6, f'Result: Ad = {results["ad"]:.2f} m^2', 0, 1)
        self.ln(3)
        
        # Environmental Factor
        self.set_font('Arial', 'B', 11)
        self.cell(0, 7, '1.2 Environmental Factor (CD)', 0, 1)
        self.set_font('Arial', '', 9)
        self.cell(0, 5, f'Environment: {inputs.get("environment", "Isolated")}', 0, 1)
        self.set_font('Arial', 'B', 9)
        self.cell(0, 6, f'Result: CD = {inputs.get("cd", 1)}', 0, 1)
        self.ln(3)
        
        # Lightning Density
        self.set_font('Arial', 'B', 11)
        self.cell(0, 7, '1.3 Lightning Ground Flash Density (NG)', 0, 1)
        self.set_font('Arial', '', 9)
        self.cell(0, 5, 'Formula: NG = 0.1 x Td', 0, 1)
        self.set_font('Arial', 'B', 9)
        self.cell(0, 6, f'Result: NG = {results.get("ng", 1)} flashes/km^2/year', 0, 1)
        self.ln(3)
        
        # Expected Frequency
        self.set_font('Arial', 'B', 11)
        self.cell(0, 7, '1.4 Expected Annual Frequency (Nd)', 0, 1)
        self.set_font('Arial', '', 9)
        self.cell(0, 5, 'Formula: Nd = NG x Ad x CD x 10^-6', 0, 1)
        self.set_font('Arial', 'B', 9)
        self.cell(0, 6, f'Result: Nd = {results.get("nd", 0):.6f} events/year', 0, 1)
        self.ln(3)
        
        # Protection Level
        self.set_font('Arial', 'B', 11)
        self.cell(0, 7, '1.5 Protection Level', 0, 1)
        self.set_font('Arial', 'B', 9)
        self.cell(0, 6, f'Result: {results.get("lpl", "Class III")}', 0, 1)
        self.cell(0, 6, f'Rolling Sphere: {results.get("sphere", 45)}m', 0, 1)
        self.ln(3)
        
        # Air Terminals
        self.set_font('Arial', 'B', 11)
        self.cell(0, 7, '1.6 Air Terminals Required', 0, 1)
        self.set_font('Arial', 'B', 9)
        self.cell(0, 6, f'Result: {results.get("air_terminals", 4)} air terminals', 0, 1)

# ========== SIDEBAR ==========
with st.sidebar:
    st.markdown("### ⚡ Lightning Protection Systems")
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
    
    # Common Project Information
    st.markdown("### 📋 Project Information")
    
    st.session_state.project_info['company'] = st.text_input("Company Name", st.session_state.project_info['company'])
    st.session_state.project_info['contractor'] = st.text_input("Contractor Name", st.session_state.project_info['contractor'])
    st.session_state.project_info['project_title'] = st.text_area("Project Title", st.session_state.project_info['project_title'], height=60)
    st.session_state.project_info['document_number'] = st.text_input("Document Number", st.session_state.project_info['document_number'])
    st.session_state.project_info['project_number'] = st.text_input("Project Number", st.session_state.project_info['project_number'])
    
    st.markdown("---")
    
    # Logo Uploads
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

# ========== MAIN CONTENT ==========
st.title(f"⚡ {st.session_state.selected_calculator} Calculator")

# Lightning Protection Calculator
if st.session_state.selected_calculator == "⚡ Lightning Protection":
    
    lp_tabs = st.tabs([
        "🏢 Title Page", 
        "📊 Risk Assessment", 
        "🔧 Protection Design", 
        "📋 Calculations",
        "📝 Revision History",
        "📥 PDF Report"
    ])
    
    # TAB 1: Title Page
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
    
    # TAB 2: Risk Assessment
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
            st.info(f"**CD = {cd}**")
            
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
            
            ng = 0.1 * td_days
            nd = ng * ad * cd * 1e-6
            
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
            
            col_a, col_b, col_c = st.columns(3)
            with col_a:
                st.metric("Collection Area", f"{ad:.0f} m²")
            with col_b:
                st.metric("Protection Level", lpl)
                st.metric("Efficiency", f"{efficiency:.1%}")
            with col_c:
                st.metric("Rolling Sphere", f"{sphere}m")
                st.metric("Air Terminals", air_terminals)
            
            st.session_state.calc_results = {
                'ad': ad, 'ng': ng, 'nd': nd, 'efficiency': efficiency,
                'lpl': lpl, 'sphere': sphere, 'air_terminals': air_terminals
            }
            st.session_state.input_values = {
                'length': length, 'width': width, 'height': height,
                'td_days': td_days, 'environment': environment, 'cd': cd
            }
            st.session_state.calc_done = True
    
    # TAB 3: Protection Design
    with lp_tabs[2]:
        st.markdown('<div class="report-header">', unsafe_allow_html=True)
        st.markdown("## PROTECTION DESIGN")
        st.markdown('</div>', unsafe_allow_html=True)
        
        if not st.session_state.calc_done:
            st.warning("⚠️ Please complete Risk Assessment first!")
        else:
            results = st.session_state.calc_results
            st.success(f"✅ Designing for: **{results['lpl']}**")
            
            col1, col2 = st.columns(2)
            with col1:
                st.metric("Air Terminals", results['air_terminals'])
                st.metric("Rolling Sphere", f"{results['sphere']}m")
    
    # TAB 4: Calculations - WITH DETAILED CALCULATIONS
    with lp_tabs[3]:
        st.markdown('<div class="report-header">', unsafe_allow_html=True)
        st.markdown("## DETAILED CALCULATIONS WITH FORMULAS")
        st.markdown('</div>', unsafe_allow_html=True)
        
        if not st.session_state.calc_done:
            st.warning("⚠️ Please complete Risk Assessment first!")
        else:
            results = st.session_state.calc_results
            inputs = st.session_state.input_values
            
            # 1. Collection Area
            with st.expander("1. Collection Area (Ad)", expanded=True):
                st.markdown('<div class="formula-box">', unsafe_allow_html=True)
                st.markdown("**Formula:** Ad = L × W + 2 × (3H) × (L + W) + π × (3H)²")
                st.markdown("**Reference:** IEC 62305-2 Annex A.2.1.1, Equation A.2")
                if inputs.get('width', 0) == 0:  # Column
                    st.markdown(f"**For Column:** Ad = π × 9 × H²")
                    st.markdown(f"**Calculation:** Ad = π × 9 × ({inputs['height']})²")
                else:
                    st.markdown(f"**Calculation:** Ad = {inputs['length']} × {inputs['width']} + 2 × (3 × {inputs['height']}) × ({inputs['length']} + {inputs['width']}) + π × (3 × {inputs['height']})²")
                st.markdown(f"**Result:** Ad = **{results['ad']:.2f} m²**")
                st.markdown('</div>', unsafe_allow_html=True)
            
            # 2. Environmental Factor
            with st.expander("2. Environmental Factor (CD)"):
                st.markdown('<div class="formula-box">', unsafe_allow_html=True)
                st.markdown("**Reference:** IEC 62305-2 Table A.1 - Location Factor")
                st.markdown("**Values:** Surrounded=0.25, Similar height=0.5, Isolated=1, Hilltop=2")
                st.markdown(f"**Selected Environment:** {inputs.get('environment', 'Isolated')}")
                st.markdown(f"**Result:** CD = **{inputs.get('cd', 1)}**")
                st.markdown('</div>', unsafe_allow_html=True)
            
            # 3. Lightning Density
            with st.expander("3. Lightning Ground Flash Density (NG)"):
                st.markdown('<div class="formula-box">', unsafe_allow_html=True)
                st.markdown("**Formula:** NG = 0.1 × Td")
                st.markdown("**Reference:** IEC 62305-2 Annex A.1, Equation A.1")
                st.markdown(f"**Calculation:** NG = 0.1 × {inputs.get('td_days', 10)}")
                st.markdown(f"**Result:** NG = **{results.get('ng', 1)} flashes/km²/year**")
                st.markdown('</div>', unsafe_allow_html=True)
            
            # 4. Expected Frequency
            with st.expander("4. Expected Annual Frequency (Nd)"):
                st.markdown('<div class="formula-box">', unsafe_allow_html=True)
                st.markdown("**Formula:** Nd = NG × Ad × CD × 10⁻⁶")
                st.markdown("**Reference:** IEC 62305-2 Annex A.2.4, Equation A.4")
                st.markdown(f"**Calculation:** Nd = {results.get('ng', 1)} × {results['ad']:.0f} × {inputs.get('cd', 1)} × 10⁻⁶")
                st.markdown(f"**Result:** Nd = **{results.get('nd', 0):.6f} events/year**")
                st.markdown('</div>', unsafe_allow_html=True)
            
            # 5. Protection Level
            with st.expander("5. Protection Level Determination"):
                st.markdown('<div class="formula-box">', unsafe_allow_html=True)
                st.markdown("**Reference:** IEC 62305-1 Table 1 & Figure 1")
                st.markdown(f"**Protection Efficiency:** {results.get('efficiency', 0):.1%}")
                st.markdown(f"**Result:** **{results.get('lpl', 'Class III')}**")
                st.markdown(f"**Rolling Sphere Radius:** {results.get('sphere', 45)}m (IEC 62305-3 Table 2)")
                st.markdown('</div>', unsafe_allow_html=True)
            
            # 6. Air Terminals
            with st.expander("6. Air Terminals Required"):
                st.markdown('<div class="formula-box">', unsafe_allow_html=True)
                st.markdown("**Method:** Rolling Sphere Method")
                st.markdown("**Reference:** IEC 62305-3 Clause 5.2.2, Table 2")
                if inputs.get('height', 0) <= results.get('sphere', 45):
                    st.markdown("**Using:** Protection Width Method")
                else:
                    st.markdown("**Using:** Mesh Method for tall structures")
                st.markdown(f"**Result:** **{results.get('air_terminals', 4)} air terminals required**")
                st.markdown('</div>', unsafe_allow_html=True)
    
    # TAB 5: Revision History
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
    
    # TAB 6: PDF Report
    with lp_tabs[5]:
        st.markdown('<div class="report-header">', unsafe_allow_html=True)
        st.markdown("## GENERATE PDF REPORT")
        st.markdown('</div>', unsafe_allow_html=True)
        
        if not st.session_state.calc_done:
            st.warning("⚠️ Please complete Risk Assessment first!")
        else:
            if st.button("📥 GENERATE PDF REPORT", type="primary", use_container_width=True):
                with st.spinner("Generating PDF report..."):
                    
                    # Save logos temporarily
                    company_logo_path = ""
                    contractor_logo_path = ""
                    
                    if st.session_state.company_logo is not None:
                        company_logo_path = "temp_company_logo.png"
                        st.session_state.company_logo.save(company_logo_path)
                    
                    if st.session_state.contractor_logo is not None:
                        contractor_logo_path = "temp_contractor_logo.png"
                        st.session_state.contractor_logo.save(contractor_logo_path)
                    
                    # Create PDF
                    pdf = PDF_Report()
                    pdf.add_title_page(company_logo_path, contractor_logo_path)
                    pdf.add_calculations(st.session_state.calc_results, st.session_state.input_values)
                    
                    # Clean up temp files
                    if company_logo_path and os.path.exists(company_logo_path):
                        os.remove(company_logo_path)
                    if contractor_logo_path and os.path.exists(contractor_logo_path):
                        os.remove(contractor_logo_path)
                    
                    # Generate PDF
                    pdf_output = pdf.output(dest='S')
                    b64 = base64.b64encode(pdf_output).decode()
                    
                    # Create download button
                    filename = f"LPS_Report_{st.session_state.cover_details['revision']}_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf"
                    
                    st.markdown(f"""
                    <div style="text-align: center; margin: 20px 0;">
                        <a href="data:application/pdf;base64,{b64}" download="{filename}" 
                           style="display: inline-block; padding: 15px 30px; background-color: #4CAF50; 
                                  color: white; text-decoration: none; border-radius: 5px; font-size: 18px;">
                            📥 CLICK HERE TO DOWNLOAD PDF
                        </a>
                    </div>
                    """, unsafe_allow_html=True)
                    
                    st.success("✅ PDF Generated Successfully! Click the button above to download.")

# Other Calculators (Placeholders)
elif st.session_state.selected_calculator == "🔌 Cable Sizing":
    st.info("🔧 Cable sizing calculator will be implemented here")

elif st.session_state.selected_calculator == "⚙️ Transformer Sizing":
    st.info("⚙️ Transformer sizing calculator will be implemented here")

elif st.session_state.selected_calculator == "📊 Load Flow Analysis":
    st.info("📊 Load flow analysis calculator will be implemented here")

elif st.session_state.selected_calculator == "🔧 Short Circuit":
    st.info("🔧 Short circuit calculation will be implemented here")

elif st.session_state.selected_calculator == "📈 Voltage Drop":
    st.info("📈 Voltage drop calculator will be implemented here")

# Footer
st.markdown("---")
st.markdown(f"<div style='text-align: center; color: gray;'>⚡ Professional Engineering Tools | Version 3.0 | {datetime.now().strftime('%Y-%m-%d %H:%M')}</div>", unsafe_allow_html=True)