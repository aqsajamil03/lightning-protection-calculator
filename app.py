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
        'title': 'ELECTRICAL CABLE SIZING CALCULATION',
        'revision': 'A',
        'date': '09 Sep 2025',
        'purpose': 'ISSUED FOR APPROVAL',
        'prepared_by': 'ASZ',
        'reviewed_by': 'SHZ',
        'approved_by': 'SHD'
    }
if 'project_info' not in st.session_state:
    st.session_state.project_info = {
        'company': 'COMPANY',
        'contractor': 'BOILEN ENERGY DMCC',
        'contractor_address': 'Office 2707B, JBC2 Tower, Cluster V, JLT, Dubai, UAE',
        'contractor_note': 'This document is property of Boilen Energy DMCC. Any unauthorized use, reproduction, or distribution of the document, whether in whole or in part, is expressly prohibited without prior written consent.',
        'project_title': 'BASIC AND DETAIL ENGINEERING DESIGN SERVICES FOR\n70,000 BPD CDU & LPG UNIT FOR MAYSAN REFINERY',
        'document_number': 'B049-BED-MAY-100-EL-CAL-0004',
        'project_number': '2024B049'
    }
if 'revision_history' not in st.session_state:
    st.session_state.revision_history = [
        {'rev': 'A', 'date': '09-Sep-2025', 'purpose': 'ISSUED FOR APPROVAL', 'prpd': 'ASZ', 'revd': 'SHZ', 'appd': 'SHD'}
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
            self.cell(0, 10, 'Electrical Cable Sizing Calculation', 0, 0, 'L')
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
        
        # ROW 1: 3 Columns - Company Logo (50mm), Title (80mm), Contractor Logo (50mm)
        # Column 1 - Company Logo (smaller)
        self.set_xy(15, 15)
        self.rect(15, 15, 50, 25)
        if company_logo_path and os.path.exists(company_logo_path):
            try:
                self.image(company_logo_path, 18, 17, 44, 21)
            except:
                self.set_xy(20, 22)
                self.set_font('Arial', 'B', 8)
                self.cell(40, 10, 'COMPANY', 0, 1, 'C')
        else:
            self.set_xy(20, 22)
            self.set_font('Arial', 'B', 8)
            self.cell(40, 10, 'COMPANY', 0, 1, 'C')
        
        # Column 2 - Project Title (wider)
        self.set_xy(70, 15)  # Start after 50mm + 5mm gap
        self.rect(70, 15, 80, 25)
        self.set_xy(72, 20)
        self.set_font('Arial', 'B', 7)
        self.multi_cell(76, 3.5, st.session_state.project_info['project_title'], 0, 'C')
        
        # Column 3 - Contractor Logo (smaller)
        self.set_xy(155, 15)  # Start after 70+80+5mm
        self.rect(155, 15, 50, 25)
        if contractor_logo_path and os.path.exists(contractor_logo_path):
            try:
                self.image(contractor_logo_path, 158, 17, 44, 21)
            except:
                self.set_xy(160, 22)
                self.set_font('Arial', 'B', 8)
                self.cell(40, 10, 'CONTRACTOR', 0, 1, 'C')
        else:
            self.set_xy(160, 22)
            self.set_font('Arial', 'B', 8)
            self.cell(40, 10, 'CONTRACTOR', 0, 1, 'C')
        
        # ROW 2: 3 Columns - Rev, Title, Date (with same dimensions)
        # Column 1 - Revision
        self.set_xy(15, 45)
        self.rect(15, 45, 50, 15)
        self.set_xy(20, 48)
        self.set_font('Arial', 'B', 10)
        self.cell(40, 8, f"Rev: {st.session_state.cover_details['revision']}", 0, 1, 'C')
        
        # Column 2 - Document Title
        self.set_xy(70, 45)
        self.rect(70, 45, 80, 15)
        self.set_xy(72, 48)
        self.set_font('Arial', 'B', 9)
        self.cell(76, 8, st.session_state.cover_details['title'], 0, 1, 'C')
        
        # Column 3 - Date
        self.set_xy(155, 45)
        self.rect(155, 45, 50, 15)
        self.set_xy(160, 48)
        self.set_font('Arial', 'B', 9)
        self.cell(40, 8, st.session_state.cover_details['date'], 0, 1, 'C')
        
        # Space after boxes
        self.set_y(70)
        
        # Main Title (not needed as we already have document title)
        
        # ===== REVISION LEGEND TABLE - EMPTY BLOCKS FOR FUTURE =====
        self.set_xy(15, 80)
        self.rect(15, 80, 165, 40)
        self.line(70, 80, 70, 120)
        
        # Empty blocks for future expansion - exactly as reference
        self.set_font('Arial', '', 9)
        y_pos = 85
        # Row 1 - Empty
        self.set_xy(20, y_pos)
        self.cell(45, 6, '', 0, 0, 'C')
        self.set_xy(75, y_pos)
        self.cell(100, 6, '', 0, 1, 'L')
        y_pos += 8
        # Row 2 - Empty
        self.set_xy(20, y_pos)
        self.cell(45, 6, '', 0, 0, 'C')
        self.set_xy(75, y_pos)
        self.cell(100, 6, '', 0, 1, 'L')
        y_pos += 8
        # Row 3 - Empty
        self.set_xy(20, y_pos)
        self.cell(45, 6, '', 0, 0, 'C')
        self.set_xy(75, y_pos)
        self.cell(100, 6, '', 0, 1, 'L')
        y_pos += 8
        # Row 4 - Empty
        self.set_xy(20, y_pos)
        self.cell(45, 6, '', 0, 0, 'C')
        self.set_xy(75, y_pos)
        self.cell(100, 6, '', 0, 1, 'L')
        
        # ===== REVISION HISTORY TABLE - EXACTLY AS REFERENCE =====
        self.set_xy(15, 130)
        self.rect(15, 130, 180, 35)
        
        # Table Header
        self.set_xy(17, 133)
        self.set_font('Arial', 'B', 8)
        self.set_fill_color(240, 240, 240)
        
        self.cell(15, 6, 'REV', 1, 0, 'C', 1)
        self.cell(25, 6, 'DATE', 1, 0, 'C', 1)
        self.cell(60, 6, 'ISSUE PURPOSE', 1, 0, 'C', 1)
        self.cell(25, 6, 'PRPD BY', 1, 0, 'C', 1)
        self.cell(25, 6, 'REVD BY', 1, 0, 'C', 1)
        self.cell(25, 6, 'APPD BY', 1, 1, 'C', 1)
        
        # Table Data - with empty cells for future
        self.set_font('Arial', '', 8)
        self.set_xy(17, 139)
        rev = st.session_state.revision_history[0]
        self.cell(15, 6, rev['rev'], 1, 0, 'C')
        self.cell(25, 6, rev['date'], 1, 0, 'C')
        self.cell(60, 6, rev['purpose'], 1, 0, 'C')
        self.cell(25, 6, rev['prpd'], 1, 0, 'C')
        self.cell(25, 6, rev['revd'], 1, 0, 'C')
        self.cell(25, 6, rev['appd'], 1, 1, 'C')
        
        # Empty rows for future revisions (3 empty rows)
        y_pos = 145
        for i in range(3):
            self.set_xy(17, y_pos)
            self.cell(15, 6, '', 1, 0, 'C')
            self.cell(25, 6, '', 1, 0, 'C')
            self.cell(60, 6, '', 1, 0, 'C')
            self.cell(25, 6, '', 1, 0, 'C')
            self.cell(25, 6, '', 1, 0, 'C')
            self.cell(25, 6, '', 1, 1, 'C')
            y_pos += 6
        
        # ===== CONTRACTOR INFORMATION - EXACTLY AS REFERENCE =====
        self.set_y(175)
        self.set_font('Arial', 'B', 10)
        self.cell(0, 5, st.session_state.project_info['contractor'], 0, 1, 'L')
        
        self.set_font('Arial', '', 8)
        self.multi_cell(0, 4, st.session_state.project_info['contractor_address'], 0, 'L')
        self.ln(2)
        self.set_font('Arial', 'I', 7)
        self.multi_cell(0, 3.5, st.session_state.project_info['contractor_note'], 0, 'L')
        
        # ===== DOCUMENT AND PROJECT NUMBERS =====
        self.set_y(220)
        self.set_font('Arial', 'B', 9)
        self.cell(60, 5, 'DOCUMENT NUMBER', 0, 0)
        self.cell(60, 5, '', 0, 0)
        self.cell(60, 5, 'PROJECT NUMBER', 0, 1)
        
        self.set_font('Arial', '', 9)
        self.cell(60, 5, st.session_state.project_info['document_number'], 0, 0)
        self.cell(60, 5, '', 0, 0)
        self.cell(60, 5, st.session_state.project_info['project_number'], 0, 1)
        
        # Page number
        self.set_y(250)
        self.set_font('Arial', 'I', 8)
        self.cell(0, 5, 'Page 1 of 9', 0, 1, 'C')

    def add_calculations(self, results, inputs):
        self.add_page()
        
        self.set_font('Arial', 'B', 16)
        self.set_text_color(0, 51, 102)
        self.cell(0, 10, 'CABLE SIZING CALCULATIONS', 0, 1)
        self.ln(5)
        
        # Placeholder for cable sizing calculations
        self.set_font('Arial', '', 10)
        self.cell(0, 5, 'Cable sizing calculations will be implemented here.', 0, 1)
        self.cell(0, 5, 'This is a placeholder for future development.', 0, 1)

# ========== SIDEBAR ==========
with st.sidebar:
    st.markdown("### ⚡ CES-Electrical Design Calculations")
    st.markdown("---")
    
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
    
    st.markdown("### 📋 Project Information")
    
    st.session_state.project_info['company'] = st.text_input("Company Name", st.session_state.project_info['company'])
    st.session_state.project_info['contractor'] = st.text_input("Contractor Name", st.session_state.project_info['contractor'])
    st.session_state.project_info['contractor_address'] = st.text_input("Contractor Address", st.session_state.project_info['contractor_address'])
    st.session_state.project_info['project_title'] = st.text_area("Project Title", st.session_state.project_info['project_title'], height=60)
    st.session_state.project_info['document_number'] = st.text_input("Document Number", st.session_state.project_info['document_number'])
    st.session_state.project_info['project_number'] = st.text_input("Project Number", st.session_state.project_info['project_number'])
    
    st.markdown("---")
    
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
            st.session_state.cover_details['prepared_by'] = st.text_input("Prepared By", st.session_state.cover_details['prepared_by'])
            st.session_state.cover_details['reviewed_by'] = st.text_input("Reviewed By", st.session_state.cover_details['reviewed_by'])
            st.session_state.cover_details['approved_by'] = st.text_input("Approved By", st.session_state.cover_details['approved_by'])
        
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
            
            with col2:
                st.metric("Rod Diameter", "12.7 mm")
                st.metric("Down Conductor", "50 mm²")
    
    # TAB 4: Calculations
    with lp_tabs[3]:
        st.markdown('<div class="report-header">', unsafe_allow_html=True)
        st.markdown("## DETAILED CALCULATIONS")
        st.markdown('</div>', unsafe_allow_html=True)
        
        if not st.session_state.calc_done:
            st.warning("⚠️ Please complete Risk Assessment first!")
        else:
            results = st.session_state.calc_results
            inputs = st.session_state.input_values
            
            with st.expander("1. Collection Area (Ad)", expanded=True):
                st.markdown('<div class="formula-box">', unsafe_allow_html=True)
                st.markdown("**Formula:** Ad = L × W + 2 × (3H) × (L + W) + π × (3H)²")
                st.markdown("**Reference:** IEC 62305-2 Annex A.2.1.1, Equation A.2")
                if inputs.get('width', 0) == 0:
                    st.markdown(f"**For Column:** Ad = π × 9 × H²")
                st.markdown(f"**Result:** Ad = **{results['ad']:.2f} m²**")
                st.markdown('</div>', unsafe_allow_html=True)
            
            with st.expander("2. Environmental Factor (CD)"):
                st.markdown('<div class="formula-box">', unsafe_allow_html=True)
                st.markdown("**Reference:** IEC 62305-2 Table A.1")
                st.markdown(f"**Result:** CD = **{inputs.get('cd', 1)}**")
                st.markdown('</div>', unsafe_allow_html=True)
            
            with st.expander("3. Lightning Density (NG)"):
                st.markdown('<div class="formula-box">', unsafe_allow_html=True)
                st.markdown("**Formula:** NG = 0.1 × Td")
                st.markdown(f"**Result:** NG = **{results.get('ng', 1)} flashes/km²/year**")
                st.markdown('</div>', unsafe_allow_html=True)
            
            with st.expander("4. Expected Frequency (Nd)"):
                st.markdown('<div class="formula-box">', unsafe_allow_html=True)
                st.markdown("**Formula:** Nd = NG × Ad × CD × 10⁻⁶")
                st.markdown(f"**Result:** Nd = **{results.get('nd', 0):.6f} events/year**")
                st.markdown('</div>', unsafe_allow_html=True)
            
            with st.expander("5. Protection Level"):
                st.markdown('<div class="formula-box">', unsafe_allow_html=True)
                st.markdown(f"**Efficiency:** {results.get('efficiency', 0):.1%}")
                st.markdown(f"**Result:** **{results.get('lpl', 'Class III')}**")
                st.markdown('</div>', unsafe_allow_html=True)
    
    # TAB 5: Revision History
    with lp_tabs[4]:
        st.markdown('<div class="report-header">', unsafe_allow_html=True)
        st.markdown("## REVISION HISTORY")
        st.markdown('</div>', unsafe_allow_html=True)
        
        st.markdown("### Current Revision")
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
                    
                    company_logo_path = ""
                    contractor_logo_path = ""
                    
                    if st.session_state.company_logo is not None:
                        company_logo_path = "temp_company_logo.png"
                        try:
                            st.session_state.company_logo.save(company_logo_path)
                        except:
                            company_logo_path = ""
                    
                    if st.session_state.contractor_logo is not None:
                        contractor_logo_path = "temp_contractor_logo.png"
                        try:
                            st.session_state.contractor_logo.save(contractor_logo_path)
                        except:
                            contractor_logo_path = ""
                    
                    pdf = PDF_Report()
                    pdf.add_title_page(company_logo_path, contractor_logo_path)
                    pdf.add_calculations(st.session_state.calc_results, st.session_state.input_values)
                    
                    if company_logo_path and os.path.exists(company_logo_path):
                        os.remove(company_logo_path)
                    if contractor_logo_path and os.path.exists(contractor_logo_path):
                        os.remove(contractor_logo_path)
                    
                    pdf_output = pdf.output(dest='S')
                    b64 = base64.b64encode(pdf_output).decode()
                    
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
                    
                    st.success("✅ PDF Generated Successfully!")

# Other Calculators
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

st.markdown("---")
st.markdown(f"<div style='text-align: center; color: gray;'>⚡ CES-Electrical Design Calculations | Version 8.0 | {datetime.now().strftime('%Y-%m-%d %H:%M')}</div>", unsafe_allow_html=True)