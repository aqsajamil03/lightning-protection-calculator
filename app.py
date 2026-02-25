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
import PyPDF2
from PyPDF2 import PdfReader, PdfWriter
import img2pdf
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

st.set_page_config(page_title="Professional Engineering Tools", page_icon="⚡", layout="wide")

# Custom CSS - UPDATED with much larger tab font
st.markdown("""
<style>
    .report-header {
        background-color: #1E3A8A;
        color: white;
        padding: 20px;
        border-radius: 10px;
        text-align: center;
        margin-bottom: 20px;
        font-size: 24px;
    }
    .formula-box {
        background-color: #F3F4F6;
        padding: 15px;
        border-radius: 8px;
        border-left: 5px solid #1E3A8A;
        margin: 10px 0;
        font-family: 'Courier New', monospace;
    }
    .success-box {
        background-color: #D4EDDA;
        color: #155724;
        padding: 15px;
        border-radius: 8px;
        border-left: 5px solid #28A745;
        margin: 10px 0;
    }
    .warning-box {
        background-color: #FFF3CD;
        color: #856404;
        padding: 15px;
        border-radius: 8px;
        border-left: 5px solid #FFC107;
        margin: 10px 0;
    }
    .download-btn {
        display: inline-block;
        padding: 12px 24px;
        margin: 10px;
        color: white;
        text-decoration: none;
        border-radius: 5px;
        font-size: 16px;
        font-weight: bold;
        transition: all 0.3s;
        text-align: center;
    }
    .download-btn:hover {
        transform: scale(1.05);
        color: white;
    }
    .pdf-btn {
        background-color: #dc3545;
    }
    .word-btn {
        background-color: #1e3a8a;
    }
    /* MUCH LARGER TAB FONT SIZE */
    .stTabs [data-baseweb="tab-list"] button [data-testid="stMarkdownContainer"] p {
        font-size: 24px !important;
        font-weight: 700 !important;
    }
    .stTabs [data-baseweb="tab-list"] {
        gap: 15px;
    }
    .stTabs [data-baseweb="tab"] {
        padding: 15px 25px;
        background-color: #f0f2f6;
        border-radius: 8px;
    }
    /* Make the tab text even larger on hover */
    .stTabs [data-baseweb="tab"]:hover {
        background-color: #e0e2e6;
        transform: scale(1.02);
    }
</style>
""", unsafe_allow_html=True)

# ========== Word Document Generator Class ==========
class Word_Report:
    def __init__(self):
        self.doc = Document()
        self.doc.core_properties.title = "Lightning Protection Calculation"
        self.doc.core_properties.author = "CES-Electrical"
    
    def add_calculations(self, results, inputs):
        self.doc.add_heading('LIGHTNING PROTECTION CALCULATIONS', 0)
        
        # 1. Collection Area (Ad)
        self.doc.add_heading('1.1 Collection Area (Ad)', level=1)
        self.doc.add_paragraph('Formula: Ad = L × W + 2 × (3H) × (L + W) + π × (3H)²')
        self.doc.add_paragraph('Reference: IEC 62305-2 Annex A.2.1.1, Equation A.2')
        p = self.doc.add_paragraph()
        p.add_run('Result: ').bold = True
        p.add_run(f'Ad = {results["ad"]:.2f} m²')
        
        # 2. Near Strike Collection Area (Am)
        self.doc.add_heading('1.2 Near Strike Collection Area (Am)', level=1)
        self.doc.add_paragraph('Formula: Am = 2 × 500 × (L + W) + π × 500²')
        self.doc.add_paragraph('Reference: IEC 62305-2 Annex A.3, Equation A.7')
        p = self.doc.add_paragraph()
        p.add_run('Result: ').bold = True
        p.add_run(f'Am = {results["am"]:.2f} m²')
        
        # 3. Environmental Factor
        self.doc.add_heading('1.3 Environmental Factor (CD)', level=1)
        self.doc.add_paragraph('Reference: IEC 62305-2 Table A.1')
        self.doc.add_paragraph('• Surrounded by taller structures: CD = 0.25')
        self.doc.add_paragraph('• Similar height structures: CD = 0.5')
        self.doc.add_paragraph('• Isolated structure: CD = 1.0')
        self.doc.add_paragraph('• Hilltop or knoll: CD = 2.0')
        self.doc.add_paragraph(f'Selected Environment: {inputs.get("environment", "Isolated")}')
        p = self.doc.add_paragraph()
        p.add_run('Result: ').bold = True
        p.add_run(f'CD = {inputs.get("cd", 1)}')
        
        # 4. Lightning Density
        self.doc.add_heading('1.4 Lightning Ground Flash Density (NG)', level=1)
        self.doc.add_paragraph('Formula: NG = 0.1 × Td')
        self.doc.add_paragraph('Reference: IEC 62305-2 Annex A.1, Equation A.1')
        p = self.doc.add_paragraph()
        p.add_run('Result: ').bold = True
        p.add_run(f'NG = {results.get("ng", 1)} flashes/km²/year')
        
        # 5. Lightning Frequencies
        self.doc.add_heading('1.5 Lightning Frequencies', level=1)
        self.doc.add_paragraph('Direct Strike Frequency (Nd):')
        self.doc.add_paragraph('Formula: Nd = NG × Ad × CD × 10⁻⁶')
        self.doc.add_paragraph('Reference: IEC 62305-2 Annex A.2.4, Equation A.4')
        p = self.doc.add_paragraph()
        p.add_run('Result: ').bold = True
        p.add_run(f'Nd = {results.get("nd", 0):.6f} events/year')
        
        self.doc.add_paragraph()
        self.doc.add_paragraph('Near Strike Frequency (Nm):')
        self.doc.add_paragraph('Formula: Nm = NG × Am × 10⁻⁶')
        self.doc.add_paragraph('Reference: IEC 62305-2 Annex A.3, Equation A.6')
        p = self.doc.add_paragraph()
        p.add_run('Result: ').bold = True
        p.add_run(f'Nm = {results.get("nm", 0):.6f} events/year')
        
        # 6. Protection Level
        self.doc.add_heading('1.6 Protection Level', level=1)
        self.doc.add_paragraph('Reference: IEC 62305-1 Table 1 and Figure 1')
        self.doc.add_paragraph(f'Protection Efficiency: {results.get("efficiency", 0):.1%}')
        p = self.doc.add_paragraph()
        p.add_run('Result: ').bold = True
        p.add_run(f'{results.get("lpl", "Class III")}')
        self.doc.add_paragraph(f'Rolling Sphere Radius: {results.get("sphere", 45)}m (IEC 62305-3 Table 2)')
        
        # 7. Air Terminals
        self.doc.add_heading('1.7 Air Terminals Required', level=1)
        self.doc.add_paragraph('Method: Rolling Sphere Method')
        self.doc.add_paragraph('Reference: IEC 62305-3 Clause 5.2.2 Table 2')
        p = self.doc.add_paragraph()
        p.add_run('Result: ').bold = True
        p.add_run(f'{results.get("air_terminals", 4)} air terminals required')
        
        # Summary Table
        self.doc.add_page_break()
        self.doc.add_heading('SUMMARY OF RESULTS', level=1)
        
        table = self.doc.add_table(rows=1, cols=2)
        table.style = 'Light Grid Accent 1'
        
        hdr_cells = table.rows[0].cells
        hdr_cells[0].text = 'Parameter'
        hdr_cells[1].text = 'Value'
        
        for cell in hdr_cells:
            cell.paragraphs[0].runs[0].bold = True
        
        summary_data = [
            ('Collection Area (Ad)', f"{results['ad']:.2f} m²"),
            ('Near Strike Area (Am)', f"{results['am']:.2f} m²"),
            ('Environmental Factor (CD)', str(inputs.get('cd', 1))),
            ('Lightning Density (NG)', f"{results.get('ng', 1)} flashes/km²/year"),
            ('Direct Frequency (Nd)', f"{results.get('nd', 0):.6f} events/year"),
            ('Near Frequency (Nm)', f"{results.get('nm', 0):.6f} events/year"),
            ('Protection Efficiency', f"{results.get('efficiency', 0):.1%}"),
            ('Protection Level', results.get('lpl', 'Class III')),
            ('Rolling Sphere Radius', f"{results.get('sphere', 45)} m"),
            ('Air Terminals Required', str(results.get('air_terminals', 4)))
        ]
        
        for param, value in summary_data:
            row_cells = table.add_row().cells
            row_cells[0].text = param
            row_cells[1].text = value
        
        footer = self.doc.add_paragraph()
        footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
        footer.add_run(f'Generated by CES-Electrical Design Calculations on {datetime.now().strftime("%Y-%m-%d %H:%M")}').italic = True
    
    def save(self, filename):
        self.doc.save(filename)

# ========== PDF Report Generator Class ==========
class PDF_Report(FPDF):
    def __init__(self):
        super().__init__()
        self.set_auto_page_break(auto=True, margin=15)
        
    def header(self):
        if self.page_no() > 1:
            self.set_font('Arial', 'I', 10)
            self.cell(0, 12, 'Lightning Protection Calculation', 0, 0, 'L')
            self.cell(0, 12, f'Page {self.page_no()}', 0, 0, 'R')
            self.ln(18)
    
    def footer(self):
        self.set_y(-15)
        self.set_font('Arial', 'I', 8)
        self.cell(0, 10, f'Page {self.page_no()}', 0, 0, 'C')
    
    def add_calculations(self, results, inputs):
        self.add_page()
        
        # Calculations Title
        self.set_font('Arial', 'B', 18)
        self.set_text_color(0, 51, 102)
        self.cell(0, 15, 'LIGHTNING PROTECTION CALCULATIONS', 0, 1, 'C')
        self.ln(8)
        
        # 1. Collection Area (Ad)
        self.set_font('Arial', 'B', 14)
        self.cell(0, 10, '1.1 Collection Area (Ad)', 0, 1)
        self.set_font('Arial', '', 11)
        self.multi_cell(0, 7, 'Formula: Ad = L x W + 2 x (3H) x (L + W) + pi x (3H)^2')
        self.cell(0, 7, 'Reference: IEC 62305-2 Annex A.2.1.1 Equation A.2', 0, 1)
        
        if inputs.get('width', 0) == 0:
            self.cell(0, 7, 'For Column: Ad = pi x 9 x H^2', 0, 1)
            self.cell(0, 7, f'Calculation: Ad = pi x 9 x ({inputs["height"]})^2', 0, 1)
        else:
            self.cell(0, 7, f'Calculation: Ad = {inputs["length"]} x {inputs["width"]} + 2 x (3 x {inputs["height"]}) x ({inputs["length"]} + {inputs["width"]}) + pi x (3 x {inputs["height"]})^2', 0, 1)
        
        self.set_font('Arial', 'B', 12)
        self.cell(0, 8, f'Result: Ad = {results["ad"]:.2f} m^2', 0, 1)
        self.ln(8)
        
        # 2. Near Strike Collection Area (Am)
        self.set_font('Arial', 'B', 14)
        self.cell(0, 10, '1.2 Near Strike Collection Area (Am)', 0, 1)
        self.set_font('Arial', '', 11)
        self.multi_cell(0, 7, 'Formula: Am = 2 x 500 x (L + W) + pi x 500^2')
        self.cell(0, 7, 'Reference: IEC 62305-2 Annex A.3, Equation A.7', 0, 1)
        self.cell(0, 7, f'Calculation: Am = 2 x 500 x ({inputs["length"]} + {inputs["width"]}) + pi x 500^2', 0, 1)
        self.set_font('Arial', 'B', 12)
        self.cell(0, 8, f'Result: Am = {results["am"]:.2f} m^2', 0, 1)
        self.ln(8)
        
        # 3. Environmental Factor
        self.set_font('Arial', 'B', 14)
        self.cell(0, 10, '1.3 Environmental Factor (CD)', 0, 1)
        self.set_font('Arial', '', 11)
        self.cell(0, 7, 'Reference: IEC 62305-2 Table A.1', 0, 1)
        
        self.set_font('Arial', '', 10)
        self.cell(0, 6, 'Surrounded by taller structures  CD = 0.25', 0, 1)
        self.cell(0, 6, 'Similar height structures  CD = 0.5', 0, 1)
        self.cell(0, 6, 'Isolated structure  CD = 1.0', 0, 1)
        self.cell(0, 6, 'Hilltop or knoll  CD = 2.0', 0, 1)
        self.ln(4)
        
        self.set_font('Arial', '', 11)
        self.cell(0, 7, f'Selected Environment: {inputs.get("environment", "Isolated")}', 0, 1)
        self.set_font('Arial', 'B', 12)
        self.cell(0, 8, f'Result: CD = {inputs.get("cd", 1)}', 0, 1)
        self.ln(8)
        
        # Page 2
        self.add_page()
        
        # 4. Lightning Density
        self.set_font('Arial', 'B', 14)
        self.cell(0, 10, '1.4 Lightning Ground Flash Density (NG)', 0, 1)
        self.set_font('Arial', '', 11)
        self.cell(0, 7, 'Formula: NG = 0.1 x Td', 0, 1)
        self.cell(0, 7, 'Reference: IEC 62305-2 Annex A.1 Equation A.1', 0, 1)
        self.cell(0, 7, f'Calculation: NG = 0.1 x {inputs.get("td_days", 10)}', 0, 1)
        self.set_font('Arial', 'B', 12)
        self.cell(0, 8, f'Result: NG = {results.get("ng", 1)} flashes/km^2/year', 0, 1)
        self.ln(8)
        
        # 5. Lightning Frequencies
        self.set_font('Arial', 'B', 14)
        self.cell(0, 10, '1.5 Lightning Frequencies', 0, 1)
        self.set_font('Arial', '', 11)
        
        # Nd
        self.cell(0, 7, 'Direct Strike Frequency (Nd):', 0, 1)
        self.cell(0, 7, 'Formula: Nd = NG x Ad x CD x 10^-6', 0, 1)
        self.cell(0, 7, 'Reference: IEC 62305-2 Annex A.2.4 Equation A.4', 0, 1)
        self.cell(0, 7, f'Calculation: Nd = {results.get("ng", 1)} x {results["ad"]:.0f} x {inputs.get("cd", 1)} x 10^-6', 0, 1)
        self.set_font('Arial', 'B', 12)
        self.cell(0, 8, f'Result: Nd = {results.get("nd", 0):.6f} events/year', 0, 1)
        self.ln(4)
        
        # Nm
        self.set_font('Arial', '', 11)
        self.cell(0, 7, 'Near Strike Frequency (Nm):', 0, 1)
        self.cell(0, 7, 'Formula: Nm = NG x Am x 10^-6', 0, 1)
        self.cell(0, 7, 'Reference: IEC 62305-2 Annex A.3 Equation A.6', 0, 1)
        self.cell(0, 7, f'Calculation: Nm = {results.get("ng", 1)} x {results["am"]:.0f} x 10^-6', 0, 1)
        self.set_font('Arial', 'B', 12)
        self.cell(0, 8, f'Result: Nm = {results.get("nm", 0):.6f} events/year', 0, 1)
        self.ln(8)
        
        # 6. Protection Level
        self.set_font('Arial', 'B', 14)
        self.cell(0, 10, '1.6 Protection Level Determination', 0, 1)
        self.set_font('Arial', '', 11)
        self.cell(0, 7, 'Reference: IEC 62305-1 Table 1 and Figure 1', 0, 1)
        self.cell(0, 7, f'Protection Efficiency: {results.get("efficiency", 0):.1%}', 0, 1)
        
        if results.get("efficiency", 0) > 0.98:
            lpl_text = "Class I (Maximum Protection)"
        elif results.get("efficiency", 0) > 0.95:
            lpl_text = "Class II (High Protection)"
        elif results.get("efficiency", 0) > 0.90:
            lpl_text = "Class III (Standard Protection)"
        else:
            lpl_text = "Class IV (Basic Protection)"
        
        self.set_font('Arial', 'B', 12)
        self.cell(0, 8, f'Result: {lpl_text}', 0, 1)
        self.cell(0, 8, f'Rolling Sphere Radius: {results.get("sphere", 45)}m (IEC 62305-3 Table 2)', 0, 1)
        self.ln(8)
        
        # Page 3
        self.add_page()
        
        # 7. Air Terminals
        self.set_font('Arial', 'B', 14)
        self.cell(0, 10, '1.7 Air Terminals Required', 0, 1)
        self.set_font('Arial', '', 11)
        self.cell(0, 7, 'Method: Rolling Sphere Method', 0, 1)
        self.cell(0, 7, 'Reference: IEC 62305-3 Clause 5.2.2 Table 2', 0, 1)
        
        if inputs.get('height', 0) <= results.get('sphere', 45):
            self.cell(0, 7, 'Using: Protection Width Method', 0, 1)
        else:
            self.cell(0, 7, 'Using: Mesh Method for tall structures', 0, 1)
        
        self.set_font('Arial', 'B', 12)
        self.cell(0, 8, f'Result: {results.get("air_terminals", 4)} air terminals required', 0, 1)
        self.ln(10)
        
        # Summary Section
        self.set_font('Arial', 'B', 16)
        self.set_text_color(0, 51, 102)
        self.cell(0, 12, 'SUMMARY OF RESULTS', 0, 1, 'C')
        self.ln(6)
        
        # Summary Table
        self.set_font('Arial', 'B', 11)
        self.set_fill_color(240, 240, 240)
        self.cell(80, 8, 'Parameter', 1, 0, 'C', 1)
        self.cell(90, 8, 'Value', 1, 1, 'C', 1)
        
        self.set_font('Arial', '', 10)
        summary_data = [
            ('Collection Area (Ad)', f"{results['ad']:.2f} m^2"),
            ('Near Strike Area (Am)', f"{results['am']:.2f} m^2"),
            ('Environmental Factor (CD)', str(inputs.get('cd', 1))),
            ('Lightning Density (NG)', f"{results.get('ng', 1)} flashes/km^2/year"),
            ('Direct Frequency (Nd)', f"{results.get('nd', 0):.6f} events/year"),
            ('Near Frequency (Nm)', f"{results.get('nm', 0):.6f} events/year"),
            ('Protection Efficiency', f"{results.get('efficiency', 0):.1%}"),
            ('Protection Level', results.get('lpl', 'Class III')),
            ('Rolling Sphere Radius', f"{results.get('sphere', 45)} m"),
            ('Air Terminals Required', str(results.get('air_terminals', 4)))
        ]
        
        for param, value in summary_data:
            self.cell(80, 7, param, 1)
            self.cell(90, 7, value, 1)
            self.ln()

# Initialize session state
if 'calc_results' not in st.session_state:
    st.session_state.calc_results = {}
if 'calc_done' not in st.session_state:
    st.session_state.calc_done = False
if 'selected_calculator' not in st.session_state:
    st.session_state.selected_calculator = "⚡ Lightning Protection"
if 'cover_details' not in st.session_state:
    st.session_state.cover_details = {
        'revision': 'A',
        'date': '09 Sep 2025',
    }
if 'project_info' not in st.session_state:
    st.session_state.project_info = {
        'project_title': 'BASIC AND DETAIL ENGINEERING DESIGN SERVICES FOR\n70,000 BPD CDU and LPG UNIT FOR MAYSAN REFINERY',
        'document_number': 'B049-BED-MAY-100-EL-CAL-0004',
        'project_number': '2024B049'
    }
if 'input_values' not in st.session_state:
    st.session_state.input_values = {}

# ========== SIDEBAR ==========
with st.sidebar:
    st.markdown("### ⚡ CES-Electrical Design Calculations")
    st.markdown("---")
    
    # Calculator Navigation
    st.markdown("### 📌 Select Calculator")
    
    calculators = [
        "⚡ Lightning Protection",
        "🔌 Cable Sizing",
        "⚙️ Transformer Sizing",
        "⚡ Generator Sizing",
        "🌍 Earthing System Design",
        "💡 Lighting Calculation",
        "📊 Load Flow Analysis",
        "⚡ Short Circuit",
        "📉 Voltage Drop"
    ]
    
    for calc in calculators:
        if st.button(calc, key=f"nav_{calc}", use_container_width=True):
            st.session_state.selected_calculator = calc
            st.rerun()

# ========== MAIN CONTENT ==========
st.title(f"⚡ {st.session_state.selected_calculator} Calculator")

# ========== LIGHTNING PROTECTION CALCULATOR ==========
if st.session_state.selected_calculator == "⚡ Lightning Protection":
    
    lp_tabs = st.tabs([
        "📊 Risk Assessment", 
        "🔧 Protection Design", 
        "📋 Calculations",
        "📥 Download Report"
    ])
    
    # TAB 1: Risk Assessment
    with lp_tabs[0]:
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
            st.markdown("### 📊 Environmental Factor (CD)")
            cd_values = {"Surrounded": 0.25, "Similar height": 0.5, "Isolated": 1, "Hilltop": 2}
            cd = cd_values[environment]
            
            st.markdown("**IEC 62305-2 Table A.1 Values:**")
            st.markdown("• Surrounded by taller structures: **CD = 0.25**")
            st.markdown("• Similar height structures: **CD = 0.5**")
            st.markdown("• Isolated structure: **CD = 1.0**")
            st.markdown("• Hilltop or knoll: **CD = 2.0**")
            st.markdown("---")
            st.success(f"**Selected: {environment} → CD = {cd}**")
            
            st.markdown("### 📊 Other Coefficients")
            if structure_type == "Column 4-C01":
                c2, c3, c4, c5 = 0.5, 2.0, 3.0, 10.0
            else:
                c2, c3, c4, c5 = 1.0, 3.0, 1.0, 5.0
            
            st.metric("C2 - Type", c2)
            st.metric("C3 - Content", c3)
            st.metric("C4 - Occupancy", c4)
            st.metric("C5 - Consequence", c5)
        
        if st.button("🔧 CALCULATE RISK", type="primary", use_container_width=True):
            
            # Ad Calculation
            if structure_type == "Column 4-C01":
                ad = math.pi * 9 * height**2
            else:
                ad = length * width + 2 * (3 * height) * (length + width) + math.pi * (3 * height)**2
            
            # Am Calculation
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
                'efficiency': efficiency,
                'lpl': lpl, 'sphere': sphere, 'air_terminals': air_terminals
            }
            st.session_state.input_values = {
                'length': length, 'width': width, 'height': height,
                'td_days': td_days, 'environment': environment, 'cd': cd
            }
            st.session_state.calc_done = True
    
    # TAB 2: Protection Design (unchanged)
    with lp_tabs[1]:
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
                if results['lpl'] in ["Class I", "Class II"]:
                    st.metric("Rod Diameter", "12.7 mm")
                    st.metric("Down Conductor", "58 mm²")
                else:
                    st.metric("Rod Diameter", "9.5 mm")
                    st.metric("Down Conductor", "29 mm²")
    
    # TAB 3: Calculations (unchanged)
    with lp_tabs[2]:
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
            
            with st.expander("2. Near Strike Collection Area (Am)", expanded=True):
                st.markdown('<div class="formula-box">', unsafe_allow_html=True)
                st.markdown("**Formula:** Am = 2 × 500 × (L + W) + π × 500²")
                st.markdown("**Reference:** IEC 62305-2 Annex A.3, Equation A.7")
                st.markdown(f"**Calculation:** Am = 2 × 500 × ({inputs['length']} + {inputs['width']}) + π × 500²")
                st.markdown(f"**Result:** Am = **{results['am']:.2f} m²**")
                st.markdown('</div>', unsafe_allow_html=True)
            
            with st.expander("3. Environmental Factor (CD)"):
                st.markdown('<div class="formula-box">', unsafe_allow_html=True)
                st.markdown("**Reference:** IEC 62305-2 Table A.1")
                st.markdown("**Values:**")
                st.markdown("• Surrounded by taller structures: **0.25**")
                st.markdown("• Similar height structures: **0.5**")
                st.markdown("• Isolated structure: **1.0**")
                st.markdown("• Hilltop or knoll: **2.0**")
                st.markdown(f"**Selected:** {inputs.get('environment', 'Isolated')} → **{inputs.get('cd', 1)}**")
                st.markdown('</div>', unsafe_allow_html=True)
            
            with st.expander("4. Lightning Density (NG)"):
                st.markdown('<div class="formula-box">', unsafe_allow_html=True)
                st.markdown("**Formula:** NG = 0.1 × Td")
                st.markdown(f"**Result:** NG = **{results.get('ng', 1)} flashes/km²/year**")
                st.markdown('</div>', unsafe_allow_html=True)
            
            with st.expander("5. Lightning Frequencies"):
                st.markdown('<div class="formula-box">', unsafe_allow_html=True)
                st.markdown("**Nd (Direct Strike Frequency):**")
                st.markdown("Formula: Nd = NG × Ad × CD × 10⁻⁶")
                st.markdown(f"Result: **{results.get('nd', 0):.6f} events/year**")
                st.markdown("---")
                st.markdown("**Nm (Near Strike Frequency):**")
                st.markdown("Formula: Nm = NG × Am × 10⁻⁶")
                st.markdown(f"Result: **{results.get('nm', 0):.6f} events/year**")
                st.markdown('</div>', unsafe_allow_html=True)
            
            with st.expander("6. Protection Level"):
                st.markdown('<div class="formula-box">', unsafe_allow_html=True)
                st.markdown(f"**Efficiency:** {results.get('efficiency', 0):.1%}")
                st.markdown(f"**Result:** **{results.get('lpl', 'Class III')}**")
                st.markdown('</div>', unsafe_allow_html=True)
    
    # TAB 4: Download Report - WITHOUT FRONT PAGE
    with lp_tabs[3]:
        st.markdown('<div class="report-header">', unsafe_allow_html=True)
        st.markdown("## DOWNLOAD REPORT")
        st.markdown('</div>', unsafe_allow_html=True)
        
        if not st.session_state.calc_done:
            st.warning("⚠️ Please complete Risk Assessment first!")
        else:
            st.markdown("### 📥 Select Format")
            
            col1, col2 = st.columns(2)
            
            with col1:
                st.markdown("#### 📄 PDF Format")
                if st.button("📥 Generate PDF", key="pdf_btn", use_container_width=True):
                    with st.spinner("Generating PDF report..."):
                        pdf = PDF_Report()
                        # NO TITLE PAGE - Direct calculations
                        pdf.add_calculations(st.session_state.calc_results, st.session_state.input_values)
                        
                        pdf_output = pdf.output(dest='S')
                        b64 = base64.b64encode(pdf_output).decode()
                        
                        filename = f"LPS_Report_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf"
                        page_count = "3 pages (Calculations only)"
                        
                        st.markdown(f'<a href="data:application/pdf;base64,{b64}" download="{filename}" class="download-btn pdf-btn">📥 Click to Download PDF ({page_count})</a>', unsafe_allow_html=True)
                        st.success("✅ PDF generated successfully!")
            
            with col2:
                st.markdown("#### 📝 Word Format")
                if st.button("📥 Generate Word", key="word_btn", use_container_width=True):
                    with st.spinner("Generating Word report..."):
                        try:
                            word = Word_Report()
                            # NO TITLE PAGE - Direct calculations
                            word.add_calculations(st.session_state.calc_results, st.session_state.input_values)
                            
                            word_path = "temp_report.docx"
                            word.save(word_path)
                            
                            with open(word_path, "rb") as f:
                                word_bytes = f.read()
                            
                            b64 = base64.b64encode(word_bytes).decode()
                            
                            filename = f"LPS_Report_{datetime.now().strftime('%Y%m%d_%H%M')}.docx"
                            
                            if os.path.exists(word_path):
                                os.remove(word_path)
                            
                            st.markdown(f'<a href="data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,{b64}" download="{filename}" class="download-btn word-btn">📥 Click to Download Word</a>', unsafe_allow_html=True)
                            st.success("✅ Word document generated successfully!")
                        except Exception as e:
                            st.error(f"Error generating Word document: {str(e)}")
                            st.info("Please make sure python-docx is installed: `pip install python-docx`")

# ========== OTHER CALCULATORS (Placeholders) ==========
elif st.session_state.selected_calculator == "🔌 Cable Sizing":
    st.markdown('<div class="report-header">', unsafe_allow_html=True)
    st.markdown("## CABLE SIZING CALCULATOR")
    st.markdown('</div>', unsafe_allow_html=True)
    st.info("🔧 Cable sizing calculator will be implemented in next phase")

elif st.session_state.selected_calculator == "⚙️ Transformer Sizing":
    st.markdown('<div class="report-header">', unsafe_allow_html=True)
    st.markdown("## TRANSFORMER SIZING CALCULATOR")
    st.markdown('</div>', unsafe_allow_html=True)
    st.info("⚙️ Transformer sizing calculator will be implemented in next phase")

elif st.session_state.selected_calculator == "⚡ Generator Sizing":
    st.markdown('<div class="report-header">', unsafe_allow_html=True)
    st.markdown("## GENERATOR SIZING CALCULATOR")
    st.markdown('</div>', unsafe_allow_html=True)
    st.info("⚡ Generator sizing calculator will be implemented in next phase")

elif st.session_state.selected_calculator == "🌍 Earthing System Design":
    st.markdown('<div class="report-header">', unsafe_allow_html=True)
    st.markdown("## EARTHING SYSTEM DESIGN")
    st.markdown('</div>', unsafe_allow_html=True)
    st.info("🌍 Earthing system design calculator will be implemented in next phase")

elif st.session_state.selected_calculator == "💡 Lighting Calculation":
    st.markdown('<div class="report-header">', unsafe_allow_html=True)
    st.markdown("## LIGHTING CALCULATION")
    st.markdown('</div>', unsafe_allow_html=True)
    st.info("💡 Lighting calculation will be implemented in next phase")

elif st.session_state.selected_calculator == "📊 Load Flow Analysis":
    st.markdown('<div class="report-header">', unsafe_allow_html=True)
    st.markdown("## LOAD FLOW ANALYSIS")
    st.markdown('</div>', unsafe_allow_html=True)
    st.info("📊 Load flow analysis will be implemented in next phase")

elif st.session_state.selected_calculator == "⚡ Short Circuit":
    st.markdown('<div class="report-header">', unsafe_allow_html=True)
    st.markdown("## SHORT CIRCUIT CALCULATION")
    st.markdown('</div>', unsafe_allow_html=True)
    st.info("⚡ Short circuit calculation will be implemented in next phase")

elif st.session_state.selected_calculator == "📉 Voltage Drop":
    st.markdown('<div class="report-header">', unsafe_allow_html=True)
    st.markdown("## VOLTAGE DROP CALCULATOR")
    st.markdown('</div>', unsafe_allow_html=True)
    st.info("📉 Voltage drop calculator will be implemented in next phase")

# Footer
st.markdown("---")
st.markdown(f"<div style='text-align: center; color: gray;'>⚡ CES-Electrical Design Calculations | Version 32.0 | {datetime.now().strftime('%Y-%m-%d %H:%M')}</div>", unsafe_allow_html=True)