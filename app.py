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
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

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
    .reference-box {
        background-color: #E8F5E9;
        padding: 10px;
        border-radius: 5px;
        border-left: 3px solid #4CAF50;
        margin: 5px 0;
        font-size: 0.9em;
    }
    .calculation-detail {
        background-color: #F5F5F5;
        padding: 15px;
        border-radius: 8px;
        border: 1px solid #ddd;
        margin: 10px 0;
        font-family: 'Courier New', monospace;
    }
    .param-highlight {
        background-color: #FFE5B4;
        padding: 2px 5px;
        border-radius: 3px;
        font-weight: bold;
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
    .info-box {
        background-color: #E7F3FF;
        color: #004085;
        padding: 15px;
        border-radius: 8px;
        border-left: 5px solid #1E3A8A;
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
    .stTabs [data-baseweb="tab"]:hover {
        background-color: #e0e2e6;
        transform: scale(1.02);
    }
    .parameter-table {
        width: 100%;
        border-collapse: collapse;
        margin: 20px 0;
    }
    .parameter-table th {
        background-color: #1E3A8A;
        color: white;
        padding: 10px;
        text-align: center;
        font-weight: bold;
    }
    .parameter-table td {
        border: 1px solid #ddd;
        padding: 8px;
        text-align: left;
    }
    .parameter-table tr:nth-child(even) {
        background-color: #f2f2f2;
    }
    .parameter-table tr:nth-child(odd) {
        background-color: white;
    }
    .parameter-table td {
        color: black !important;
    }
</style>
""", unsafe_allow_html=True)

# ========== LIGHTNING PROTECTION CALCULATOR - EXISTING CLASSES ==========
class LightningWordReport:
    def __init__(self):
        self.doc = Document()
        self.doc.core_properties.title = "Lightning Protection Calculation"
        self.doc.core_properties.author = "CES-Electrical"
    
    def add_calculations(self, results, inputs):
        self.doc.add_heading('LIGHTNING PROTECTION CALCULATIONS', 0)
        
        # 1. Collection Area (Ad)
        self.doc.add_heading('1.1 Collection Area (Ad)', level=1)
        self.doc.add_paragraph('Formula: Ad = L x W + 2 x (3H) x (L + W) + pi x (3H)^2')
        self.doc.add_paragraph('Reference: IEC 62305-2 Annex A.2.1.1, Equation A.2')
        p = self.doc.add_paragraph()
        p.add_run('Result: ').bold = True
        p.add_run(f'Ad = {results["ad"]:.2f} m²')
        
        # 2. Near Strike Collection Area (Am)
        self.doc.add_heading('1.2 Near Strike Collection Area (Am)', level=1)
        self.doc.add_paragraph('Formula: Am = 2 x 500 x (L + W) + pi x 500^2')
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
        self.doc.add_paragraph('Formula: NG = 0.1 x Td')
        self.doc.add_paragraph('Reference: IEC 62305-2 Annex A.1, Equation A.1')
        p = self.doc.add_paragraph()
        p.add_run('Result: ').bold = True
        p.add_run(f'NG = {results.get("ng", 1)} flashes/km²/year')
        
        # 5. Lightning Frequencies
        self.doc.add_heading('1.5 Lightning Frequencies', level=1)
        self.doc.add_paragraph('Direct Strike Frequency (Nd):')
        self.doc.add_paragraph('Formula: Nd = NG x Ad x CD x 10^-6')
        self.doc.add_paragraph('Reference: IEC 62305-2 Annex A.2.4, Equation A.4')
        p = self.doc.add_paragraph()
        p.add_run('Result: ').bold = True
        p.add_run(f'Nd = {results.get("nd", 0):.6f} events/year')
        
        self.doc.add_paragraph()
        self.doc.add_paragraph('Near Strike Frequency (Nm):')
        self.doc.add_paragraph('Formula: Nm = NG x Am x 10^-6')
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

class LightningPDFReport(FPDF):
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

# ========== CIRCUIT BREAKER DATA AND CALCULATIONS ==========
# Circuit Breaker Standard Ratings (IEC 60898)
CB_RATINGS = [6, 10, 16, 20, 25, 32, 40, 50, 63, 80, 100, 125, 160, 200, 250, 315, 400, 500, 630, 800, 1000, 1250, 1600, 2000, 2500, 3200, 4000, 5000, 6300]

# Breaker Types
BREAKER_TYPES = {
    'MCB': {'min': 6, 'max': 125, 'poles': ['1P', '2P', '3P', '4P']},
    'MCCB': {'min': 125, 'max': 1600, 'poles': ['3P', '4P']},
    'ACB': {'min': 1600, 'max': 6300, 'poles': ['3P', '4P']}
}

# Manufacturers
MANUFACTURERS = {
    'Schneider Electric': {
        'MCB': 'Acti9 series',
        'MCCB': 'EasyPact EVC series',
        'ACB': 'MasterPact MTZ series'
    },
    'Siemens': {
        'MCB': '5SY series',
        'MCCB': '3VA series',
        'ACB': '3WL series'
    },
    'ABB': {
        'MCB': 'S200 series',
        'MCCB': 'Tmax XT series',
        'ACB': 'Emax 2 series'
    },
    'Legrand': {
        'MCB': 'DX³ series',
        'MCCB': 'DPX³ series',
        'ACB': 'DMX³ series'
    }
}

# ========== CABLE DATABASE ==========
# 4-Core Cable Data (Three Phase)
CABLE_4CORE_DATA = {
    'copper': {
        1.5: {'R': 12.1, 'X': 0.093, 'ampacity': 22, 'diameter': 2.9},
        2.5: {'R': 7.41, 'X': 0.088, 'ampacity': 30, 'diameter': 3.7},
        4: {'R': 4.61, 'X': 0.088, 'ampacity': 40, 'diameter': 4.3},
        6: {'R': 3.08, 'X': 0.088, 'ampacity': 51, 'diameter': 5.0},
        10: {'R': 1.83, 'X': 0.084, 'ampacity': 70, 'diameter': 6.4},
        16: {'R': 1.15, 'X': 0.077, 'ampacity': 94, 'diameter': 7.8},
        25: {'R': 0.727, 'X': 0.074, 'ampacity': 123, 'diameter': 9.8},
        35: {'R': 0.524, 'X': 0.074, 'ampacity': 150, 'diameter': 11.0},
        50: {'R': 0.387, 'X': 0.071, 'ampacity': 181, 'diameter': 12.8},
        70: {'R': 0.268, 'X': 0.069, 'ampacity': 228, 'diameter': 15.1},
        95: {'R': 0.193, 'X': 0.068, 'ampacity': 276, 'diameter': 17.2},
        120: {'R': 0.153, 'X': 0.067, 'ampacity': 318, 'diameter': 19.1},
        150: {'R': 0.124, 'X': 0.066, 'ampacity': 364, 'diameter': 21.1},
        185: {'R': 0.0991, 'X': 0.066, 'ampacity': 415, 'diameter': 23.2},
        240: {'R': 0.0754, 'X': 0.065, 'ampacity': 492, 'diameter': 26.3},
        300: {'R': 0.0601, 'X': 0.064, 'ampacity': 567, 'diameter': 29.5},
    },
    'aluminium': {
        16: {'R': 1.91, 'X': 0.077, 'ampacity': 72, 'diameter': 7.8},
        25: {'R': 1.20, 'X': 0.074, 'ampacity': 94, 'diameter': 9.8},
        35: {'R': 0.868, 'X': 0.074, 'ampacity': 115, 'diameter': 11.0},
        50: {'R': 0.641, 'X': 0.071, 'ampacity': 140, 'diameter': 12.8},
        70: {'R': 0.443, 'X': 0.069, 'ampacity': 177, 'diameter': 15.1},
        95: {'R': 0.320, 'X': 0.068, 'ampacity': 215, 'diameter': 17.2},
        120: {'R': 0.253, 'X': 0.067, 'ampacity': 247, 'diameter': 19.1},
        150: {'R': 0.206, 'X': 0.066, 'ampacity': 283, 'diameter': 21.1},
        185: {'R': 0.164, 'X': 0.066, 'ampacity': 322, 'diameter': 23.2},
        240: {'R': 0.125, 'X': 0.065, 'ampacity': 382, 'diameter': 26.3},
        300: {'R': 0.100, 'X': 0.064, 'ampacity': 440, 'diameter': 29.5},
    }
}

# 2-Core Cable Data (Single Phase)
CABLE_2CORE_DATA = {
    'copper': {
        1.5: {'R': 12.1, 'X': 0.093, 'ampacity': 22, 'diameter': 2.9},
        2.5: {'R': 7.41, 'X': 0.088, 'ampacity': 30, 'diameter': 3.7},
        4: {'R': 4.61, 'X': 0.088, 'ampacity': 40, 'diameter': 4.3},
        6: {'R': 3.08, 'X': 0.088, 'ampacity': 51, 'diameter': 5.0},
        10: {'R': 1.83, 'X': 0.084, 'ampacity': 70, 'diameter': 6.4},
        16: {'R': 1.15, 'X': 0.077, 'ampacity': 94, 'diameter': 7.8},
        25: {'R': 0.727, 'X': 0.074, 'ampacity': 123, 'diameter': 9.8},
        35: {'R': 0.524, 'X': 0.074, 'ampacity': 150, 'diameter': 11.0},
        50: {'R': 0.387, 'X': 0.071, 'ampacity': 181, 'diameter': 12.8},
        70: {'R': 0.268, 'X': 0.069, 'ampacity': 228, 'diameter': 15.1},
        95: {'R': 0.193, 'X': 0.068, 'ampacity': 276, 'diameter': 17.2},
        120: {'R': 0.153, 'X': 0.067, 'ampacity': 318, 'diameter': 19.1},
        150: {'R': 0.124, 'X': 0.066, 'ampacity': 364, 'diameter': 21.1},
        185: {'R': 0.0991, 'X': 0.066, 'ampacity': 415, 'diameter': 23.2},
    },
    'aluminium': {
        16: {'R': 1.91, 'X': 0.077, 'ampacity': 72, 'diameter': 7.8},
        25: {'R': 1.20, 'X': 0.074, 'ampacity': 94, 'diameter': 9.8},
        35: {'R': 0.868, 'X': 0.074, 'ampacity': 115, 'diameter': 11.0},
        50: {'R': 0.641, 'X': 0.071, 'ampacity': 140, 'diameter': 12.8},
        70: {'R': 0.443, 'X': 0.069, 'ampacity': 177, 'diameter': 15.1},
        95: {'R': 0.320, 'X': 0.068, 'ampacity': 215, 'diameter': 17.2},
        120: {'R': 0.253, 'X': 0.067, 'ampacity': 247, 'diameter': 19.1},
        150: {'R': 0.206, 'X': 0.066, 'ampacity': 283, 'diameter': 21.1},
        185: {'R': 0.164, 'X': 0.066, 'ampacity': 322, 'diameter': 23.2},
    }
}

# Derating Factors Tables (Based on IEC 60502)
TEMPERATURE_FACTORS = {
    90: {20: 1.07, 25: 1.04, 30: 1.00, 35: 0.96, 40: 0.91, 45: 0.87, 50: 0.82, 55: 0.76, 60: 0.71},
    70: {20: 1.08, 25: 1.04, 30: 1.00, 35: 0.94, 40: 0.87, 45: 0.79, 50: 0.71, 55: 0.61, 60: 0.50}
}

GROUPING_FACTORS = {
    'touching': {1: 1.00, 2: 0.80, 3: 0.70, 4: 0.65, 5: 0.60, 6: 0.57, 7: 0.54, 8: 0.52, 9: 0.50, 12: 0.45, 16: 0.41},
    'spaced': {1: 1.00, 2: 0.85, 3: 0.79, 4: 0.75, 5: 0.73, 6: 0.72, 7: 0.72, 8: 0.71, 9: 0.70, 12: 0.70, 16: 0.70}
}

# ========== CABLE SIZING CALCULATOR CLASS ==========
class CableSizingCalculator:
    def __init__(self):
        self.results = {}
    
    def calculate_load_current(self, power_kw, voltage_v, pf, efficiency=1.0, phase='3-phase'):
        """Calculate load current based on power and voltage
        Reference: IEC 60364-5-52 Section 523"""
        if phase == '3-phase':
            return (power_kw * 1000) / (1.732 * voltage_v * pf * efficiency)
        elif phase == '1-phase':
            return (power_kw * 1000) / (voltage_v * pf * efficiency)
        else:  # DC
            return (power_kw * 1000) / voltage_v
    
    def get_derating_factors(self, temp_c, insulation_temp=90, num_cables=1, grouping='touching'):
        """Calculate derating factors based on IEC 60502-2 Tables B.10-B.22"""
        k1 = TEMPERATURE_FACTORS[insulation_temp].get(temp_c, 1.0)
        k2 = GROUPING_FACTORS[grouping].get(min(num_cables, 16), 0.5)
        total_k = k1 * k2
        return total_k, {'k1': k1, 'k2': k2}
    
    def calculate_voltage_drop(self, current, length_m, R, X, pf, voltage_v, phase='3-phase'):
        """Calculate voltage drop based on IEC 60364-5-52 Section 525"""
        R_total = R * length_m / 1000
        X_total = X * length_m / 1000
        
        if phase == '3-phase':
            Vd = 1.732 * current * (R_total * pf + X_total * math.sin(math.acos(pf)))
        elif phase == '1-phase':
            Vd = 2 * current * (R_total * pf + X_total * math.sin(math.acos(pf)))
        else:  # DC
            Vd = 2 * current * R_total
        
        Vd_percent = (Vd / voltage_v) * 100
        return Vd, Vd_percent
    
    def get_cable_by_phase(self, phase, material):
        """Get appropriate cable database based on phase"""
        if phase == '3-phase':
            return CABLE_4CORE_DATA[material], '4-Core'
        else:
            return CABLE_2CORE_DATA[material], '2-Core'

# ========== CIRCUIT BREAKER CALCULATOR CLASS ==========
class CircuitBreakerCalculator:
    def __init__(self):
        pass
    
    def get_standard_rating(self, current, design_factor=1.25):
        """Get next higher standard CB rating (IEC 60898)"""
        required = current * design_factor
        for rating in CB_RATINGS:
            if rating >= required:
                return rating, required
        return CB_RATINGS[-1], required
    
    def get_breaker_type(self, rating):
        """Determine breaker type based on rating"""
        if rating <= 125:
            return 'MCB'
        elif rating <= 1600:
            return 'MCCB'
        else:
            return 'ACB'
    
    def get_poles(self, phase, rating):
        """Determine number of poles"""
        if phase == '1-phase':
            return '2P'
        elif phase == '3-phase':
            return '3P'
        else:
            return '2P'
    
    def calculate_cb_size(self, loads_df, design_factor=1.25, manufacturer='Schneider Electric'):
        """Calculate CB sizes for multiple loads"""
        results = []
        for idx, load in loads_df.iterrows():
            # Calculate load current
            if load['Phase'] == '3-phase':
                current = load['Power (kW)'] * 1000 / (1.732 * load['Voltage (V)'] * load['Power Factor'])
            elif load['Phase'] == '1-phase':
                current = load['Power (kW)'] * 1000 / (load['Voltage (V)'] * load['Power Factor'])
            else:  # DC
                current = load['Power (kW)'] * 1000 / load['Voltage (V)']
            
            # Get standard CB rating
            rating, required = self.get_standard_rating(current, design_factor)
            breaker_type = self.get_breaker_type(rating)
            poles = self.get_poles(load['Phase'], rating)
            
            # Get manufacturer series
            series = MANUFACTURERS[manufacturer][breaker_type]
            
            results.append({
                'Load': load['Load Name'],
                'Power (kW)': load['Power (kW)'],
                'Current (A)': current,
                'Required CB (A)': required,
                'Selected CB (A)': rating,
                'Breaker Type': breaker_type,
                'Poles': poles,
                'Manufacturer': manufacturer,
                'Series': series
            })
        
        return results
    
    def calculate_main_cb(self, loads_df, voltage=400, pf=0.8, design_factor=1.25):
        """Calculate main circuit breaker"""
        total_power = loads_df['Power (kW)'].sum()
        current = total_power * 1000 / (1.732 * voltage * pf)
        rating, required = self.get_standard_rating(current, design_factor)
        breaker_type = self.get_breaker_type(rating)
        poles = self.get_poles('3-phase', rating)
        
        return {
            'total_power': total_power,
            'current': current,
            'required_cb': required,
            'selected_cb': rating,
            'breaker_type': breaker_type,
            'poles': poles
        }

# ========== SESSION STATE INITIALIZATION ==========
if 'calc_results' not in st.session_state:
    st.session_state.calc_results = {}
if 'calc_done' not in st.session_state:
    st.session_state.calc_done = False
if 'selected_calculator' not in st.session_state:
    st.session_state.selected_calculator = "⚡ Lightning Protection"
if 'cable_results' not in st.session_state:
    st.session_state.cable_results = {}
if 'cable_all_cables' not in st.session_state:
    st.session_state.cable_all_cables = []
if 'project_info' not in st.session_state:
    st.session_state.project_info = {
        'project_title': 'BASIC AND DETAIL ENGINEERING DESIGN SERVICES FOR\n70,000 BPD CDU and LPG UNIT FOR MAYSAN REFINERY',
        'document_number': 'B049-BED-MAY-100-EL-CAL-0004',
        'project_number': '2024B049'
    }
if 'cover_details' not in st.session_state:
    st.session_state.cover_details = {'revision': 'A', 'date': '09 Sep 2025'}
if 'input_values' not in st.session_state:
    st.session_state.input_values = {}

# Initialize loads dataframe with 1 default row
if 'loads_df' not in st.session_state:
    st.session_state.loads_df = pd.DataFrame({
        'Load Name': ['Load 1'],
        'Power (kW)': [5.0],
        'Voltage (V)': [400],
        'Phase': ['3-phase'],
        'Power Factor': [0.85],
        'Length (m)': [50]
    })

# ========== SIDEBAR ==========
with st.sidebar:
    st.markdown("### ⚡ CES-Electrical Design Calculations")
    st.markdown("---")
    
    # Calculator Navigation
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
    
    # TAB 2: Protection Design
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
    
    # TAB 3: Calculations
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
    
    # TAB 4: Download Report
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
                if st.button("📥 Generate PDF", key="lp_pdf_btn", use_container_width=True):
                    with st.spinner("Generating PDF report..."):
                        pdf = LightningPDFReport()
                        pdf.add_calculations(st.session_state.calc_results, st.session_state.input_values)
                        pdf_output = pdf.output(dest='S')
                        b64 = base64.b64encode(pdf_output).decode()
                        filename = f"Lightning_Report_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf"
                        st.markdown(f'<a href="data:application/pdf;base64,{b64}" download="{filename}" class="download-btn pdf-btn">📥 Click to Download PDF</a>', unsafe_allow_html=True)
                        st.success("✅ PDF generated successfully!")
            
            with col2:
                st.markdown("#### 📝 Word Format")
                if st.button("📥 Generate Word", key="lp_word_btn", use_container_width=True):
                    with st.spinner("Generating Word report..."):
                        word = LightningWordReport()
                        word.add_calculations(st.session_state.calc_results, st.session_state.input_values)
                        word_path = "temp_lightning_report.docx"
                        word.save(word_path)
                        with open(word_path, "rb") as f:
                            word_bytes = f.read()
                        b64 = base64.b64encode(word_bytes).decode()
                        filename = f"Lightning_Report_{datetime.now().strftime('%Y%m%d_%H%M')}.docx"
                        if os.path.exists(word_path):
                            os.remove(word_path)
                        st.markdown(f'<a href="data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,{b64}" download="{filename}" class="download-btn word-btn">📥 Click to Download Word</a>', unsafe_allow_html=True)
                        st.success("✅ Word document generated successfully!")

# ========== CABLE SIZING CALCULATOR - COMPLETE FIXED VERSION ==========
elif st.session_state.selected_calculator == "🔌 Cable Sizing":
    
    cable_tabs = st.tabs([
        "📥 Loads Table", 
        "📊 Detailed Calculations", 
        "⚡ Circuit Breakers",
        "📥 Download Report"
    ])
    
    # TAB 1: LOADS TABLE INPUT - DYNAMIC ROWS
    with cable_tabs[0]:
        st.markdown('<div class="report-header">', unsafe_allow_html=True)
        st.markdown("## MULTIPLE LOADS INPUT")
        st.markdown('</div>', unsafe_allow_html=True)
        
        st.markdown('<div class="info-box">', unsafe_allow_html=True)
        st.markdown("### 📋 Enter Load Details")
        st.markdown("""
        - **Add/Delete Rows:** Use the buttons at the bottom of the table
        - **Enter any values:** No dropdown restrictions
        - **Phase:** Enter '1-phase', '3-phase', or 'DC'
        - **Single-phase loads** use 2-core cables
        - **Three-phase loads** use 4-core cables
        """)
        st.markdown('</div>', unsafe_allow_html=True)
        
        # Add row button
        col1, col2, col3 = st.columns([1, 1, 5])
        with col1:
            if st.button("➕ Add Row", use_container_width=True):
                new_row = pd.DataFrame({
                    'Load Name': [f'Load {len(st.session_state.loads_df) + 1}'],
                    'Power (kW)': [10.0],
                    'Voltage (V)': [400],
                    'Phase': ['3-phase'],
                    'Power Factor': [0.85],
                    'Length (m)': [50]
                })
                st.session_state.loads_df = pd.concat([st.session_state.loads_df, new_row], ignore_index=True)
                st.rerun()
        
        with col2:
            if st.button("🗑️ Delete Last Row", use_container_width=True):
                if len(st.session_state.loads_df) > 1:
                    st.session_state.loads_df = st.session_state.loads_df[:-1]
                    st.rerun()
                else:
                    st.warning("At least one row required")
        
        # Editable dataframe for loads - COMPLETELY FREE TEXT
        edited_df = st.data_editor(
            st.session_state.loads_df,
            num_rows="fixed",
            use_container_width=True,
            column_config={
                "Load Name": st.column_config.TextColumn("Load Name", width="medium"),
                "Power (kW)": st.column_config.NumberColumn("Power (kW)", min_value=0.0, max_value=100000.0, step=0.1, format="%.1f"),
                "Voltage (V)": st.column_config.NumberColumn("Voltage (V)", min_value=0.0, max_value=100000.0, step=1.0, format="%.0f"),
                "Phase": st.column_config.TextColumn("Phase", help="Enter '1-phase', '3-phase', or 'DC'"),
                "Power Factor": st.column_config.NumberColumn("PF", min_value=0.0, max_value=1.0, step=0.05, format="%.2f"),
                "Length (m)": st.column_config.NumberColumn("Length (m)", min_value=0.0, max_value=100000.0, step=1.0, format="%.0f")
            }
        )
        
        st.session_state.loads_df = edited_df
        
        st.markdown("### ⚙️ Installation Parameters")
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            ambient_temp = st.number_input("Ambient Temperature (°C)", value=55.0, min_value=0.0, max_value=100.0, step=1.0)
        with col2:
            num_cables = st.number_input("Number of Cables in Group", value=6, min_value=1, max_value=100, step=1)
        with col3:
            grouping = st.selectbox("Grouping Configuration", ["touching", "spaced"])
        with col4:
            voltage_drop_limit = st.number_input("Max Voltage Drop (%)", value=3.0, min_value=0.1, max_value=20.0, step=0.1)
        
        col5, col6, col7 = st.columns(3)
        with col5:
            material = st.selectbox("Conductor Material", ["copper", "aluminium"])
        with col6:
            manufacturer = st.selectbox("Manufacturer", list(MANUFACTURERS.keys()))
        with col7:
            design_factor = st.number_input("Design Factor (CB Sizing)", value=1.25, min_value=1.0, max_value=2.0, step=0.05)
        
        if st.button("🔧 CALCULATE ALL", type="primary", use_container_width=True):
            with st.spinner("Calculating..."):
                st.session_state.cable_calculated = True
                
                # Store parameters in session state
                st.session_state.ambient_temp = ambient_temp
                st.session_state.num_cables = num_cables
                st.session_state.grouping = grouping
                st.session_state.voltage_drop_limit = voltage_drop_limit
                st.session_state.material = material
                st.session_state.manufacturer = manufacturer
                st.session_state.design_factor = design_factor
                
                # Calculate CB sizes
                cb_calc = CircuitBreakerCalculator()
                st.session_state.cb_results = cb_calc.calculate_cb_size(
                    st.session_state.loads_df, design_factor, manufacturer
                )
                st.session_state.main_cb = cb_calc.calculate_main_cb(
                    st.session_state.loads_df, 400, 0.85, design_factor
                )
                
                # Calculate cable sizes
                cable_calc = CableSizingCalculator()
                total_k, factors = cable_calc.get_derating_factors(ambient_temp, 90, num_cables, grouping)
                
                # Store derating factor in session state
                st.session_state.derating_factor = total_k
                st.session_state.derating_factors = factors
                
                # Calculate for each load
                cable_results = []
                detailed_calculations = []
                
                for idx, load in st.session_state.loads_df.iterrows():
                    # Skip if invalid phase
                    if load['Phase'] not in ['1-phase', '3-phase', 'DC']:
                        continue
                        
                    # Calculate load current
                    current = cable_calc.calculate_load_current(
                        load['Power (kW)'], load['Voltage (V)'], load['Power Factor'], 1.0, load['Phase']
                    )
                    
                    # Get cable database based on phase
                    cable_db, cable_type = cable_calc.get_cable_by_phase(load['Phase'], material)
                    
                    # Find suitable cable
                    selected_cable = None
                    for size, data in cable_db.items():
                        derated = data['ampacity'] * total_k
                        if derated >= current:
                            vd_v, vd_pct = cable_calc.calculate_voltage_drop(
                                current, load['Length (m)'], data['R'], data['X'],
                                load['Power Factor'], load['Voltage (V)'], load['Phase']
                            )
                            selected_cable = {
                                'Load Name': load['Load Name'],
                                'Power (kW)': load['Power (kW)'],
                                'Voltage (V)': load['Voltage (V)'],
                                'Phase': load['Phase'],
                                'PF': load['Power Factor'],
                                'Length (m)': load['Length (m)'],
                                'Current (A)': f"{current:.2f}",
                                'Cable Size': f"{size} mm²",
                                'Type': cable_type,
                                'Base Ampacity': data['ampacity'],
                                'Derated Ampacity': f"{derated:.1f}",
                                'Voltage Drop %': f"{vd_pct:.3f}",
                                'Status': 'PASS' if vd_pct <= voltage_drop_limit else 'FAIL',
                                'R (ohm/km)': f"{data['R']:.3f}",
                                'X (ohm/km)': f"{data['X']:.3f}"
                            }
                            
                            # Store detailed calculation for this load
                            detail = f"""
### Load: {load['Load Name']}
**Step 1: Load Current Calculation**
Formula: I = P × 1000 / (√3 × V × PF) [IEC 60364-5-52 Section 523]
Calculation: I = {load['Power (kW)']} × 1000 / (1.732 × {load['Voltage (V)']} × {load['Power Factor']}) = **{current:.2f} A**

**Step 2: Derating Factors [IEC 60502-2 Tables B.10-B.22]**
k1 (Temperature): {factors['k1']:.3f} at {ambient_temp}°C
k2 (Grouping): {factors['k2']:.3f} for {num_cables} cables ({grouping})
Total Derating Factor K = k1 × k2 = {total_k:.3f}

**Step 3: Cable Selection**
Selected Cable: {size} mm² {material} ({cable_type})
Base Ampacity (Ic): {data['ampacity']} A
Derated Ampacity (Id): K × Ic = {total_k:.3f} × {data['ampacity']} = {derated:.1f} A
Check: {derated:.1f} A ≥ {current:.1f} A → **{'PASS' if derated >= current else 'FAIL'}**

**Step 4: Voltage Drop Calculation [IEC 60364-5-52 Section 525]**
R = {data['R']:.3f} ohm/km, X = {data['X']:.3f} ohm/km
Vd = {vd_v:.2f} V
Vd% = ({vd_v:.2f} / {load['Voltage (V)']}) × 100 = {vd_pct:.3f}%
Limit: {voltage_drop_limit}%
Check: {vd_pct:.3f}% ≤ {voltage_drop_limit}% → **{'PASS' if vd_pct <= voltage_drop_limit else 'FAIL'}**

**Final Status:** {'✅ PASS' if vd_pct <= voltage_drop_limit and derated >= current else '❌ FAIL'}
---
"""
                            detailed_calculations.append(detail)
                            break
                    
                    if selected_cable:
                        cable_results.append(selected_cable)
                
                st.session_state.cable_results_df = pd.DataFrame(cable_results)
                st.session_state.detailed_calculations = detailed_calculations
                st.success("✅ Calculations complete! Check other tabs for results.")
    
    # TAB 2: DETAILED CALCULATIONS - WITH FULL FORMULAS
    with cable_tabs[1]:
        st.markdown('<div class="report-header">', unsafe_allow_html=True)
        st.markdown("## DETAILED CALCULATIONS WITH REFERENCES")
        st.markdown('</div>', unsafe_allow_html=True)
        
        if 'detailed_calculations' in st.session_state:
            st.markdown("### 📋 Installation Parameters")
            params_html = f"""
            <table class="parameter-table">
                <tr>
                    <th colspan="2" style="background-color: #1E3A8A; color: white; padding: 10px; text-align: center; font-weight: bold;">
                        Installation Parameters (IEC 60502-2)
                    </th>
                </tr>
                <tr style="background-color: white;">
                    <td style="padding: 8px; border: 1px solid #ddd; font-weight: bold; width: 40%; color: black;">Ambient Temperature</td>
                    <td style="padding: 8px; border: 1px solid #ddd; color: black;">{st.session_state.ambient_temp} °C</td>
                </tr>
                <tr style="background-color: #f2f2f2;">
                    <td style="padding: 8px; border: 1px solid #ddd; font-weight: bold; color: black;">Number of Cables in Group</td>
                    <td style="padding: 8px; border: 1px solid #ddd; color: black;">{st.session_state.num_cables}</td>
                </tr>
                <tr style="background-color: white;">
                    <td style="padding: 8px; border: 1px solid #ddd; font-weight: bold; color: black;">Grouping Configuration</td>
                    <td style="padding: 8px; border: 1px solid #ddd; color: black;">{st.session_state.grouping}</td>
                </tr>
                <tr style="background-color: #f2f2f2;">
                    <td style="padding: 8px; border: 1px solid #ddd; font-weight: bold; color: black;">Voltage Drop Limit</td>
                    <td style="padding: 8px; border: 1px solid #ddd; color: black;">{st.session_state.voltage_drop_limit}% [IEC 60364-5-52 Sec 525]</td>
                </tr>
                <tr style="background-color: white;">
                    <td style="padding: 8px; border: 1px solid #ddd; font-weight: bold; color: black;">Conductor Material</td>
                    <td style="padding: 8px; border: 1px solid #ddd; color: black;">{st.session_state.material}</td>
                </tr>
            </table>
            """
            st.markdown(params_html, unsafe_allow_html=True)
            
            st.markdown("### 📊 Derating Factors [IEC 60502-2 Tables B.10-B.22]")
            factors = st.session_state.derating_factors
            factors_html = f"""
            <table class="parameter-table">
                <tr>
                    <th style="background-color: #1E3A8A; color: white; padding: 10px; text-align: center; font-weight: bold; width: 30%;">Factor</th>
                    <th style="background-color: #1E3A8A; color: white; padding: 10px; text-align: center; font-weight: bold; width: 30%;">Value</th>
                    <th style="background-color: #1E3A8A; color: white; padding: 10px; text-align: center; font-weight: bold; width: 40%;">Reference</th>
                </tr>
                <tr style="background-color: white;">
                    <td style="padding: 8px; border: 1px solid #ddd; font-weight: bold; color: black;">k1 - Temperature Correction</td>
                    <td style="padding: 8px; border: 1px solid #ddd; color: black;">{factors['k1']:.3f}</td>
                    <td style="padding: 8px; border: 1px solid #ddd; color: black;">Table B.10 at {st.session_state.ambient_temp}°C</td>
                </tr>
                <tr style="background-color: #f2f2f2;">
                    <td style="padding: 8px; border: 1px solid #ddd; font-weight: bold; color: black;">k2 - Grouping Factor</td>
                    <td style="padding: 8px; border: 1px solid #ddd; color: black;">{factors['k2']:.3f}</td>
                    <td style="padding: 8px; border: 1px solid #ddd; color: black;">Table 4C1 ({st.session_state.grouping})</td>
                </tr>
                <tr style="background-color: white;">
                    <td style="padding: 8px; border: 1px solid #ddd; font-weight: bold; color: black;"><strong>Total K = k1 × k2</strong></td>
                    <td style="padding: 8px; border: 1px solid #ddd; color: black;"><strong>{st.session_state.derating_factor:.3f}</strong></td>
                    <td style="padding: 8px; border: 1px solid #ddd; color: black;">IEC 60502-2</td>
                </tr>
            </table>
            """
            st.markdown(factors_html, unsafe_allow_html=True)
            
            st.markdown("### 🔍 Load-Wise Detailed Calculations")
            for detail in st.session_state.detailed_calculations:
                st.markdown(detail)
        else:
            st.info("👈 Enter loads in first tab and click CALCULATE")
    
    # TAB 3: CIRCUIT BREAKER RESULTS
    with cable_tabs[2]:
        st.markdown('<div class="report-header">', unsafe_allow_html=True)
        st.markdown("## CIRCUIT BREAKER SIZING")
        st.markdown('</div>', unsafe_allow_html=True)
        
        if 'cb_results' in st.session_state:
            st.markdown("### ⚡ Individual Circuit Breakers [IEC 60898]")
            cb_df = pd.DataFrame([{
                'Load': r['Load'],
                'Power (kW)': r['Power (kW)'],
                'Current (A)': f"{r['Current (A)']:.2f}",
                'Required (A)': f"{r['Required CB (A)']:.2f}",
                'Selected CB (A)': r['Selected CB (A)'],
                'Type': f"{r['Breaker Type']} {r['Poles']}",
                'Manufacturer': r['Manufacturer']
            } for r in st.session_state.cb_results])
            
            st.dataframe(cb_df, use_container_width=True, hide_index=True)
            
            # Main breaker
            st.markdown("### 🔋 Main Circuit Breaker")
            main = st.session_state.main_cb
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("Total Power", f"{main['total_power']:.1f} kW")
            with col2:
                st.metric("Current", f"{main['current']:.2f} A")
            with col3:
                st.metric("Required CB", f"{main['required_cb']:.2f} A")
            with col4:
                st.metric("Selected CB", f"{main['selected_cb']} A {main['breaker_type']} {main['poles']}")
            
            st.markdown('<div class="reference-box">', unsafe_allow_html=True)
            st.markdown("**Reference:** IEC 60898 - Circuit breakers for overcurrent protection")
            st.markdown('</div>', unsafe_allow_html=True)
        else:
            st.info("👈 Enter loads in first tab and click CALCULATE")
    
    # TAB 4: DOWNLOAD REPORT - PDF AND WORD ONLY
    with cable_tabs[3]:
        st.markdown('<div class="report-header">', unsafe_allow_html=True)
        st.markdown("## DOWNLOAD REPORT")
        st.markdown('</div>', unsafe_allow_html=True)
        
        if 'cable_results_df' in st.session_state and 'cb_results' in st.session_state:
            st.markdown("### 📥 Select Format")
            
            col1, col2 = st.columns(2)
            
            with col1:
                st.markdown("#### 📄 PDF Format")
                if st.button("📥 Generate PDF Report", key="cable_pdf_btn", use_container_width=True):
                    with st.spinner("Generating PDF report..."):
                        # Create PDF
                        pdf = FPDF()
                        pdf.add_page()
                        
                        # Title
                        pdf.set_font('Arial', 'B', 16)
                        pdf.cell(0, 10, 'CABLE SIZING & CIRCUIT BREAKER REPORT', 0, 1, 'C')
                        pdf.set_font('Arial', '', 10)
                        pdf.cell(0, 8, f'Date: {datetime.now().strftime("%Y-%m-%d %H:%M")}', 0, 1, 'R')
                        pdf.ln(10)
                        
                        # Installation Parameters
                        pdf.set_font('Arial', 'B', 12)
                        pdf.cell(0, 8, '1. INSTALLATION PARAMETERS', 0, 1)
                        pdf.ln(4)
                        pdf.set_font('Arial', '', 10)
                        pdf.cell(0, 6, f"Ambient Temperature: {st.session_state.ambient_temp} °C", 0, 1)
                        pdf.cell(0, 6, f"Number of Cables in Group: {st.session_state.num_cables}", 0, 1)
                        pdf.cell(0, 6, f"Grouping Configuration: {st.session_state.grouping}", 0, 1)
                        pdf.cell(0, 6, f"Voltage Drop Limit: {st.session_state.voltage_drop_limit}%", 0, 1)
                        pdf.ln(5)
                        
                        # Derating Factors
                        pdf.set_font('Arial', 'B', 12)
                        pdf.cell(0, 8, '2. DERATING FACTORS (IEC 60502-2)', 0, 1)
                        pdf.ln(4)
                        pdf.set_font('Arial', '', 10)
                        pdf.cell(0, 6, f"k1 (Temperature): {st.session_state.derating_factors['k1']:.3f}", 0, 1)
                        pdf.cell(0, 6, f"k2 (Grouping): {st.session_state.derating_factors['k2']:.3f}", 0, 1)
                        pdf.set_font('Arial', 'B', 10)
                        pdf.cell(0, 6, f"Total Derating Factor (K): {st.session_state.derating_factor:.3f}", 0, 1)
                        pdf.ln(5)
                        
                        # Loads Table
                        pdf.set_font('Arial', 'B', 12)
                        pdf.cell(0, 8, '3. LOAD DETAILS', 0, 1)
                        pdf.ln(4)
                        
                        # Add loads table to PDF
                        pdf.set_font('Arial', 'B', 9)
                        pdf.cell(30, 6, 'Load Name', 1, 0, 'C')
                        pdf.cell(20, 6, 'Power', 1, 0, 'C')
                        pdf.cell(20, 6, 'Voltage', 1, 0, 'C')
                        pdf.cell(20, 6, 'Phase', 1, 0, 'C')
                        pdf.cell(20, 6, 'PF', 1, 0, 'C')
                        pdf.cell(20, 6, 'Length', 1, 1, 'C')
                        
                        pdf.set_font('Arial', '', 8)
                        for idx, load in st.session_state.loads_df.iterrows():
                            pdf.cell(30, 5, load['Load Name'][:15], 1, 0, 'L')
                            pdf.cell(20, 5, f"{load['Power (kW)']:.1f}", 1, 0, 'R')
                            pdf.cell(20, 5, f"{load['Voltage (V)']:.0f}", 1, 0, 'R')
                            pdf.cell(20, 5, load['Phase'], 1, 0, 'C')
                            pdf.cell(20, 5, f"{load['Power Factor']:.2f}", 1, 0, 'R')
                            pdf.cell(20, 5, f"{load['Length (m)']:.0f}", 1, 1, 'R')
                        
                        pdf.ln(10)
                        
                        # Cable Results
                        pdf.set_font('Arial', 'B', 12)
                        pdf.cell(0, 8, '4. CABLE SIZING RESULTS', 0, 1)
                        pdf.ln(4)
                        
                        # Add cable results table
                        if 'cable_results_df' in st.session_state:
                            pdf.set_font('Arial', 'B', 8)
                            pdf.cell(25, 5, 'Load', 1, 0, 'C')
                            pdf.cell(15, 5, 'Cable', 1, 0, 'C')
                            pdf.cell(15, 5, 'Type', 1, 0, 'C')
                            pdf.cell(18, 5, 'Base A', 1, 0, 'C')
                            pdf.cell(18, 5, 'Derated A', 1, 0, 'C')
                            pdf.cell(15, 5, 'VD %', 1, 0, 'C')
                            pdf.cell(15, 5, 'Status', 1, 1, 'C')
                            
                            pdf.set_font('Arial', '', 7)
                            for idx, row in st.session_state.cable_results_df.iterrows():
                                pdf.cell(25, 4, row['Load Name'][:12], 1, 0, 'L')
                                pdf.cell(15, 4, row['Cable Size'], 1, 0, 'C')
                                pdf.cell(15, 4, row['Type'], 1, 0, 'C')
                                pdf.cell(18, 4, str(row['Base Ampacity']), 1, 0, 'R')
                                pdf.cell(18, 4, row['Derated Ampacity'], 1, 0, 'R')
                                pdf.cell(15, 4, row['Voltage Drop %'], 1, 0, 'R')
                                pdf.cell(15, 4, row['Status'], 1, 1, 'C')
                        
                        pdf.ln(10)
                        
                        # CB Results
                        pdf.set_font('Arial', 'B', 12)
                        pdf.cell(0, 8, '5. CIRCUIT BREAKER SIZING (IEC 60898)', 0, 1)
                        pdf.ln(4)
                        
                        # Add CB results table
                        pdf.set_font('Arial', 'B', 8)
                        pdf.cell(25, 5, 'Load', 1, 0, 'C')
                        pdf.cell(15, 5, 'Current', 1, 0, 'C')
                        pdf.cell(20, 5, 'Selected CB', 1, 0, 'C')
                        pdf.cell(20, 5, 'Type', 1, 0, 'C')
                        pdf.cell(30, 5, 'Manufacturer', 1, 1, 'C')
                        
                        pdf.set_font('Arial', '', 7)
                        for cb in st.session_state.cb_results:
                            pdf.cell(25, 4, cb['Load'][:12], 1, 0, 'L')
                            pdf.cell(15, 4, f"{cb['Current (A)']:.1f}", 1, 0, 'R')
                            pdf.cell(20, 4, f"{cb['Selected CB (A)']} A", 1, 0, 'C')
                            pdf.cell(20, 4, f"{cb['Breaker Type']} {cb['Poles']}", 1, 0, 'C')
                            pdf.cell(30, 4, cb['Manufacturer'], 1, 1, 'L')
                        
                        pdf.ln(5)
                        
                        # Main CB
                        main = st.session_state.main_cb
                        pdf.set_font('Arial', 'B', 10)
                        pdf.cell(0, 6, f"Main Circuit Breaker: {main['selected_cb']} A {main['breaker_type']} {main['poles']}", 0, 1)
                        
                        pdf_output = pdf.output(dest='S')
                        b64 = base64.b64encode(pdf_output).decode()
                        filename = f"Cable_CB_Report_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf"
                        st.markdown(f'<a href="data:application/pdf;base64,{b64}" download="{filename}" class="download-btn pdf-btn">📥 Download PDF Report</a>', unsafe_allow_html=True)
                        st.success("✅ PDF generated successfully!")
            
            with col2:
                st.markdown("#### 📝 Word Format")
                if st.button("📥 Generate Word Report", key="cable_word_btn", use_container_width=True):
                    with st.spinner("Generating Word report..."):
                        # Create Word document
                        from docx import Document
                        doc = Document()
                        
                        # Title
                        doc.add_heading('CABLE SIZING & CIRCUIT BREAKER REPORT', 0)
                        doc.add_paragraph(f'Date: {datetime.now().strftime("%Y-%m-%d %H:%M")}')
                        doc.add_paragraph()
                        
                        # Installation Parameters
                        doc.add_heading('1. INSTALLATION PARAMETERS', level=1)
                        doc.add_paragraph(f"Ambient Temperature: {st.session_state.ambient_temp} °C")
                        doc.add_paragraph(f"Number of Cables in Group: {st.session_state.num_cables}")
                        doc.add_paragraph(f"Grouping Configuration: {st.session_state.grouping}")
                        doc.add_paragraph(f"Voltage Drop Limit: {st.session_state.voltage_drop_limit}%")
                        doc.add_paragraph()
                        
                        # Derating Factors
                        doc.add_heading('2. DERATING FACTORS (IEC 60502-2)', level=1)
                        doc.add_paragraph(f"k1 (Temperature): {st.session_state.derating_factors['k1']:.3f}")
                        doc.add_paragraph(f"k2 (Grouping): {st.session_state.derating_factors['k2']:.3f}")
                        p = doc.add_paragraph()
                        p.add_run(f"Total Derating Factor (K): {st.session_state.derating_factor:.3f}").bold = True
                        doc.add_paragraph()
                        
                        # Loads Table
                        doc.add_heading('3. LOAD DETAILS', level=1)
                        table = doc.add_table(rows=1, cols=6)
                        table.style = 'Light Grid Accent 1'
                        hdr_cells = table.rows[0].cells
                        hdr_cells[0].text = 'Load Name'
                        hdr_cells[1].text = 'Power (kW)'
                        hdr_cells[2].text = 'Voltage (V)'
                        hdr_cells[3].text = 'Phase'
                        hdr_cells[4].text = 'PF'
                        hdr_cells[5].text = 'Length (m)'
                        
                        for idx, load in st.session_state.loads_df.iterrows():
                            row_cells = table.add_row().cells
                            row_cells[0].text = load['Load Name']
                            row_cells[1].text = f"{load['Power (kW)']:.1f}"
                            row_cells[2].text = f"{load['Voltage (V)']:.0f}"
                            row_cells[3].text = load['Phase']
                            row_cells[4].text = f"{load['Power Factor']:.2f}"
                            row_cells[5].text = f"{load['Length (m)']:.0f}"
                        
                        doc.add_paragraph()
                        
                        # Cable Results
                        doc.add_heading('4. CABLE SIZING RESULTS', level=1)
                        if 'cable_results_df' in st.session_state:
                            table = doc.add_table(rows=1, cols=7)
                            table.style = 'Light Grid Accent 1'
                            hdr_cells = table.rows[0].cells
                            hdr_cells[0].text = 'Load'
                            hdr_cells[1].text = 'Cable Size'
                            hdr_cells[2].text = 'Type'
                            hdr_cells[3].text = 'Base A'
                            hdr_cells[4].text = 'Derated A'
                            hdr_cells[5].text = 'VD %'
                            hdr_cells[6].text = 'Status'
                            
                            for idx, row in st.session_state.cable_results_df.iterrows():
                                row_cells = table.add_row().cells
                                row_cells[0].text = row['Load Name']
                                row_cells[1].text = row['Cable Size']
                                row_cells[2].text = row['Type']
                                row_cells[3].text = str(row['Base Ampacity'])
                                row_cells[4].text = row['Derated Ampacity']
                                row_cells[5].text = row['Voltage Drop %']
                                row_cells[6].text = row['Status']
                        
                        doc.add_paragraph()
                        
                        # CB Results
                        doc.add_heading('5. CIRCUIT BREAKER SIZING (IEC 60898)', level=1)
                        table = doc.add_table(rows=1, cols=5)
                        table.style = 'Light Grid Accent 1'
                        hdr_cells = table.rows[0].cells
                        hdr_cells[0].text = 'Load'
                        hdr_cells[1].text = 'Current (A)'
                        hdr_cells[2].text = 'Selected CB'
                        hdr_cells[3].text = 'Type'
                        hdr_cells[4].text = 'Manufacturer'
                        
                        for cb in st.session_state.cb_results:
                            row_cells = table.add_row().cells
                            row_cells[0].text = cb['Load']
                            row_cells[1].text = f"{cb['Current (A)']:.1f}"
                            row_cells[2].text = f"{cb['Selected CB (A)']} A"
                            row_cells[3].text = f"{cb['Breaker Type']} {cb['Poles']}"
                            row_cells[4].text = cb['Manufacturer']
                        
                        doc.add_paragraph()
                        
                        # Main CB
                        main = st.session_state.main_cb
                        doc.add_heading('6. MAIN CIRCUIT BREAKER', level=1)
                        doc.add_paragraph(f"Selected: {main['selected_cb']} A {main['breaker_type']} {main['poles']}")
                        
                        # Save Word document
                        word_path = "cable_report.docx"
                        doc.save(word_path)
                        
                        with open(word_path, "rb") as f:
                            word_bytes = f.read()
                        
                        b64 = base64.b64encode(word_bytes).decode()
                        filename = f"Cable_CB_Report_{datetime.now().strftime('%Y%m%d_%H%M')}.docx"
                        
                        if os.path.exists(word_path):
                            os.remove(word_path)
                        
                        st.markdown(f'<a href="data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,{b64}" download="{filename}" class="download-btn word-btn">📥 Download Word Report</a>', unsafe_allow_html=True)
                        st.success("✅ Word report generated successfully!")
        else:
            st.info("👈 Calculate first in Loads Table tab")

# ========== OTHER CALCULATORS (Placeholders) ==========
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
st.markdown(f"<div style='text-align: center; color: gray;'>⚡ CES-Electrical Design Calculators | Version 45.0 | {datetime.now().strftime('%Y-%m-%d %H:%M')}</div>", unsafe_allow_html=True)