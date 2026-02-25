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
    /* LARGER TAB FONT SIZE */
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
    .reasoning-box {
        background-color: #f0f7ff;
        padding: 20px;
        border-radius: 10px;
        border: 2px solid #1E3A8A;
        margin: 20px 0;
    }
    .step-highlight {
        background-color: #ffd966;
        padding: 2px 5px;
        border-radius: 3px;
        font-weight: bold;
    }
</style>
""", unsafe_allow_html=True)

# ========== LIGHTNING PROTECTION CALCULATOR - EXISTING CLASSES (UNCHANGED) ==========
class LightningWordReport:
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

# ========== CABLE SIZING CALCULATOR - WITH REFERENCES AND EXPLANATION ==========
# Cable Database (Based on Pakistan Cables Catalogue & IEC 60502)
CABLE_DATA = {
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
        400: {'R': 0.0470, 'X': 0.063, 'ampacity': 655, 'diameter': 33.2},
        500: {'R': 0.0366, 'X': 0.062, 'ampacity': 749, 'diameter': 37.1},
        630: {'R': 0.0283, 'X': 0.061, 'ampacity': 855, 'diameter': 41.4},
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
        400: {'R': 0.0778, 'X': 0.063, 'ampacity': 508, 'diameter': 33.2},
        500: {'R': 0.0605, 'X': 0.062, 'ampacity': 581, 'diameter': 37.1},
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

SOIL_RESISTIVITY_FACTORS = {0.7: 1.28, 0.8: 1.24, 0.9: 1.19, 1.0: 1.15, 1.5: 1.00, 2.0: 0.89, 2.5: 0.81, 3.0: 0.75}
DEPTH_FACTORS = {0.5: 1.04, 0.6: 1.02, 0.8: 1.00, 1.0: 0.98, 1.25: 0.96, 1.5: 0.95, 1.75: 0.94, 2.0: 0.93}

class CableSizingCalculator:
    def __init__(self):
        self.results = {}
    
    def calculate_load_current(self, power_kw, voltage_v, pf, efficiency=1.0, phase='3-phase'):
        """Calculate load current based on power and voltage
        Reference: IEC 60364-5-52, Section 523"""
        if phase == '3-phase':
            return (power_kw * 1000) / (1.732 * voltage_v * pf * efficiency)
        elif phase == '1-phase':
            return (power_kw * 1000) / (voltage_v * pf * efficiency)
        else:  # DC
            return (power_kw * 1000) / voltage_v
    
    def get_derating_factors(self, temp_c, insulation_temp=90, num_cables=1, 
                            grouping='touching', soil_resistivity=1.5, depth=0.8):
        """Calculate total derating factor based on installation conditions
        Reference: IEC 60502-2 Tables B.10-B.22"""
        k1 = TEMPERATURE_FACTORS[insulation_temp].get(temp_c, 1.0)
        k2 = GROUPING_FACTORS[grouping].get(min(num_cables, 16), 0.5)
        k3 = SOIL_RESISTIVITY_FACTORS.get(soil_resistivity, 1.0)
        k4 = DEPTH_FACTORS.get(depth, 1.0)
        
        total_k = k1 * k2 * k3 * k4
        return total_k, {'k1': k1, 'k2': k2, 'k3': k3, 'k4': k4}
    
    def select_cable_with_voltage_check(self, load_current, voltage_v, length_m, pf, 
                                        material='copper', derating_factor=1.0, 
                                        max_vd_percent=3.0, phase='3-phase'):
        """Select cable based on both ampacity and voltage drop
        Reference: IEC 60364-5-52 Section 525 (Voltage Drop)"""
        cable_data = CABLE_DATA[material]
        suitable_cables = []
        
        for size, data in cable_data.items():
            # Check ampacity
            derated_ampacity = data['ampacity'] * derating_factor
            if derated_ampacity < load_current:
                continue
            
            # Check voltage drop
            vd_volts, vd_percent = self.calculate_voltage_drop(
                load_current, length_m, data['R'], data['X'], pf, voltage_v, phase
            )
            
            if vd_percent <= max_vd_percent:
                suitable_cables.append({
                    'size': size,
                    'R': data['R'],
                    'X': data['X'],
                    'base_ampacity': data['ampacity'],
                    'derated_ampacity': derated_ampacity,
                    'diameter': data.get('diameter', 0),
                    'vd_percent': vd_percent,
                    'vd_volts': vd_volts
                })
        
        # Return smallest suitable cable (most economical)
        if suitable_cables:
            return min(suitable_cables, key=lambda x: x['size'])
        return None
    
    def calculate_voltage_drop(self, current, length_m, R, X, pf, voltage_v, phase='3-phase'):
        """Calculate voltage drop in volts and percentage
        Reference: IEC 60949, IEC 60364-5-52 Section 525"""
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
    
    def calculate_max_length(self, current, R, X, pf, voltage_v, max_vd_percent, phase='3-phase'):
        """Calculate maximum cable length for given voltage drop limit
        Reference: Derived from IEC 60364-5-52 voltage drop formula"""
        max_vd = max_vd_percent * voltage_v / 100
        
        if phase == '3-phase':
            denominator = 1.732 * current * (R/1000 * pf + X/1000 * math.sin(math.acos(pf)))
        elif phase == '1-phase':
            denominator = 2 * current * (R/1000 * pf + X/1000 * math.sin(math.acos(pf)))
        else:
            denominator = 2 * current * (R/1000)
        
        return (max_vd / denominator) / 1000 if denominator > 0 else 0

class CableWordReport:
    def __init__(self):
        self.doc = Document()
        self.doc.core_properties.title = "Cable Sizing Calculation"
        self.doc.core_properties.author = "CES-Electrical"
    
    def add_title(self, load_tag):
        title = self.doc.add_heading(f'CABLE SIZING CALCULATION - {load_tag}', 0)
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        self.doc.add_paragraph()
        self.doc.add_paragraph(f'Date: {datetime.now().strftime("%Y-%m-%d %H:%M")}')
        self.doc.add_paragraph('Reference: IEC 60502, IEC 60364-5-52')
        self.doc.add_paragraph('_' * 50)
        self.doc.add_paragraph()
    
    def add_input_parameters(self, params):
        self.doc.add_heading('INPUT PARAMETERS', level=1)
        for key, value in params.items():
            self.doc.add_paragraph(f'{key}: {value}')
        self.doc.add_paragraph()
    
    def add_derating_factors(self, factors):
        self.doc.add_heading('DERATING FACTORS (IEC 60502)', level=1)
        for key, value in factors.items():
            if key != 'total':
                self.doc.add_paragraph(f'{key}: {value:.3f}')
        p = self.doc.add_paragraph()
        p.add_run('Total Derating Factor: ').bold = True
        p.add_run(f'{factors["total"]:.3f}')
        self.doc.add_paragraph()
    
    def add_results(self, results):
        self.doc.add_heading('CABLE SELECTION RESULTS', level=1)
        self.doc.add_paragraph(f'Selected Cable Size: {results["size"]}')
        self.doc.add_paragraph(f'Base Ampacity: {results["base_ampacity"]:.1f} A')
        self.doc.add_paragraph(f'Derated Ampacity: {results["derated_ampacity"]:.1f} A')
        self.doc.add_paragraph(f'Resistance at 90°C: {results["R"]:.4f} ohm/km')
        self.doc.add_paragraph(f'Reactance: {results["X"]:.4f} ohm/km')
        self.doc.add_paragraph(f'Cable Diameter: {results["diameter"]:.1f} mm')
        self.doc.add_paragraph()
        
        if results.get('vd_volts') is not None:
            self.doc.add_heading('VOLTAGE DROP CALCULATION', level=1)
            self.doc.add_paragraph(f'Voltage Drop: {results["vd_volts"]:.2f} V')
            self.doc.add_paragraph(f'Voltage Drop: {results["vd_percent"]:.2f} %')
            if results.get('max_length') is not None:
                self.doc.add_paragraph(f'Maximum Length: {results["max_length"]:.1f} m')
            self.doc.add_paragraph()
        
        status = self.doc.add_paragraph()
        status.add_run('STATUS: ').bold = True
        status_text = '✓ PASSED' if results['status'] == 'PASS' else '✗ FAILED'
        status.add_run(status_text).bold = True
        
        ref = self.doc.add_paragraph()
        ref.add_run('References: IEC 60502, IEC 60364-5-52, IEC 60949').italic = True
    
    def save(self, filename):
        self.doc.save(filename)

class CablePDFReport(FPDF):
    def __init__(self):
        super().__init__()
        self.set_auto_page_break(auto=True, margin=15)
        
    def header(self):
        if self.page_no() > 1:
            self.set_font('Arial', 'I', 10)
            self.cell(0, 12, 'Cable Sizing Calculation', 0, 0, 'L')
            self.cell(0, 12, f'Page {self.page_no()}', 0, 0, 'R')
            self.ln(18)
    
    def footer(self):
        self.set_y(-15)
        self.set_font('Arial', 'I', 8)
        self.cell(0, 10, f'IEC 60502 Compliant', 0, 0, 'C')
    
    def add_title(self, load_tag):
        self.add_page()
        self.set_font('Arial', 'B', 24)
        self.set_text_color(0, 51, 102)
        self.cell(0, 40, f'CABLE SIZING CALCULATION', 0, 1, 'C')
        self.set_font('Arial', '', 14)
        self.cell(0, 20, f'{load_tag}', 0, 1, 'C')
        self.ln(10)
        self.set_font('Arial', '', 12)
        self.cell(0, 8, f'Date: {datetime.now().strftime("%Y-%m-%d %H:%M")}', 0, 1, 'R')
        self.cell(0, 8, 'Reference: IEC 60502, IEC 60364-5-52', 0, 1, 'R')
        self.ln(10)
    
    def add_input_parameters(self, params):
        self.set_font('Arial', 'B', 16)
        self.cell(0, 10, 'INPUT PARAMETERS', 0, 1)
        self.ln(5)
        self.set_font('Arial', '', 11)
        for key, value in params.items():
            self.cell(0, 7, f'{key}: {value}', 0, 1)
        self.ln(10)
    
    def add_derating_factors(self, factors):
        self.set_font('Arial', 'B', 16)
        self.cell(0, 10, 'DERATING FACTORS (IEC 60502)', 0, 1)
        self.ln(5)
        self.set_font('Arial', '', 11)
        for key, value in factors.items():
            if key != 'total':
                self.cell(0, 7, f'{key}: {value:.3f}', 0, 1)
        self.set_font('Arial', 'B', 12)
        self.cell(0, 8, f'Total Derating Factor: {factors["total"]:.3f}', 0, 1)
        self.ln(10)
    
    def add_results(self, results):
        self.set_font('Arial', 'B', 16)
        self.cell(0, 10, 'CABLE SELECTION RESULTS', 0, 1)
        self.ln(5)
        self.set_font('Arial', '', 11)
        self.cell(0, 7, f'Selected Cable Size: {results["size"]}', 0, 1)
        self.cell(0, 7, f'Base Ampacity: {results["base_ampacity"]:.1f} A', 0, 1)
        self.cell(0, 7, f'Derated Ampacity: {results["derated_ampacity"]:.1f} A', 0, 1)
        # FIXED: No special characters
        self.cell(0, 7, f'Resistance at 90°C: {results["R"]:.4f} ohm/km', 0, 1)
        self.cell(0, 7, f'Reactance: {results["X"]:.4f} ohm/km', 0, 1)
        self.cell(0, 7, f'Cable Diameter: {results["diameter"]:.1f} mm', 0, 1)
        self.ln(5)
        
        if results.get('vd_volts') is not None:
            self.set_font('Arial', 'B', 16)
            self.cell(0, 10, 'VOLTAGE DROP CALCULATION', 0, 1)
            self.ln(5)
            self.set_font('Arial', '', 11)
            self.cell(0, 7, f'Voltage Drop: {results["vd_volts"]:.2f} V', 0, 1)
            self.cell(0, 7, f'Voltage Drop: {results["vd_percent"]:.2f} %', 0, 1)
            if results.get('max_length') is not None:
                self.cell(0, 7, f'Maximum Length: {results["max_length"]:.1f} m', 0, 1)
            self.ln(5)
        
        # Status
        self.set_font('Arial', 'B', 14)
        self.set_text_color(0, 128, 0) if results['status'] == 'PASS' else self.set_text_color(255, 0, 0)
        status_text = "✓ CABLE SELECTION PASSED" if results['status'] == 'PASS' else "✗ CABLE SELECTION FAILED"
        self.cell(0, 10, status_text, 0, 1)
        
        # References
        self.ln(5)
        self.set_font('Arial', 'I', 9)
        self.set_text_color(100, 100, 100)
        self.cell(0, 5, 'References: IEC 60502 (Cable Construction & Ampacity)', 0, 1)
        self.cell(0, 5, 'IEC 60364-5-52 (Voltage Drop Requirements)', 0, 1)
        self.cell(0, 5, 'IEC 60949 (Short Circuit Calculation)', 0, 1)

# ========== SESSION STATE INITIALIZATION ==========
if 'calc_results' not in st.session_state:
    st.session_state.calc_results = {}
if 'calc_done' not in st.session_state:
    st.session_state.calc_done = False
if 'selected_calculator' not in st.session_state:
    st.session_state.selected_calculator = "⚡ Lightning Protection"
if 'cable_results' not in st.session_state:
    st.session_state.cable_results = {}
if 'cable_selection_reason' not in st.session_state:
    st.session_state.cable_selection_reason = ""
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

# ========== LIGHTNING PROTECTION CALCULATOR (UNCHANGED) ==========
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

# ========== CABLE SIZING CALCULATOR ==========
elif st.session_state.selected_calculator == "🔌 Cable Sizing":
    
    cable_tabs = st.tabs([
        "📥 Input Parameters", 
        "📊 Results", 
        "📋 Detailed Calculations",
        "📥 Download Report"
    ])
    
    # Initialize cable calculator
    cable_calc = CableSizingCalculator()
    
    # TAB 1: INPUT PARAMETERS
    with cable_tabs[0]:
        st.markdown('<div class="report-header">', unsafe_allow_html=True)
        st.markdown("## CABLE SIZING - INPUT PARAMETERS")
        st.markdown('</div>', unsafe_allow_html=True)
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown('<div class="info-box">', unsafe_allow_html=True)
            st.markdown("### ⚡ Load Parameters")
            
            load_type = st.selectbox("Load Type", ["3-Phase Motor", "3-Phase Other", "1-Phase", "DC"], key="cable_load_type")
            power_kw = st.number_input("Power (kW)", value=560.0, step=10.0, key="cable_power")
            voltage_v = st.number_input("Voltage (V)", value=3300.0, step=100.0, key="cable_voltage")
            pf = st.number_input("Power Factor", value=0.85, min_value=0.0, max_value=1.0, step=0.05, key="cable_pf")
            efficiency = st.number_input("Efficiency", value=0.95, min_value=0.0, max_value=1.0, step=0.05, key="cable_eff")
            length_m = st.number_input("Cable Length (m)", value=227.0, step=10.0, key="cable_length")
            
            st.markdown('</div>', unsafe_allow_html=True)
            
            st.markdown('<div class="info-box">', unsafe_allow_html=True)
            st.markdown("### 🔧 Cable Parameters")
            
            material = st.selectbox("Conductor Material", ["copper", "aluminium"], key="cable_material")
            insulation = st.selectbox("Insulation Type", ["XLPE (90°C)", "PVC (70°C)"], key="cable_insulation")
            installation = st.selectbox("Installation Method", 
                                       ["Aboveground - Tray", "Aboveground - Conduit", 
                                        "Underground - Direct Burial", "Underground - Duct Bank"], key="cable_install")
            
            st.markdown('</div>', unsafe_allow_html=True)
        
        with col2:
            st.markdown('<div class="info-box">', unsafe_allow_html=True)
            st.markdown("### 🌡️ Environmental Parameters")
            
            ambient_temp = st.number_input("Ambient Temperature (°C)", value=40, min_value=-10, max_value=80, key="cable_temp")
            num_cables = st.number_input("Number of Cables in Group", value=12, min_value=1, max_value=30, key="cable_num")
            grouping = st.selectbox("Grouping Configuration", ["touching", "spaced"], key="cable_grouping")
            
            if "Underground" in installation:
                soil_resistivity = st.selectbox("Soil Thermal Resistivity (K.m/W)", 
                                               [0.7, 0.8, 0.9, 1.0, 1.5, 2.0, 2.5, 3.0], index=4, key="cable_soil")
                depth = st.selectbox("Depth of Laying (m)", 
                                    [0.5, 0.6, 0.8, 1.0, 1.25, 1.5, 1.75, 2.0], index=2, key="cable_depth")
            else:
                soil_resistivity = 1.5
                depth = 0.8
            
            voltage_drop_limit = st.number_input("Max Voltage Drop (%)", value=3.0, min_value=0.1, max_value=20.0, step=0.5, key="cable_vd_limit")
            
            st.markdown('</div>', unsafe_allow_html=True)
            
            st.markdown('<div class="info-box">', unsafe_allow_html=True)
            st.markdown("### 📝 Load Identification")
            
            load_tag = st.text_input("Load Tag", value="4-P02 A", key="cable_tag")
            load_description = st.text_input("Load Description", value="Desalted Crude Oil Pump", key="cable_desc")
            
            st.markdown('</div>', unsafe_allow_html=True)
        
        if st.button("🔧 CALCULATE CABLE SIZE", type="primary", use_container_width=True):
            with st.spinner("Calculating..."):
                # Determine phase
                if load_type in ["3-Phase Motor", "3-Phase Other"]:
                    phase = '3-phase'
                elif load_type == "1-Phase":
                    phase = '1-phase'
                else:
                    phase = 'dc'
                
                # Calculate load current
                load_current = cable_calc.calculate_load_current(power_kw, voltage_v, pf, efficiency, phase)
                
                # Get insulation temperature
                insulation_temp = 90 if "XLPE" in insulation else 70
                
                # Calculate derating factors
                total_k, factors = cable_calc.get_derating_factors(
                    ambient_temp, insulation_temp, num_cables, 
                    grouping, soil_resistivity, depth
                )
                
                # Select cable with voltage drop check
                cable_data = cable_calc.select_cable_with_voltage_check(
                    load_current, voltage_v, length_m, pf, 
                    material, total_k, voltage_drop_limit, phase
                )
                
                if cable_data:
                    # Calculate max length
                    max_length = cable_calc.calculate_max_length(
                        load_current, cable_data['R'], cable_data['X'], 
                        pf, voltage_v, voltage_drop_limit, phase
                    )
                    
                    # Status already PASS if we got here
                    status = "PASS"
                    
                    # Generate selection reasoning
                    reasoning = f"""
### 🔍 Cable Selection Reasoning

**Step 1: Load Current Calculation**
- Power: {power_kw} kW
- Voltage: {voltage_v} V
- Power Factor: {pf}
- **Load Current (IL) = {load_current:.2f} A**
- Reference: IEC 60364-5-52 Section 523

**Step 2: Derating Factors (IEC 60502-2)**
- k1 (Temperature): {factors['k1']:.3f} at {ambient_temp}°C
- k2 (Grouping): {factors['k2']:.3f} for {num_cables} cables ({grouping})
{f"- k3 (Soil Resistivity): {factors['k3']:.3f}" if 'k3' in factors else ""}
{f"- k4 (Depth): {factors['k4']:.3f}" if 'k4' in factors else ""}
- **Total Derating Factor K = {total_k:.3f}**

**Step 3: Cable Selection Process**
The calculator checks each standard cable size from smallest to largest:

| Size (mm²) | Base Ampacity | Derated Ampacity | Check | Voltage Drop | Check |
|------------|---------------|------------------|-------|--------------|-------|
"""
                    
                    # Add comparison table for all cables
                    all_sizes = []
                    for size, data in CABLE_DATA[material].items():
                        derated = data['ampacity'] * total_k
                        vd_v, vd_pct = cable_calc.calculate_voltage_drop(
                            load_current, length_m, data['R'], data['X'], pf, voltage_v, phase
                        )
                        amp_ok = "✓" if derated >= load_current else "✗"
                        vd_ok = "✓" if vd_pct <= voltage_drop_limit else "✗"
                        
                        if derated >= load_current and vd_pct <= voltage_drop_limit:
                            is_selected = " ← **SELECTED**" if size == cable_data['size'] else ""
                            all_sizes.append(f"| {size} | {data['ampacity']} | {derated:.1f} | {amp_ok} | {vd_pct:.2f}% | {vd_ok} |{is_selected}")
                    
                    reasoning += "\n".join(all_sizes[-10:])  # Show last 10 sizes
                    
                    reasoning += f"""

**Step 4: Selected Cable**
- **Selected Size:** {cable_data['size']} mm² {material}
- **Base Ampacity (Ic):** {cable_data['base_ampacity']} A
- **Derated Ampacity (Id):** K × Ic = {total_k:.3f} × {cable_data['base_ampacity']} = **{cable_data['derated_ampacity']:.1f} A**
- **Ampacity Check:** {cable_data['derated_ampacity']:.1f} A ≥ {load_current:.1f} A → **✓ PASS**

**Step 5: Voltage Drop Check**
- **Voltage Drop:** {cable_data['vd_percent']:.2f}%
- **Maximum Allowed:** {voltage_drop_limit}%
- **Check:** {cable_data['vd_percent']:.2f}% ≤ {voltage_drop_limit}% → **✓ PASS**

**Step 6: Final Selection**
This is the **smallest standard cable size** that satisfies both:
1. Ampacity requirement (Id ≥ IL)
2. Voltage drop limit (Vd ≤ {voltage_drop_limit}%)

**References:**
- IEC 60502-2: Cable construction and ampacity tables
- IEC 60364-5-52: Selection of cables - current-carrying capacity
- IEC 60364-5-52 Section 525: Voltage drop requirements
"""
                    
                    st.session_state.cable_selection_reason = reasoning
                    
                    # Store results
                    st.session_state.cable_results = {
                        'load_tag': load_tag,
                        'load_description': load_description,
                        'load_current': load_current,
                        'selected_size': cable_data['size'],
                        'cable_data': cable_data,
                        'derating_factors': {**factors, 'total': total_k},
                        'vd_volts': cable_data['vd_volts'],
                        'vd_percent': cable_data['vd_percent'],
                        'max_length': max_length,
                        'voltage_drop_limit': voltage_drop_limit,
                        'status': status,
                        'phase': phase,
                        'material': material,
                        'insulation': insulation,
                        'installation': installation,
                        'ambient_temp': ambient_temp,
                        'num_cables': num_cables,
                        'grouping': grouping,
                        'power_kw': power_kw,
                        'voltage_v': voltage_v,
                        'pf': pf,
                        'length_m': length_m
                    }
                    
                    st.success("✅ Calculation complete! Go to Results tab to view detailed reasoning.")
                else:
                    st.error("❌ No suitable cable found! Try larger sizes or better conditions.")
    
    # TAB 2: RESULTS
    with cable_tabs[1]:
        st.markdown('<div class="report-header">', unsafe_allow_html=True)
        st.markdown("## CABLE SIZING RESULTS")
        st.markdown('</div>', unsafe_allow_html=True)
        
        if st.session_state.cable_results:
            r = st.session_state.cable_results
            
            # Show selection reasoning
            if st.session_state.cable_selection_reason:
                st.markdown('<div class="reasoning-box">', unsafe_allow_html=True)
                st.markdown(st.session_state.cable_selection_reason)
                st.markdown('</div>', unsafe_allow_html=True)
            
            col1, col2 = st.columns(2)
            
            with col1:
                st.markdown('<div class="success-box">', unsafe_allow_html=True)
                st.markdown("### 📊 Summary")
                st.markdown(f"**Load Tag:** {r['load_tag']}")
                st.markdown(f"**Description:** {r['load_description']}")
                st.markdown(f"**Load Current:** {r['load_current']:.2f} A")
                st.markdown(f"**Selected Cable:** {r['selected_size']} mm² {r['material']}")
                st.markdown('</div>', unsafe_allow_html=True)
                
                st.markdown('<div class="info-box">', unsafe_allow_html=True)
                st.markdown("### ⚡ Ampacity Check")
                st.markdown(f"**Base Ampacity:** {r['cable_data']['base_ampacity']:.1f} A")
                st.markdown(f"**Derating Factor:** {r['derating_factors']['total']:.3f}")
                st.markdown(f"**Derated Ampacity:** {r['cable_data']['derated_ampacity']:.1f} A")
                
                if r['cable_data']['derated_ampacity'] >= r['load_current']:
                    st.success(f"✓ Derated ampacity > Load current ({r['cable_data']['derated_ampacity']:.1f} A > {r['load_current']:.1f} A)")
                else:
                    st.error(f"✗ Derated ampacity < Load current ({r['cable_data']['derated_ampacity']:.1f} A < {r['load_current']:.1f} A)")
                st.markdown('</div>', unsafe_allow_html=True)
            
            with col2:
                st.markdown('<div class="info-box">', unsafe_allow_html=True)
                st.markdown("### 📉 Voltage Drop")
                st.markdown(f"**Voltage Drop:** {r['vd_volts']:.2f} V")
                st.markdown(f"**Voltage Drop:** {r['vd_percent']:.2f} %")
                st.markdown(f"**Limit:** {r['voltage_drop_limit']} %")
                
                if r['vd_percent'] <= r['voltage_drop_limit']:
                    st.success(f"✓ Voltage drop within limit ({r['vd_percent']:.2f}% ≤ {r['voltage_drop_limit']}%)")
                else:
                    st.error(f"✗ Voltage drop exceeds limit ({r['vd_percent']:.2f}% > {r['voltage_drop_limit']}%)")
                
                st.markdown(f"**Maximum Length:** {r['max_length']:.1f} m")
                st.markdown(f"**Actual Length:** {r['length_m']:.1f} m")
                if r['length_m'] <= r['max_length']:
                    st.success(f"✓ Length within limit")
                else:
                    st.error(f"✗ Length exceeds maximum")
                st.markdown('</div>', unsafe_allow_html=True)
            
            st.markdown("---")
            if r['status'] == 'PASS' and r['cable_data']['derated_ampacity'] >= r['load_current']:
                st.markdown('<div class="success-box">', unsafe_allow_html=True)
                st.markdown(f"## ✅ CABLE SELECTION PASSED")
                st.markdown(f"**Cable Size:** {r['selected_size']} mm² {r['material']} is suitable for {r['load_tag']}")
                st.markdown('</div>', unsafe_allow_html=True)
            else:
                st.markdown('<div class="warning-box">', unsafe_allow_html=True)
                st.markdown(f"## ❌ CABLE SELECTION FAILED")
                st.markdown(f"**Cable Size:** {r['selected_size']} mm² {r['material']} is NOT suitable")
                st.markdown('</div>', unsafe_allow_html=True)
        else:
            st.info("👈 Enter parameters in Input tab and click CALCULATE")
    
    # TAB 3: DETAILED CALCULATIONS
    with cable_tabs[2]:
        st.markdown('<div class="report-header">', unsafe_allow_html=True)
        st.markdown("## DETAILED CALCULATIONS")
        st.markdown('</div>', unsafe_allow_html=True)
        
        if st.session_state.cable_results:
            r = st.session_state.cable_results
            
            with st.expander("1. Load Current Calculation", expanded=True):
                st.markdown('<div class="formula-box">', unsafe_allow_html=True)
                st.markdown("**Reference:** IEC 60364-5-52 Section 523")
                if r['phase'] == '3-phase':
                    st.markdown("**Formula:** I = P × 1000 / (√3 × V × PF × η)")
                    st.markdown(f"**Calculation:** I = {r['power_kw']} × 1000 / (1.732 × {r['voltage_v']} × {r['pf']} × 1.0)")
                elif r['phase'] == '1-phase':
                    st.markdown("**Formula:** I = P × 1000 / (V × PF × η)")
                    st.markdown(f"**Calculation:** I = {r['power_kw']} × 1000 / ({r['voltage_v']} × {r['pf']} × 1.0)")
                else:
                    st.markdown("**Formula:** I = P × 1000 / V")
                    st.markdown(f"**Calculation:** I = {r['power_kw']} × 1000 / {r['voltage_v']}")
                st.markdown(f"**Result:** I = **{r['load_current']:.2f} A**")
                st.markdown('</div>', unsafe_allow_html=True)
            
            with st.expander("2. Derating Factors (IEC 60502)"):
                st.markdown('<div class="formula-box">', unsafe_allow_html=True)
                st.markdown("**Reference:** IEC 60502-2 Tables B.10-B.22")
                st.markdown(f"**k1 - Temperature Correction:** {r['derating_factors']['k1']:.3f} (at {r['ambient_temp']}°C)")
                st.markdown(f"**k2 - Grouping Factor:** {r['derating_factors']['k2']:.3f} ({r['num_cables']} cables, {r['grouping']})")
                if 'k3' in r['derating_factors']:
                    st.markdown(f"**k3 - Soil Resistivity:** {r['derating_factors']['k3']:.3f}")
                if 'k4' in r['derating_factors']:
                    st.markdown(f"**k4 - Depth Factor:** {r['derating_factors']['k4']:.3f}")
                st.markdown(f"**Total K =** {r['derating_factors']['k1']:.3f} × {r['derating_factors']['k2']:.3f} = **{r['derating_factors']['total']:.3f}**")
                st.markdown('</div>', unsafe_allow_html=True)
            
            with st.expander("3. Ampacity Check"):
                st.markdown('<div class="formula-box">', unsafe_allow_html=True)
                st.markdown("**Reference:** IEC 60502-2")
                st.markdown(f"**Base Ampacity (Ic):** {r['cable_data']['base_ampacity']} A")
                st.markdown(f"**Derated Ampacity (Id):** Ic × K = {r['cable_data']['base_ampacity']} × {r['derating_factors']['total']:.3f} = **{r['cable_data']['derated_ampacity']:.1f} A**")
                st.markdown(f"**Load Current (IL):** {r['load_current']:.1f} A")
                if r['cable_data']['derated_ampacity'] >= r['load_current']:
                    st.success(f"✓ {r['cable_data']['derated_ampacity']:.1f} A ≥ {r['load_current']:.1f} A - OK")
                else:
                    st.error(f"✗ {r['cable_data']['derated_ampacity']:.1f} A < {r['load_current']:.1f} A - NOT OK")
                st.markdown('</div>', unsafe_allow_html=True)
            
            with st.expander("4. Voltage Drop Calculation"):
                st.markdown('<div class="formula-box">', unsafe_allow_html=True)
                st.markdown("**Reference:** IEC 60364-5-52 Section 525, IEC 60949")
                if r['phase'] == '3-phase':
                    st.markdown("**Formula:** Vd = √3 × I × L × (R cosφ + X sinφ) / 1000")
                elif r['phase'] == '1-phase':
                    st.markdown("**Formula:** Vd = 2 × I × L × (R cosφ + X sinφ) / 1000")
                else:
                    st.markdown("**Formula:** Vd = 2 × I × L × R / 1000")
                
                st.markdown(f"**R:** {r['cable_data']['R']:.4f} ohm/km")
                st.markdown(f"**X:** {r['cable_data']['X']:.4f} ohm/km")
                st.markdown(f"**Z = R cosφ + X sinφ =** {r['cable_data']['R']:.4f}×{r['pf']:.2f} + {r['cable_data']['X']:.4f}×{math.sin(math.acos(r['pf'])):.3f}")
                st.markdown(f"**Voltage Drop:** {r['vd_volts']:.2f} V")
                st.markdown(f"**Percentage:** ({r['vd_volts']:.2f} / {r['voltage_v']}) × 100 = **{r['vd_percent']:.2f}%**")
                
                if r['vd_percent'] <= r['voltage_drop_limit']:
                    st.success(f"✓ {r['vd_percent']:.2f}% ≤ {r['voltage_drop_limit']}% - OK")
                else:
                    st.error(f"✗ {r['vd_percent']:.2f}% > {r['voltage_drop_limit']}% - NOT OK")
                st.markdown('</div>', unsafe_allow_html=True)
            
            with st.expander("5. Maximum Length Calculation"):
                st.markdown('<div class="formula-box">', unsafe_allow_html=True)
                st.markdown("**Derived from IEC 60364-5-52 voltage drop formula**")
                st.markdown(f"**Maximum Length:** {r['max_length']:.1f} m")
                st.markdown(f"**Actual Length:** {r['length_m']:.1f} m")
                if r['length_m'] <= r['max_length']:
                    st.success(f"✓ Length within limit")
                else:
                    st.error(f"✗ Length exceeds maximum")
                st.markdown('</div>', unsafe_allow_html=True)
        else:
            st.info("👈 Calculate cable size first in Input tab")
    
    # TAB 4: DOWNLOAD REPORT - WITH PDF AND WORD OPTIONS
    with cable_tabs[3]:
        st.markdown('<div class="report-header">', unsafe_allow_html=True)
        st.markdown("## DOWNLOAD REPORT")
        st.markdown('</div>', unsafe_allow_html=True)
        
        if st.session_state.cable_results:
            r = st.session_state.cable_results
            
            st.markdown("### 📥 Select Format")
            
            col1, col2 = st.columns(2)
            
            with col1:
                st.markdown("#### 📄 PDF Format")
                if st.button("📥 Generate PDF Report", key="cable_pdf_btn", use_container_width=True):
                    with st.spinner("Generating PDF report..."):
                        pdf = CablePDFReport()
                        pdf.add_title(r['load_tag'])
                        
                        params = {
                            'Load Tag': r['load_tag'],
                            'Description': r['load_description'],
                            'Power': f"{r['power_kw']} kW",
                            'Voltage': f"{r['voltage_v']} V",
                            'Power Factor': f"{r['pf']}",
                            'Length': f"{r['length_m']} m",
                            'Installation': r['installation'],
                            'Material': r['material'],
                            'Ambient Temp': f"{r['ambient_temp']}°C"
                        }
                        pdf.add_input_parameters(params)
                        pdf.add_derating_factors(r['derating_factors'])
                        
                        results_dict = {
                            'size': f"{r['selected_size']} mm²",
                            'base_ampacity': r['cable_data']['base_ampacity'],
                            'derated_ampacity': r['cable_data']['derated_ampacity'],
                            'R': r['cable_data']['R'],
                            'X': r['cable_data']['X'],
                            'diameter': r['cable_data'].get('diameter', 0),
                            'vd_volts': r['vd_volts'],
                            'vd_percent': r['vd_percent'],
                            'max_length': r['max_length'],
                            'status': r['status']
                        }
                        pdf.add_results(results_dict)
                        
                        pdf_output = pdf.output(dest='S')
                        b64 = base64.b64encode(pdf_output).decode()
                        filename = f"Cable_Sizing_{r['load_tag']}_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf"
                        st.markdown(f'<a href="data:application/pdf;base64,{b64}" download="{filename}" class="download-btn pdf-btn">📥 Click to Download PDF Report</a>', unsafe_allow_html=True)
                        st.success("✅ PDF generated successfully!")
            
            with col2:
                st.markdown("#### 📝 Word Format")
                if st.button("📥 Generate Word Report", key="cable_word_btn", use_container_width=True):
                    with st.spinner("Generating Word report..."):
                        word = CableWordReport()
                        word.add_title(r['load_tag'])
                        
                        params = {
                            'Load Tag': r['load_tag'],
                            'Description': r['load_description'],
                            'Power': f"{r['power_kw']} kW",
                            'Voltage': f"{r['voltage_v']} V",
                            'Power Factor': f"{r['pf']}",
                            'Length': f"{r['length_m']} m",
                            'Installation': r['installation'],
                            'Material': r['material'],
                            'Ambient Temp': f"{r['ambient_temp']}°C"
                        }
                        word.add_input_parameters(params)
                        word.add_derating_factors(r['derating_factors'])
                        
                        results_dict = {
                            'size': f"{r['selected_size']} mm²",
                            'base_ampacity': r['cable_data']['base_ampacity'],
                            'derated_ampacity': r['cable_data']['derated_ampacity'],
                            'R': r['cable_data']['R'],
                            'X': r['cable_data']['X'],
                            'diameter': r['cable_data'].get('diameter', 0),
                            'vd_volts': r['vd_volts'],
                            'vd_percent': r['vd_percent'],
                            'max_length': r['max_length'],
                            'status': r['status']
                        }
                        word.add_results(results_dict)
                        
                        word_path = "temp_cable_report.docx"
                        word.save(word_path)
                        
                        with open(word_path, "rb") as f:
                            word_bytes = f.read()
                        
                        b64 = base64.b64encode(word_bytes).decode()
                        filename = f"Cable_Sizing_{r['load_tag']}_{datetime.now().strftime('%Y%m%d_%H%M')}.docx"
                        
                        if os.path.exists(word_path):
                            os.remove(word_path)
                        
                        st.markdown(f'<a href="data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,{b64}" download="{filename}" class="download-btn word-btn">📥 Click to Download Word Report</a>', unsafe_allow_html=True)
                        st.success("✅ Word document generated successfully!")
        else:
            st.info("👈 Calculate cable size first in Input tab")

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
st.markdown(f"<div style='text-align: center; color: gray;'>⚡ CES-Electrical Design Calculations | Version 35.0 | {datetime.now().strftime('%Y-%m-%d %H:%M')}</div>", unsafe_allow_html=True)