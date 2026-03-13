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
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

st.set_page_config(page_title="Professional Engineering Tools", page_icon="🔌", layout="wide")

# ========== CUSTOM CSS ==========
st.markdown("""
<style>
    /* Main Theme Colors */
    :root {
        --primary: #1E3A8A;
        --primary-light: #3B5BA6;
        --primary-dark: #0D1B4A;
        --secondary: #00A86B;
        --success: #28A745;
        --danger: #DC3545;
        --warning: #FFC107;
        --info: #17A2B8;
        --light: #F8F9FA;
        --dark: #343A40;
        --white: #FFFFFF;
        --gray-100: #F8F9FA;
        --gray-200: #E9ECEF;
        --gray-300: #DEE2E6;
        --gray-400: #CED4DA;
        --gray-500: #ADB5BD;
        --gray-600: #6C757D;
    }
    
    /* Report Header */
    .report-header {
        background: linear-gradient(135deg, var(--primary) 0%, var(--primary-light) 100%);
        color: white;
        padding: 25px;
        border-radius: 12px;
        text-align: center;
        margin-bottom: 25px;
        font-size: 28px;
        font-weight: bold;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        border-bottom: 4px solid var(--secondary);
    }
    
    /* Section Headers */
    .section-header {
        color: var(--primary);
        font-size: 22px;
        font-weight: 600;
        margin: 20px 0 15px 0;
        padding-bottom: 10px;
        border-bottom: 3px solid var(--secondary);
    }
    
    /* Card Style */
    .card {
        background-color: var(--white);
        padding: 20px;
        border-radius: 10px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        border: 1px solid var(--gray-300);
        margin-bottom: 20px;
    }
    
    /* Formula Box */
    .formula-box {
        background: linear-gradient(135deg, var(--gray-100) 0%, var(--white) 100%);
        padding: 20px;
        border-radius: 10px;
        border-left: 6px solid var(--secondary);
        margin: 15px 0;
        font-family: 'Courier New', monospace;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
        color: #000000 !important;
    }
    .formula-box * {
        color: #000000 !important;
    }
    
    /* Info Box */
    .info-box {
        background: linear-gradient(135deg, #E7F3FF 0%, #D4E6FF 100%);
        color: #004085 !important;
        padding: 15px;
        border-radius: 8px;
        border-left: 5px solid var(--primary);
        margin: 10px 0;
    }
    .info-box * {
        color: #004085 !important;
    }
    
    /* Upload Section Styling */
    .upload-section {
        background-color: #F0F7FF;
        padding: 30px;
        border-radius: 10px;
        border: 3px dashed #1E3A8A;
        margin: 20px 0;
        text-align: center;
    }
    .upload-section h3 {
        color: #1E3A8A;
        margin-top: 0;
        font-size: 24px;
    }
    .upload-section p {
        color: #2C3E50;
        font-size: 16px;
    }
    
    /* DataFrame Styling */
    .stDataFrame {
        color: #000000 !important;
    }
    .stDataFrame table {
        color: #000000 !important;
        border: 2px solid #1E3A8A !important;
    }
    .stDataFrame th {
        background: linear-gradient(135deg, #1E3A8A 0%, #3B5BA6 100%) !important;
        color: white !important;
        font-weight: 600 !important;
        padding: 12px !important;
    }
    .stDataFrame td {
        color: #000000 !important;
        padding: 10px !important;
        background-color: white !important;
    }
    .stDataFrame tr:nth-child(even) td {
        background-color: #F8F9FA !important;
        color: #000000 !important;
    }
    .stDataFrame tr:nth-child(odd) td {
        background-color: white !important;
        color: #000000 !important;
    }
    
    /* Sidebar Navigation */
    .sidebar-nav {
        background: linear-gradient(135deg, #1E3A8A 0%, #3B5BA6 100%);
        color: white;
        padding: 15px;
        border-radius: 10px;
        margin-bottom: 20px;
        text-align: center;
    }
    .sidebar-nav h2 {
        color: white !important;
        margin: 0;
    }
    
    /* Download Buttons */
    .download-btn {
        display: inline-block;
        padding: 14px 28px;
        margin: 10px;
        color: white !important;
        text-decoration: none;
        border-radius: 8px;
        font-size: 16px;
        font-weight: bold;
        transition: all 0.3s;
        text-align: center;
        border: none;
        cursor: pointer;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
    }
    .download-btn:hover {
        transform: translateY(-2px);
        box-shadow: 0 4px 8px rgba(0,0,0,0.2);
    }
    .pdf-btn {
        background: linear-gradient(135deg, #DC3545 0%, #C82333 100%);
    }
    .word-btn {
        background: linear-gradient(135deg, #1E3A8A 0%, #3B5BA6 100%);
    }
</style>
""", unsafe_allow_html=True)

# ========== CIRCUIT BREAKER DATA ==========
CB_RATINGS = [6, 10, 16, 20, 25, 32, 40, 50, 63, 80, 100, 125, 160, 200, 250, 315, 400, 500, 630, 800, 1000, 1250, 1600]

BREAKER_TYPES = {
    'MCB': {'min': 6, 'max': 125, 'standard': 'IEC 60898', 'application': 'Miniature Circuit Breaker - For final circuits'},
    'MCCB': {'min': 125, 'max': 1600, 'standard': 'IEC 60947-2', 'application': 'Moulded Case Circuit Breaker - For distribution'},
    'ACB': {'min': 1600, 'max': 6300, 'standard': 'IEC 60947-2', 'application': 'Air Circuit Breaker - For main incomers'}
}

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
    }
}

# ========== CABLE DATABASE ==========
LV_CABLE_DATA = {
    'unarmoured': {
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
    'armoured': {
        1.5: {'R': 12.1, 'X': 0.098, 'ampacity': 25, 'diameter': 3.2},
        2.5: {'R': 7.41, 'X': 0.093, 'ampacity': 34, 'diameter': 4.0},
        4: {'R': 4.61, 'X': 0.092, 'ampacity': 45, 'diameter': 4.7},
        6: {'R': 3.08, 'X': 0.091, 'ampacity': 57, 'diameter': 5.5},
        10: {'R': 1.83, 'X': 0.088, 'ampacity': 78, 'diameter': 7.0},
        16: {'R': 1.15, 'X': 0.082, 'ampacity': 105, 'diameter': 8.5},
        25: {'R': 0.727, 'X': 0.079, 'ampacity': 138, 'diameter': 10.5},
        35: {'R': 0.524, 'X': 0.078, 'ampacity': 168, 'diameter': 12.0},
        50: {'R': 0.387, 'X': 0.075, 'ampacity': 203, 'diameter': 14.0},
        70: {'R': 0.268, 'X': 0.073, 'ampacity': 255, 'diameter': 16.5},
        95: {'R': 0.193, 'X': 0.072, 'ampacity': 310, 'diameter': 19.0},
        120: {'R': 0.153, 'X': 0.071, 'ampacity': 357, 'diameter': 21.0},
        150: {'R': 0.124, 'X': 0.070, 'ampacity': 408, 'diameter': 23.5},
        185: {'R': 0.0991, 'X': 0.070, 'ampacity': 466, 'diameter': 26.0},
        240: {'R': 0.0754, 'X': 0.069, 'ampacity': 553, 'diameter': 29.0},
        300: {'R': 0.0601, 'X': 0.068, 'ampacity': 637, 'diameter': 32.5},
    }
}

MV_CABLE_DATA = {
    'unarmoured': {
        25: {'R': 0.727, 'X': 0.120, 'ampacity': 145, 'diameter': 18.5},
        35: {'R': 0.524, 'X': 0.115, 'ampacity': 175, 'diameter': 20.0},
        50: {'R': 0.387, 'X': 0.110, 'ampacity': 210, 'diameter': 22.0},
        70: {'R': 0.268, 'X': 0.105, 'ampacity': 265, 'diameter': 24.5},
        95: {'R': 0.193, 'X': 0.100, 'ampacity': 320, 'diameter': 27.0},
        120: {'R': 0.153, 'X': 0.095, 'ampacity': 370, 'diameter': 29.5},
        150: {'R': 0.124, 'X': 0.092, 'ampacity': 420, 'diameter': 32.0},
        185: {'R': 0.0991, 'X': 0.090, 'ampacity': 475, 'diameter': 34.5},
        240: {'R': 0.0754, 'X': 0.088, 'ampacity': 560, 'diameter': 38.0},
        300: {'R': 0.0601, 'X': 0.086, 'ampacity': 645, 'diameter': 42.0},
    },
    'armoured': {
        25: {'R': 0.727, 'X': 0.135, 'ampacity': 155, 'diameter': 21.0},
        35: {'R': 0.524, 'X': 0.130, 'ampacity': 188, 'diameter': 23.0},
        50: {'R': 0.387, 'X': 0.125, 'ampacity': 225, 'diameter': 25.0},
        70: {'R': 0.268, 'X': 0.120, 'ampacity': 285, 'diameter': 27.5},
        95: {'R': 0.193, 'X': 0.115, 'ampacity': 345, 'diameter': 30.0},
        120: {'R': 0.153, 'X': 0.110, 'ampacity': 398, 'diameter': 33.0},
        150: {'R': 0.124, 'X': 0.107, 'ampacity': 452, 'diameter': 36.0},
        185: {'R': 0.0991, 'X': 0.105, 'ampacity': 512, 'diameter': 39.0},
        240: {'R': 0.0754, 'X': 0.102, 'ampacity': 605, 'diameter': 43.0},
        300: {'R': 0.0601, 'X': 0.100, 'ampacity': 695, 'diameter': 47.0},
    }
}

# ========== DERATING FACTORS ==========
TEMPERATURE_FACTORS = {
    90: {20: 1.07, 25: 1.04, 30: 1.00, 35: 0.96, 40: 0.91, 
         45: 0.87, 50: 0.82, 55: 0.76, 60: 0.71, 65: 0.65, 70: 0.58}
}

GROUPING_FACTORS = {
    'touching': {1: 1.00, 2: 0.80, 3: 0.70, 4: 0.65, 5: 0.60, 6: 0.57,
                 7: 0.54, 8: 0.52, 9: 0.50, 10: 0.48, 11: 0.46, 12: 0.45,
                 13: 0.44, 14: 0.43, 15: 0.42, 16: 0.41, 17: 0.40, 18: 0.39},
    'spaced_1d': {1: 1.00, 2: 0.90, 3: 0.85, 4: 0.82, 5: 0.80, 6: 0.78,
                  7: 0.76, 8: 0.74, 9: 0.72, 10: 0.70, 11: 0.68, 12: 0.66},
    'spaced_2d': {1: 1.00, 2: 0.95, 3: 0.92, 4: 0.90, 5: 0.88, 6: 0.86,
                  7: 0.84, 8: 0.82, 9: 0.80, 10: 0.78, 11: 0.76, 12: 0.74},
    'spaced_3d': {1: 1.00, 2: 0.98, 3: 0.96, 4: 0.94, 5: 0.92, 6: 0.90,
                  7: 0.88, 8: 0.86, 9: 0.84, 10: 0.82, 11: 0.80, 12: 0.78},
    'cleated': {1: 1.00, 2: 0.95, 3: 0.90, 4: 0.85, 5: 0.82, 6: 0.80,
                7: 0.78, 8: 0.76, 9: 0.74, 10: 0.72, 11: 0.70, 12: 0.68}
}

FORMATION_FACTORS = {'flat': 1.00, 'trefoil': 0.95, 'single': 1.00}
INSTALLATION_FACTORS = {'air': 1.00, 'surface': 0.98, 'tray': 0.95, 'ladder': 0.96, 
                        'trench': 0.90, 'buried': 0.85, 'duct': 0.82, 'conduit': 0.80}
SOIL_RESISTIVITY_FACTORS = {0.7: 1.28, 0.8: 1.24, 0.9: 1.19, 1.0: 1.15,
                            1.5: 1.00, 2.0: 0.89, 2.5: 0.81, 3.0: 0.75}
DEPTH_FACTORS = {0.5: 1.04, 0.6: 1.02, 0.7: 1.01, 0.8: 1.00,
                 0.9: 0.99, 1.0: 0.98, 1.25: 0.96, 1.5: 0.95,
                 1.75: 0.94, 2.0: 0.93}

def get_grouping_factor(num_cables, spacing_mm, cable_diameter, arrangement='touching'):
    if arrangement == 'touching':
        return GROUPING_FACTORS['touching'].get(min(num_cables, 18), 0.39)
    elif arrangement == 'cleated':
        return GROUPING_FACTORS['cleated'].get(min(num_cables, 12), 0.68)
    else:
        if cable_diameter > 0:
            spacing_ratio = spacing_mm / cable_diameter
            if spacing_ratio < 0.5:
                return GROUPING_FACTORS['touching'].get(min(num_cables, 18), 0.39)
            elif spacing_ratio < 1.5:
                return GROUPING_FACTORS['spaced_1d'].get(min(num_cables, 12), 0.66)
            elif spacing_ratio < 2.5:
                return GROUPING_FACTORS['spaced_2d'].get(min(num_cables, 12), 0.74)
            else:
                return GROUPING_FACTORS['spaced_3d'].get(min(num_cables, 12), 0.78)
        else:
            return GROUPING_FACTORS['touching'].get(min(num_cables, 18), 0.39)

# ========== LIGHTNING PROTECTION CLASSES ==========
class LightningWordReport:
    def __init__(self):
        self.doc = Document()
        style = self.doc.styles['Normal']
        style.font.name = 'Arial'
        style.font.size = Pt(11)
    
    def add_calculations(self, results, inputs):
        title = self.doc.add_heading('LIGHTNING PROTECTION CALCULATIONS', 0)
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        self.doc.add_paragraph(f'Date: {datetime.now().strftime("%Y-%m-%d %H:%M")}')
        self.doc.add_paragraph('Reference: IEC 62305-2')
        self.doc.add_paragraph()
        
        self.doc.add_heading('1.1 Collection Area (Ad)', level=1)
        self.doc.add_paragraph('Formula: Ad = L x W + 2 x (3H) x (L + W) + pi x (3H)^2')
        self.doc.add_paragraph('Reference: IEC 62305-2 Annex A.2.1.1')
        p = self.doc.add_paragraph()
        p.add_run('Result: ').bold = True
        p.add_run(f'Ad = {results["ad"]:.2f} m²')
        
        self.doc.add_heading('1.2 Near Strike Collection Area (Am)', level=1)
        self.doc.add_paragraph('Formula: Am = 2 x 500 x (L + W) + pi x 500^2')
        self.doc.add_paragraph('Reference: IEC 62305-2 Annex A.3')
        p = self.doc.add_paragraph()
        p.add_run('Result: ').bold = True
        p.add_run(f'Am = {results["am"]:.2f} m²')
        
        self.doc.add_heading('1.3 Environmental Factor (CD)', level=1)
        self.doc.add_paragraph(f'Selected Environment: {inputs.get("environment", "Isolated")}')
        p = self.doc.add_paragraph()
        p.add_run('Result: ').bold = True
        p.add_run(f'CD = {inputs.get("cd", 1)}')
        
        self.doc.add_heading('1.4 Lightning Ground Flash Density (NG)', level=1)
        self.doc.add_paragraph('Formula: NG = 0.1 x Td')
        p = self.doc.add_paragraph()
        p.add_run('Result: ').bold = True
        p.add_run(f'NG = {results.get("ng", 1)} flashes/km²/year')
        
        self.doc.add_heading('1.5 Lightning Frequencies', level=1)
        p = self.doc.add_paragraph()
        p.add_run('Nd (Direct): ').bold = True
        p.add_run(f'{results.get("nd", 0):.6f} events/year')
        p = self.doc.add_paragraph()
        p.add_run('Nm (Near): ').bold = True
        p.add_run(f'{results.get("nm", 0):.6f} events/year')
        
        self.doc.add_heading('1.6 Protection Level', level=1)
        p = self.doc.add_paragraph()
        p.add_run('Result: ').bold = True
        p.add_run(f'{results.get("lpl", "Class III")}')
        self.doc.add_paragraph(f'Rolling Sphere Radius: {results.get("sphere", 45)}m')
        
        self.doc.add_page_break()
        self.doc.add_heading('SUMMARY OF RESULTS', level=1)
        table = self.doc.add_table(rows=1, cols=2)
        table.style = 'Light Grid Accent 1'
        hdr_cells = table.rows[0].cells
        hdr_cells[0].text = 'Parameter'
        hdr_cells[1].text = 'Value'
        
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
        footer.add_run(f'Generated by CES-Electrical on {datetime.now().strftime("%Y-%m-%d %H:%M")}').italic = True
    
    def save(self, filename):
        self.doc.save(filename)

class LightningPDFReport(FPDF):
    def __init__(self):
        super().__init__()
        self.set_auto_page_break(auto=True, margin=25)
    
    def header(self):
        if self.page_no() > 1:
            self.set_font('Arial', 'I', 10)
            self.cell(0, 12, 'Lightning Protection Calculation', 0, 0, 'L')
            self.cell(0, 12, f'Page {self.page_no()}', 0, 0, 'R')
            self.ln(18)
    
    def footer(self):
        self.set_y(-20)
        self.set_font('Arial', 'I', 8)
        self.cell(0, 10, f'Page {self.page_no()}', 0, 0, 'C')
    
    def add_calculations(self, results, inputs):
        self.add_page()
        self.set_font('Arial', 'B', 20)
        self.set_text_color(0, 51, 102)
        self.cell(0, 20, 'LIGHTNING PROTECTION CALCULATIONS', 0, 1, 'C')
        self.ln(10)
        
        self.set_font('Arial', '', 10)
        self.set_text_color(0, 0, 0)
        self.cell(0, 6, f'Date: {datetime.now().strftime("%Y-%m-%d %H:%M")}', 0, 1, 'R')
        self.cell(0, 6, 'Reference: IEC 62305-2', 0, 1, 'R')
        self.ln(10)
        
        self.set_font('Arial', 'B', 14)
        self.set_text_color(0, 51, 102)
        self.cell(0, 10, '1.1 Collection Area (Ad)', 0, 1)
        self.set_font('Arial', '', 11)
        self.set_text_color(0, 0, 0)
        self.multi_cell(0, 7, 'Formula: Ad = L x W + 2 x (3H) x (L + W) + pi x (3H)^2')
        self.cell(0, 7, 'Reference: IEC 62305-2 Annex A.2.1.1', 0, 1)
        self.ln(2)
        if inputs.get('width', 0) == 0:
            self.cell(0, 7, f'Calculation: Ad = pi x 9 x ({inputs["height"]})^2', 0, 1)
        else:
            self.cell(0, 7, f'Calculation: Ad = {inputs["length"]} x {inputs["width"]} + 2 x (3 x {inputs["height"]}) x ({inputs["length"]} + {inputs["width"]}) + pi x (3 x {inputs["height"]})^2', 0, 1)
        self.set_font('Arial', 'B', 12)
        self.cell(0, 8, f'Result: Ad = {results["ad"]:.2f} m²', 0, 1)
        self.ln(8)
        
        self.set_font('Arial', 'B', 14)
        self.set_text_color(0, 51, 102)
        self.cell(0, 10, '1.2 Near Strike Collection Area (Am)', 0, 1)
        self.set_font('Arial', '', 11)
        self.set_text_color(0, 0, 0)
        self.multi_cell(0, 7, 'Formula: Am = 2 x 500 x (L + W) + pi x 500^2')
        self.cell(0, 7, 'Reference: IEC 62305-2 Annex A.3', 0, 1)
        self.ln(2)
        self.cell(0, 7, f'Calculation: Am = 2 x 500 x ({inputs["length"]} + {inputs["width"]}) + pi x 500^2', 0, 1)
        self.set_font('Arial', 'B', 12)
        self.cell(0, 8, f'Result: Am = {results["am"]:.2f} m²', 0, 1)
        self.ln(8)
        
        self.set_font('Arial', 'B', 14)
        self.set_text_color(0, 51, 102)
        self.cell(0, 10, '1.3 Environmental Factor (CD)', 0, 1)
        self.set_font('Arial', '', 11)
        self.set_text_color(0, 0, 0)
        self.cell(0, 7, 'Reference: IEC 62305-2 Table A.1', 0, 1)
        self.cell(0, 7, f'Selected Environment: {inputs.get("environment", "Isolated")}', 0, 1)
        self.set_font('Arial', 'B', 12)
        self.cell(0, 8, f'Result: CD = {inputs.get("cd", 1)}', 0, 1)
        self.ln(8)
        
        self.set_font('Arial', 'B', 14)
        self.set_text_color(0, 51, 102)
        self.cell(0, 10, '1.4 Lightning Ground Flash Density (NG)', 0, 1)
        self.set_font('Arial', '', 11)
        self.set_text_color(0, 0, 0)
        self.cell(0, 7, 'Formula: NG = 0.1 x Td', 0, 1)
        self.cell(0, 7, f'Calculation: NG = 0.1 x {inputs.get("td_days", 10)}', 0, 1)
        self.set_font('Arial', 'B', 12)
        self.cell(0, 8, f'Result: NG = {results.get("ng", 1)} flashes/km²/year', 0, 1)
        self.ln(8)
        
        self.set_font('Arial', 'B', 14)
        self.set_text_color(0, 51, 102)
        self.cell(0, 10, '1.5 Lightning Frequencies', 0, 1)
        self.set_font('Arial', '', 11)
        self.set_text_color(0, 0, 0)
        self.cell(0, 7, 'Direct Strike Frequency (Nd):', 0, 1)
        self.cell(0, 7, f'Calculation: Nd = {results.get("ng", 1)} x {results["ad"]:.0f} x {inputs.get("cd", 1)} x 10^-6', 0, 1)
        self.set_font('Arial', 'B', 12)
        self.cell(0, 8, f'Result: Nd = {results.get("nd", 0):.6f} events/year', 0, 1)
        self.ln(4)
        self.set_font('Arial', '', 11)
        self.cell(0, 7, 'Near Strike Frequency (Nm):', 0, 1)
        self.cell(0, 7, f'Calculation: Nm = {results.get("ng", 1)} x {results["am"]:.0f} x 10^-6', 0, 1)
        self.set_font('Arial', 'B', 12)
        self.cell(0, 8, f'Result: Nm = {results.get("nm", 0):.6f} events/year', 0, 1)
        self.ln(8)
        
        self.set_font('Arial', 'B', 14)
        self.set_text_color(0, 51, 102)
        self.cell(0, 10, '1.6 Protection Level Determination', 0, 1)
        self.set_font('Arial', '', 11)
        self.set_text_color(0, 0, 0)
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
        self.cell(0, 8, f'Rolling Sphere Radius: {results.get("sphere", 45)}m', 0, 1)
        self.ln(8)
        
        self.set_font('Arial', 'B', 14)
        self.set_text_color(0, 51, 102)
        self.cell(0, 10, '1.7 Air Terminals Required', 0, 1)
        self.set_font('Arial', '', 11)
        self.set_text_color(0, 0, 0)
        self.cell(0, 7, 'Method: Rolling Sphere Method', 0, 1)
        self.set_font('Arial', 'B', 12)
        self.cell(0, 8, f'Result: {results.get("air_terminals", 4)} air terminals required', 0, 1)
        self.ln(10)
        
        self.add_page()
        self.set_font('Arial', 'B', 16)
        self.set_text_color(0, 51, 102)
        self.cell(0, 12, 'SUMMARY OF RESULTS', 0, 1, 'C')
        self.ln(6)
        
        self.set_font('Arial', 'B', 11)
        self.set_fill_color(240, 240, 240)
        self.cell(80, 8, 'Parameter', 1, 0, 'C', 1)
        self.cell(90, 8, 'Value', 1, 1, 'C', 1)
        
        self.set_font('Arial', '', 10)
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
        
        fill = False
        for param, value in summary_data:
            self.cell(80, 7, param, 1, 0, 'L', fill)
            self.cell(90, 7, value, 1, 1, 'R', fill)
            fill = not fill

# ========== CABLE SIZING CALCULATOR CLASS ==========
class CableSizingCalculator:
    def __init__(self):
        self.results = {}
    
    def calculate_load_current(self, power_kw, voltage_v, pf, efficiency=1.0, phase='3-phase'):
        if phase == '3-phase':
            return (power_kw * 1000) / (1.732 * voltage_v * pf * efficiency)
        elif phase == '1-phase':
            return (power_kw * 1000) / (voltage_v * pf * efficiency)
        else:
            return (power_kw * 1000) / voltage_v
    
    def get_all_derating_factors(self, temp_c, insulation_temp=90, num_cables=1, 
                                  arrangement='touching', spacing_mm=0, cable_diameter=0,
                                  formation='flat', installation='air',
                                  soil_resistivity=1.5, depth=0.8):
        k1 = TEMPERATURE_FACTORS[insulation_temp].get(temp_c, 1.0)
        
        if arrangement == 'touching':
            k2 = GROUPING_FACTORS['touching'].get(min(num_cables, 18), 0.39)
        elif arrangement == 'cleated':
            k2 = GROUPING_FACTORS['cleated'].get(min(num_cables, 12), 0.68)
        else:
            if cable_diameter > 0:
                spacing_ratio = spacing_mm / cable_diameter
                if spacing_ratio < 0.5:
                    k2 = GROUPING_FACTORS['touching'].get(min(num_cables, 18), 0.39)
                elif spacing_ratio < 1.5:
                    k2 = GROUPING_FACTORS['spaced_1d'].get(min(num_cables, 12), 0.66)
                elif spacing_ratio < 2.5:
                    k2 = GROUPING_FACTORS['spaced_2d'].get(min(num_cables, 12), 0.74)
                else:
                    k2 = GROUPING_FACTORS['spaced_3d'].get(min(num_cables, 12), 0.78)
            else:
                k2 = GROUPING_FACTORS['touching'].get(min(num_cables, 18), 0.39)
        
        k_formation = FORMATION_FACTORS.get(formation, 1.0)
        k_install = INSTALLATION_FACTORS.get(installation, 1.0)
        
        if installation in ['buried', 'duct', 'trench']:
            k3 = SOIL_RESISTIVITY_FACTORS.get(soil_resistivity, 1.0)
            k4 = DEPTH_FACTORS.get(depth, 1.0)
        else:
            k3 = 1.0
            k4 = 1.0
        
        total_k = k1 * k2 * k_formation * k_install * k3 * k4
        
        factors = {
            'k1 (Temperature)': {'value': k1, 'reference': 'IEC 60502-2 Table B.10'},
            'k2 (Grouping/Spacing)': {'value': k2, 'reference': f'IEC 60502-2 Table 4C1'},
            'k_formation (Formation)': {'value': k_formation, 'reference': f'IEC 60502-2'},
            'k_install (Installation)': {'value': k_install, 'reference': f'IEC 60502-2'},
            'k3 (Soil Resistivity)': {'value': k3, 'reference': 'IEC 60502-2 Table B.14'},
            'k4 (Depth)': {'value': k4, 'reference': 'IEC 60502-2 Table B.12'},
            'total': total_k
        }
        return total_k, factors
    
    def calculate_voltage_drop(self, current, length_m, R, X, pf, voltage_v, phase='3-phase'):
        R_total = R * length_m / 1000
        X_total = X * length_m / 1000
        
        if phase == '3-phase':
            Vd = 1.732 * current * (R_total * pf + X_total * math.sin(math.acos(pf)))
        elif phase == '1-phase':
            Vd = 2 * current * (R_total * pf + X_total * math.sin(math.acos(pf)))
        else:
            Vd = 2 * current * R_total
        
        Vd_percent = (Vd / voltage_v) * 100
        return Vd, Vd_percent
    
    def calculate_short_circuit(self, size_mm2, duration_s=1.0):
        K = 143
        Isc = K * size_mm2 / math.sqrt(duration_s)
        return Isc
    
    def get_cable_type(self, voltage_v):
        if voltage_v <= 1000:
            return 'LV (0.6/1kV)', LV_CABLE_DATA
        else:
            return 'MV (3.6/6kV - 12/20kV)', MV_CABLE_DATA

# ========== CIRCUIT BREAKER CALCULATOR CLASS ==========
class CircuitBreakerCalculator:
    def __init__(self):
        pass
    
    def get_standard_rating(self, current, design_factor=1.25):
        required = current * design_factor
        for rating in CB_RATINGS:
            if rating >= required:
                return rating, required
        return CB_RATINGS[-1], required
    
    def get_breaker_type(self, rating):
        if rating <= 125:
            return 'MCB', 'IEC 60898'
        elif rating <= 1600:
            return 'MCCB', 'IEC 60947-2'
        else:
            return 'ACB', 'IEC 60947-2'
    
    def select_poles(self, phase, system_type='TN-S'):
        if phase == '1-phase':
            if system_type in ['TN-S', 'TN-C-S', 'TT']:
                return '2P', 'Phase + Neutral protection required for TN/TT systems.'
            else:
                return '1P', 'Phase only protection - For IT systems only.'
        elif phase == '3-phase':
            if system_type == 'TN-S':
                return '4P', '4-Pole required for TN-S systems with separate neutral.'
            elif system_type == 'TN-C':
                return '3P', '3-Pole for TN-C systems (PEN conductor).'
            else:
                return '3P', '3-Pole standard for 3-wire systems.'
        else:
            return '2P', '2-Pole required for DC circuits as per IEC 60947-2.'
    
    def calculate_cb_size(self, loads_df, design_factor=1.25, manufacturer='Schneider Electric', system_type='TN-S'):
        results = []
        detailed_reasons = []
        
        for idx, load in loads_df.iterrows():
            if load['Phase'] == '3-phase':
                current = load['Power (kW)'] * 1000 / (1.732 * load['Voltage (V)'] * load['Power Factor'])
                phase_desc = "Three-phase"
            elif load['Phase'] == '1-phase':
                current = load['Power (kW)'] * 1000 / (load['Voltage (V)'] * load['Power Factor'])
                phase_desc = "Single-phase"
            else:
                current = load['Power (kW)'] * 1000 / load['Voltage (V)']
                phase_desc = "DC"
            
            rating, required = self.get_standard_rating(current, design_factor)
            breaker_type, standard = self.get_breaker_type(rating)
            poles, reason = self.select_poles(load['Phase'], system_type)
            series = MANUFACTURERS[manufacturer][breaker_type]
            
            results.append({
                'Load': load['Load Name'],
                'Power (kW)': load['Power (kW)'],
                'Voltage (V)': load['Voltage (V)'],
                'Phase': load['Phase'],
                'Current (A)': current,
                'Required CB (A)': required,
                'Selected CB (A)': rating,
                'Breaker Type': breaker_type,
                'Standard': standard,
                'Poles': poles,
                'Manufacturer': manufacturer,
                'Series': series
            })
            
            detailed_reasons.append({
                'load_name': load['Load Name'],
                'phase_desc': phase_desc,
                'current': current,
                'required': required,
                'selected': rating,
                'breaker_type': breaker_type,
                'standard': standard,
                'poles': poles,
                'pole_reason': reason,
                'design_factor': design_factor,
                'system_type': system_type,
                'manufacturer': manufacturer,
                'series': series
            })
        
        return results, detailed_reasons
    
    def calculate_main_cb(self, loads_df, voltage=400, pf=0.85, design_factor=1.25, system_type='TN-S'):
        total_power = loads_df['Power (kW)'].sum()
        current = total_power * 1000 / (1.732 * voltage * pf)
        rating, required = self.get_standard_rating(current, design_factor)
        breaker_type, standard = self.get_breaker_type(rating)
        poles, reason = self.select_poles('3-phase', system_type)
        
        detailed_reason = f"""
MAIN CIRCUIT BREAKER DETAILED CALCULATION

STEP 1: TOTAL LOAD ANALYSIS
- Total Connected Load: {total_power:.2f} kW
- System Voltage: {voltage} V (3-phase)
- Power Factor: {pf}
- System Type: {system_type}

STEP 2: TOTAL CURRENT CALCULATION [IEC 60364-5-52]
Formula: I = P x 1000 / (1.732 x V x PF)
I = {total_power:.2f} x 1000 / (1.732 x {voltage} x {pf})
I = {current:.2f} A

STEP 3: CIRCUIT BREAKER SIZING [IEC 60898/IEC 60947-2]
- Design Factor: {design_factor}
- Required Rating = {current:.2f} x {design_factor} = {required:.2f} A
- Selected Standard Rating: {rating} A

STEP 4: BREAKER TYPE SELECTION
- Based on rating {rating} A -> {breaker_type} ({standard})
- Application: {BREAKER_TYPES[breaker_type]['application']}

STEP 5: POLE SELECTION [IEC 60364-5-53]
- System Type: {system_type}
- Selected Poles: {poles}
- Reason: {reason}

STEP 6: MANUFACTURER SELECTION
- Manufacturer: Schneider Electric
- Series: {MANUFACTURERS['Schneider Electric'][breaker_type]}

FINAL SELECTION: {rating} A {breaker_type} {poles}
"""
        
        return {
            'total_power': total_power,
            'current': current,
            'required_cb': required,
            'selected_cb': rating,
            'breaker_type': breaker_type,
            'poles': poles,
            'standard': standard,
            'detailed_reason': detailed_reason
        }

# ========== PDF REPORT CLASSES ==========
class CablePDFReport(FPDF):
    def __init__(self):
        super().__init__()
        self.set_auto_page_break(auto=True, margin=25)
    
    def header(self):
        self.set_font('Arial', 'B', 12)
        self.set_text_color(0, 51, 102)
        self.cell(0, 10, 'CABLE SIZING & CIRCUIT BREAKER REPORT', 0, 1, 'C')
        self.line(10, 25, 200, 25)
        self.ln(10)
    
    def footer(self):
        self.set_y(-20)
        self.set_font('Arial', 'I', 8)
        self.set_text_color(128, 128, 128)
        self.cell(0, 10, f'Page {self.page_no()} | Generated on {datetime.now().strftime("%Y-%m-%d")}', 0, 0, 'C')
    
    def add_title(self):
        self.add_page()
        self.set_font('Arial', 'B', 20)
        self.set_text_color(0, 51, 102)
        self.cell(0, 20, 'CABLE SIZING & CIRCUIT BREAKER REPORT', 0, 1, 'C')
        self.ln(5)
        self.set_font('Arial', '', 10)
        self.set_text_color(0, 0, 0)
        self.cell(0, 6, f'Date: {datetime.now().strftime("%Y-%m-%d %H:%M")}', 0, 1, 'R')
        self.ln(10)
    
    def add_installation_parameters(self, params):
        self.set_font('Arial', 'B', 14)
        self.set_text_color(0, 51, 102)
        self.cell(0, 10, '1. INSTALLATION PARAMETERS', 0, 1)
        self.ln(2)
        self.set_font('Arial', '', 11)
        self.set_text_color(0, 0, 0)
        for key, value in params.items():
            self.cell(50, 7, key + ':', 0, 0)
            self.cell(0, 7, value, 0, 1)
        self.ln(10)
    
    def add_derating_factors(self, factors):
        self.set_font('Arial', 'B', 14)
        self.set_text_color(0, 51, 102)
        self.cell(0, 10, '2. DERATING FACTORS (IEC 60502-2)', 0, 1)
        self.ln(2)
        
        self.set_font('Arial', 'B', 10)
        self.set_fill_color(240, 240, 240)
        self.cell(50, 8, 'Factor', 1, 0, 'C', 1)
        self.cell(30, 8, 'Value', 1, 0, 'C', 1)
        self.cell(100, 8, 'Reference', 1, 1, 'C', 1)
        
        self.set_font('Arial', '', 10)
        fill = False
        for key, data in factors.items():
            if key != 'total':
                self.cell(50, 7, key, 1, 0, 'L', fill)
                self.cell(30, 7, f"{data['value']:.3f}", 1, 0, 'C', fill)
                self.cell(100, 7, data['reference'], 1, 1, 'L', fill)
                fill = not fill
        
        self.set_font('Arial', 'B', 10)
        self.set_fill_color(0, 51, 102)
        self.set_text_color(255, 255, 255)
        self.cell(80, 8, 'Total Derating Factor K', 1, 0, 'L', 1)
        self.cell(100, 8, f"{factors['total']:.3f}", 1, 1, 'C', 1)
        self.set_text_color(0, 0, 0)
        self.ln(10)
    
    def add_load_details(self, loads_df):
        self.set_font('Arial', 'B', 14)
        self.set_text_color(0, 51, 102)
        self.cell(0, 10, '3. LOAD DETAILS', 0, 1)
        self.ln(2)
        
        self.set_font('Arial', 'B', 9)
        self.set_fill_color(240, 240, 240)
        self.cell(35, 8, 'Load Name', 1, 0, 'C', 1)
        self.cell(20, 8, 'Power', 1, 0, 'C', 1)
        self.cell(20, 8, 'Voltage', 1, 0, 'C', 1)
        self.cell(20, 8, 'Phase', 1, 0, 'C', 1)
        self.cell(15, 8, 'PF', 1, 0, 'C', 1)
        self.cell(20, 8, 'Length', 1, 1, 'C', 1)
        
        self.set_font('Arial', '', 8)
        fill = False
        for idx, load in loads_df.iterrows():
            self.cell(35, 6, load['Load Name'][:20], 1, 0, 'L', fill)
            self.cell(20, 6, f"{load['Power (kW)']:.1f} kW", 1, 0, 'R', fill)
            self.cell(20, 6, f"{load['Voltage (V)']:.0f} V", 1, 0, 'R', fill)
            self.cell(20, 6, load['Phase'], 1, 0, 'C', fill)
            self.cell(15, 6, f"{load['Power Factor']:.2f}", 1, 0, 'R', fill)
            self.cell(20, 6, f"{load['Length (m)']:.0f} m", 1, 1, 'R', fill)
            fill = not fill
        self.ln(10)
    
    def add_cable_results(self, cable_df):
        self.set_font('Arial', 'B', 14)
        self.set_text_color(0, 51, 102)
        self.cell(0, 10, '4. CABLE SIZING RESULTS', 0, 1)
        self.ln(2)
        
        self.set_font('Arial', 'B', 7)
        self.set_fill_color(240, 240, 240)
        headers = ['Load', 'Size', 'Type', 'Base A', 'Derated', 'VD %', 'SC kA', 'Eff%', 'Status']
        widths = [25, 12, 12, 15, 15, 12, 12, 12, 15]
        
        for i, header in enumerate(headers):
            self.cell(widths[i], 8, header, 1, 0, 'C', 1)
        self.ln()
        
        self.set_font('Arial', '', 6)
        fill = False
        for idx, row in cable_df.iterrows():
            data = [
                row['Load Name'][:15],
                str(row['Size (mm²)']),
                'Cu',
                str(row['Base Ampacity (A)']),
                str(row['Derated Ampacity (A)']).replace(' A', ''),
                str(row['Voltage Drop (%)']).replace('%', ''),
                str(row['Short Circuit (kA)']).replace(' kA', ''),
                str(row['Efficiency (%)']).replace('%', ''),
                row['Status']
            ]
            
            for i, value in enumerate(data):
                align = 'R' if i > 2 else 'L'
                self.cell(widths[i], 5, value, 1, 0, align, fill)
            self.ln()
            fill = not fill
        self.ln(10)
    
    def add_detailed_cable_calculations(self, detailed_calcs):
        self.add_page()
        self.set_font('Arial', 'B', 16)
        self.set_text_color(0, 51, 102)
        self.cell(0, 15, '5. DETAILED CABLE CALCULATIONS WITH REFERENCES', 0, 1, 'C')
        self.ln(5)
        
        for i, calc in enumerate(detailed_calcs):
            if self.get_y() > 250:
                self.add_page()
            
            self.set_font('Arial', 'B', 14)
            self.set_text_color(0, 51, 102)
            self.cell(0, 10, f'LOAD {i+1}: {calc["load_name"]}', 0, 1)
            self.ln(2)
            
            self.set_font('Arial', 'B', 12)
            self.set_text_color(0, 51, 102)
            self.cell(0, 8, 'STEP 1: LOAD CURRENT CALCULATION', 0, 1)
            self.set_font('Arial', '', 10)
            self.set_text_color(0, 0, 0)
            self.cell(0, 6, 'Reference: IEC 60364-5-52 Section 523', 0, 1)
            self.cell(0, 6, 'Formula: I = P x 1000 / (1.732 x V x PF) for 3-phase', 0, 1)
            self.cell(0, 6, f'P = {calc["power"]} kW, V = {calc["voltage"]} V, PF = {calc["pf"]}', 0, 1)
            self.cell(0, 6, f'I = {calc["power"]} x 1000 / (1.732 x {calc["voltage"]} x {calc["pf"]})', 0, 1)
            self.set_font('Arial', 'B', 10)
            self.cell(0, 6, f'LOAD CURRENT = {calc["current"]:.1f} A', 0, 1)
            self.ln(3)
            
            self.set_font('Arial', 'B', 12)
            self.set_text_color(0, 51, 102)
            self.cell(0, 8, 'STEP 2: DERATING FACTORS CALCULATION', 0, 1)
            self.set_font('Arial', '', 10)
            self.set_text_color(0, 0, 0)
            self.cell(0, 6, 'Reference: IEC 60502-2 Tables B.10-B.22', 0, 1)
            self.cell(0, 6, f'k1 (Temperature Correction) : {calc["k1"]:.3f} - Table B.10 at {calc["ambient_temp"]}°C', 0, 1)
            self.cell(0, 6, f'k2 (Grouping/Spacing)      : {calc["k2"]:.3f} - {calc["arrangement"]}, spacing={calc["spacing"]}mm', 0, 1)
            self.cell(0, 6, f'k_formation (Formation)    : {calc["k_formation"]:.3f} - {calc["formation"]} formation', 0, 1)
            self.cell(0, 6, f'k_install (Installation)   : {calc["k_install"]:.3f} - {calc["installation"]} method', 0, 1)
            self.cell(0, 6, f'k3 (Soil Resistivity)      : {calc["k3"]:.3f} - Table B.14', 0, 1)
            self.cell(0, 6, f'k4 (Depth)                 : {calc["k4"]:.3f} - Table B.12', 0, 1)
            self.set_font('Arial', 'B', 10)
            self.cell(0, 6, f'TOTAL K = {calc["total_k"]:.3f}', 0, 1)
            self.ln(3)
            
            self.set_font('Arial', 'B', 12)
            self.set_text_color(0, 51, 102)
            self.cell(0, 8, 'STEP 3: CABLE SELECTION', 0, 1)
            self.set_font('Arial', '', 10)
            self.set_text_color(0, 0, 0)
            self.cell(0, 6, f'Cable Category: {calc["cable_category"]}', 0, 1)
            self.cell(0, 6, f'Cable Type: {calc["cable_type"]} copper', 0, 1)
            self.cell(0, 6, f'Selected Cable Size: {calc["size"]} mm²', 0, 1)
            self.cell(0, 6, f'Base Ampacity: {calc["base_amp"]} A', 0, 1)
            self.cell(0, 6, f'Derated Ampacity = Base Ampacity x K', 0, 1)
            self.cell(0, 6, f'Derated Ampacity = {calc["base_amp"]} x {calc["total_k"]:.3f} = {calc["derated_amp"]:.1f} A', 0, 1)
            status1 = 'PASS' if calc['derated_amp'] >= calc['current'] else 'FAIL'
            self.set_font('Arial', 'B', 10)
            self.cell(0, 6, f'Check: {calc["derated_amp"]:.1f} A >= {calc["current"]:.1f} A ? {status1}', 0, 1)
            self.ln(3)
            
            self.set_font('Arial', 'B', 12)
            self.set_text_color(0, 51, 102)
            self.cell(0, 8, 'STEP 4: VOLTAGE DROP CALCULATION', 0, 1)
            self.set_font('Arial', '', 10)
            self.set_text_color(0, 0, 0)
            self.cell(0, 6, 'Reference: IEC 60364-5-52 Section 525', 0, 1)
            self.cell(0, 6, f'Cable Length: {calc["length"]} m', 0, 1)
            self.cell(0, 6, f'Voltage Drop = {calc["vd_pct"]:.3f}%', 0, 1)
            self.cell(0, 6, 'Maximum Allowable Voltage Drop: 2.5%', 0, 1)
            status2 = 'PASS' if calc['vd_pct'] <= 2.5 else 'FAIL'
            self.set_font('Arial', 'B', 10)
            self.cell(0, 6, f'Check: {calc["vd_pct"]:.3f}% <= 2.5% ? {status2}', 0, 1)
            self.ln(3)
            
            self.set_font('Arial', 'B', 12)
            self.set_text_color(0, 51, 102)
            self.cell(0, 8, 'STEP 5: SHORT CIRCUIT CALCULATION', 0, 1)
            self.set_font('Arial', '', 10)
            self.set_text_color(0, 0, 0)
            self.cell(0, 6, 'Reference: IEC 60949', 0, 1)
            self.cell(0, 6, 'Formula: Isc = K x S / sqrt(t)', 0, 1)
            self.cell(0, 6, 'K = 143 for Copper, XLPE insulated', 0, 1)
            self.cell(0, 6, f'S = {calc["size"]} mm²', 0, 1)
            self.cell(0, 6, f't = 1.0 s', 0, 1)
            self.set_font('Arial', 'B', 10)
            self.cell(0, 6, f'SHORT CIRCUIT CAPACITY = {calc["sc"]:.2f} kA', 0, 1)
            self.ln(3)
            
            self.set_font('Arial', 'B', 12)
            self.set_text_color(0, 51, 102)
            self.cell(0, 8, 'STEP 6: EFFICIENCY CALCULATION', 0, 1)
            self.set_font('Arial', '', 10)
            self.set_text_color(0, 0, 0)
            self.cell(0, 6, f'Input Power = 1.732 x V x I = 1.732 x {calc["voltage"]} x {calc["current"]:.1f} / 1000 = {calc["input_power"]:.1f} kW', 0, 1)
            self.cell(0, 6, f'Output Power = {calc["power"]} kW', 0, 1)
            self.set_font('Arial', 'B', 10)
            self.cell(0, 6, f'EFFICIENCY = {calc["efficiency"]:.1f}%', 0, 1)
            self.ln(3)
            
            self.set_font('Arial', 'B', 14)
            if calc['status'] == 'PASS':
                self.set_text_color(0, 128, 0)
                self.cell(0, 8, f'FINAL STATUS: PASS', 0, 1)
            else:
                self.set_text_color(255, 0, 0)
                self.cell(0, 8, f'FINAL STATUS: FAIL', 0, 1)
            self.set_text_color(0, 0, 0)
            
            self.ln(5)
            self.set_draw_color(200, 200, 200)
            self.line(10, self.get_y(), 200, self.get_y())
            self.ln(5)
    
    def add_detailed_cb_calculations(self, cb_details, main_cb):
        self.add_page()
        self.set_font('Arial', 'B', 16)
        self.set_text_color(0, 51, 102)
        self.cell(0, 15, '6. DETAILED CIRCUIT BREAKER CALCULATIONS', 0, 1, 'C')
        self.ln(5)
        
        self.set_font('Arial', 'B', 14)
        self.set_text_color(0, 51, 102)
        self.cell(0, 10, '6.1 Individual Circuit Breakers - Detailed Calculations', 0, 1)
        self.ln(2)
        
        for i, detail in enumerate(cb_details):
            if self.get_y() > 250:
                self.add_page()
            
            self.set_font('Arial', 'B', 12)
            self.set_text_color(0, 51, 102)
            self.cell(0, 8, f'LOAD {i+1}: {detail["load_name"]}', 0, 1)
            self.ln(1)
            
            self.set_font('Arial', 'B', 11)
            self.set_text_color(0, 0, 0)
            self.cell(0, 6, 'STEP 1: LOAD ANALYSIS', 0, 1)
            self.set_font('Arial', '', 10)
            self.cell(0, 5, f'  - Load Type: {detail["phase_desc"]}', 0, 1)
            self.cell(0, 5, f'  - Load Current: {detail["current"]:.2f} A', 0, 1)
            self.ln(1)
            
            self.set_font('Arial', 'B', 11)
            self.cell(0, 6, 'STEP 2: RATING CALCULATION [IEC 60364]', 0, 1)
            self.set_font('Arial', '', 10)
            self.cell(0, 5, f'  - Design Factor: {detail["design_factor"]}', 0, 1)
            self.cell(0, 5, f'  - Required Rating = {detail["current"]:.2f} x {detail["design_factor"]} = {detail["required"]:.2f} A', 0, 1)
            self.cell(0, 5, f'  - Selected Standard Rating: {detail["selected"]} A', 0, 1)
            self.ln(1)
            
            self.set_font('Arial', 'B', 11)
            self.cell(0, 6, 'STEP 3: BREAKER TYPE SELECTION', 0, 1)
            self.set_font('Arial', '', 10)
            self.cell(0, 5, f'  - Type: {detail["breaker_type"]} ({detail["standard"]})', 0, 1)
            self.cell(0, 5, f'  - Application: {BREAKER_TYPES[detail["breaker_type"]]["application"]}', 0, 1)
            self.ln(1)
            
            self.set_font('Arial', 'B', 11)
            self.cell(0, 6, 'STEP 4: POLE SELECTION [IEC 60364-5-53]', 0, 1)
            self.set_font('Arial', '', 10)
            self.cell(0, 5, f'  - System Type: {detail["system_type"]}', 0, 1)
            self.cell(0, 5, f'  - Selected Poles: {detail["poles"]}', 0, 1)
            self.cell(0, 5, f'  - Reason: {detail["pole_reason"]}', 0, 1)
            self.ln(1)
            
            self.set_font('Arial', 'B', 11)
            self.cell(0, 6, 'STEP 5: MANUFACTURER SELECTION', 0, 1)
            self.set_font('Arial', '', 10)
            self.cell(0, 5, f'  - Manufacturer: {detail["manufacturer"]}', 0, 1)
            self.cell(0, 5, f'  - Series: {detail["series"]}', 0, 1)
            self.ln(2)
            
            self.set_font('Arial', 'B', 11)
            self.set_text_color(0, 51, 102)
            self.cell(0, 6, f'FINAL SELECTION: {detail["selected"]} A {detail["breaker_type"]} {detail["poles"]}', 0, 1)
            self.set_text_color(0, 0, 0)
            self.ln(3)
            
            self.set_draw_color(200, 200, 200)
            self.line(10, self.get_y(), 200, self.get_y())
            self.ln(3)
        
        if self.get_y() > 220:
            self.add_page()
        
        self.set_font('Arial', 'B', 14)
        self.set_text_color(0, 51, 102)
        self.cell(0, 10, '6.2 Main Circuit Breaker - Detailed Calculation', 0, 1)
        self.ln(2)
        
        self.set_font('Arial', '', 10)
        self.set_text_color(0, 0, 0)
        lines = main_cb['detailed_reason'].split('\n')
        for line in lines:
            if line.strip():
                clean_line = line.strip()
                self.cell(0, 5, clean_line, 0, 1)
        self.ln(5)

class CableWordReport:
    def __init__(self):
        self.doc = Document()
        style = self.doc.styles['Normal']
        style.font.name = 'Arial'
        style.font.size = Pt(11)
        
        sections = self.doc.sections
        for section in sections:
            section.top_margin = Cm(2.5)
            section.bottom_margin = Cm(2.5)
            section.left_margin = Cm(2.5)
            section.right_margin = Cm(2.5)
    
    def add_title(self):
        title = self.doc.add_heading('CABLE SIZING & CIRCUIT BREAKER REPORT', 0)
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        title.runs[0].font.size = Pt(20)
        title.runs[0].font.color.rgb = RGBColor(0, 51, 102)
        
        p = self.doc.add_paragraph()
        p.add_run(f'Date: {datetime.now().strftime("%Y-%m-%d %H:%M")}').italic = True
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        self.doc.add_paragraph()
    
    def add_installation_parameters(self, params):
        heading = self.doc.add_heading('1. INSTALLATION PARAMETERS', level=1)
        heading.runs[0].font.color.rgb = RGBColor(0, 51, 102)
        
        table = self.doc.add_table(rows=len(params), cols=2)
        table.style = 'Light Grid Accent 1'
        
        for i, (key, value) in enumerate(params.items()):
            row = table.rows[i].cells
            row[0].text = key
            row[1].text = value
            row[0].paragraphs[0].runs[0].bold = True
        
        self.doc.add_paragraph()
    
    def add_derating_factors(self, factors):
        heading = self.doc.add_heading('2. DERATING FACTORS (IEC 60502-2)', level=1)
        heading.runs[0].font.color.rgb = RGBColor(0, 51, 102)
        
        table = self.doc.add_table(rows=len(factors), cols=3)
        table.style = 'Light Grid Accent 1'
        
        header_cells = table.rows[0].cells
        headers = ['Factor', 'Value', 'Reference']
        for i, header in enumerate(headers):
            header_cells[i].text = header
            header_cells[i].paragraphs[0].runs[0].bold = True
        
        for i, (key, data) in enumerate(factors.items()):
            if i < len(factors) - 1:
                row = table.add_row().cells
                row[0].text = key
                row[1].text = f"{data['value']:.3f}"
                row[2].text = data['reference']
        
        total_row = table.add_row().cells
        total_row[0].text = 'Total Derating Factor K'
        total_row[1].text = f"{factors['total']:.3f}"
        total_row[2].text = 'IEC 60502-2'
        for cell in total_row:
            cell.paragraphs[0].runs[0].bold = True
        
        self.doc.add_paragraph()
    
    def add_load_details(self, loads_df):
        heading = self.doc.add_heading('3. LOAD DETAILS', level=1)
        heading.runs[0].font.color.rgb = RGBColor(0, 51, 102)
        
        table = self.doc.add_table(rows=1, cols=6)
        table.style = 'Light Grid Accent 1'
        
        headers = ['Load Name', 'Power (kW)', 'Voltage (V)', 'Phase', 'PF', 'Length (m)']
        for i, header in enumerate(headers):
            table.rows[0].cells[i].text = header
            table.rows[0].cells[i].paragraphs[0].runs[0].bold = True
        
        for idx, load in loads_df.iterrows():
            row = table.add_row().cells
            row[0].text = load['Load Name']
            row[1].text = f"{load['Power (kW)']:.1f}"
            row[2].text = f"{load['Voltage (V)']:.0f}"
            row[3].text = load['Phase']
            row[4].text = f"{load['Power Factor']:.2f}"
            row[5].text = f"{load['Length (m)']:.0f}"
        
        self.doc.add_paragraph()
    
    def add_cable_results(self, cable_df):
        heading = self.doc.add_heading('4. CABLE SIZING RESULTS', level=1)
        heading.runs[0].font.color.rgb = RGBColor(0, 51, 102)
        
        table = self.doc.add_table(rows=1, cols=9)
        table.style = 'Light Grid Accent 1'
        
        headers = ['Load', 'Size', 'Type', 'Base A', 'Derated', 'VD %', 'SC kA', 'Eff%', 'Status']
        for i, header in enumerate(headers):
            table.rows[0].cells[i].text = header
            table.rows[0].cells[i].paragraphs[0].runs[0].bold = True
        
        for idx, row in cable_df.iterrows():
            new_row = table.add_row().cells
            new_row[0].text = row['Load Name']
            new_row[1].text = str(row['Size (mm²)'])
            new_row[2].text = 'Cu'
            new_row[3].text = str(row['Base Ampacity (A)'])
            new_row[4].text = str(row['Derated Ampacity (A)']).replace(' A', '')
            new_row[5].text = str(row['Voltage Drop (%)']).replace('%', '')
            new_row[6].text = str(row['Short Circuit (kA)']).replace(' kA', '')
            new_row[7].text = str(row['Efficiency (%)']).replace('%', '')
            new_row[8].text = row['Status']
        
        self.doc.add_paragraph()
    
    def add_detailed_cable_calculations(self, detailed_calcs):
        self.doc.add_page_break()
        heading = self.doc.add_heading('5. DETAILED CABLE CALCULATIONS WITH REFERENCES', level=1)
        heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
        heading.runs[0].font.color.rgb = RGBColor(0, 51, 102)
        
        for i, calc in enumerate(detailed_calcs):
            self.doc.add_heading(f'LOAD {i+1}: {calc["load_name"]}', level=2)
            
            self.doc.add_heading('STEP 1: LOAD CURRENT CALCULATION [IEC 60364-5-52 Section 523]', level=3)
            p = self.doc.add_paragraph()
            p.add_run('Formula: ').bold = True
            p.add_run('I = P x 1000 / (1.732 x V x PF) for 3-phase')
            p = self.doc.add_paragraph()
            p.add_run('Calculation: ').bold = True
            p.add_run(f'I = {calc["power"]} x 1000 / (1.732 x {calc["voltage"]} x {calc["pf"]}) = {calc["current"]:.1f} A')
            
            self.doc.add_heading('STEP 2: DERATING FACTORS [IEC 60502-2 Tables B.10-B.22]', level=3)
            self.doc.add_paragraph(f'k1 (Temperature Correction): {calc["k1"]:.3f} - Table B.10 at {calc["ambient_temp"]}°C')
            self.doc.add_paragraph(f'k2 (Grouping/Spacing)      : {calc["k2"]:.3f} - {calc["arrangement"]}, spacing={calc["spacing"]}mm')
            self.doc.add_paragraph(f'k_formation (Formation)    : {calc["k_formation"]:.3f} - {calc["formation"]} formation')
            self.doc.add_paragraph(f'k_install (Installation)   : {calc["k_install"]:.3f} - {calc["installation"]} method')
            self.doc.add_paragraph(f'k3 (Soil Resistivity)      : {calc["k3"]:.3f} - Table B.14')
            self.doc.add_paragraph(f'k4 (Depth)                 : {calc["k4"]:.3f} - Table B.12')
            p = self.doc.add_paragraph()
            p.add_run('Total K = ').bold = True
            p.add_run(f'{calc["total_k"]:.3f}')
            
            self.doc.add_heading('STEP 3: CABLE SELECTION', level=3)
            p = self.doc.add_paragraph()
            p.add_run('Selected Cable: ').bold = True
            p.add_run(f'{calc["size"]} mm² {calc["cable_type"]} copper ({calc["cable_category"]})')
            p = self.doc.add_paragraph()
            p.add_run('Base Ampacity: ').bold = True
            p.add_run(f'{calc["base_amp"]} A')
            p = self.doc.add_paragraph()
            p.add_run('Derated Ampacity: ').bold = True
            p.add_run(f'{calc["derated_amp"]:.1f} A')
            p = self.doc.add_paragraph()
            status = 'PASS' if calc['derated_amp'] >= calc['current'] else 'FAIL'
            p.add_run('Check: ').bold = True
            check = p.add_run(f'{calc["derated_amp"]:.1f} A >= {calc["current"]:.1f} A ? {status}')
            if status == 'PASS':
                check.font.color.rgb = RGBColor(0, 128, 0)
            else:
                check.font.color.rgb = RGBColor(255, 0, 0)
            
            self.doc.add_heading('STEP 4: VOLTAGE DROP [IEC 60364-5-52 Section 525]', level=3)
            p = self.doc.add_paragraph()
            p.add_run('Voltage Drop: ').bold = True
            p.add_run(f'{calc["vd_pct"]:.3f}%')
            p = self.doc.add_paragraph()
            p.add_run('Limit: ').bold = True
            p.add_run('2.5%')
            p = self.doc.add_paragraph()
            status = 'PASS' if calc['vd_pct'] <= 2.5 else 'FAIL'
            p.add_run('Check: ').bold = True
            check = p.add_run(f'{calc["vd_pct"]:.3f}% <= 2.5% ? {status}')
            if status == 'PASS':
                check.font.color.rgb = RGBColor(0, 128, 0)
            else:
                check.font.color.rgb = RGBColor(255, 0, 0)
            
            self.doc.add_heading('STEP 5: SHORT CIRCUIT [IEC 60949]', level=3)
            p = self.doc.add_paragraph()
            p.add_run('Formula: ').bold = True
            p.add_run('Isc = K x S / sqrt(t)')
            p = self.doc.add_paragraph()
            p.add_run('Calculation: ').bold = True
            p.add_run(f'Isc = 143 x {calc["size"]} / sqrt(1.0) = {calc["sc"]:.2f} kA')
            
            self.doc.add_heading('STEP 6: EFFICIENCY', level=3)
            p = self.doc.add_paragraph()
            p.add_run('Efficiency: ').bold = True
            p.add_run(f'{calc["efficiency"]:.1f}%')
            
            self.doc.add_heading('FINAL STATUS', level=3)
            p = self.doc.add_paragraph()
            if calc['status'] == 'PASS':
                final_status = p.add_run(f'PASS')
                final_status.font.color.rgb = RGBColor(0, 128, 0)
            else:
                final_status = p.add_run(f'FAIL')
                final_status.font.color.rgb = RGBColor(255, 0, 0)
            final_status.font.size = Pt(14)
            final_status.font.bold = True
            
            self.doc.add_paragraph('_' * 50)
            self.doc.add_paragraph()
    
    def add_detailed_cb_calculations(self, cb_details, main_cb):
        self.doc.add_page_break()
        heading = self.doc.add_heading('6. DETAILED CIRCUIT BREAKER CALCULATIONS', level=1)
        heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
        heading.runs[0].font.color.rgb = RGBColor(0, 51, 102)
        
        self.doc.add_heading('6.1 Individual Circuit Breakers - Detailed Calculations', level=2)
        
        for i, detail in enumerate(cb_details):
            self.doc.add_heading(f'LOAD {i+1}: {detail["load_name"]}', level=3)
            
            self.doc.add_heading('STEP 1: LOAD ANALYSIS', level=4)
            self.doc.add_paragraph(f'  - Load Type: {detail["phase_desc"]}')
            self.doc.add_paragraph(f'  - Load Current: {detail["current"]:.2f} A')
            
            self.doc.add_heading('STEP 2: RATING CALCULATION [IEC 60364]', level=4)
            self.doc.add_paragraph(f'  - Design Factor: {detail["design_factor"]}')
            self.doc.add_paragraph(f'  - Required Rating = {detail["current"]:.2f} x {detail["design_factor"]} = {detail["required"]:.2f} A')
            self.doc.add_paragraph(f'  - Selected Standard Rating: {detail["selected"]} A')
            
            self.doc.add_heading('STEP 3: BREAKER TYPE SELECTION', level=4)
            self.doc.add_paragraph(f'  - Type: {detail["breaker_type"]} ({detail["standard"]})')
            self.doc.add_paragraph(f'  - Application: {BREAKER_TYPES[detail["breaker_type"]]["application"]}')
            
            self.doc.add_heading('STEP 4: POLE SELECTION [IEC 60364-5-53]', level=4)
            self.doc.add_paragraph(f'  - Selected Poles: {detail["poles"]}')
            self.doc.add_paragraph(f'  - Reason: {detail["pole_reason"]}')
            
            self.doc.add_heading('STEP 5: MANUFACTURER SELECTION', level=4)
            self.doc.add_paragraph(f'  - Manufacturer: {detail["manufacturer"]}')
            self.doc.add_paragraph(f'  - Series: {detail["series"]}')
            
            p = self.doc.add_paragraph()
            p.add_run('FINAL SELECTION: ').bold = True
            p.add_run(f'{detail["selected"]} A {detail["breaker_type"]} {detail["poles"]}')
            
            self.doc.add_paragraph()
        
        self.doc.add_heading('6.2 Main Circuit Breaker - Detailed Calculation', level=2)
        for line in main_cb['detailed_reason'].split('\n'):
            if line.strip():
                self.doc.add_paragraph(line.strip())
        
        self.doc.add_paragraph()
    
    def save(self, filename):
        self.doc.save(filename)

# ========== TRANSFORMER PDF REPORT CLASS ==========
class TransformerPDFReport(FPDF):
    def __init__(self):
        super().__init__()
        self.set_auto_page_break(auto=True, margin=25)
    
    def header(self):
        self.set_font('Arial', 'B', 12)
        self.set_text_color(0, 51, 102)
        self.cell(0, 10, 'TRANSFORMER SIZING REPORT', 0, 1, 'C')
        self.line(10, 25, 200, 25)
        self.ln(10)
    
    def footer(self):
        self.set_y(-20)
        self.set_font('Arial', 'I', 8)
        self.set_text_color(128, 128, 128)
        self.cell(0, 10, f'Page {self.page_no()} | Generated on {datetime.now().strftime("%Y-%m-%d")}', 0, 0, 'C')
    
    def add_title(self):
        self.add_page()
        self.set_font('Arial', 'B', 20)
        self.set_text_color(0, 51, 102)
        self.cell(0, 20, 'TRANSFORMER SIZING CALCULATIONS', 0, 1, 'C')
        self.ln(5)
        self.set_font('Arial', '', 10)
        self.set_text_color(0, 0, 0)
        self.cell(0, 6, f'Date: {datetime.now().strftime("%Y-%m-%d %H:%M")}', 0, 1, 'R')
        self.ln(10)
    
    def add_load_analysis(self, loads_df, tx_calc):
        self.set_font('Arial', 'B', 14)
        self.set_text_color(0, 51, 102)
        self.cell(0, 10, '1. LOAD ANALYSIS', 0, 1)
        self.ln(2)
        
        # Load details table
        self.set_font('Arial', 'B', 10)
        self.set_fill_color(240, 240, 240)
        self.cell(50, 8, 'Load Description', 1, 0, 'C', 1)
        self.cell(20, 8, 'Qty', 1, 0, 'C', 1)
        self.cell(25, 8, 'Rating (kW)', 1, 0, 'C', 1)
        self.cell(25, 8, 'Connected', 1, 0, 'C', 1)
        self.cell(25, 8, 'Diversity', 1, 0, 'C', 1)
        self.cell(25, 8, 'P (kW)', 1, 1, 'C', 1)
        
        self.set_font('Arial', '', 9)
        fill = False
        total_p = 0
        
        for idx, load in loads_df.iterrows():
            connected = load['Rating (kW)'] * load['Quantity']
            p = load['Rating (kW)'] * load['Quantity'] * load['Diversity Factor']
            total_p += p
            
            self.cell(50, 6, load['Load Description'][:20], 1, 0, 'L', fill)
            self.cell(20, 6, str(load['Quantity']), 1, 0, 'C', fill)
            self.cell(25, 6, f"{load['Rating (kW)']:.0f}", 1, 0, 'R', fill)
            self.cell(25, 6, f"{connected:.0f} kW", 1, 0, 'R', fill)
            self.cell(25, 6, f"{load['Diversity Factor']:.1f}", 1, 0, 'C', fill)
            self.cell(25, 6, f"{p:.1f}", 1, 1, 'R', fill)
            fill = not fill
        
        self.set_font('Arial', 'B', 10)
        self.set_fill_color(0, 51, 102)
        self.set_text_color(255, 255, 255)
        self.cell(145, 8, 'TOTAL REAL POWER (P)', 1, 0, 'R', 1)
        self.cell(25, 8, f"{total_p:.1f} kW", 1, 1, 'R', 1)
        self.set_text_color(0, 0, 0)
        self.ln(10)
    
    def add_step_by_step(self, loads_df, tx_calc):
        self.add_page()
        self.set_font('Arial', 'B', 14)
        self.set_text_color(0, 51, 102)
        self.cell(0, 10, '2. STEP-BY-STEP P, Q, S CALCULATIONS', 0, 1)
        self.ln(2)
        
        total_p = 0
        total_q = 0
        
        for idx, load in loads_df.iterrows():
            if self.get_y() > 250:
                self.add_page()
            
            connected = load['Rating (kW)'] * load['Quantity']
            p = tx_calc.calculate_p(load['Rating (kW)'], load['Quantity'], load['Diversity Factor'])
            phi = math.acos(load['Power Factor'])
            tan_phi = math.tan(phi)
            q = tx_calc.calculate_q(p, load['Power Factor'])
            s = tx_calc.calculate_s(p, q)
            
            self.set_font('Arial', 'B', 12)
            self.set_text_color(0, 51, 102)
            self.cell(0, 8, f'Load {idx+1}: {load["Load Description"]}', 0, 1)
            self.ln(1)
            
            self.set_font('Arial', '', 10)
            self.set_text_color(0, 0, 0)
            self.cell(0, 5, f'Step 1 - Connected Power: {load["Rating (kW)"]:.0f} kW × {load["Quantity"]} = {connected:.0f} kW', 0, 1)
            self.cell(0, 5, f'Step 2 - Demand Power (P): {connected:.0f} kW × {load["Diversity Factor"]} = {p:.1f} kW', 0, 1)
            self.cell(0, 5, f'Step 3 - Angle φ: acos({load["Power Factor"]}) = {math.degrees(phi):.1f}°', 0, 1)
            self.cell(0, 5, f'Step 4 - tan(φ): tan({math.degrees(phi):.1f}°) = {tan_phi:.3f}', 0, 1)
            self.cell(0, 5, f'Step 5 - Reactive Power (Q): {p:.1f} kW × {tan_phi:.3f} = {q:.1f} kVAR', 0, 1)
            self.set_font('Arial', 'B', 10)
            self.cell(0, 5, f'Step 6 - Apparent Power (S): √({p:.1f}² + {q:.1f}²) = {s:.1f} kVA', 0, 1)
            self.ln(3)
            
            total_p += p
            total_q += q
            
            self.set_draw_color(200, 200, 200)
            self.line(10, self.get_y(), 200, self.get_y())
            self.ln(3)
        
        st.session_state.total_p = total_p
        st.session_state.total_q = total_q
    
    def add_largest_equipment(self, loads_df, tx_calc, total_p, total_s):
        self.add_page()
        self.set_font('Arial', 'B', 14)
        self.set_text_color(0, 51, 102)
        self.cell(0, 10, '3. LARGEST EQUIPMENT ANALYSIS', 0, 1)
        self.ln(2)
        
        # Find largest equipment
        max_p = 0
        max_load = None
        max_idx = -1
        
        for idx, load in loads_df.iterrows():
            p_connected = load['Rating (kW)'] * load['Quantity']
            if p_connected > max_p:
                max_p = p_connected
                max_load = load
                max_idx = idx
        
        if max_load is not None:
            p_largest = tx_calc.calculate_p(max_load['Rating (kW)'], max_load['Quantity'], max_load['Diversity Factor'])
            q_largest = tx_calc.calculate_q(p_largest, max_load['Power Factor'])
            s_largest = tx_calc.calculate_s(p_largest, q_largest)
            
            self.set_font('Arial', 'B', 12)
            self.set_text_color(0, 51, 102)
            self.cell(0, 8, f'Largest Equipment: {max_load["Load Description"]}', 0, 1)
            self.ln(2)
            
            self.set_font('Arial', '', 10)
            self.set_text_color(0, 0, 0)
            self.cell(0, 6, f'Connected Power: {max_p:.0f} kW ({max_load["Rating (kW)"]:.0f} kW × {max_load["Quantity"]})', 0, 1)
            self.cell(0, 6, f'Demand Power (P): {p_largest:.1f} kW (after diversity factor {max_load["Diversity Factor"]})', 0, 1)
            self.cell(0, 6, f'Reactive Power (Q): {q_largest:.1f} kVAR (PF = {max_load["Power Factor"]})', 0, 1)
            self.set_font('Arial', 'B', 10)
            self.cell(0, 6, f'Apparent Power (S): {s_largest:.1f} kVA', 0, 1)
            self.ln(5)
            
            # Impact analysis
            self.set_font('Arial', 'B', 11)
            self.set_text_color(0, 51, 102)
            self.cell(0, 7, 'Impact on Total System:', 0, 1)
            self.set_font('Arial', '', 10)
            self.set_text_color(0, 0, 0)
            
            p_pct = (p_largest / total_p) * 100 if total_p > 0 else 0
            s_pct = (s_largest / total_s) * 100 if total_s > 0 else 0
            
            self.cell(0, 6, f'• Contributes {p_pct:.1f}% of total real power (P)', 0, 1)
            self.cell(0, 6, f'• Contributes {s_pct:.1f}% of total apparent power (S)', 0, 1)
            self.cell(0, 6, f'• Starting this motor would cause approx. {s_pct:.1f}% voltage dip', 0, 1)
            self.ln(5)
            
            # Contribution table
            self.set_font('Arial', 'B', 10)
            self.set_fill_color(240, 240, 240)
            self.cell(60, 7, 'Load', 1, 0, 'C', 1)
            self.cell(30, 7, 'P (kW)', 1, 0, 'C', 1)
            self.cell(30, 7, '% of P', 1, 0, 'C', 1)
            self.cell(30, 7, 'S (kVA)', 1, 0, 'C', 1)
            self.cell(30, 7, '% of S', 1, 1, 'C', 1)
            
            self.set_font('Arial', '', 9)
            fill = False
            for idx, load in loads_df.iterrows():
                p = tx_calc.calculate_p(load['Rating (kW)'], load['Quantity'], load['Diversity Factor'])
                q = tx_calc.calculate_q(p, load['Power Factor'])
                s = tx_calc.calculate_s(p, q)
                
                p_pct = (p / total_p) * 100 if total_p > 0 else 0
                s_pct = (s / total_s) * 100 if total_s > 0 else 0
                
                self.cell(60, 6, load['Load Description'][:15], 1, 0, 'L', fill)
                self.cell(30, 6, f"{p:.1f}", 1, 0, 'R', fill)
                self.cell(30, 6, f"{p_pct:.1f}%", 1, 0, 'R', fill)
                self.cell(30, 6, f"{s:.1f}", 1, 0, 'R', fill)
                self.cell(30, 6, f"{s_pct:.1f}%", 1, 1, 'R', fill)
                fill = not fill
    
    def add_transformer_selection(self, total_p, total_q, future_expansion, selected_kva, with_future, total_s):
        self.add_page()
        self.set_font('Arial', 'B', 14)
        self.set_text_color(0, 51, 102)
        self.cell(0, 10, '4. TRANSFORMER SELECTION [IEC 60076]', 0, 1)
        self.ln(2)
        
        self.set_font('Arial', '', 10)
        self.set_text_color(0, 0, 0)
        self.cell(0, 7, f'Total Real Power (P) = {total_p:.1f} kW', 0, 1)
        self.cell(0, 7, f'Total Reactive Power (Q) = {total_q:.1f} kVAR', 0, 1)
        self.set_font('Arial', 'B', 11)
        self.cell(0, 7, f'Total Apparent Power (S) = √({total_p:.1f}² + {total_q:.1f}²) = {total_s:.1f} kVA', 0, 1)
        self.ln(3)
        
        self.set_font('Arial', '', 10)
        self.cell(0, 7, f'Future Expansion: +{future_expansion}%', 0, 1)
        self.cell(0, 7, f'Required with future = {total_s:.1f} × 1.{future_expansion/100:.0f} = {with_future:.1f} kVA', 0, 1)
        self.ln(3)
        
        self.set_font('Arial', 'B', 12)
        self.set_text_color(0, 51, 102)
        self.cell(0, 8, 'Standard Ratings [IEC 60076]:', 0, 1)
        self.set_font('Arial', '', 10)
        self.set_text_color(0, 0, 0)
        
        ratings = [50, 100, 160, 200, 250, 315, 400, 500, 630, 800, 1000, 1250, 1600, 2000, 2500, 3150]
        ratings_str = ', '.join(str(r) for r in ratings[:10]) + '...'
        self.multi_cell(0, 5, ratings_str)
        self.ln(3)
        
        self.set_fill_color(0, 51, 102)
        self.set_text_color(255, 255, 255)
        self.set_font('Arial', 'B', 14)
        self.cell(0, 12, f'SELECTED TRANSFORMER: {selected_kva} kVA', 0, 1, 'C', 1)

class TransformerWordReport:
    def __init__(self):
        self.doc = Document()
        style = self.doc.styles['Normal']
        style.font.name = 'Arial'
        style.font.size = Pt(11)
        
        sections = self.doc.sections
        for section in sections:
            section.top_margin = Cm(2.5)
            section.bottom_margin = Cm(2.5)
            section.left_margin = Cm(2.5)
            section.right_margin = Cm(2.5)
    
    def add_title(self):
        title = self.doc.add_heading('TRANSFORMER SIZING REPORT', 0)
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        title.runs[0].font.size = Pt(20)
        title.runs[0].font.color.rgb = RGBColor(0, 51, 102)
        
        p = self.doc.add_paragraph()
        p.add_run(f'Date: {datetime.now().strftime("%Y-%m-%d %H:%M")}').italic = True
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        self.doc.add_paragraph()
    
    def add_load_analysis(self, loads_df):
        heading = self.doc.add_heading('1. LOAD ANALYSIS', level=1)
        heading.runs[0].font.color.rgb = RGBColor(0, 51, 102)
        
        table = self.doc.add_table(rows=1, cols=6)
        table.style = 'Light Grid Accent 1'
        
        headers = ['Load Description', 'Qty', 'Rating (kW)', 'Connected (kW)', 'Diversity', 'P (kW)']
        for i, header in enumerate(headers):
            table.rows[0].cells[i].text = header
            table.rows[0].cells[i].paragraphs[0].runs[0].bold = True
        
        total_p = 0
        for idx, load in loads_df.iterrows():
            connected = load['Rating (kW)'] * load['Quantity']
            p = connected * load['Diversity Factor']
            total_p += p
            
            row = table.add_row().cells
            row[0].text = load['Load Description']
            row[1].text = str(load['Quantity'])
            row[2].text = f"{load['Rating (kW)']:.0f}"
            row[3].text = f"{connected:.0f}"
            row[4].text = f"{load['Diversity Factor']:.1f}"
            row[5].text = f"{p:.1f}"
        
        p_row = table.add_row().cells
        p_row[0].text = 'TOTAL REAL POWER (P)'
        p_row[0].paragraphs[0].runs[0].bold = True
        p_row[1].text = ''
        p_row[2].text = ''
        p_row[3].text = ''
        p_row[4].text = ''
        p_row[5].text = f"{total_p:.1f} kW"
        p_row[5].paragraphs[0].runs[0].bold = True
        
        self.doc.add_paragraph()
        return total_p
    
    def add_step_by_step(self, loads_df, tx_calc):
        self.doc.add_page_break()
        heading = self.doc.add_heading('2. STEP-BY-STEP P, Q, S CALCULATIONS', level=1)
        heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
        heading.runs[0].font.color.rgb = RGBColor(0, 51, 102)
        
        total_p = 0
        total_q = 0
        
        for idx, load in loads_df.iterrows():
            self.doc.add_heading(f'Load {idx+1}: {load["Load Description"]}', level=2)
            
            connected = load['Rating (kW)'] * load['Quantity']
            p = tx_calc.calculate_p(load['Rating (kW)'], load['Quantity'], load['Diversity Factor'])
            phi = math.acos(load['Power Factor'])
            tan_phi = math.tan(phi)
            q = tx_calc.calculate_q(p, load['Power Factor'])
            s = tx_calc.calculate_s(p, q)
            
            self.doc.add_paragraph(f'Step 1 - Connected Power: {load["Rating (kW)"]:.0f} kW × {load["Quantity"]} = {connected:.0f} kW')
            self.doc.add_paragraph(f'Step 2 - Demand Power (P): {connected:.0f} kW × {load["Diversity Factor"]} = {p:.1f} kW')
            self.doc.add_paragraph(f'Step 3 - Angle φ: acos({load["Power Factor"]}) = {math.degrees(phi):.1f}°')
            self.doc.add_paragraph(f'Step 4 - tan(φ): tan({math.degrees(phi):.1f}°) = {tan_phi:.3f}')
            self.doc.add_paragraph(f'Step 5 - Reactive Power (Q): {p:.1f} kW × {tan_phi:.3f} = {q:.1f} kVAR')
            p_step = self.doc.add_paragraph()
            p_step.add_run('Step 6 - Apparent Power (S): ').bold = True
            p_step.add_run(f'√({p:.1f}² + {q:.1f}²) = {s:.1f} kVA')
            
            total_p += p
            total_q += q
            
            self.doc.add_paragraph('_' * 50)
        
        return total_p, total_q
    
    def add_largest_equipment(self, loads_df, tx_calc, total_p, total_s):
        self.doc.add_page_break()
        heading = self.doc.add_heading('3. LARGEST EQUIPMENT ANALYSIS', level=1)
        heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
        heading.runs[0].font.color.rgb = RGBColor(0, 51, 102)
        
        # Find largest equipment
        max_p = 0
        max_load = None
        max_idx = -1
        
        for idx, load in loads_df.iterrows():
            p_connected = load['Rating (kW)'] * load['Quantity']
            if p_connected > max_p:
                max_p = p_connected
                max_load = load
                max_idx = idx
        
        if max_load is not None:
            p_largest = tx_calc.calculate_p(max_load['Rating (kW)'], max_load['Quantity'], max_load['Diversity Factor'])
            q_largest = tx_calc.calculate_q(p_largest, max_load['Power Factor'])
            s_largest = tx_calc.calculate_s(p_largest, q_largest)
            
            self.doc.add_heading(f'Largest Equipment: {max_load["Load Description"]}', level=2)
            
            self.doc.add_paragraph(f'Connected Power: {max_p:.0f} kW ({max_load["Rating (kW)"]:.0f} kW × {max_load["Quantity"]})')
            self.doc.add_paragraph(f'Demand Power (P): {p_largest:.1f} kW (after diversity factor {max_load["Diversity Factor"]})')
            self.doc.add_paragraph(f'Reactive Power (Q): {q_largest:.1f} kVAR (PF = {max_load["Power Factor"]})')
            p = self.doc.add_paragraph()
            p.add_run('Apparent Power (S): ').bold = True
            p.add_run(f'{s_largest:.1f} kVA')
            
            self.doc.add_heading('Impact on Total System:', level=3)
            p_pct = (p_largest / total_p) * 100 if total_p > 0 else 0
            s_pct = (s_largest / total_s) * 100 if total_s > 0 else 0
            
            self.doc.add_paragraph(f'• Contributes {p_pct:.1f}% of total real power (P)')
            self.doc.add_paragraph(f'• Contributes {s_pct:.1f}% of total apparent power (S)')
            self.doc.add_paragraph(f'• Starting this motor would cause approx. {s_pct:.1f}% voltage dip')
            self.doc.add_paragraph()
            
            # Contribution table
            self.doc.add_heading('Load Contribution Analysis:', level=3)
            table = self.doc.add_table(rows=1, cols=5)
            table.style = 'Light Grid Accent 1'
            
            headers = ['Load', 'P (kW)', '% of P', 'S (kVA)', '% of S']
            for i, header in enumerate(headers):
                table.rows[0].cells[i].text = header
                table.rows[0].cells[i].paragraphs[0].runs[0].bold = True
            
            for idx, load in loads_df.iterrows():
                p = tx_calc.calculate_p(load['Rating (kW)'], load['Quantity'], load['Diversity Factor'])
                q = tx_calc.calculate_q(p, load['Power Factor'])
                s = tx_calc.calculate_s(p, q)
                
                p_pct = (p / total_p) * 100 if total_p > 0 else 0
                s_pct = (s / total_s) * 100 if total_s > 0 else 0
                
                row = table.add_row().cells
                row[0].text = load['Load Description']
                row[1].text = f"{p:.1f}"
                row[2].text = f"{p_pct:.1f}%"
                row[3].text = f"{s:.1f}"
                row[4].text = f"{s_pct:.1f}%"
    
    def add_transformer_selection(self, total_p, total_q, future_expansion, selected_kva, with_future, total_s):
        self.doc.add_page_break()
        heading = self.doc.add_heading('4. TRANSFORMER SELECTION [IEC 60076]', level=1)
        heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
        heading.runs[0].font.color.rgb = RGBColor(0, 51, 102)
        
        self.doc.add_paragraph(f'Total Real Power (P) = {total_p:.1f} kW')
        self.doc.add_paragraph(f'Total Reactive Power (Q) = {total_q:.1f} kVAR')
        p = self.doc.add_paragraph()
        p.add_run('Total Apparent Power (S) = ').bold = True
        p.add_run(f'√({total_p:.1f}² + {total_q:.1f}²) = {total_s:.1f} kVA')
        
        self.doc.add_paragraph()
        self.doc.add_paragraph(f'Future Expansion: +{future_expansion}%')
        self.doc.add_paragraph(f'Required with future = {total_s:.1f} × 1.{future_expansion/100:.0f} = {with_future:.1f} kVA')
        self.doc.add_paragraph()
        
        self.doc.add_heading('Standard Ratings [IEC 60076]:', level=3)
        ratings = [50, 100, 160, 200, 250, 315, 400, 500, 630, 800, 1000, 1250, 1600, 2000, 2500, 3150]
        ratings_str = ', '.join(str(r) for r in ratings)
        self.doc.add_paragraph(ratings_str)
        
        final_heading = self.doc.add_heading('', level=2)
        final_heading.add_run(f'SELECTED TRANSFORMER: {selected_kva} kVA').bold = True
        final_heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    def save(self, filename):
        self.doc.save(filename)

# ========== SIMPLIFIED TRANSFORMER CALCULATOR ==========
class SimpleTransformerCalculator:
    def __init__(self):
        # IEC 60076 Standard Ratings
        self.standard_ratings = [50, 100, 160, 200, 250, 315, 400, 500, 630, 800, 
                                  1000, 1250, 1600, 2000, 2500, 3150, 4000, 5000, 
                                  6300, 8000, 10000, 12500, 16000, 20000, 25000, 
                                  31500, 40000, 50000, 63000]
    
    def calculate_p(self, rating_kw, quantity, diversity):
        """Calculate Real Power P (kW)"""
        if pd.isna(rating_kw) or pd.isna(quantity) or pd.isna(diversity):
            return 0
        return rating_kw * quantity * diversity
    
    def calculate_q(self, p_kw, pf):
        """Calculate Reactive Power Q (kVAR)"""
        if pd.isna(p_kw) or pd.isna(pf) or pf >= 1.0:
            return 0
        phi = math.acos(pf)
        return p_kw * math.tan(phi)
    
    def calculate_s(self, p_kw, q_kvar):
        """Calculate Apparent Power S (kVA)"""
        if pd.isna(p_kw) or pd.isna(q_kvar):
            return 0
        return math.sqrt(p_kw**2 + q_kvar**2)
    
    def get_standard_rating(self, required_kva):
        """Get next higher standard rating from IEC 60076"""
        if pd.isna(required_kva) or required_kva <= 0:
            return 50
        for rating in self.standard_ratings:
            if rating >= required_kva:
                return rating
        return self.standard_ratings[-1]
    
    def find_largest_equipment(self, loads):
        """Find the largest equipment by connected power"""
        max_p = 0
        max_load = None
        max_idx = -1
        
        for idx, load in loads.iterrows():
            p_connected = load['Rating (kW)'] * load['Quantity']
            if p_connected > max_p:
                max_p = p_connected
                max_load = load
                max_idx = idx
        
        return max_idx, max_load, max_p

# ========== NEW: Excel Upload Only ==========
if 'uploaded_data' not in st.session_state:
    st.session_state.uploaded_data = None

# Cable sizing loads
if 'loads_df' not in st.session_state:
    st.session_state.loads_df = pd.DataFrame({
        'Load Name': ['Motor 1', 'Motor 2', 'Lighting', 'HVAC'],
        'Power (kW)': [75, 50, 25, 40],
        'Voltage (V)': [415, 415, 230, 415],
        'Phase': ['3-phase', '3-phase', '1-phase', '3-phase'],
        'Power Factor': [0.85, 0.85, 0.95, 0.80],
        'Length (m)': [50, 60, 30, 45]
    })

if 'cable_results_df' not in st.session_state:
    st.session_state.cable_results_df = pd.DataFrame()
if 'detailed_calcs' not in st.session_state:
    st.session_state.detailed_calcs = []
if 'derating_factors' not in st.session_state:
    st.session_state.derating_factors = None
if 'cb_results' not in st.session_state:
    st.session_state.cb_results = []
if 'cb_details' not in st.session_state:
    st.session_state.cb_details = []
if 'main_cb' not in st.session_state:
    st.session_state.main_cb = None
if 'calc_results' not in st.session_state:
    st.session_state.calc_results = {}
if 'calc_done' not in st.session_state:
    st.session_state.calc_done = False

# ========== SIDEBAR NAVIGATION ==========
with st.sidebar:
    st.markdown('<div class="sidebar-nav"><h2>⚡ CES-Electrical</h2></div>', unsafe_allow_html=True)
    
    if 'selected_calculator' not in st.session_state:
        st.session_state.selected_calculator = "📋 LOAD LIST"
    
    calculators = [
        "📋 LOAD LIST",
        "⚡ Lightning Protection",
        "🔌 Cable Sizing",
        "⚙️ Transformer Sizing",
        "⚡ Generator Sizing",
        "🌍 Earthing System Design"
    ]
    
    for calc in calculators:
        if st.button(calc, key=f"nav_{calc}", use_container_width=True):
            st.session_state.selected_calculator = calc
            st.rerun()

# ========== MAIN CONTENT ==========
st.title(st.session_state.selected_calculator)

# ========== TAB 1: LOAD LIST WITH EXCEL UPLOAD ONLY ==========
if st.session_state.selected_calculator == "📋 LOAD LIST":
    
    st.markdown('<div class="report-header">📋 LOAD LIST</div>', unsafe_allow_html=True)
    
    # ===== Excel Upload Section (No default data) =====
    st.markdown('<div class="upload-section">', unsafe_allow_html=True)
    st.markdown("### 📤 Upload Excel File")
    st.markdown("Upload your Excel file to view and edit the data")
    
    uploaded_file = st.file_uploader(
        "Choose an Excel file (.xlsx, .xls)", 
        type=['xlsx', 'xls'],
        help="Upload your Excel file to see the data here"
    )
    
    if uploaded_file is not None:
        try:
            # Read the uploaded Excel file
            df = pd.read_excel(uploaded_file)
            st.session_state.uploaded_data = df
            st.success(f"✅ Successfully loaded {len(df)} rows and {len(df.columns)} columns")
            
        except Exception as e:
            st.error(f"Error reading file: {e}")
            st.session_state.uploaded_data = None
    
    st.markdown('</div>', unsafe_allow_html=True)
    
    # ===== Display Uploaded Data =====
    if st.session_state.uploaded_data is not None:
        st.markdown("### 📋 Load Data")
        st.markdown("*Scroll horizontally to see all columns*")
        
        edited_df = st.data_editor(
            st.session_state.uploaded_data,
            num_rows="dynamic",  # Allows adding/removing rows
            use_container_width=True
        )
        
        # Update the data in session state
        st.session_state.uploaded_data = edited_df
        
        # ===== Download Options =====
        st.markdown("### 📥 Download Options")
        col1, col2 = st.columns(2)
        
        with col1:
            csv = edited_df.to_csv(index=False).encode('utf-8')
            st.download_button(
                label="📥 Download as CSV",
                data=csv,
                file_name=f"load_list_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
                mime="text/csv",
                use_container_width=True
            )
        
        with col2:
            # Create Excel file in memory
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                edited_df.to_excel(writer, index=False, sheet_name='Load List')
            excel_data = output.getvalue()
            
            st.download_button(
                label="📥 Download as Excel",
                data=excel_data,
                file_name=f"load_list_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        
        # Simple summary
        st.markdown("### 📊 Summary")
        col1, col2 = st.columns(2)
        with col1:
            st.metric("Total Rows", len(edited_df))
        with col2:
            st.metric("Total Columns", len(edited_df.columns))
    
    else:
        # Show message when no file is uploaded
        st.info("👆 Please upload an Excel file to view and edit data")

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
    
    st.markdown("""
    <div class="info-box">
        <h4>📌 Cable Sizing Calculator</h4>
        <p>Upload your load data in the LOAD LIST tab first, then import here.</p>
    </div>
    """, unsafe_allow_html=True)
    
    if st.session_state.uploaded_data is not None:
        st.info(f"✅ Load data available with {len(st.session_state.uploaded_data)} rows")
    else:
        st.warning("⚠️ Please upload load data in the LOAD LIST tab first")
    
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
        st.markdown("### 📋 Load Details")
        st.markdown("""
        <div class="info-box">
            <p>To add or modify loads, please go to the main <b>📋 LOAD LIST</b> tab and upload your Excel file.</p>
        </div>
        """, unsafe_allow_html=True)
        
        if st.button("📥 Import from LOAD LIST", use_container_width=True):
            if st.session_state.uploaded_data is not None:
                # Try to map common column names
                df = st.session_state.uploaded_data
                new_loads = []
                
                # Look for power column
                power_col = None
                for col in df.columns:
                    if 'power' in str(col).lower() or 'kw' in str(col).lower() or 'motor' in str(col).lower():
                        power_col = col
                        break
                
                # Look for voltage column
                voltage_col = None
                for col in df.columns:
                    if 'voltage' in str(col).lower() or 'v' in str(col).lower():
                        voltage_col = col
                        break
                
                # Create basic loads
                for idx, row in df.iterrows():
                    power = row[power_col] if power_col else 50
                    voltage = row[voltage_col] if voltage_col else 415
                    
                    new_loads.append({
                        'Load Name': f"Load {idx+1}",
                        'Power (kW)': float(power) if pd.notna(power) else 50,
                        'Voltage (V)': float(voltage) if pd.notna(voltage) else 415,
                        'Phase': '3-phase' if float(voltage) > 300 else '1-phase',
                        'Power Factor': 0.85,
                        'Length (m)': 50
                    })
                
                st.session_state.loads_df = pd.DataFrame(new_loads)
                st.success(f"✅ Imported {len(new_loads)} loads successfully!")
                st.rerun()
            else:
                st.error("No data in LOAD LIST. Please upload an Excel file first.")
        
        st.markdown("### Current Cable Sizing Loads")
        edited_df = st.data_editor(
            st.session_state.loads_df,
            num_rows="fixed",
            use_container_width=True,
            column_config={
                "Load Name": st.column_config.TextColumn("Load Name"),
                "Power (kW)": st.column_config.NumberColumn("Power (kW)", min_value=0.0, max_value=10000.0, step=1.0),
                "Voltage (V)": st.column_config.NumberColumn("Voltage (V)", min_value=0, max_value=11000, step=1),
                "Phase": st.column_config.SelectboxColumn("Phase", options=['1-phase', '3-phase', 'DC']),
                "Power Factor": st.column_config.NumberColumn("PF", min_value=0.5, max_value=1.0, step=0.05),
                "Length (m)": st.column_config.NumberColumn("Length (m)", min_value=1.0, max_value=5000.0, step=1.0)
            }
        )
        st.session_state.loads_df = edited_df
        
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
    
    st.markdown("""
    <div class="info-box">
        <h4>📌 Using loads from LOAD LIST</h4>
        <p>Upload your load data in the LOAD LIST tab first.</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Create main tabs
    tx_main_tabs = st.tabs([
        "📊 Load Analysis",
        "📈 Largest Equipment Analysis",
        "📥 Download Report"
    ])
    
    tx_calc = SimpleTransformerCalculator()
    
    # Convert uploaded loads to transformer format
    transformer_loads = pd.DataFrame()
    if st.session_state.uploaded_data is not None:
        # Try to find power column
        power_col = None
        for col in st.session_state.uploaded_data.columns:
            if 'power' in str(col).lower() or 'kw' in str(col).lower() or 'motor' in str(col).lower():
                power_col = col
                break
        
        if power_col:
            transformer_loads = pd.DataFrame({
                'Load Description': [f"Load {i+1}" for i in range(len(st.session_state.uploaded_data))],
                'Quantity': [1] * len(st.session_state.uploaded_data),
                'Rating (kW)': st.session_state.uploaded_data[power_col].values,
                'Power Factor': [0.85] * len(st.session_state.uploaded_data),
                'Diversity Factor': [0.8] * len(st.session_state.uploaded_data)
            })
    
    # TAB 1: LOAD ANALYSIS
    with tx_main_tabs[0]:
        load_sub_tabs = st.tabs([
            "📋 Step-by-Step P, Q, S",
            "📊 Summary Table"
        ])
        
        with load_sub_tabs[0]:
            st.markdown("### 📋 Step-by-Step Calculations for Each Load")
            
            if len(transformer_loads) > 0:
                total_p = 0
                total_q = 0
                
                for idx, load in transformer_loads.iterrows():
                    # Calculate connected power
                    connected = load['Rating (kW)'] * load['Quantity']
                    
                    # Step 1: Calculate P (Real Power) with diversity
                    p = tx_calc.calculate_p(load['Rating (kW)'], load['Quantity'], load['Diversity Factor'])
                    
                    # Step 2: Calculate tan(acos(PF))
                    phi = math.acos(load['Power Factor'])
                    tan_phi = math.tan(phi)
                    
                    # Step 3: Calculate Q (Reactive Power)
                    q = tx_calc.calculate_q(p, load['Power Factor'])
                    
                    # Step 4: Calculate S (Apparent Power)
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
            else:
                st.info("No load data available. Please upload an Excel file in LOAD LIST tab.")
        
        with load_sub_tabs[1]:
            st.markdown("### 📊 Load Summary Table")
            
            if len(transformer_loads) > 0:
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
            else:
                st.info("No load data available. Please upload an Excel file in LOAD LIST tab.")
    
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
        else:
            st.info("No load data available. Please upload an Excel file in LOAD LIST tab.")
    
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

# ========== OTHER CALCULATORS ==========
elif st.session_state.selected_calculator == "⚡ Generator Sizing":
    st.markdown('<div class="report-header">⚡ GENERATOR SIZING</div>', unsafe_allow_html=True)
    st.info("⚡ Coming soon!")

elif st.session_state.selected_calculator == "🌍 Earthing System Design":
    st.markdown('<div class="report-header">🌍 EARTHING SYSTEM DESIGN</div>', unsafe_allow_html=True)
    st.info("🌍 Coming soon!")

# Footer
st.markdown("---")
st.markdown(f"<div style='text-align: center; color: gray;'>🔌 CES-Electrical | Version 84.0 | {datetime.now().strftime('%Y-%m-%d %H:%M')}</div>", unsafe_allow_html=True)