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

st.set_page_config(page_title="Professional Engineering Tools", page_icon="⚡", layout="wide")

# ========== CUSTOM CSS ==========
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
        color: #000000 !important;
    }
    .calculation-detail {
        background-color: #F5F5F5;
        padding: 15px;
        border-radius: 8px;
        border: 1px solid #ddd;
        margin: 10px 0;
        font-family: 'Courier New', monospace;
        color: #000000 !important;
    }
    .info-box {
        background-color: #E7F3FF;
        color: #004085 !important;
        padding: 15px;
        border-radius: 8px;
        border-left: 5px solid #1E3A8A;
        margin: 10px 0;
    }
    .download-btn {
        display: inline-block;
        padding: 12px 24px;
        margin: 10px;
        color: white !important;
        text-decoration: none;
        border-radius: 5px;
        font-size: 16px;
        font-weight: bold;
        transition: all 0.3s;
        text-align: center;
    }
    .pdf-btn {
        background-color: #dc3545;
    }
    .word-btn {
        background-color: #1e3a8a;
    }
    .parameter-table {
        width: 100%;
        border-collapse: collapse;
        margin: 20px 0;
    }
    .parameter-table th {
        background-color: #1E3A8A;
        color: white !important;
        padding: 10px;
        text-align: center;
        font-weight: bold;
    }
    .parameter-table td {
        border: 1px solid #ddd;
        padding: 8px;
        text-align: left;
        color: #000000 !important;
    }
    .stDataFrame {
        color: #000000 !important;
    }
    .stDataFrame td {
        color: #000000 !important;
    }
    .param-table {
        width: 100%;
        border-collapse: collapse;
        margin: 15px 0;
        font-family: Arial, sans-serif;
    }
    .param-table th {
        background-color: #1E3A8A;
        color: white !important;
        padding: 12px;
        text-align: left;
    }
    .param-table td {
        padding: 10px;
        border: 1px solid #ddd;
        color: #000000 !important;
    }
    .param-table tr:nth-child(even) {
        background-color: #f9f9f9;
    }
    .cb-selection-reason {
        background-color: #fff3cd;
        border-left: 5px solid #ffc107;
        padding: 15px;
        margin: 10px 0;
        border-radius: 5px;
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

POLE_GUIDE = {
    '1-phase': {
        '2P': 'Phase + Neutral protection - Required for single-phase circuits',
        '1P': 'Phase only protection - Not recommended for final circuits'
    },
    '3-phase': {
        '3P': '3-Pole - For 3-wire systems (no neutral)',
        '4P': '4-Pole - For 4-wire systems with neutral protection'
    },
    'DC': {
        '2P': 'Both poles protection - Required for DC circuits'
    }
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
    'touching': {
        1: 1.00, 2: 0.80, 3: 0.70, 4: 0.65, 5: 0.60, 6: 0.57,
        7: 0.54, 8: 0.52, 9: 0.50, 10: 0.48, 11: 0.46, 12: 0.45,
        13: 0.44, 14: 0.43, 15: 0.42, 16: 0.41, 17: 0.40, 18: 0.39
    },
    'spaced': {
        1: 1.00, 2: 0.85, 3: 0.79, 4: 0.75, 5: 0.73, 6: 0.72,
        7: 0.71, 8: 0.70, 9: 0.70, 10: 0.70, 11: 0.70, 12: 0.70
    }
}

SOIL_RESISTIVITY_FACTORS = {
    0.7: 1.28, 0.8: 1.24, 0.9: 1.19, 1.0: 1.15,
    1.5: 1.00, 2.0: 0.89, 2.5: 0.81, 3.0: 0.75
}

DEPTH_FACTORS = {
    0.5: 1.04, 0.6: 1.02, 0.7: 1.01, 0.8: 1.00,
    0.9: 0.99, 1.0: 0.98, 1.25: 0.96, 1.5: 0.95,
    1.75: 0.94, 2.0: 0.93
}

CABLE_LAYING_FACTORS = {
    'air': 1.00, 'surface': 0.98, 'buried': 0.95, 'duct': 0.92
}

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
        
        # Collection Area
        self.doc.add_heading('1.1 Collection Area (Ad)', level=1)
        self.doc.add_paragraph('Formula: Ad = L x W + 2 x (3H) x (L + W) + pi x (3H)^2')
        self.doc.add_paragraph('Reference: IEC 62305-2 Annex A.2.1.1')
        p = self.doc.add_paragraph()
        p.add_run('Result: ').bold = True
        p.add_run(f'Ad = {results["ad"]:.2f} m²')
        
        # Near Strike Area
        self.doc.add_heading('1.2 Near Strike Collection Area (Am)', level=1)
        self.doc.add_paragraph('Formula: Am = 2 x 500 x (L + W) + pi x 500^2')
        self.doc.add_paragraph('Reference: IEC 62305-2 Annex A.3')
        p = self.doc.add_paragraph()
        p.add_run('Result: ').bold = True
        p.add_run(f'Am = {results["am"]:.2f} m²')
        
        # Environmental Factor
        self.doc.add_heading('1.3 Environmental Factor (CD)', level=1)
        self.doc.add_paragraph(f'Selected Environment: {inputs.get("environment", "Isolated")}')
        p = self.doc.add_paragraph()
        p.add_run('Result: ').bold = True
        p.add_run(f'CD = {inputs.get("cd", 1)}')
        
        # Lightning Density
        self.doc.add_heading('1.4 Lightning Ground Flash Density (NG)', level=1)
        self.doc.add_paragraph('Formula: NG = 0.1 x Td')
        p = self.doc.add_paragraph()
        p.add_run('Result: ').bold = True
        p.add_run(f'NG = {results.get("ng", 1)} flashes/km²/year')
        
        # Frequencies
        self.doc.add_heading('1.5 Lightning Frequencies', level=1)
        p = self.doc.add_paragraph()
        p.add_run('Nd (Direct): ').bold = True
        p.add_run(f'{results.get("nd", 0):.6f} events/year')
        p = self.doc.add_paragraph()
        p.add_run('Nm (Near): ').bold = True
        p.add_run(f'{results.get("nm", 0):.6f} events/year')
        
        # Protection Level
        self.doc.add_heading('1.6 Protection Level', level=1)
        p = self.doc.add_paragraph()
        p.add_run('Result: ').bold = True
        p.add_run(f'{results.get("lpl", "Class III")}')
        self.doc.add_paragraph(f'Rolling Sphere Radius: {results.get("sphere", 45)}m')
        
        # Summary Table
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
        
        # Collection Area
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
        
        # Near Strike Area
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
        
        # Environmental Factor
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
        
        # Lightning Density
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
        
        # Frequencies
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
        
        # Protection Level
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
        
        # Air Terminals
        self.set_font('Arial', 'B', 14)
        self.set_text_color(0, 51, 102)
        self.cell(0, 10, '1.7 Air Terminals Required', 0, 1)
        self.set_font('Arial', '', 11)
        self.set_text_color(0, 0, 0)
        self.cell(0, 7, 'Method: Rolling Sphere Method', 0, 1)
        self.set_font('Arial', 'B', 12)
        self.cell(0, 8, f'Result: {results.get("air_terminals", 4)} air terminals required', 0, 1)
        self.ln(10)
        
        # Summary
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
    
    def get_all_derating_factors(self, temp_c, insulation_temp=90, num_cables=1, grouping='touching',
                                 soil_resistivity=1.5, depth=0.8, laying='air'):
        k1 = TEMPERATURE_FACTORS[insulation_temp].get(temp_c, 1.0)
        k2 = GROUPING_FACTORS[grouping].get(min(num_cables, 18), 0.5)
        k3 = SOIL_RESISTIVITY_FACTORS.get(soil_resistivity, 1.0)
        k4 = DEPTH_FACTORS.get(depth, 1.0)
        k5 = CABLE_LAYING_FACTORS.get(laying, 1.0)
        
        total_k = k1 * k2 * k3 * k4 * k5
        
        factors = {
            'k1 (Temperature)': {'value': k1, 'reference': 'IEC 60502-2 Table B.10'},
            'k2 (Grouping)': {'value': k2, 'reference': f'IEC 60502-2 Table 4C1 ({grouping})'},
            'k3 (Soil Resistivity)': {'value': k3, 'reference': 'IEC 60502-2 Table B.14'},
            'k4 (Depth)': {'value': k4, 'reference': 'IEC 60502-2 Table B.12'},
            'k5 (Laying)': {'value': k5, 'reference': 'IEC 60502-2 Table B.5'},
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
        K = 143  # Copper constant for XLPE
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
### Main Circuit Breaker Calculation

**Step 1: Total Load Calculation**
- Total Connected Load: {total_power:.2f} kW
- System Voltage: {voltage} V, 3-phase
- Power Factor: {pf}

**Step 2: Load Current Calculation**
I = {total_power:.2f} × 1000 / (1.732 × {voltage} × {pf}) = {current:.2f} A

**Step 3: Circuit Breaker Sizing [IEC 60898/IEC 60947-2]**
- Design Factor: {design_factor}
- Required Rating = {current:.2f} × {design_factor} = {required:.2f} A
- Selected Standard Rating: {rating} A

**Step 4: Breaker Type Selection**
- Based on rating {rating} A → {breaker_type} ({standard})
- Application: {BREAKER_TYPES[breaker_type]['application']}

**Step 5: Pole Selection [IEC 60364-5-53]**
- System Type: {system_type}
- Selected Poles: {poles}
- Reason: {reason}

**Final Selection: {rating} A {breaker_type} {poles}**
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
    
    def add_detailed_calculations(self, detailed_calcs):
        self.add_page()
        self.set_font('Arial', 'B', 16)
        self.set_text_color(0, 51, 102)
        self.cell(0, 15, '5. DETAILED CABLE CALCULATIONS', 0, 1, 'C')
        self.ln(5)
        
        for i, calc in enumerate(detailed_calcs):
            if self.get_y() > 250:
                self.add_page()
            
            self.set_font('Arial', 'B', 12)
            self.set_text_color(0, 51, 102)
            self.cell(0, 8, f'Load {i+1}: {calc["load_name"]}', 0, 1)
            self.ln(2)
            
            self.set_font('Arial', 'B', 10)
            self.set_text_color(0, 0, 0)
            self.cell(0, 6, 'Step 1: Load Current Calculation', 0, 1)
            self.set_font('Arial', '', 9)
            self.cell(0, 5, f'I = {calc["power"]} x 1000 / (1.732 x {calc["voltage"]} x {calc["pf"]}) = {calc["current"]:.1f} A', 0, 1)
            self.ln(2)
            
            self.set_font('Arial', 'B', 10)
            self.cell(0, 6, 'Step 2: Derating Factors', 0, 1)
            self.set_font('Arial', '', 9)
            self.cell(0, 5, f'k1={calc["k1"]:.3f}, k2={calc["k2"]:.3f}, k3={calc["k3"]:.3f}, k4={calc["k4"]:.3f}, k5={calc["k5"]:.3f}', 0, 1)
            self.cell(0, 5, f'Total K = {calc["total_k"]:.3f}', 0, 1)
            self.ln(2)
            
            self.set_font('Arial', 'B', 10)
            self.cell(0, 6, 'Step 3: Cable Selection', 0, 1)
            self.set_font('Arial', '', 9)
            self.cell(0, 5, f'Selected: {calc["size"]} mm² {calc["cable_type"]}', 0, 1)
            self.cell(0, 5, f'Derated Ampacity: {calc["derated_amp"]:.1f} A', 0, 1)
            self.ln(2)
            
            self.set_font('Arial', 'B', 10)
            self.cell(0, 6, 'Step 4: Voltage Drop', 0, 1)
            self.set_font('Arial', '', 9)
            self.cell(0, 5, f'VD = {calc["vd_pct"]:.3f}% (Limit: 2.5%)', 0, 1)
            self.ln(2)
            
            self.set_font('Arial', 'B', 10)
            self.cell(0, 6, 'Step 5: Short Circuit', 0, 1)
            self.set_font('Arial', '', 9)
            self.cell(0, 5, f'Isc = {calc["sc"]:.2f} kA', 0, 1)
            self.ln(2)
            
            self.set_font('Arial', 'B', 10)
            self.cell(0, 6, f'Final Status: {calc["status"]}', 0, 1)
            self.ln(5)
            self.line(10, self.get_y(), 200, self.get_y())
            self.ln(5)
    
    def add_cb_results(self, cb_results, cb_details, main_cb):
        self.add_page()
        self.set_font('Arial', 'B', 16)
        self.set_text_color(0, 51, 102)
        self.cell(0, 15, '6. CIRCUIT BREAKER SIZING', 0, 1, 'C')
        self.ln(5)
        
        # Individual CB Table
        self.set_font('Arial', 'B', 12)
        self.set_text_color(0, 51, 102)
        self.cell(0, 10, '6.1 Individual Circuit Breakers', 0, 1)
        self.ln(2)
        
        self.set_font('Arial', 'B', 8)
        self.set_fill_color(240, 240, 240)
        headers = ['Load', 'Current', 'Required', 'Selected', 'Type', 'Poles', 'Standard']
        widths = [30, 20, 20, 20, 20, 15, 25]
        
        for i, header in enumerate(headers):
            self.cell(widths[i], 8, header, 1, 0, 'C', 1)
        self.ln()
        
        self.set_font('Arial', '', 8)
        fill = False
        for r in cb_results:
            data = [
                r['Load'][:15],
                f"{r['Current (A)']:.1f} A",
                f"{r['Required CB (A)']:.1f} A",
                f"{r['Selected CB (A)']} A",
                r['Breaker Type'],
                r['Poles'],
                r['Standard']
            ]
            for i, value in enumerate(data):
                self.cell(widths[i], 6, value, 1, 0, 'L', fill)
            self.ln()
            fill = not fill
        self.ln(10)
        
        # Main CB
        self.set_font('Arial', 'B', 12)
        self.set_text_color(0, 51, 102)
        self.cell(0, 10, '6.2 Main Circuit Breaker', 0, 1)
        self.ln(2)
        
        self.set_font('Arial', '', 10)
        self.set_text_color(0, 0, 0)
        self.multi_cell(0, 6, main_cb['detailed_reason'])
        self.ln(5)
        
        # Detailed CB Selection Reasons
        self.set_font('Arial', 'B', 12)
        self.set_text_color(0, 51, 102)
        self.cell(0, 10, '6.3 Detailed Selection Reasons', 0, 1)
        self.ln(2)
        
        for detail in cb_details:
            self.set_font('Arial', 'B', 10)
            self.set_text_color(0, 51, 102)
            self.cell(0, 6, f'Load: {detail["load_name"]}', 0, 1)
            self.set_font('Arial', '', 9)
            self.set_text_color(0, 0, 0)
            self.multi_cell(0, 5, f"""
• Load Type: {detail['phase_desc']}
• Load Current: {detail['current']:.2f} A
• Design Factor: {detail['design_factor']}
• Required Rating: {detail['required']:.2f} A
• Selected Standard Rating: {detail['selected']} A
• Breaker Type: {detail['breaker_type']} ({detail['standard']})
• Poles: {detail['poles']}
• Reason: {detail['pole_reason']}
• Manufacturer: {detail['manufacturer']} - {detail['series']}
            """)
            self.ln(3)

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
    
    def add_detailed_calculations(self, detailed_calcs):
        self.doc.add_page_break()
        heading = self.doc.add_heading('5. DETAILED CABLE CALCULATIONS', level=1)
        heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
        heading.runs[0].font.color.rgb = RGBColor(0, 51, 102)
        
        for i, calc in enumerate(detailed_calcs):
            self.doc.add_heading(f'Load {i+1}: {calc["load_name"]}', level=2)
            
            self.doc.add_heading('Step 1: Load Current', level=3)
            p = self.doc.add_paragraph()
            p.add_run('I = ').bold = True
            p.add_run(f'{calc["power"]} × 1000 / (1.732 × {calc["voltage"]} × {calc["pf"]}) = {calc["current"]:.1f} A')
            
            self.doc.add_heading('Step 2: Derating Factors', level=3)
            p = self.doc.add_paragraph()
            p.add_run(f'k1={calc["k1"]:.3f}, k2={calc["k2"]:.3f}, k3={calc["k3"]:.3f}, k4={calc["k4"]:.3f}, k5={calc["k5"]:.3f}').bold = False
            p = self.doc.add_paragraph()
            p.add_run('Total K = ').bold = True
            p.add_run(f'{calc["total_k"]:.3f}')
            
            self.doc.add_heading('Step 3: Cable Selection', level=3)
            p = self.doc.add_paragraph()
            p.add_run('Selected: ').bold = True
            p.add_run(f'{calc["size"]} mm² {calc["cable_type"]}')
            p = self.doc.add_paragraph()
            p.add_run('Derated Ampacity: ').bold = True
            p.add_run(f'{calc["derated_amp"]:.1f} A')
            
            self.doc.add_heading('Step 4: Voltage Drop', level=3)
            p = self.doc.add_paragraph()
            p.add_run('VD = ').bold = True
            p.add_run(f'{calc["vd_pct"]:.3f}% (Limit: 2.5%)')
            
            self.doc.add_heading('Step 5: Short Circuit', level=3)
            p = self.doc.add_paragraph()
            p.add_run('Isc = ').bold = True
            p.add_run(f'{calc["sc"]:.2f} kA')
            
            self.doc.add_heading('Final Status', level=3)
            p = self.doc.add_paragraph()
            p.add_run(f'{calc["status"]}').bold = True
            
            self.doc.add_paragraph('_' * 50)
            self.doc.add_paragraph()
    
    def add_cb_results(self, cb_results, cb_details, main_cb):
        self.doc.add_page_break()
        heading = self.doc.add_heading('6. CIRCUIT BREAKER SIZING', level=1)
        heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
        heading.runs[0].font.color.rgb = RGBColor(0, 51, 102)
        
        # Individual CB Table
        self.doc.add_heading('6.1 Individual Circuit Breakers', level=2)
        table = self.doc.add_table(rows=1, cols=7)
        table.style = 'Light Grid Accent 1'
        
        headers = ['Load', 'Current', 'Required', 'Selected', 'Type', 'Poles', 'Standard']
        for i, header in enumerate(headers):
            table.rows[0].cells[i].text = header
            table.rows[0].cells[i].paragraphs[0].runs[0].bold = True
        
        for r in cb_results:
            row = table.add_row().cells
            row[0].text = r['Load']
            row[1].text = f"{r['Current (A)']:.1f} A"
            row[2].text = f"{r['Required CB (A)']:.1f} A"
            row[3].text = f"{r['Selected CB (A)']} A"
            row[4].text = r['Breaker Type']
            row[5].text = r['Poles']
            row[6].text = r['Standard']
        
        self.doc.add_paragraph()
        
        # Main CB
        self.doc.add_heading('6.2 Main Circuit Breaker', level=2)
        for line in main_cb['detailed_reason'].split('\n'):
            if line.strip():
                self.doc.add_paragraph(line.strip())
        
        self.doc.add_paragraph()
        
        # Detailed Reasons
        self.doc.add_heading('6.3 Detailed Selection Reasons', level=2)
        for detail in cb_details:
            self.doc.add_heading(f'Load: {detail["load_name"]}', level=3)
            p = self.doc.add_paragraph()
            p.add_run(f'Load Type: {detail["phase_desc"]}').bold = True
            self.doc.add_paragraph(f'• Load Current: {detail["current"]:.2f} A')
            self.doc.add_paragraph(f'• Design Factor: {detail["design_factor"]}')
            self.doc.add_paragraph(f'• Required Rating: {detail["required"]:.2f} A')
            self.doc.add_paragraph(f'• Selected Standard Rating: {detail["selected"]} A')
            self.doc.add_paragraph(f'• Breaker Type: {detail["breaker_type"]} ({detail["standard"]})')
            self.doc.add_paragraph(f'• Poles: {detail["poles"]}')
            p = self.doc.add_paragraph()
            p.add_run('Reason: ').bold = True
            p.add_run(detail["pole_reason"])
            self.doc.add_paragraph(f'• Manufacturer: {detail["manufacturer"]} - {detail["series"]}')
            self.doc.add_paragraph()
    
    def save(self, filename):
        self.doc.save(filename)

# ========== SESSION STATE INITIALIZATION ==========
if 'calc_results' not in st.session_state:
    st.session_state.calc_results = {}
if 'calc_done' not in st.session_state:
    st.session_state.calc_done = False
if 'selected_calculator' not in st.session_state:
    st.session_state.selected_calculator = "⚡ Lightning Protection"
if 'loads_df' not in st.session_state:
    st.session_state.loads_df = pd.DataFrame({
        'Load Name': ['Load 1'],
        'Power (kW)': [5.0],
        'Voltage (V)': [400],
        'Phase': ['3-phase'],
        'Power Factor': [0.85],
        'Length (m)': [50]
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

# ========== SIDEBAR ==========
with st.sidebar:
    st.markdown("### ⚡ CES-Electrical Design Calculations")
    st.markdown("---")
    
    calculators = [
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
st.title(f"⚡ {st.session_state.selected_calculator} Calculator")

# ========== LIGHTNING PROTECTION ==========
if st.session_state.selected_calculator == "⚡ Lightning Protection":
    
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

# ========== CABLE SIZING CALCULATOR (WITH CB INSIDE) - FIXED SESSION STATE ==========
elif st.session_state.selected_calculator == "🔌 Cable Sizing":
    
    cable_tabs = st.tabs([
        "📥 Loads Input", 
        "📊 Derating Factors", 
        "⚡ Cable Selection",
        "🔧 Short Circuit",
        "⚡ Circuit Breakers",
        "📥 Download Report"
    ])
    
    # TAB 1: LOADS INPUT - ALL VARIABLES DEFINED HERE
    with cable_tabs[0]:
        st.markdown('<div class="report-header">CABLE SIZING - LOADS INPUT</div>', unsafe_allow_html=True)
        
        # Load Table Controls
        col1, col2 = st.columns([1, 1])
        with col1:
            if st.button("➕ Add Load", use_container_width=True):
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
            if st.button("🗑️ Delete Last Load", use_container_width=True):
                if len(st.session_state.loads_df) > 1:
                    st.session_state.loads_df = st.session_state.loads_df[:-1]
                    st.rerun()
                else:
                    st.warning("At least one row required")
        
        # Load Table Editor
        edited_df = st.data_editor(
            st.session_state.loads_df,
            num_rows="fixed",
            use_container_width=True,
            column_config={
                "Load Name": st.column_config.TextColumn("Load Name"),
                "Power (kW)": st.column_config.NumberColumn("Power (kW)", min_value=0.0, max_value=10000.0, step=0.1),
                "Voltage (V)": st.column_config.NumberColumn("Voltage (V)", min_value=0.0, max_value=33000.0, step=1.0),
                "Phase": st.column_config.SelectboxColumn("Phase", options=['1-phase', '3-phase', 'DC']),
                "Power Factor": st.column_config.NumberColumn("PF", min_value=0.5, max_value=1.0, step=0.05),
                "Length (m)": st.column_config.NumberColumn("Length (m)", min_value=1.0, max_value=5000.0, step=1.0)
            }
        )
        st.session_state.loads_df = edited_df
        
        st.markdown("### ⚙️ Installation Parameters")
        
        # PARAMETERS DEFINED HERE - INSIDE THE TAB
        col1, col2 = st.columns(2)
        with col1:
            cable_type = st.selectbox("Cable Type", ['armoured', 'unarmoured'], key="cable_type_select")
            ambient_temp = st.number_input("Ambient Temp (°C)", value=30.0, step=5.0, key="ambient_temp_input")
            num_cables = st.number_input("Cables in Group", value=3, min_value=1, max_value=18, key="num_cables_input")
            grouping = st.selectbox("Grouping", ['touching', 'spaced'], key="grouping_select")
        
        with col2:
            laying = st.selectbox("Laying Method", ['air', 'surface', 'buried', 'duct'], key="laying_select")
            soil_res = st.number_input("Soil Resistivity (K.m/W)", value=1.5, step=0.5, min_value=0.5, max_value=3.0, key="soil_res_input")
            depth = st.number_input("Burial Depth (m)", value=0.8, step=0.1, min_value=0.3, max_value=2.0, key="depth_input")
            system_type = st.selectbox("System Type", ['TN-S', 'TN-C', 'TN-C-S', 'TT'], key="system_type_select")
        
        # STORE IN SESSION STATE - HERE, AFTER VARIABLES ARE DEFINED
        st.session_state.cable_type = cable_type
        st.session_state.ambient_temp = ambient_temp
        st.session_state.num_cables = num_cables
        st.session_state.grouping = grouping
        st.session_state.laying = laying
        st.session_state.soil_res = soil_res
        st.session_state.depth = depth
        st.session_state.system_type = system_type
        
        # Display Current Settings
        st.markdown("### 📊 Current Installation Settings")
        st.markdown(f"""
        <table class="param-table">
            <tr><th colspan="2">Installation Parameters</th></tr>
            <tr><td>Cable Type</td><td>{cable_type} copper</td></tr>
            <tr><td>Ambient Temperature</td><td>{ambient_temp} °C</td></tr>
            <tr><td>Cables in Group</td><td>{num_cables}</td></tr>
            <tr><td>Grouping</td><td>{grouping}</td></tr>
            <tr><td>Laying Method</td><td>{laying}</td></tr>
            <tr><td>Soil Resistivity</td><td>{soil_res} K.m/W</td></tr>
            <tr><td>Burial Depth</td><td>{depth} m</td></tr>
            <tr><td>System Type</td><td>{system_type}</td></tr>
        </table>
        """, unsafe_allow_html=True)
        
        # CALCULATE BUTTON - USE VALUES FROM SESSION STATE
        if st.button("🔧 CALCULATE", type="primary", use_container_width=True):
            with st.spinner("Calculating..."):
                # Get values from session state
                cable_type = st.session_state.cable_type
                ambient_temp = st.session_state.ambient_temp
                num_cables = st.session_state.num_cables
                grouping = st.session_state.grouping
                laying = st.session_state.laying
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
                    
                    total_k, factors = cable_calc.get_all_derating_factors(
                        ambient_temp, 90, num_cables, grouping, soil_res, depth, laying
                    )
                    
                    st.session_state.derating_factors = factors
                    
                    found = False
                    for size, data in db.items():
                        if found:
                            break
                        derated = data['ampacity'] * total_k
                        if derated >= current:
                            vd_v, vd_pct = cable_calc.calculate_voltage_drop(
                                current, load['Length (m)'], data['R'], data['X'],
                                load['Power Factor'], load['Voltage (V)'], load['Phase']
                            )
                            
                            isc = cable_calc.calculate_short_circuit(size, 1.0)
                            
                            if load['Phase'] == '3-phase':
                                input_power = 1.732 * load['Voltage (V)'] * current / 1000
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
                                'k2': factors['k2 (Grouping)']['value'],
                                'k3': factors['k3 (Soil Resistivity)']['value'],
                                'k4': factors['k4 (Depth)']['value'],
                                'k5': factors['k5 (Laying)']['value'],
                                'total_k': total_k,
                                'status': 'PASS' if vd_pct <= 2.5 and derated >= current else 'FAIL'
                            })
                            found = True
                    
                    if not found:
                        st.warning(f"No cable found for {load['Load Name']}")
                
                st.session_state.cable_results_df = pd.DataFrame(cable_results)
                st.session_state.detailed_calcs = detailed_calcs
                
                # Calculate Circuit Breakers
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
        st.markdown('<div class="report-header">CABLE SELECTION RESULTS</div>', unsafe_allow_html=True)
        st.markdown("### ⚡ Voltage Drop Limit: **2.5%** [IEC 60364-5-52]")
        
        if not st.session_state.cable_results_df.empty:
            st.dataframe(st.session_state.cable_results_df, use_container_width=True, hide_index=True)
            st.markdown("### 📋 DETAILED CALCULATIONS")
            
            for calc in st.session_state.detailed_calcs:
                with st.expander(f"🔍 {calc['load_name']}"):
                    st.markdown(f"""
**Step 1: Load Current**  
I = {calc['power']} × 1000 / (1.732 × {calc['voltage']} × {calc['pf']}) = **{calc['current']:.1f} A**

**Step 2: Derating Factors**  
k1={calc['k1']:.3f}, k2={calc['k2']:.3f}, k3={calc['k3']:.3f}, k4={calc['k4']:.3f}, k5={calc['k5']:.3f}  
Total K = **{calc['total_k']:.3f}**

**Step 3: Cable Selection**  
Selected: {calc['size']} mm² {calc['cable_type']}  
Derated Ampacity = {calc['base_amp']} × {calc['total_k']:.3f} = **{calc['derated_amp']:.1f} A**

**Step 4: Voltage Drop**  
VD = **{calc['vd_pct']:.3f}%** (Limit: 2.5%)

**Step 5: Short Circuit**  
Isc = **{calc['sc']:.2f} kA**

**Step 6: Efficiency**  
Efficiency = **{calc['efficiency']:.1f}%**

**Final Status: {'✅ PASS' if calc['status'] == 'PASS' else '❌ FAIL'}**
""")
        else:
            st.info("👈 Calculate loads first")
    
    # TAB 4: SHORT CIRCUIT
    with cable_tabs[3]:
        st.markdown('<div class="report-header">SHORT CIRCUIT CALCULATIONS</div>', unsafe_allow_html=True)
        st.markdown("""
        **Reference:** IEC 60949  
        **Formula:** Isc = K × S / √t, K=143 for Copper XLPE
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
        st.markdown('<div class="report-header">CIRCUIT BREAKER SIZING</div>', unsafe_allow_html=True)
        
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
            # Individual Breakers
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
            
            # Main Breaker
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
            
            # Detailed Calculations
            st.markdown("### 📋 Detailed Selection Calculations")
            
            with st.expander("Main Circuit Breaker Calculation", expanded=True):
                st.markdown(main['detailed_reason'])
            
            st.markdown("### 📋 Individual Breaker Selection Reasons")
            for detail in st.session_state.cb_details:
                with st.expander(f"Load: {detail['load_name']}"):
                    st.markdown(f"""
**Step 1: Load Analysis**
- Load Type: {detail['phase_desc']}
- Load Current: {detail['current']:.2f} A

**Step 2: Rating Calculation**
- Design Factor: {detail['design_factor']} (IEC 60364)
- Required Rating = {detail['current']:.2f} × {detail['design_factor']} = {detail['required']:.2f} A
- Selected Standard Rating: {detail['selected']} A

**Step 3: Breaker Type Selection**
- Based on rating {detail['selected']} A → {detail['breaker_type']} ({detail['standard']})
- Application: {BREAKER_TYPES[detail['breaker_type']]['application']}

**Step 4: Pole Selection [IEC 60364-5-53]**
- System Type: {st.session_state.system_type}
- Selected Poles: {detail['poles']}
- Reason: {detail['pole_reason']}

**Step 5: Manufacturer Selection**
- Manufacturer: {detail['manufacturer']}
- Series: {detail['series']}

**Final Selection: {detail['selected']} A {detail['breaker_type']} {detail['poles']}**
""")
        else:
            st.info("👈 Calculate cable sizes first to see circuit breaker results")
    
    # TAB 6: DOWNLOAD REPORT
    with cable_tabs[5]:
        st.markdown('<div class="report-header">DOWNLOAD REPORT</div>', unsafe_allow_html=True)
        
        if not st.session_state.cable_results_df.empty and st.session_state.cb_results:
            col1, col2 = st.columns(2)
            
            with col1:
                if st.button("📥 Generate PDF Report", key="cable_pdf", use_container_width=True):
                    with st.spinner("Generating PDF with cable and CB details..."):
                        pdf = CablePDFReport()
                        pdf.add_title()
                        
                        params = {
                            'Cable Type': f'{st.session_state.cable_type} copper',
                            'Ambient Temperature': f'{st.session_state.ambient_temp}°C',
                            'Cables in Group': str(st.session_state.num_cables),
                            'Grouping': st.session_state.grouping,
                            'Laying Method': st.session_state.laying,
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
                            pdf.add_detailed_calculations(st.session_state.detailed_calcs)
                        
                        if st.session_state.cb_results and st.session_state.main_cb:
                            pdf.add_cb_results(st.session_state.cb_results, st.session_state.cb_details, st.session_state.main_cb)
                        
                        pdf_output = pdf.output(dest='S')
                        b64 = base64.b64encode(pdf_output).decode()
                        filename = f"Cable_CB_Report_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf"
                        st.markdown(f'<a href="data:application/pdf;base64,{b64}" download="{filename}" class="download-btn pdf-btn">📥 Download PDF</a>', unsafe_allow_html=True)
                        st.success("✅ PDF generated with cable and CB details!")
            
            with col2:
                if st.button("📥 Generate Word Report", key="cable_word", use_container_width=True):
                    with st.spinner("Generating Word with cable and CB details..."):
                        word = CableWordReport()
                        word.add_title()
                        
                        params = {
                            'Cable Type': f'{st.session_state.cable_type} copper',
                            'Ambient Temperature': f'{st.session_state.ambient_temp}°C',
                            'Cables in Group': str(st.session_state.num_cables),
                            'Grouping': st.session_state.grouping,
                            'Laying Method': st.session_state.laying,
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
                            word.add_detailed_calculations(st.session_state.detailed_calcs)
                        
                        if st.session_state.cb_results and st.session_state.main_cb:
                            word.add_cb_results(st.session_state.cb_results, st.session_state.cb_details, st.session_state.main_cb)
                        
                        word_path = "temp_cable_report.docx"
                        word.save(word_path)
                        with open(word_path, "rb") as f:
                            word_bytes = f.read()
                        b64 = base64.b64encode(word_bytes).decode()
                        filename = f"Cable_CB_Report_{datetime.now().strftime('%Y%m%d_%H%M')}.docx"
                        os.remove(word_path)
                        st.markdown(f'<a href="data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,{b64}" download="{filename}" class="download-btn word-btn">📥 Download Word</a>', unsafe_allow_html=True)
                        st.success("✅ Word generated with cable and CB details!")
        else:
            st.info("👈 Calculate cable sizes first to generate report")

# ========== OTHER CALCULATORS ==========
elif st.session_state.selected_calculator == "⚙️ Transformer Sizing":
    st.markdown('<div class="report-header">TRANSFORMER SIZING</div>', unsafe_allow_html=True)
    st.info("⚙️ Coming soon!")

elif st.session_state.selected_calculator == "⚡ Generator Sizing":
    st.markdown('<div class="report-header">GENERATOR SIZING</div>', unsafe_allow_html=True)
    st.info("⚡ Coming soon!")

elif st.session_state.selected_calculator == "🌍 Earthing System Design":
    st.markdown('<div class="report-header">EARTHING SYSTEM DESIGN</div>', unsafe_allow_html=True)
    st.info("🌍 Coming soon!")

# Footer
st.markdown("---")
st.markdown(f"<div style='text-align: center; color: gray;'>⚡ CES-Electrical | Version 53.0 | {datetime.now().strftime('%Y-%m-%d %H:%M')}</div>", unsafe_allow_html=True)