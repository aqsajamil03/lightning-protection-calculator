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

# ========== CUSTOM CSS WITH FIXED TEXT COLORS ==========
st.markdown("""
<style>
    /* HEADER STYLES */
    .report-header {
        background-color: #1E3A8A;
        color: white;
        padding: 20px;
        border-radius: 10px;
        text-align: center;
        margin-bottom: 20px;
        font-size: 24px;
    }
    
    /* BOX STYLES - ALL TEXT BLACK */
    .formula-box {
        background-color: #F3F4F6;
        padding: 15px;
        border-radius: 8px;
        border-left: 5px solid #1E3A8A;
        margin: 10px 0;
        font-family: 'Courier New', monospace;
        color: #000000 !important;
    }
    .formula-box * {
        color: #000000 !important;
    }
    
    .reference-box {
        background-color: #E8F5E9;
        padding: 10px;
        border-radius: 5px;
        border-left: 3px solid #4CAF50;
        margin: 5px 0;
        font-size: 0.9em;
        color: #000000 !important;
    }
    .reference-box * {
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
    .calculation-detail * {
        color: #000000 !important;
    }
    .calculation-detail h3, 
    .calculation-detail h4, 
    .calculation-detail p, 
    .calculation-detail li,
    .calculation-detail table,
    .calculation-detail td,
    .calculation-detail th {
        color: #000000 !important;
    }
    
    .param-highlight {
        background-color: #FFE5B4;
        padding: 2px 5px;
        border-radius: 3px;
        font-weight: bold;
        color: #000000 !important;
    }
    
    /* STATUS BOXES */
    .success-box {
        background-color: #D4EDDA;
        color: #155724 !important;
        padding: 15px;
        border-radius: 8px;
        border-left: 5px solid #28A745;
        margin: 10px 0;
    }
    .success-box * {
        color: #155724 !important;
    }
    
    .warning-box {
        background-color: #FFF3CD;
        color: #856404 !important;
        padding: 15px;
        border-radius: 8px;
        border-left: 5px solid #FFC107;
        margin: 10px 0;
    }
    .warning-box * {
        color: #856404 !important;
    }
    
    .info-box {
        background-color: #E7F3FF;
        color: #004085 !important;
        padding: 15px;
        border-radius: 8px;
        border-left: 5px solid #1E3A8A;
        margin: 10px 0;
    }
    .info-box * {
        color: #004085 !important;
    }
    
    /* DOWNLOAD BUTTONS */
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
    .download-btn:hover {
        transform: scale(1.05);
        color: white !important;
    }
    .pdf-btn {
        background-color: #dc3545;
    }
    .word-btn {
        background-color: #1e3a8a;
    }
    
    /* TABS STYLES */
    .stTabs [data-baseweb="tab-list"] button [data-testid="stMarkdownContainer"] p {
        font-size: 24px !important;
        font-weight: 700 !important;
        color: #000000 !important;
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
    
    /* PARAMETER TABLE - FIXED TEXT COLORS */
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
    .parameter-table tr:nth-child(even) {
        background-color: #f2f2f2;
    }
    .parameter-table tr:nth-child(even) td {
        color: #000000 !important;
        background-color: #f2f2f2;
    }
    .parameter-table tr:nth-child(odd) {
        background-color: white;
    }
    .parameter-table tr:nth-child(odd) td {
        color: #000000 !important;
        background-color: white;
    }
    
    /* STREAMILT DATAFRAME FIXES */
    .stDataFrame {
        color: #000000 !important;
    }
    .stDataFrame table {
        color: #000000 !important;
    }
    .stDataFrame td {
        color: #000000 !important;
        background-color: white !important;
    }
    .stDataFrame th {
        color: white !important;
        background-color: #1E3A8A !important;
    }
    .stDataFrame tr:nth-child(even) td {
        background-color: #f9f9f9 !important;
        color: #000000 !important;
    }
    .stDataFrame tr:nth-child(odd) td {
        background-color: white !important;
        color: #000000 !important;
    }
    
    /* EXPANDER CONTENT FIX */
    .streamlit-expanderContent {
        color: #000000 !important;
        background-color: white;
    }
    .streamlit-expanderContent p, 
    .streamlit-expanderContent span, 
    .streamlit-expanderContent div,
    .streamlit-expanderContent li,
    .streamlit-expanderContent table,
    .streamlit-expanderContent td,
    .streamlit-expanderContent th {
        color: #000000 !important;
    }
    
    /* METRIC CARDS */
    [data-testid="stMetricValue"] {
        color: #000000 !important;
    }
    [data-testid="stMetricLabel"] {
        color: #333333 !important;
    }
    
    /* MAIN CONTENT TEXT */
    .main .block-container {
        color: #000000 !important;
    }
    .main .block-container p,
    .main .block-container span,
    .main .block-container div,
    .main .block-container li {
        color: #000000 !important;
    }
    
    /* FORCE ALL TABLE TEXT TO BLACK */
    table, tr, td, th, tbody, thead {
        color: #000000 !important;
    }
    td p, th p {
        color: #000000 !important;
    }
    
    /* SPECIFIC FIX FOR CABLE RESULTS TABLE */
    div[data-testid="stHorizontalBlock"] table,
    div[data-testid="stDataFrame"] table,
    .stDataFrame table,
    [data-testid="column"] table {
        color: #000000 !important;
    }
    div[data-testid="stHorizontalBlock"] table td,
    div[data-testid="stDataFrame"] table td,
    .stDataFrame table td,
    [data-testid="column"] table td {
        color: #000000 !important;
        background-color: white;
    }
    div[data-testid="stHorizontalBlock"] table tr:nth-child(even) td,
    div[data-testid="stDataFrame"] table tr:nth-child(even) td,
    .stDataFrame table tr:nth-child(even) td {
        background-color: #f9f9f9 !important;
        color: #000000 !important;
    }
    
    /* MARKDOWN TEXT FIX */
    .stMarkdown {
        color: #000000 !important;
    }
    .stMarkdown p,
    .stMarkdown h1,
    .stMarkdown h2,
    .stMarkdown h3,
    .stMarkdown h4,
    .stMarkdown li {
        color: #000000 !important;
    }
</style>
""", unsafe_allow_html=True)

# ========== CIRCUIT BREAKER DATA AND CALCULATIONS ==========
# Circuit Breaker Standard Ratings (IEC 60898 / IEC 60947-2)
CB_RATINGS = [6, 10, 16, 20, 25, 32, 40, 50, 63, 80, 100, 125, 160, 200, 250, 315, 400, 500, 630, 800, 1000, 1250, 1600]

# Breaker Types and Pole Selection (IEC 60947-2)
BREAKER_TYPES = {
    'MCB': {'min': 6, 'max': 125, 'standard': 'IEC 60898', 'application': 'Miniature Circuit Breaker - For final circuits'},
    'MCCB': {'min': 125, 'max': 1600, 'standard': 'IEC 60947-2', 'application': 'Moulded Case Circuit Breaker - For distribution'},
    'ACB': {'min': 1600, 'max': 6300, 'standard': 'IEC 60947-2', 'application': 'Air Circuit Breaker - For main incomers'}
}

# Pole Selection Guide (IEC 60364-5-53)
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

# ========== CABLE DATABASE (COPPER ONLY - XLPE Insulated) ==========
# LV Cables - 0.6/1kV, XLPE Insulated, Copper Conductors
LV_CABLE_DATA = {
    'unarmoured': {  # PVC Sheathed, Unarmoured
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
    'armoured': {  # SWA - Steel Wire Armoured
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

# MV Cables - 3.6/6kV to 12/20kV, XLPE Insulated, Copper Conductors
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

# ========== COMPLETE DERATING FACTORS TABLES (IEC 60502-2) ==========
TEMPERATURE_FACTORS = {
    90: {  # XLPE Insulation
        20: 1.07, 25: 1.04, 30: 1.00, 35: 0.96, 40: 0.91, 
        45: 0.87, 50: 0.82, 55: 0.76, 60: 0.71, 65: 0.65, 70: 0.58
    }
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
    'air': 1.00,  # Cables in air
    'surface': 0.98,  # Cables on surface
    'buried': 0.95,  # Direct buried
    'duct': 0.92  # In ducts
}

# ========== LIGHTNING PROTECTION REPORT CLASSES ==========
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
    
    def get_all_derating_factors(self, temp_c, insulation_temp=90, num_cables=1, grouping='touching',
                                 soil_resistivity=1.5, depth=0.8, laying='air'):
        """Calculate ALL derating factors with IEC references"""
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
    
    def calculate_short_circuit(self, size_mm2, duration_s=1.0, material='copper'):
        """Calculate short circuit current capacity
        Reference: IEC 60949, IEC 60364-5-54"""
        K = 143  # Copper constant for XLPE (IEC 60949)
        Isc = K * size_mm2 / math.sqrt(duration_s)
        return Isc
    
    def get_cable_type(self, voltage_v):
        """Determine cable type based on voltage"""
        if voltage_v <= 1000:
            return 'LV (0.6/1kV)', LV_CABLE_DATA
        else:
            return 'MV (3.6/6kV - 12/20kV)', MV_CABLE_DATA

# ========== CIRCUIT BREAKER CALCULATOR CLASS ==========
class CircuitBreakerCalculator:
    def __init__(self):
        pass
    
    def get_standard_rating(self, current, design_factor=1.25):
        """Get next higher standard CB rating (IEC 60898 / IEC 60947-2)"""
        required = current * design_factor
        for rating in CB_RATINGS:
            if rating >= required:
                return rating, required
        return CB_RATINGS[-1], required
    
    def get_breaker_type(self, rating):
        """Determine breaker type based on rating (IEC 60898 / IEC 60947-2)"""
        if rating <= 125:
            return 'MCB', 'IEC 60898'
        elif rating <= 1600:
            return 'MCCB', 'IEC 60947-2'
        else:
            return 'ACB', 'IEC 60947-2'
    
    def select_poles(self, phase, system_type='TN-S'):
        """Select number of poles based on system type
        Reference: IEC 60364-5-53, Table 53A"""
        
        if phase == '1-phase':
            if system_type in ['TN-S', 'TN-C-S', 'TT']:
                return '2P', 'Phase + Neutral protection - Required for TN/TT systems'
            else:
                return '1P', 'Phase only - For IT systems (not recommended)'
        
        elif phase == '3-phase':
            if system_type == 'TN-S':
                return '4P', '4-Pole - For TN-S systems with separate neutral'
            elif system_type == 'TN-C':
                return '3P', '3-Pole - For TN-C systems (PEN conductor)'
            else:
                return '3P', '3-Pole - Standard for 3-wire systems'
        
        else:  # DC
            return '2P', '2-Pole - Required for DC circuits (IEC 60947-2)'
    
    def calculate_cb_size(self, loads_df, design_factor=1.25, manufacturer='Schneider Electric', system_type='TN-S'):
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
            breaker_type, standard = self.get_breaker_type(rating)
            poles, reason = self.select_poles(load['Phase'], system_type)
            
            # Get manufacturer series
            series = MANUFACTURERS[manufacturer][breaker_type]
            
            results.append({
                'Load': load['Load Name'],
                'Power (kW)': load['Power (kW)'],
                'Current (A)': current,
                'Required CB (A)': required,
                'Selected CB (A)': rating,
                'Breaker Type': breaker_type,
                'Standard': standard,
                'Poles': poles,
                'Pole Selection Reason': reason,
                'Manufacturer': manufacturer,
                'Series': series
            })
        
        return results
    
    def calculate_main_cb(self, loads_df, voltage=400, pf=0.8, design_factor=1.25):
        """Calculate main circuit breaker"""
        total_power = loads_df['Power (kW)'].sum()
        current = total_power * 1000 / (1.732 * voltage * pf)
        rating, required = self.get_standard_rating(current, design_factor)
        breaker_type, standard = self.get_breaker_type(rating)
        poles, reason = self.select_poles('3-phase', 'TN-S')
        
        return {
            'total_power': total_power,
            'current': current,
            'required_cb': required,
            'selected_cb': rating,
            'breaker_type': breaker_type,
            'poles': poles,
            'standard': standard
        }

# ========== PDF Report Classes ==========
class CablePDFReport(FPDF):
    def __init__(self):
        super().__init__()
        self.set_auto_page_break(auto=True, margin=15)
    
    def header(self):
        self.set_font('Arial', 'B', 12)
        self.cell(0, 10, 'CABLE SIZING CALCULATION REPORT', 0, 1, 'C')
        self.ln(5)
    
    def footer(self):
        self.set_y(-15)
        self.set_font('Arial', 'I', 8)
        self.cell(0, 10, f'Page {self.page_no()}', 0, 0, 'C')
    
    def add_title(self):
        self.add_page()
        self.set_font('Arial', 'B', 16)
        self.set_text_color(0, 51, 102)
        self.cell(0, 15, 'CABLE SIZING & CIRCUIT BREAKER REPORT', 0, 1, 'C')
        self.set_font('Arial', '', 10)
        self.cell(0, 8, f'Date: {datetime.now().strftime("%Y-%m-%d %H:%M")}', 0, 1, 'R')
        self.ln(10)
    
    def add_installation_parameters(self, params):
        self.set_font('Arial', 'B', 12)
        self.cell(0, 8, '1. INSTALLATION PARAMETERS', 0, 1)
        self.ln(4)
        self.set_font('Arial', '', 10)
        for key, value in params.items():
            self.cell(0, 6, f"{key}: {value}", 0, 1)
        self.ln(5)
    
    def add_derating_factors(self, factors):
        self.set_font('Arial', 'B', 12)
        self.cell(0, 8, '2. DERATING FACTORS (IEC 60502-2)', 0, 1)
        self.ln(4)
        self.set_font('Arial', '', 10)
        for key, data in factors.items():
            if key != 'total':
                self.cell(0, 6, f"{key}: {data['value']:.3f} ({data['reference']})", 0, 1)
        self.set_font('Arial', 'B', 10)
        self.cell(0, 6, f"Total Derating Factor (K): {factors['total']:.3f}", 0, 1)
        self.ln(5)
    
    def add_load_details(self, loads_df):
        self.set_font('Arial', 'B', 12)
        self.cell(0, 8, '3. LOAD DETAILS', 0, 1)
        self.ln(4)
        
        self.set_font('Arial', 'B', 9)
        self.cell(30, 6, 'Load Name', 1, 0, 'C')
        self.cell(20, 6, 'Power', 1, 0, 'C')
        self.cell(20, 6, 'Voltage', 1, 0, 'C')
        self.cell(20, 6, 'Phase', 1, 0, 'C')
        self.cell(15, 6, 'PF', 1, 0, 'C')
        self.cell(20, 6, 'Length', 1, 1, 'C')
        
        self.set_font('Arial', '', 8)
        for idx, load in loads_df.iterrows():
            self.cell(30, 5, load['Load Name'][:15], 1, 0, 'L')
            self.cell(20, 5, f"{load['Power (kW)']:.1f}", 1, 0, 'R')
            self.cell(20, 5, f"{load['Voltage (V)']:.0f}", 1, 0, 'R')
            self.cell(20, 5, load['Phase'], 1, 0, 'C')
            self.cell(15, 5, f"{load['Power Factor']:.2f}", 1, 0, 'R')
            self.cell(20, 5, f"{load['Length (m)']:.0f}", 1, 1, 'R')
        self.ln(10)
    
    def add_cable_results(self, cable_df):
        self.set_font('Arial', 'B', 12)
        self.cell(0, 8, '4. CABLE SIZING RESULTS', 0, 1)
        self.ln(4)
        
        self.set_font('Arial', 'B', 8)
        self.cell(25, 5, 'Load', 1, 0, 'C')
        self.cell(15, 5, 'Size', 1, 0, 'C')
        self.cell(15, 5, 'Type', 1, 0, 'C')
        self.cell(15, 5, 'Base A', 1, 0, 'C')
        self.cell(15, 5, 'Derated', 1, 0, 'C')
        self.cell(15, 5, 'VD %', 1, 0, 'C')
        self.cell(15, 5, 'SC kA', 1, 0, 'C')
        self.cell(15, 5, 'Efficiency', 1, 0, 'C')
        self.cell(15, 5, 'Status', 1, 1, 'C')
        
        self.set_font('Arial', '', 7)
        for idx, row in cable_df.iterrows():
            self.cell(25, 4, row['Load Name'][:12], 1, 0, 'L')
            self.cell(15, 4, str(row['Size (mm²)']), 1, 0, 'C')
            self.cell(15, 4, 'Cu', 1, 0, 'C')
            self.cell(15, 4, str(row['Base Ampacity (A)']), 1, 0, 'R')
            self.cell(15, 4, str(row['Derated Ampacity (A)']).replace(' A', ''), 1, 0, 'R')
            self.cell(15, 4, str(row['Voltage Drop (%)']).replace('%', ''), 1, 0, 'R')
            self.cell(15, 4, str(row['Short Circuit (kA)']).replace(' kA', ''), 1, 0, 'R')
            self.cell(15, 4, str(row['Efficiency (%)']).replace('%', ''), 1, 0, 'R')
            self.cell(15, 4, row['Status'], 1, 1, 'C')

class CableWordReport:
    def __init__(self):
        self.doc = Document()
    
    def add_title(self):
        self.doc.add_heading('CABLE SIZING & CIRCUIT BREAKER REPORT', 0)
        self.doc.add_paragraph(f'Date: {datetime.now().strftime("%Y-%m-%d %H:%M")}')
        self.doc.add_paragraph()
    
    def add_installation_parameters(self, params):
        self.doc.add_heading('1. INSTALLATION PARAMETERS', level=1)
        for key, value in params.items():
            self.doc.add_paragraph(f"{key}: {value}")
        self.doc.add_paragraph()
    
    def add_derating_factors(self, factors):
        self.doc.add_heading('2. DERATING FACTORS (IEC 60502-2)', level=1)
        for key, data in factors.items():
            if key != 'total':
                self.doc.add_paragraph(f"{key}: {data['value']:.3f} ({data['reference']})")
        p = self.doc.add_paragraph()
        p.add_run(f"Total Derating Factor (K): {factors['total']:.3f}").bold = True
        self.doc.add_paragraph()
    
    def add_load_details(self, loads_df):
        self.doc.add_heading('3. LOAD DETAILS', level=1)
        table = self.doc.add_table(rows=1, cols=6)
        table.style = 'Light Grid Accent 1'
        hdr_cells = table.rows[0].cells
        hdr_cells[0].text = 'Load Name'
        hdr_cells[1].text = 'Power (kW)'
        hdr_cells[2].text = 'Voltage (V)'
        hdr_cells[3].text = 'Phase'
        hdr_cells[4].text = 'PF'
        hdr_cells[5].text = 'Length (m)'
        
        for idx, load in loads_df.iterrows():
            row_cells = table.add_row().cells
            row_cells[0].text = load['Load Name']
            row_cells[1].text = f"{load['Power (kW)']:.1f}"
            row_cells[2].text = f"{load['Voltage (V)']:.0f}"
            row_cells[3].text = load['Phase']
            row_cells[4].text = f"{load['Power Factor']:.2f}"
            row_cells[5].text = f"{load['Length (m)']:.0f}"
        self.doc.add_paragraph()
    
    def add_cable_results(self, cable_df):
        self.doc.add_heading('4. CABLE SIZING RESULTS', level=1)
        table = self.doc.add_table(rows=1, cols=9)
        table.style = 'Light Grid Accent 1'
        hdr_cells = table.rows[0].cells
        hdr_cells[0].text = 'Load'
        hdr_cells[1].text = 'Size'
        hdr_cells[2].text = 'Type'
        hdr_cells[3].text = 'Base A'
        hdr_cells[4].text = 'Derated'
        hdr_cells[5].text = 'VD %'
        hdr_cells[6].text = 'SC kA'
        hdr_cells[7].text = 'Efficiency'
        hdr_cells[8].text = 'Status'
        
        for idx, row in cable_df.iterrows():
            row_cells = table.add_row().cells
            row_cells[0].text = row['Load Name']
            row_cells[1].text = str(row['Size (mm²)'])
            row_cells[2].text = 'Cu'
            row_cells[3].text = str(row['Base Ampacity (A)'])
            row_cells[4].text = str(row['Derated Ampacity (A)']).replace(' A', '')
            row_cells[5].text = str(row['Voltage Drop (%)']).replace('%', '')
            row_cells[6].text = str(row['Short Circuit (kA)']).replace(' kA', '')
            row_cells[7].text = str(row['Efficiency (%)']).replace('%', '')
            row_cells[8].text = row['Status']
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

# ========== SIDEBAR ==========
with st.sidebar:
    st.markdown("### ⚡ CES-Electrical Design Calculations")
    st.markdown("---")
    
    # Calculator Navigation
    calculators = [
        "⚡ Lightning Protection",
        "🔌 Cable Sizing",
        "⚡ Circuit Breaker Sizing",
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
        "📥 Loads Input", 
        "📊 Derating Factors", 
        "⚡ Cable Selection",
        "🔧 Short Circuit",
        "📥 Download Report"
    ])
    
    # TAB 1: LOADS INPUT
    with cable_tabs[0]:
        st.markdown('<div class="report-header">', unsafe_allow_html=True)
        st.markdown("## CABLE SIZING - LOADS INPUT")
        st.markdown('</div>', unsafe_allow_html=True)
        
        st.markdown('<div class="info-box">', unsafe_allow_html=True)
        st.markdown("### 📋 Enter Load Details")
        st.markdown("""
        - **Add/Delete Rows:** Use the buttons below
        - **Phase Options:** 1-phase, 3-phase, DC
        - **Voltage:** LV (≤1000V) or MV (>1000V)
        """)
        st.markdown('</div>', unsafe_allow_html=True)
        
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
        
        edited_df = st.data_editor(
            st.session_state.loads_df,
            num_rows="fixed",
            use_container_width=True,
            column_config={
                "Load Name": st.column_config.TextColumn("Load Name", width="medium"),
                "Power (kW)": st.column_config.NumberColumn("Power (kW)", min_value=0.0, max_value=10000.0, step=0.1),
                "Voltage (V)": st.column_config.NumberColumn("Voltage (V)", min_value=0.0, max_value=33000.0, step=1.0),
                "Phase": st.column_config.SelectboxColumn("Phase", options=['1-phase', '3-phase', 'DC']),
                "Power Factor": st.column_config.NumberColumn("PF", min_value=0.5, max_value=1.0, step=0.05),
                "Length (m)": st.column_config.NumberColumn("Length (m)", min_value=1.0, max_value=5000.0, step=1.0)
            }
        )
        st.session_state.loads_df = edited_df
        
        st.markdown("### ⚙️ Installation Parameters")
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            cable_type = st.selectbox("Cable Type", ['armoured', 'unarmoured'])
        with col2:
            ambient_temp = st.number_input("Ambient Temp (°C)", value=30.0, step=5.0)
        with col3:
            num_cables = st.number_input("Cables in Group", value=3, min_value=1, max_value=18)
        with col4:
            grouping = st.selectbox("Grouping", ['touching', 'spaced'])
        
        col5, col6, col7, col8 = st.columns(4)
        with col5:
            laying = st.selectbox("Laying Method", ['air', 'surface', 'buried', 'duct'])
        with col6:
            soil_res = st.number_input("Soil Resistivity (K.m/W)", value=1.5, step=0.5, min_value=0.5, max_value=3.0)
        with col7:
            depth = st.number_input("Burial Depth (m)", value=0.8, step=0.1, min_value=0.3, max_value=2.0)
        with col8:
            system_type = st.selectbox("System Type", ['TN-S', 'TN-C', 'TN-C-S', 'TT'])
        
        # Store parameters in session state
        st.session_state.cable_type = cable_type
        st.session_state.ambient_temp = ambient_temp
        st.session_state.num_cables = num_cables
        st.session_state.grouping = grouping
        st.session_state.laying = laying
        st.session_state.soil_res = soil_res
        st.session_state.depth = depth
        
        if st.button("🔧 CALCULATE", type="primary", use_container_width=True):
            with st.spinner("Calculating..."):
                cable_calc = CableSizingCalculator()
                
                # Calculate for each load
                cable_results = []
                detailed_calcs = []
                
                for idx, load in st.session_state.loads_df.iterrows():
                    # Determine cable category (LV/MV)
                    cable_category, cable_db = cable_calc.get_cable_type(load['Voltage (V)'])
                    db = cable_db[cable_type]
                    
                    # Load current
                    current = cable_calc.calculate_load_current(
                        load['Power (kW)'], load['Voltage (V)'], load['Power Factor'], 1.0, load['Phase']
                    )
                    
                    # Get ALL derating factors
                    total_k, factors = cable_calc.get_all_derating_factors(
                        ambient_temp, 90, num_cables, grouping, soil_res, depth, laying
                    )
                    
                    # Store factors in session state
                    st.session_state.derating_factors = factors
                    
                    # Find suitable cable
                    found = False
                    for size, data in db.items():
                        if found:
                            break
                        derated = data['ampacity'] * total_k
                        if derated >= current:
                            # Voltage drop calculation
                            vd_v, vd_pct = cable_calc.calculate_voltage_drop(
                                current, load['Length (m)'], data['R'], data['X'],
                                load['Power Factor'], load['Voltage (V)'], load['Phase']
                            )
                            
                            # Short circuit calculation
                            isc = cable_calc.calculate_short_circuit(size, 1.0)
                            
                            # Efficiency calculation
                            if load['Phase'] == '3-phase':
                                input_power = 1.732 * load['Voltage (V)'] * current / 1000
                            elif load['Phase'] == '1-phase':
                                input_power = load['Voltage (V)'] * current / 1000
                            else:  # DC
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
                            
                            # Store detailed calculation
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
                        st.warning(f"No suitable cable found for {load['Load Name']}")
                
                st.session_state.cable_results_df = pd.DataFrame(cable_results)
                st.session_state.detailed_calcs = detailed_calcs
                st.success("✅ Calculations complete! Check other tabs for results.")
    
    # TAB 2: DERATING FACTORS
    with cable_tabs[1]:
        st.markdown('<div class="report-header">', unsafe_allow_html=True)
        st.markdown("## ALL DERATING FACTORS (IEC 60502-2)")
        st.markdown('</div>', unsafe_allow_html=True)
        
        if st.session_state.derating_factors:
            factors = st.session_state.derating_factors
            factors_html = "<table class='parameter-table'>"
            factors_html += "<tr><th>Factor</th><th>Value</th><th>Reference</th><th>Description</th></tr>"
            
            desc_map = {
                'k1 (Temperature)': 'Ambient temperature correction',
                'k2 (Grouping)': 'Number of cables grouped together',
                'k3 (Soil Resistivity)': 'Soil thermal resistivity',
                'k4 (Depth)': 'Depth of laying correction',
                'k5 (Laying)': 'Installation method correction'
            }
            
            for key, data in factors.items():
                if key != 'total':
                    factors_html += f"<tr><td>{key}</td><td>{data['value']:.3f}</td><td>{data['reference']}</td><td>{desc_map.get(key, '')}</td></tr>"
            
            factors_html += f"<tr style='background-color: #1E3A8A; color: white;'><td colspan='4'><strong>Total K = {factors['total']:.3f}</strong></td></tr>"
            factors_html += "</table>"
            
            st.markdown(factors_html, unsafe_allow_html=True)
        else:
            st.info("👈 Calculate loads first")
    
    # TAB 3: CABLE SELECTION - WITH DETAILED CALCULATIONS
    with cable_tabs[2]:
        st.markdown('<div class="report-header">', unsafe_allow_html=True)
        st.markdown("## CABLE SELECTION RESULTS")
        st.markdown('</div>', unsafe_allow_html=True)
        
        st.markdown("### ⚡ Voltage Drop Limit: **2.5%** [IEC 60364-5-52 Section 525]")
        
        if not st.session_state.cable_results_df.empty:
            # Display main results table
            st.dataframe(st.session_state.cable_results_df, use_container_width=True, hide_index=True)
            
            st.markdown("---")
            st.markdown("### 📋 DETAILED CALCULATIONS WITH REFERENCES")
            
            # Display detailed calculations for each load
            if st.session_state.detailed_calcs:
                for calc in st.session_state.detailed_calcs:
                    with st.expander(f"🔍 Detailed Calculation for {calc['load_name']}", expanded=False):
                        
                        st.markdown(f"""
<div class="calculation-detail">

### 📊 Load: {calc['load_name']}

#### **Step 1: Load Current Calculation [IEC 60364-5-52 Section 523]**

**Formula:** I = P × 1000 / (√3 × V × PF)

**Calculation:** I = {calc['power']} × 1000 / (1.732 × {calc['voltage']} × {calc['pf']}) = **{calc['current']:.1f} A**

---

#### **Step 2: Derating Factors [IEC 60502-2 Tables B.10-B.22]**

| Factor | Value | Reference |
|--------|-------|-----------|
| k1 - Temperature Correction | {calc['k1']:.3f} | Table B.10 at {st.session_state.get('ambient_temp', 30)}°C |
| k2 - Grouping Factor | {calc['k2']:.3f} | Table 4C1 ({st.session_state.get('grouping', 'touching')}) |
| k3 - Soil Resistivity | {calc['k3']:.3f} | Table B.14 |
| k4 - Depth Factor | {calc['k4']:.3f} | Table B.12 |
| k5 - Laying Method | {calc['k5']:.3f} | Table B.5 |

**Total Derating Factor K = k1 × k2 × k3 × k4 × k5 = {calc['total_k']:.3f}**

---

#### **Step 3: Cable Selection**

**Selected Cable:** {calc['size']} mm² {calc['cable_type']} copper ({calc['cable_category']})

**Base Ampacity (Ic):** {calc['base_amp']} A

**Derated Ampacity (Id):** K × Ic = {calc['total_k']:.3f} × {calc['base_amp']} = **{calc['derated_amp']:.1f} A**

**Check:** {calc['derated_amp']:.1f} A ≥ {calc['current']:.1f} A → **{'✅ PASS' if calc['derated_amp'] >= calc['current'] else '❌ FAIL'}**

---

#### **Step 4: Voltage Drop Calculation [IEC 60364-5-52 Section 525]**

**Voltage Drop:** {calc['vd_pct']:.3f}%

**Limit:** 2.5% (Fixed per IEC 60364-5-52)

**Check:** {calc['vd_pct']:.3f}% ≤ 2.5% → **{'✅ PASS' if calc['vd_pct'] <= 2.5 else '❌ FAIL'}**

---

#### **Step 5: Short Circuit Calculation [IEC 60949]**

**Formula:** Isc = K × S / √t

**K = 143** (Copper, XLPE insulated, 90°C operating, 250°C short-circuit)

**S =** {calc['size']} mm²

**t =** 1.0 s (Typical clearing time)

**Isc =** 143 × {calc['size']} / √1.0 = **{calc['sc']:.2f} kA**

---

#### **Step 6: Efficiency Calculation**

**Input Power:** {calc['input_power']:.1f} kW

**Output Power:** {calc['power']} kW

**Efficiency η =** (Output / Input) × 100 = ({calc['power']} / {calc['input_power']:.1f}) × 100 = **{calc['efficiency']:.1f}%**

---

#### **Final Status:** {'✅ PASS' if calc['vd_pct'] <= 2.5 and calc['derated_amp'] >= calc['current'] else '❌ FAIL'}

</div>
""", unsafe_allow_html=True)
            else:
                st.info("No detailed calculations available")
        else:
            st.info("👈 Calculate loads first")
    
    # TAB 4: SHORT CIRCUIT
    with cable_tabs[3]:
        st.markdown('<div class="report-header">', unsafe_allow_html=True)
        st.markdown("## SHORT CIRCUIT CALCULATIONS")
        st.markdown('</div>', unsafe_allow_html=True)
        
        st.markdown("""
        **Reference:** IEC 60949 - Calculation of thermally permissible short-circuit currents
        **Formula:** Isc = K × S / √t
        - K = 143 for Copper, XLPE insulated (90°C operating, 250°C short-circuit)
        - S = Conductor cross-sectional area (mm²)
        - t = Duration of short circuit (seconds)
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
            st.markdown("### 📊 Calculated Cables Short Circuit Capacity")
            df = st.session_state.cable_results_df[['Load Name', 'Size (mm²)', 'Short Circuit (kA)']]
            st.dataframe(df, use_container_width=True, hide_index=True)
    
    # TAB 5: DOWNLOAD REPORT
    with cable_tabs[4]:
        st.markdown('<div class="report-header">', unsafe_allow_html=True)
        st.markdown("## DOWNLOAD REPORT")
        st.markdown('</div>', unsafe_allow_html=True)
        
        if not st.session_state.cable_results_df.empty:
            st.markdown("### 📥 Download Options")
            
            col1, col2 = st.columns(2)
            
            with col1:
                st.markdown("#### 📄 PDF Format")
                if st.button("📥 Generate PDF Report", key="cable_pdf_btn", use_container_width=True):
                    with st.spinner("Generating PDF report..."):
                        pdf = CablePDFReport()
                        pdf.add_title()
                        
                        # Installation parameters
                        params = {
                            'Cable Type': f'{st.session_state.get("cable_type", "armoured")} copper',
                            'Ambient Temperature': f'{st.session_state.get("ambient_temp", 30)}°C',
                            'Cables in Group': str(st.session_state.get("num_cables", 3)),
                            'Grouping': st.session_state.get("grouping", "touching"),
                            'Laying Method': st.session_state.get("laying", "air"),
                            'Soil Resistivity': f'{st.session_state.get("soil_res", 1.5)} K.m/W',
                            'Burial Depth': f'{st.session_state.get("depth", 0.8)} m'
                        }
                        pdf.add_installation_parameters(params)
                        
                        # Derating factors
                        if st.session_state.derating_factors:
                            pdf.add_derating_factors(st.session_state.derating_factors)
                        
                        # Load details
                        pdf.add_load_details(st.session_state.loads_df)
                        
                        # Cable results
                        pdf.add_cable_results(st.session_state.cable_results_df)
                        
                        pdf_output = pdf.output(dest='S')
                        b64 = base64.b64encode(pdf_output).decode()
                        filename = f"Cable_Sizing_Report_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf"
                        st.markdown(f'<a href="data:application/pdf;base64,{b64}" download="{filename}" class="download-btn pdf-btn">📥 Click to Download PDF</a>', unsafe_allow_html=True)
                        st.success("✅ PDF generated successfully!")
            
            with col2:
                st.markdown("#### 📝 Word Format")
                if st.button("📥 Generate Word Report", key="cable_word_btn", use_container_width=True):
                    with st.spinner("Generating Word report..."):
                        word = CableWordReport()
                        word.add_title()
                        
                        # Installation parameters
                        params = {
                            'Cable Type': f'{st.session_state.get("cable_type", "armoured")} copper',
                            'Ambient Temperature': f'{st.session_state.get("ambient_temp", 30)}°C',
                            'Cables in Group': str(st.session_state.get("num_cables", 3)),
                            'Grouping': st.session_state.get("grouping", "touching"),
                            'Laying Method': st.session_state.get("laying", "air"),
                            'Soil Resistivity': f'{st.session_state.get("soil_res", 1.5)} K.m/W',
                            'Burial Depth': f'{st.session_state.get("depth", 0.8)} m'
                        }
                        word.add_installation_parameters(params)
                        
                        # Derating factors
                        if st.session_state.derating_factors:
                            word.add_derating_factors(st.session_state.derating_factors)
                        
                        # Load details
                        word.add_load_details(st.session_state.loads_df)
                        
                        # Cable results
                        word.add_cable_results(st.session_state.cable_results_df)
                        
                        word_path = "temp_cable_report.docx"
                        word.save(word_path)
                        
                        with open(word_path, "rb") as f:
                            word_bytes = f.read()
                        
                        b64 = base64.b64encode(word_bytes).decode()
                        filename = f"Cable_Sizing_Report_{datetime.now().strftime('%Y%m%d_%H%M')}.docx"
                        
                        if os.path.exists(word_path):
                            os.remove(word_path)
                        
                        st.markdown(f'<a href="data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,{b64}" download="{filename}" class="download-btn word-btn">📥 Click to Download Word</a>', unsafe_allow_html=True)
                        st.success("✅ Word report generated successfully!")
        else:
            st.info("👈 Calculate cable sizes first")

# ========== CIRCUIT BREAKER SIZING CALCULATOR ==========
elif st.session_state.selected_calculator == "⚡ Circuit Breaker Sizing":
    
    st.markdown('<div class="report-header">', unsafe_allow_html=True)
    st.markdown("## CIRCUIT BREAKER SIZING")
    st.markdown('</div>', unsafe_allow_html=True)
    
    st.markdown("""
    ### 🔍 Circuit Breaker Selection Criteria
    
    **Pole Selection Guide (IEC 60364-5-53):**
    - **1P (Single Pole):** Phase protection only - for IT systems
    - **2P (Double Pole):** Phase + Neutral protection - Required for TN/TT systems
    - **3P (Three Pole):** For 3-wire systems (no neutral)
    - **4P (Four Pole):** For 4-wire systems with neutral protection
    
    **Breaker Types (IEC 60898 / IEC 60947-2):**
    - **MCB:** ≤ 125A - Miniature Circuit Breakers
    - **MCCB:** 125A - 1600A - Moulded Case Circuit Breakers
    - **ACB:** ≥ 1600A - Air Circuit Breakers
    """)
    
    if 'loads_df' in st.session_state and not st.session_state.loads_df.empty:
        cb_calc = CircuitBreakerCalculator()
        system_type = st.selectbox("System Type", ['TN-S', 'TN-C', 'TN-C-S', 'TT'])
        manufacturer = st.selectbox("Manufacturer", list(MANUFACTURERS.keys()))
        
        cb_results = cb_calc.calculate_cb_size(st.session_state.loads_df, 1.25, manufacturer, system_type)
        
        cb_df = pd.DataFrame([{
            'Load': r['Load'],
            'Power (kW)': r['Power (kW)'],
            'Current (A)': f"{r['Current (A)']:.1f}",
            'Required CB (A)': f"{r['Required CB (A)']:.1f}",
            'Selected CB (A)': r['Selected CB (A)'],
            'Type': r['Breaker Type'],
            'Standard': r['Standard'],
            'Poles': r['Poles'],
            'Selection Reason': r['Pole Selection Reason']
        } for r in cb_results])
        
        st.dataframe(cb_df, use_container_width=True, hide_index=True)
        
        # Main Circuit Breaker
        st.markdown("### 🔋 Main Circuit Breaker")
        main_cb = cb_calc.calculate_main_cb(st.session_state.loads_df, 400, 0.85, 1.25)
        
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("Total Power", f"{main_cb['total_power']:.1f} kW")
        with col2:
            st.metric("Total Current", f"{main_cb['current']:.1f} A")
        with col3:
            st.metric("Required CB", f"{main_cb['required_cb']:.1f} A")
        with col4:
            st.metric("Selected CB", f"{main_cb['selected_cb']} A {main_cb['breaker_type']} {main_cb['poles']}")
    else:
        st.info("👈 Add loads in Cable Sizing calculator first")

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
st.markdown(f"<div style='text-align: center; color: gray;'>⚡ CES-Electrical Design Calculators | IEC Compliant | Version 48.0 | {datetime.now().strftime('%Y-%m-%d %H:%M')}</div>", unsafe_allow_html=True)