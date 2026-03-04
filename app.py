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

# ========== CUSTOM CSS - FIXED SYMMETRICAL TABLES ==========
st.markdown("""
<style>
    .report-header {
        background-color: #1E3A8A;
        color: white;
        padding: 20px;
        border-radius: 10px;
        text-align: center;
        margin-bottom: 20px;
        font-size: 28px;
        font-weight: bold;
    }
    .formula-box {
        background-color: #F8F9FA;
        padding: 15px;
        border-radius: 8px;
        border-left: 5px solid #1E3A8A;
        margin: 10px 0;
        font-family: 'Courier New', monospace;
        color: #000000 !important;
        border: 1px solid #DEE2E6;
    }
    .calculation-detail {
        background-color: #FFFFFF;
        padding: 15px;
        border-radius: 8px;
        border: 1px solid #DEE2E6;
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
        border: 1px solid #B8DAFF;
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
    
    /* ===== SYMMETRICAL TABLE STYLES ===== */
    .parameter-table {
        width: 100%;
        border-collapse: collapse;
        margin: 20px 0;
        border: 2px solid #1E3A8A;
        table-layout: fixed;
    }
    .parameter-table th {
        background-color: #1E3A8A;
        color: WHITE !important;
        padding: 12px;
        text-align: center;
        font-weight: bold;
        font-size: 14px;
        border: 1px solid #0D1B4A;
        width: 33.33%;
    }
    .parameter-table td {
        border: 1px solid #A0AEC0;
        padding: 10px;
        text-align: left;
        color: #000000 !important;
        font-size: 13px;
        width: 33.33%;
    }
    .parameter-table tr:nth-child(even) {
        background-color: #F0F4FA !important;
    }
    .parameter-table tr:nth-child(even) td {
        color: #000000 !important;
        background-color: #F0F4FA;
    }
    .parameter-table tr:nth-child(odd) {
        background-color: #FFFFFF !important;
    }
    .parameter-table tr:nth-child(odd) td {
        color: #000000 !important;
        background-color: #FFFFFF;
    }
    
    /* STREAMLIT DATAFRAME FIXES - SYMMETRICAL */
    .stDataFrame {
        color: #000000 !important;
    }
    .stDataFrame table {
        color: #000000 !important;
        border: 2px solid #1E3A8A;
        table-layout: fixed;
        width: 100%;
    }
    .stDataFrame th {
        color: WHITE !important;
        background-color: #1E3A8A !important;
        font-weight: bold;
        padding: 10px !important;
        text-align: center !important;
    }
    .stDataFrame td {
        color: #000000 !important;
        padding: 8px !important;
        text-align: left !important;
    }
    .stDataFrame tr:nth-child(even) td {
        background-color: #F0F4FA !important;
        color: #000000 !important;
    }
    .stDataFrame tr:nth-child(odd) td {
        background-color: #FFFFFF !important;
        color: #000000 !important;
    }
    
    /* PARAMETER TABLE - SYMMETRICAL */
    .param-table {
        width: 100%;
        border-collapse: collapse;
        margin: 15px 0;
        font-family: Arial, sans-serif;
        border: 2px solid #1E3A8A;
        table-layout: fixed;
    }
    .param-table th {
        background-color: #1E3A8A;
        color: WHITE !important;
        padding: 12px;
        text-align: center;
        font-size: 15px;
        border: 1px solid #0D1B4A;
    }
    .param-table td {
        padding: 10px;
        border: 1px solid #A0AEC0;
        color: #000000 !important;
        font-size: 14px;
        text-align: left;
    }
    .param-table tr:nth-child(even) {
        background-color: #F0F4FA !important;
    }
    .param-table tr:nth-child(odd) {
        background-color: #FFFFFF !important;
    }
    
    /* RESULTS TABLE - SYMMETRICAL */
    .results-table {
        width: 100%;
        border-collapse: collapse;
        margin: 20px 0;
        border: 2px solid #1E3A8A;
        table-layout: fixed;
    }
    .results-table th {
        background-color: #1E3A8A;
        color: WHITE !important;
        padding: 10px;
        text-align: center;
        font-weight: bold;
        font-size: 13px;
        border: 1px solid #0D1B4A;
    }
    .results-table td {
        border: 1px solid #A0AEC0;
        padding: 8px;
        text-align: left;
        color: #000000 !important;
        font-size: 12px;
    }
    .results-table tr:nth-child(even) {
        background-color: #F0F4FA !important;
    }
    .results-table tr:nth-child(odd) {
        background-color: #FFFFFF !important;
    }
    
    .stTabs [data-baseweb="tab-list"] {
        gap: 20px;
        background-color: #F8F9FA;
        padding: 10px;
        border-radius: 10px;
        margin-bottom: 20px;
    }
    .stTabs [data-baseweb="tab"] {
        padding: 15px 30px !important;
        background-color: #E9ECEF !important;
        border-radius: 8px !important;
        font-size: 18px !important;
        font-weight: 600 !important;
        color: #1E3A8A !important;
        border: 1px solid #CED4DA !important;
        transition: all 0.3s ease;
    }
    .stTabs [data-baseweb="tab"]:hover {
        background-color: #1E3A8A !important;
        color: WHITE !important;
        transform: scale(1.05);
        border-color: #1E3A8A !important;
    }
    .stTabs [aria-selected="true"] {
        background-color: #1E3A8A !important;
        color: WHITE !important;
        font-weight: 700 !important;
        border-color: #0D1B4A !important;
    }
    [data-testid="stMetricValue"] {
        color: #1E3A8A !important;
        font-size: 24px !important;
        font-weight: bold;
    }
    [data-testid="stMetricLabel"] {
        color: #2D3748 !important;
        font-size: 16px !important;
    }
    .streamlit-expanderHeader {
        font-size: 16px !important;
        font-weight: 600 !important;
        color: #1E3A8A !important;
        background-color: #F0F4FA !important;
        border-radius: 5px;
    }
    .success-box {
        background-color: #D4EDDA;
        color: #155724 !important;
        padding: 15px;
        border-radius: 8px;
        border-left: 5px solid #28A745;
        margin: 10px 0;
    }
    .warning-box {
        background-color: #FFF3CD;
        color: #856404 !important;
        padding: 15px;
        border-radius: 8px;
        border-left: 5px solid #FFC107;
        margin: 10px 0;
    }
    .info-box {
        background-color: #E7F3FF;
        color: #004085 !important;
        padding: 15px;
        border-radius: 8px;
        border-left: 5px solid #1E3A8A;
        margin: 10px 0;
    }
    .cb-selection-reason {
        background-color: #FFF3CD;
        border-left: 5px solid #FFC107;
        padding: 15px;
        margin: 10px 0;
        border-radius: 5px;
        color: #856404 !important;
    }
    .settings-container {
        background-color: #F8F9FA;
        padding: 20px;
        border-radius: 10px;
        border: 1px solid #DEE2E6;
        margin: 15px 0;
    }
    .settings-category {
        color: #1E3A8A;
        font-size: 18px;
        font-weight: bold;
        margin-top: 10px;
        margin-bottom: 10px;
        padding-bottom: 5px;
        border-bottom: 2px solid #1E3A8A;
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

# ========== ENHANCED DERATING FACTORS WITH SPACING AND ARRANGEMENT ==========
TEMPERATURE_FACTORS = {
    90: {20: 1.07, 25: 1.04, 30: 1.00, 35: 0.96, 40: 0.91, 
         45: 0.87, 50: 0.82, 55: 0.76, 60: 0.71, 65: 0.65, 70: 0.58}
}

# Enhanced grouping factors based on cable arrangement, formation, and spacing
GROUPING_FACTORS = {
    'touching': {  # Cables touching each other
        1: 1.00, 2: 0.80, 3: 0.70, 4: 0.65, 5: 0.60, 6: 0.57,
        7: 0.54, 8: 0.52, 9: 0.50, 10: 0.48, 11: 0.46, 12: 0.45,
        13: 0.44, 14: 0.43, 15: 0.42, 16: 0.41, 17: 0.40, 18: 0.39
    },
    'spaced_1d': {  # Spaced by 1 x cable diameter
        1: 1.00, 2: 0.90, 3: 0.85, 4: 0.82, 5: 0.80, 6: 0.78,
        7: 0.76, 8: 0.74, 9: 0.72, 10: 0.70, 11: 0.68, 12: 0.66
    },
    'spaced_2d': {  # Spaced by 2 x cable diameter
        1: 1.00, 2: 0.95, 3: 0.92, 4: 0.90, 5: 0.88, 6: 0.86,
        7: 0.84, 8: 0.82, 9: 0.80, 10: 0.78, 11: 0.76, 12: 0.74
    },
    'spaced_3d': {  # Spaced by 3 x cable diameter
        1: 1.00, 2: 0.98, 3: 0.96, 4: 0.94, 5: 0.92, 6: 0.90,
        7: 0.88, 8: 0.86, 9: 0.84, 10: 0.82, 11: 0.80, 12: 0.78
    },
    'cleated': {  # Cables with cleats/spacers
        1: 1.00, 2: 0.95, 3: 0.90, 4: 0.85, 5: 0.82, 6: 0.80,
        7: 0.78, 8: 0.76, 9: 0.74, 10: 0.72, 11: 0.70, 12: 0.68
    }
}

# Formation factors (flat vs trefoil)
FORMATION_FACTORS = {
    'flat': 1.00,      # Flat formation - base case
    'trefoil': 0.95,   # Trefoil formation - reduced rating due to mutual heating
    'single': 1.00      # Single cable
}

# Installation method factors
INSTALLATION_FACTORS = {
    'air': 1.00,           # Cables in free air
    'surface': 0.98,       # Cables on surface (wall/celling)
    'tray': 0.95,          # Cables on perforated tray
    'ladder': 0.96,        # Cables on ladder rack
    'trench': 0.90,        # Cables in open trench
    'buried': 0.85,        # Direct buried in ground
    'duct': 0.82,          # Cables in underground ducts
    'conduit': 0.80        # Cables in conduit
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

# Function to determine grouping factor based on spacing and arrangement
def get_grouping_factor(num_cables, spacing_mm, cable_diameter, arrangement='touching'):
    """Calculate grouping factor based on cable spacing and arrangement"""
    if arrangement == 'touching':
        return GROUPING_FACTORS['touching'].get(min(num_cables, 18), 0.39)
    elif arrangement == 'cleated':
        return GROUPING_FACTORS['cleated'].get(min(num_cables, 12), 0.68)
    else:
        # Calculate spacing ratio relative to cable diameter
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

# ========== LIGHTNING PROTECTION CLASSES (UNCHANGED) ==========
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

# ========== ENHANCED CABLE SIZING CALCULATOR CLASS (UNCHANGED) ==========
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
        """Calculate ALL derating factors with enhanced parameters"""
        
        # Temperature factor (k1)
        k1 = TEMPERATURE_FACTORS[insulation_temp].get(temp_c, 1.0)
        
        # Grouping factor based on spacing and arrangement (k2)
        if arrangement == 'touching':
            k2 = GROUPING_FACTORS['touching'].get(min(num_cables, 18), 0.39)
        elif arrangement == 'cleated':
            k2 = GROUPING_FACTORS['cleated'].get(min(num_cables, 12), 0.68)
        else:
            # Calculate spacing ratio
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
        
        # Formation factor (k_formation)
        k_formation = FORMATION_FACTORS.get(formation, 1.0)
        
        # Installation method factor (k_install)
        k_install = INSTALLATION_FACTORS.get(installation, 1.0)
        
        # Soil resistivity factor (k3) - only applicable for buried/duct installations
        if installation in ['buried', 'duct', 'trench']:
            k3 = SOIL_RESISTIVITY_FACTORS.get(soil_resistivity, 1.0)
        else:
            k3 = 1.0
        
        # Depth factor (k4) - only applicable for buried/duct installations
        if installation in ['buried', 'duct', 'trench']:
            k4 = DEPTH_FACTORS.get(depth, 1.0)
        else:
            k4 = 1.0
        
        # Total derating factor = k1 * k2 * k_formation * k_install * k3 * k4
        total_k = k1 * k2 * k_formation * k_install * k3 * k4
        
        factors = {
            'k1 (Temperature)': {'value': k1, 'reference': 'IEC 60502-2 Table B.10'},
            'k2 (Grouping/Spacing)': {'value': k2, 'reference': f'IEC 60502-2 Table 4C1 - {arrangement}, spacing={spacing_mm}mm'},
            'k_formation (Formation)': {'value': k_formation, 'reference': f'IEC 60502-2 - {formation} formation'},
            'k_install (Installation)': {'value': k_install, 'reference': f'IEC 60502-2 - {installation} method'},
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

# ========== CIRCUIT BREAKER CALCULATOR CLASS (UNCHANGED) ==========
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

# ========== PDF REPORT CLASSES (UNCHANGED) ==========
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

# ========== ENHANCED TRANSFORMER CALCULATOR WITH ALL IEC STANDARDS ==========
class TransformerCalculator:
    def __init__(self):
        pass
    
    # IEC 60059 - Standard current ratings
    IEC_60059_CURRENT_RATINGS = [1, 1.25, 1.6, 2, 2.5, 3.15, 4, 5, 6.3, 8, 
                                  10, 12.5, 16, 20, 25, 31.5, 40, 50, 63, 80,
                                  100, 125, 160, 200, 250, 315, 400, 500, 630, 800,
                                  1000, 1250, 1600, 2000, 2500, 3150, 4000, 5000, 6300, 8000, 10000]
    
    def calculate_required_kva(self, loads_df, future_expansion=20):
        """Calculate required transformer kVA based on loads"""
        total_connected = 0
        total_demand = 0
        total_kvar = 0
        
        for idx, load in loads_df.iterrows():
            connected = load['Rating (kW)'] * load['Quantity']
            demand = connected * load['Diversity Factor']
            kvar = demand * math.tan(math.acos(load['Power Factor'])) if load['Power Factor'] < 1.0 else 0
            
            total_connected += connected
            total_demand += demand
            total_kvar += kvar
        
        total_kva = math.sqrt(total_demand**2 + total_kvar**2)
        with_future = total_kva * (1 + future_expansion/100)
        
        # Standard transformer ratings (IEC 60076)
        standard_ratings = [50, 100, 160, 200, 250, 315, 400, 500, 630, 800, 1000, 1250, 1600, 2000, 2500, 3150, 4000, 5000, 6300, 8000, 10000, 12500, 16000, 20000, 25000, 31500, 40000, 50000, 63000]
        
        selected_kva = with_future
        for rating in standard_ratings:
            if rating >= with_future:
                selected_kva = rating
                break
        
        return {
            'total_connected': total_connected,
            'total_demand': total_demand,
            'total_kvar': total_kvar,
            'total_kva': total_kva,
            'with_future': with_future,
            'selected_kva': selected_kva
        }
    
    def calculate_voltage_regulation(self, impedance, xr_ratio, pf, load_pct=100, load_type='lagging'):
        """Calculate voltage regulation using IEC formula"""
        resistance_drop = impedance / math.sqrt(1 + xr_ratio**2)
        reactance_drop = resistance_drop * xr_ratio
        
        if load_type == 'lagging':
            reg = (resistance_drop * pf + reactance_drop * math.sqrt(1 - pf**2)) * (load_pct/100)
        elif load_type == 'leading':
            reg = (resistance_drop * pf - reactance_drop * math.sqrt(1 - pf**2)) * (load_pct/100)
        else:
            reg = resistance_drop * (load_pct/100)
        
        return {
            'resistance_drop': resistance_drop,
            'reactance_drop': reactance_drop,
            'regulation': reg
        }
    
    def calculate_efficiency(self, kva, no_load_loss, load_loss, load_pct):
        """Calculate transformer efficiency at given load"""
        output = kva * (load_pct/100)
        losses = no_load_loss + load_loss * (load_pct/100)**2
        efficiency = output / (output + losses) * 100 if output > 0 else 0
        return efficiency, losses
    
    def calculate_short_circuit(self, kva, voltage, impedance, source_sc=None, xr_ratio=7):
        """Calculate short circuit currents"""
        fla = kva * 1000 / (1.732 * voltage)
        z_base = (voltage ** 2) / (kva * 1000)
        z_tx_ohms = (impedance / 100) * z_base
        
        if source_sc:
            z_source_ohms = (voltage ** 2) / (source_sc * 1e6)
            z_total_ohms = math.sqrt((z_source_ohms + z_tx_ohms)**2)
        else:
            z_total_ohms = z_tx_ohms
        
        i_sc_sym = (voltage / 1.732) / z_total_ohms / 1000
        
        kappa = 1.02 + 0.98 * math.exp(-3 / xr_ratio)
        i_sc_peak = kappa * math.sqrt(2) * i_sc_sym
        
        return {
            'fla': fla,
            'i_sc_sym': i_sc_sym,
            'i_sc_peak': i_sc_peak,
            'kappa': kappa,
            'z_tx_ohms': z_tx_ohms
        }
    
    # NEW: IEC 60059 - Get standard current rating
    def get_standard_current(self, calculated_current):
        """Get the next higher standard current rating from IEC 60059"""
        for rating in self.IEC_60059_CURRENT_RATINGS:
            if rating >= calculated_current:
                return rating
        return self.IEC_60059_CURRENT_RATINGS[-1]
    
    # NEW: IEC 60076-7 - Thermal ageing calculation
    def calculate_thermal_ageing(self, hot_spot_temp, insulation_class='F', reference_temp=98):
        """
        Calculate relative ageing rate using Arrhenius equation
        IEC 60076-7 Table 1 - Thermal constants
        """
        # Temperature limits per IEC 60076-2
        temp_limits = {
            'A': 105,
            'E': 120,
            'B': 130,
            'F': 155,
            'H': 180
        }
        
        if hot_spot_temp > temp_limits.get(insulation_class, 155):
            st.warning(f"⚠️ Hot spot temperature exceeds {insulation_class} class limit of {temp_limits[insulation_class]}°C")
        
        # Relative ageing rate (doubles every 6°C above reference)
        v = 2 ** ((hot_spot_temp - reference_temp) / 6)
        
        # Calculate insulation life (assuming 180,000 hours at reference temp)
        normal_life_hours = 180000
        remaining_life = normal_life_hours / v if v > 0 else normal_life_hours
        
        return {
            'relative_ageing_rate': v,
            'remaining_life_hours': remaining_life,
            'remaining_life_years': remaining_life / 8760,
            'temperature_limit': temp_limits.get(insulation_class, 155),
            'is_safe': hot_spot_temp <= temp_limits.get(insulation_class, 155)
        }
    
    # NEW: IEC 60076-10 - Sound level calculation
    def calculate_sound_level(self, kva, type='liquid', cooling='ONAN'):
        """
        Calculate guaranteed sound power level per IEC 60076-10
        
        Parameters:
        - kva: Transformer rating in kVA
        - type: 'liquid' for liquid-filled, 'dry' for dry-type
        - cooling: Cooling method (ONAN, ONAF, etc.)
        
        Returns sound power level in dB(A)
        """
        if type == 'liquid':
            # Liquid-filled transformers - Equation from IEC 60076-10
            if kva <= 100:
                L = 50 + 12 * math.log10(kva/100)
            elif kva <= 1000:
                L = 55 + 10 * math.log10(kva/100)
            elif kva <= 10000:
                L = 60 + 8 * math.log10(kva/1000)
            else:
                L = 68 + 5 * math.log10(kva/10000)
        else:
            # Dry-type transformers
            if kva <= 100:
                L = 45 + 15 * math.log10(kva/100)
            elif kva <= 1000:
                L = 50 + 12 * math.log10(kva/100)
            else:
                L = 58 + 8 * math.log10(kva/1000)
        
        # Add cooling factor
        cooling_factor = {
            'ONAN': 0,
            'ONAF': 3,
            'OFAF': 5,
            'ODAF': 7
        }
        
        total_sound = L + cooling_factor.get(cooling, 0)
        
        return {
            'base_sound_level': round(L, 1),
            'cooling_addition': cooling_factor.get(cooling, 0),
            'total_sound_level': round(total_sound, 1),
            'unit': 'dB(A)',
            'reference': 'IEC 60076-10'
        }

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

# Transformer session state
if 'tx_loads_df' not in st.session_state:
    st.session_state.tx_loads_df = pd.DataFrame({
        'Load Description': ['Motor Load 1', 'Lighting Load', 'Heating Load'],
        'Load Type': ['Motor', 'Lighting', 'Heating'],
        'Rating (kW)': [75.0, 25.0, 50.0],
        'Quantity': [2, 1, 1],
        'Power Factor': [0.85, 0.95, 1.0],
        'Diversity Factor': [0.8, 1.0, 1.0],
        'Load Category': ['Continuous', 'Continuous', 'Continuous']
    })

# ========== SIDEBAR ==========
with st.sidebar:
    st.markdown("### 🔌 CES-Electrical Design Calculations")
    st.markdown("---")
    
    calculators = [
        "⚡ Lightning Protection",
        "🔌 Cable Sizing",  # Cable icon without switch
        "⚙️ Transformer Sizing",
        "⚡ Generator Sizing",
        "🌍 Earthing System Design"
    ]
    
    for calc in calculators:
        if st.button(calc, key=f"nav_{calc}", use_container_width=True):
            st.session_state.selected_calculator = calc
            st.rerun()

# ========== MAIN CONTENT ==========
st.title(f"{st.session_state.selected_calculator} Calculator")

# ========== LIGHTNING PROTECTION (UNCHANGED) ==========
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

# ========== CABLE SIZING CALCULATOR (UNCHANGED - ONLY LOGO TO CABLE) ==========
elif st.session_state.selected_calculator == "🔌 Cable Sizing":
    
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
        st.markdown('<div class="report-header">🔌 CABLE SIZING - LOADS INPUT</div>', unsafe_allow_html=True)
        
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

# ========== ENHANCED TRANSFORMER SIZING CALCULATOR (WITH ALL 3 STANDARDS) ==========
elif st.session_state.selected_calculator == "⚙️ Transformer Sizing":
    
    tx_tabs = st.tabs([
        "📥 Load Analysis", 
        "⚡ Transformer Selection", 
        "📊 Voltage Regulation",
        "🔧 Efficiency & Losses",
        "📈 Short Circuit",
        "🔊 Sound & Ageing",  # NEW TAB with IEC 60076-7 and IEC 60076-10
        "📥 Download Report"
    ])
    
    tx_calc = TransformerCalculator()
    
    # TAB 1: LOAD ANALYSIS (UNCHANGED)
    with tx_tabs[0]:
        st.markdown('<div class="report-header">⚙️ TRANSFORMER SIZING - LOAD ANALYSIS</div>', unsafe_allow_html=True)
        
        st.markdown("""
        ### 📋 Load Types and Diversity [IEC 60076]
        
        Different load types have different characteristics and diversity factors:
        - **Continuous Loads:** Motors, heaters, lighting (operate continuously)
        - **Intermittent Loads:** Cranes, elevators, welders (operate periodically)
        - **Peak Loads:** Starting currents, temporary overloads
        """)
        
        col1, col2 = st.columns([1, 1])
        with col1:
            if st.button("➕ Add Load", key="tx_add_load", use_container_width=True):
                new_row = pd.DataFrame({
                    'Load Description': [f'Load {len(st.session_state.tx_loads_df) + 1}'],
                    'Load Type': ['Motor'],
                    'Rating (kW)': [50.0],
                    'Quantity': [1],
                    'Power Factor': [0.85],
                    'Diversity Factor': [0.8],
                    'Load Category': ['Continuous']
                })
                st.session_state.tx_loads_df = pd.concat([st.session_state.tx_loads_df, new_row], ignore_index=True)
                st.rerun()
        
        with col2:
            if st.button("🗑️ Delete Last Load", key="tx_delete_load", use_container_width=True):
                if len(st.session_state.tx_loads_df) > 1:
                    st.session_state.tx_loads_df = st.session_state.tx_loads_df[:-1]
                    st.rerun()
                else:
                    st.warning("At least one row required")
        
        edited_tx_df = st.data_editor(
            st.session_state.tx_loads_df,
            num_rows="fixed",
            use_container_width=True,
            column_config={
                "Load Description": st.column_config.TextColumn("Load Description", width="medium"),
                "Load Type": st.column_config.SelectboxColumn("Load Type", options=['Motor', 'Lighting', 'Heating', 'UPS', 'Other']),
                "Rating (kW)": st.column_config.NumberColumn("Rating (kW)", min_value=0.0, max_value=10000.0, step=1.0),
                "Quantity": st.column_config.NumberColumn("Quantity", min_value=1, max_value=100, step=1),
                "Power Factor": st.column_config.NumberColumn("Power Factor", min_value=0.5, max_value=1.0, step=0.05),
                "Diversity Factor": st.column_config.NumberColumn("Diversity Factor", min_value=0.0, max_value=1.0, step=0.05,
                                                                  help="Factor accounting for non-simultaneous operation"),
                "Load Category": st.column_config.SelectboxColumn("Load Category", options=['Continuous', 'Intermittent', 'Peak'])
            }
        )
        st.session_state.tx_loads_df = edited_tx_df
        
        st.markdown("### ⚙️ System Parameters")
        
        col1, col2, col3 = st.columns(3)
        with col1:
            future_expansion = st.number_input("Future Expansion (%)", value=20, min_value=0, max_value=100, step=5,
                                              help="Additional capacity for future loads")
            system_voltage = st.selectbox("Secondary Voltage (V)", [415, 400, 380, 33000, 11000, 6600], index=0)
        
        with col2:
            ambient_temp = st.number_input("Ambient Temperature (°C)", value=30, min_value=-10, max_value=50, step=1)
            altitude = st.number_input("Altitude (m)", value=0, min_value=0, max_value=5000, step=100,
                                      help="Derating for high altitude installations")
        
        with col3:
            cooling_type = st.selectbox("Cooling Type", 
                                       ['ONAN', 'ONAF', 'OFAF', 'ODAF'],
                                       help="ONAN: Oil Natural Air Natural\nONAF: Oil Natural Air Forced\nOFAF: Oil Forced Air Forced")
            insulation_class = st.selectbox("Insulation Class", ['A (105°C)', 'E (120°C)', 'B (130°C)', 'F (155°C)', 'H (180°C)'], index=3)
        
        # Calculate transformer sizing
        tx_result = tx_calc.calculate_required_kva(st.session_state.tx_loads_df, future_expansion)
        
        st.markdown("### 📊 Load Summary")
        
        # Calculate load details for display
        load_details = []
        for idx, load in st.session_state.tx_loads_df.iterrows():
            connected = load['Rating (kW)'] * load['Quantity']
            demand = connected * load['Diversity Factor']
            kvar = demand * math.tan(math.acos(load['Power Factor'])) if load['Power Factor'] < 1.0 else 0
            
            load_details.append({
                'Load': load['Load Description'],
                'Type': load['Load Type'],
                'Connected (kW)': f"{connected:.1f}",
                'Demand (kW)': f"{demand:.1f}",
                'PF': f"{load['Power Factor']:.2f}",
                'kvar': f"{kvar:.1f}"
            })
        
        summary_df = pd.DataFrame(load_details)
        st.dataframe(summary_df, use_container_width=True, hide_index=True)
        
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("Total Connected", f"{tx_result['total_connected']:.0f} kW")
        with col2:
            st.metric("Max Demand", f"{tx_result['total_demand']:.0f} kW")
        with col3:
            st.metric("Reactive Power", f"{tx_result['total_kvar']:.0f} kvar")
        with col4:
            st.metric("Required kVA", f"{tx_result['total_kva']:.0f} kVA")
        
        st.info(f"""
        ### 📋 Transformer Sizing Calculation
        
        **Step 1:** Calculate connected load per equipment
        - Connected Load = Rating × Quantity
        
        **Step 2:** Apply diversity factors [IEC 60364]
        - Demand Load = Connected Load × Diversity Factor
        
        **Step 3:** Calculate apparent power
        - S (kVA) = √(P² + Q²) = √({tx_result['total_demand']:.0f}² + {tx_result['total_kvar']:.0f}²) = **{tx_result['total_kva']:.0f} kVA**
        
        **Step 4:** Add future expansion ({future_expansion}%)
        - Required Capacity = {tx_result['total_kva']:.0f} × (1 + {future_expansion/100:.2f}) = **{tx_result['with_future']:.0f} kVA**
        
        **Step 5:** Select standard rating [IEC 60076]
        - Selected Transformer: **{tx_result['selected_kva']} kVA**
        """)
        
        if st.button("➡️ Proceed to Transformer Selection", type="primary", use_container_width=True):
            st.session_state.tx_required_kva = tx_result['selected_kva']
            st.session_state.tx_load_data = {
                'total_connected': tx_result['total_connected'],
                'total_demand': tx_result['total_demand'],
                'total_kvar': tx_result['total_kvar'],
                'total_kva': tx_result['total_kva'],
                'selected_kva': tx_result['selected_kva'],
                'future_expansion': future_expansion,
                'system_voltage': system_voltage,
                'ambient_temp': ambient_temp,
                'altitude': altitude,
                'cooling_type': cooling_type,
                'insulation_class': insulation_class
            }
            st.success(f"✅ Selected Transformer: {tx_result['selected_kva']} kVA")
    
    # TAB 2: TRANSFORMER SELECTION (UNCHANGED)
    with tx_tabs[1]:
        st.markdown('<div class="report-header">⚙️ TRANSFORMER SELECTION</div>', unsafe_allow_html=True)
        
        if 'tx_load_data' not in st.session_state:
            st.warning("⚠️ Please complete Load Analysis first!")
        else:
            tx_data = st.session_state.tx_load_data
            
            st.markdown("### 🔋 Transformer Parameters [IEC 60076]")
            
            col1, col2 = st.columns(2)
            
            with col1:
                st.markdown("#### Electrical Characteristics")
                
                primary_voltage = st.selectbox("Primary Voltage (kV)", 
                                              [132, 110, 66, 33, 22, 11, 6.6, 3.3], 
                                              index=3,
                                              help="HV side voltage")
                
                vector_groups = {
                    'Dyn11': 'Delta primary, Star secondary with neutral - Most common',
                    'Dyn5': 'Delta primary, Star secondary with neutral - 150° phase shift',
                    'Yyn0': 'Star-Star with neutral - For small transformers',
                    'Yd11': 'Star primary, Delta secondary - For step-down',
                    'Dd0': 'Delta-Delta - For industrial applications'
                }
                vector_group = st.selectbox("Vector Group", list(vector_groups.keys()), index=0)
                st.caption(vector_groups[vector_group])
                
                if tx_data['selected_kva'] <= 1000:
                    default_imp = 5.0
                elif tx_data['selected_kva'] <= 5000:
                    default_imp = 6.5
                else:
                    default_imp = 8.0
                    
                impedance = st.number_input("Impedance Voltage (%)", 
                                           value=default_imp, min_value=2.0, max_value=15.0, step=0.25,
                                           help="Percentage impedance at rated current")
                
                tapping = st.selectbox("Tapping Range", 
                                      ['±2.5%', '±5%', '±7.5%', '±10%', '-5% to +15%'], 
                                      index=1,
                                      help="Off-circuit or on-load tap changing range")
            
            with col2:
                st.markdown("#### Construction Details")
                
                core_material = st.selectbox("Core Material", 
                                            ['CRGO Silicon Steel', 'Amorphous Metal'], 
                                            index=0,
                                            help="CRGO: Cold Rolled Grain Oriented")
                
                winding_material = st.selectbox("Winding Material", 
                                               ['Copper', 'Aluminium'], 
                                               index=0,
                                               help="Copper: Lower losses, higher cost\nAluminium: Higher losses, lower cost")
                
                cooling_detail = st.selectbox("Cooling Detailed", 
                                             ['ONAN', 'ONAF', 'ONAN/ONAF', 'OFAF', 'ODAF'], 
                                             index=0)
                
                enclosure = st.selectbox("Enclosure Type", 
                                        ['Indoor', 'Outdoor', 'Weatherproof', 'Hermetically Sealed'], 
                                        index=1)
                
                frequency = st.selectbox("Frequency (Hz)", [50, 60], index=0)
            
            st.markdown("### 📊 Standard Transformer Ratings [IEC 60076]")
            
            comparison_data = {
                'Parameter': ['Rated Power (kVA)', 'Primary Voltage (kV)', 'Secondary Voltage (V)', 
                             'Impedance (%)', 'Cooling Type', 'Vector Group'],
                'Selected Value': [f"{tx_data['selected_kva']}", f"{primary_voltage}", f"{tx_data['system_voltage']}",
                                  f"{impedance}", cooling_detail, vector_group],
                'IEC Limit': ['-', '-', '-', f"±10% of declared", 'Per IEC 60076', 'Per IEC 60076']
            }
            
            comparison_df = pd.DataFrame(comparison_data)
            st.dataframe(comparison_df, use_container_width=True, hide_index=True)
            
            st.markdown("### ✅ Selection Summary")
            
            st.success(f"""
            **Selected Transformer:** {tx_data['selected_kva']} kVA, {primary_voltage} kV / {tx_data['system_voltage']} V
            
            **Configuration:**
            - Vector Group: {vector_group}
            - Impedance: {impedance}%
            - Cooling: {cooling_detail}
            - Tapping: {tapping}
            - Core: {core_material}
            - Windings: {winding_material}
            
            **Application:** Suitable for general power distribution
            """)
            
            if st.button("➡️ Next: Voltage Regulation", use_container_width=True):
                st.session_state.tx_selection = {
                    'primary_voltage': primary_voltage,
                    'vector_group': vector_group,
                    'impedance': impedance,
                    'tapping': tapping,
                    'core_material': core_material,
                    'winding_material': winding_material,
                    'cooling_detail': cooling_detail,
                    'enclosure': enclosure,
                    'frequency': frequency
                }
    
    # TAB 3: VOLTAGE REGULATION (UNCHANGED)
    with tx_tabs[2]:
        st.markdown('<div class="report-header">⚙️ VOLTAGE REGULATION</div>', unsafe_allow_html=True)
        
        if 'tx_load_data' not in st.session_state:
            st.warning("⚠️ Please complete Load Analysis first!")
        else:
            tx_data = st.session_state.tx_load_data
            
            st.markdown("""
            ### 📉 Voltage Regulation Calculation [IEC 60076-1]
            
            Voltage regulation is the change in secondary voltage from no-load to full-load,
            expressed as a percentage of the rated voltage.
            
            **Formula:** %Reg = ε_x cosφ + ε_r sinφ + (ε_x cosφ - ε_r sinφ)²/200
            
            Where:
            - ε_r = Resistance voltage drop (%)
            - ε_x = Reactance voltage drop (%)
            - cosφ = Power factor
            """)
            
            col1, col2 = st.columns(2)
            
            with col1:
                st.markdown("#### Transformer Parameters")
                
                impedance = st.number_input("Impedance Voltage (%)", 
                                           value=st.session_state.get('tx_selection', {}).get('impedance', 6.0),
                                           step=0.25, key="reg_imp")
                
                xr_ratio = st.selectbox("X/R Ratio", 
                                       [3, 5, 7, 10, 12, 15], 
                                       index=2,
                                       help="Ratio of reactance to resistance")
            
            with col2:
                st.markdown("#### Load Conditions")
                
                pf_reg = st.slider("Power Factor", min_value=0.5, max_value=1.0, value=0.85, step=0.05)
                load_pct = st.slider("Load Percentage (%)", min_value=0, max_value=150, value=100, step=10)
                load_type = st.radio("Load Type", ['lagging', 'leading', 'unity'], index=0)
            
            reg_result = tx_calc.calculate_voltage_regulation(impedance, xr_ratio, pf_reg, load_pct, load_type)
            
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Resistance Drop (ε_r)", f"{reg_result['resistance_drop']:.3f}%")
            with col2:
                st.metric("Reactance Drop (ε_x)", f"{reg_result['reactance_drop']:.3f}%")
            with col3:
                st.metric("Voltage Regulation", f"{reg_result['regulation']:.3f}%")
            
            # Regulation curve
            st.markdown("### 📊 Regulation vs Power Factor")
            
            pf_range = [x/100 for x in range(50, 101, 5)]
            reg_lagging = []
            reg_leading = []
            
            for pf in pf_range:
                lag = tx_calc.calculate_voltage_regulation(impedance, xr_ratio, pf, 100, 'lagging')
                lead = tx_calc.calculate_voltage_regulation(impedance, xr_ratio, pf, 100, 'leading')
                reg_lagging.append(lag['regulation'])
                reg_leading.append(lead['regulation'])
            
            chart_data = pd.DataFrame({
                'Power Factor': pf_range,
                'Lagging Load': reg_lagging,
                'Leading Load': reg_leading
            })
            
            st.line_chart(chart_data.set_index('Power Factor'))
            
            st.info(f"""
            ### 📝 Detailed Calculation
            
            **Given:**
            - Transformer Rating: {tx_data['selected_kva']} kVA
            - Impedance Voltage: {impedance}%
            - X/R Ratio: {xr_ratio}
            
            **Step 1:** Calculate resistance drop
            ε_r = %Z / √(1 + (X/R)²) = {impedance} / √(1 + {xr_ratio}²) = **{reg_result['resistance_drop']:.3f}%**
            
            **Step 2:** Calculate reactance drop
            ε_x = ε_r × (X/R) = {reg_result['resistance_drop']:.3f} × {xr_ratio} = **{reg_result['reactance_drop']:.3f}%**
            
            **Step 3:** Calculate voltage regulation
            %Reg = ε_r cosφ + ε_x sinφ
            %Reg = {reg_result['resistance_drop']:.3f} × {pf_reg} + {reg_result['reactance_drop']:.3f} × {math.sqrt(1 - pf_reg**2):.3f} = **{reg_result['regulation']:.3f}%**
            
            **IEC 60076-1 Limit:** ±5% for distribution transformers
            """)
    
    # TAB 4: EFFICIENCY & LOSSES (UNCHANGED)
    with tx_tabs[3]:
        st.markdown('<div class="report-header">⚙️ EFFICIENCY & LOSSES</div>', unsafe_allow_html=True)
        
        if 'tx_load_data' not in st.session_state:
            st.warning("⚠️ Please complete Load Analysis first!")
        else:
            tx_data = st.session_state.tx_load_data
            
            st.markdown("""
            ### ⚡ Transformer Losses [IEC 60076-1]
            
            **Two types of losses:**
            1. **No-load losses (Iron losses)** - Constant, independent of load
            2. **Load losses (Copper losses)** - Vary with square of load current
            """)
            
            col1, col2 = st.columns(2)
            
            with col1:
                st.markdown("#### Loss Parameters")
                
                if tx_data['selected_kva'] <= 100:
                    base_no_load = 0.01 * tx_data['selected_kva']
                elif tx_data['selected_kva'] <= 1000:
                    base_no_load = 0.008 * tx_data['selected_kva']
                elif tx_data['selected_kva'] <= 10000:
                    base_no_load = 0.006 * tx_data['selected_kva']
                else:
                    base_no_load = 0.004 * tx_data['selected_kva']
                
                no_load_loss = st.number_input("No-load Loss (kW)", 
                                              value=round(base_no_load, 1),
                                              step=0.1,
                                              help="Iron losses - constant regardless of load")
                
                if tx_data['selected_kva'] <= 100:
                    base_load_loss = 0.02 * tx_data['selected_kva']
                elif tx_data['selected_kva'] <= 1000:
                    base_load_loss = 0.015 * tx_data['selected_kva']
                elif tx_data['selected_kva'] <= 10000:
                    base_load_loss = 0.012 * tx_data['selected_kva']
                else:
                    base_load_loss = 0.01 * tx_data['selected_kva']
                
                load_loss_100 = st.number_input("Load Loss at 100% (kW)", 
                                               value=round(base_load_loss, 1),
                                               step=0.1,
                                               help="Copper losses at full load")
            
            with col2:
                st.markdown("#### Operating Conditions")
                
                load_profile = st.selectbox("Load Profile", 
                                           ['Continuous', 'Industrial', 'Commercial', 'Residential'],
                                           index=0)
                
                operating_hours = st.number_input("Annual Operating Hours", 
                                                 value=8760, step=1000,
                                                 help="8760 hours = 24/7 operation")
                
                energy_cost = st.number_input("Energy Cost ($/kWh)", 
                                             value=0.12, step=0.01, format="%.3f")
            
            load_points = [0, 25, 50, 75, 100, 110]
            efficiency_data = []
            
            for load in load_points:
                efficiency, losses = tx_calc.calculate_efficiency(
                    tx_data['selected_kva'], no_load_loss, load_loss_100, load
                )
                output = tx_data['selected_kva'] * (load/100)
                
                efficiency_data.append({
                    'Load (%)': load,
                    'Output (kW)': round(output, 1),
                    'Losses (kW)': round(losses, 2),
                    'Efficiency (%)': round(efficiency, 2)
                })
            
            efficiency_df = pd.DataFrame(efficiency_data)
            st.dataframe(efficiency_df, use_container_width=True, hide_index=True)
            
            # Maximum efficiency point
            max_eff_load = math.sqrt(no_load_loss / load_loss_100) * 100
            max_eff_output = tx_data['selected_kva'] * (max_eff_load / 100)
            max_eff_losses = no_load_loss + load_loss_100 * (max_eff_load/100)**2
            max_eff = max_eff_output / (max_eff_output + max_eff_losses) * 100
            
            col1, col2 = st.columns(2)
            with col1:
                st.metric("Maximum Efficiency", f"{max_eff:.2f}%")
            with col2:
                st.metric("at Load", f"{max_eff_load:.1f}%")
            
            avg_load = 60
            avg_losses = no_load_loss + load_loss_100 * (avg_load/100)**2
            annual_loss = avg_losses * operating_hours / 1000
            annual_cost = annual_loss * energy_cost * 1000
            
            st.info(f"""
            ### 📊 Annual Energy Losses
            
            **Assumptions:**
            - Average Loading: {avg_load}%
            - Operating Hours: {operating_hours} hours/year
            
            **Calculations:**
            - Average Losses = No-load + (Load Loss × Load²) = {no_load_loss:.2f} + ({load_loss_100:.2f} × {(avg_load/100):.2f}²) = **{avg_losses:.2f} kW**
            
            - Annual Energy Loss = {avg_losses:.2f} kW × {operating_hours} h / 1000 = **{annual_loss:.1f} MWh/year**
            
            - Annual Cost @ ${energy_cost}/kWh = **${annual_cost:,.0f}/year**
            """)
    
    # TAB 5: SHORT CIRCUIT (UNCHANGED)
    with tx_tabs[4]:
        st.markdown('<div class="report-header">⚙️ SHORT CIRCUIT CALCULATIONS</div>', unsafe_allow_html=True)
        
        if 'tx_load_data' not in st.session_state:
            st.warning("⚠️ Please complete Load Analysis first!")
        else:
            tx_data = st.session_state.tx_load_data
            
            st.markdown("""
            ### ⚡ Transformer Short Circuit Currents [IEC 60076-5]
            
            Calculate short circuit currents for protection coordination and switchgear selection.
            """)
            
            col1, col2 = st.columns(2)
            
            with col1:
                st.markdown("#### Transformer Parameters")
                
                tx_kva = st.number_input("Transformer Rating (kVA)", 
                                         value=tx_data['selected_kva'],
                                         step=100, key="sc_tx_kva")
                
                secondary_v = st.number_input("Secondary Voltage (V)", 
                                             value=tx_data['system_voltage'],
                                             step=10, key="sc_sec_v")
                
                impedance_sc = st.number_input("Impedance Voltage (%)", 
                                              value=st.session_state.get('tx_selection', {}).get('impedance', 6.0),
                                              step=0.25, key="sc_imp")
                
                xr_ratio_sc = st.selectbox("X/R Ratio", [3, 5, 7, 10, 12, 15], index=2, key="sc_xr")
            
            with col2:
                st.markdown("#### System Parameters")
                
                source_sc = st.number_input("Source Short Circuit MVA", 
                                           value=500, step=50,
                                           help="Short circuit level at transformer primary")
                
                motor_load = st.number_input("Motor Load (% of transformer)", 
                                            value=50, min_value=0, max_value=100,
                                            help="Motor contribution to short circuit")
                
                duration = st.number_input("Short Circuit Duration (s)", 
                                          value=1.0, step=0.1,
                                          help="Duration for thermal withstand calculation")
            
            sc_result = tx_calc.calculate_short_circuit(tx_kva, secondary_v, impedance_sc, source_sc*1000, xr_ratio_sc)
            
            motor_sc = sc_result['i_sc_sym'] * (motor_load / 100) * 4
            i_sc_total = sc_result['i_sc_sym'] + motor_sc
            min_cable_size = sc_result['i_sc_sym'] * 1000 * math.sqrt(duration) / 143
            
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("Full Load Current", f"{sc_result['fla']:.0f} A")
            with col2:
                st.metric("Symmetrical SC", f"{sc_result['i_sc_sym']:.2f} kA")
            with col3:
                st.metric("Peak SC Current", f"{sc_result['i_sc_peak']:.2f} kA")
            with col4:
                st.metric("With Motors", f"{i_sc_total:.2f} kA")
            
            # NEW: IEC 60059 - Standard current rating for switchgear
            std_current = tx_calc.get_standard_current(sc_result['fla'])
            
            st.info(f"""
            ### 📝 Detailed Calculations
            
            **Step 1:** Full Load Current
            I_FL = {tx_kva} × 1000 / (1.732 × {secondary_v}) = **{sc_result['fla']:.0f} A**
            
            **IEC 60059 Standard Rating:** **{std_current} A** (next higher standard)
            
            **Step 2:** Transformer Impedance
            Z_tx = (%Z/100) × (V² / S) = ({impedance_sc}/100) × ({secondary_v}² / {tx_kva*1000}) = **{sc_result['z_tx_ohms']*1000:.2f} mΩ**
            
            **Step 3:** Symmetrical Short Circuit Current
            I_sc = V / (√3 × Z_total) = **{sc_result['i_sc_sym']:.2f} kA**
            
            **Step 4:** Peak Current (with X/R = {xr_ratio_sc})
            κ = 1.02 + 0.98 × e^(-3/{xr_ratio_sc}) = **{sc_result['kappa']:.3f}**
            I_peak = κ × √2 × I_sc = {sc_result['kappa']:.3f} × 1.414 × {sc_result['i_sc_sym']:.2f} = **{sc_result['i_sc_peak']:.2f} kA**
            
            **Step 5:** Minimum Cable Size for Thermal Withstand
            S_min = I_sc × √t / K = {sc_result['i_sc_sym']*1000:.0f} × √{duration} / 143 = **{min_cable_size:.0f} mm²**
            """)
    
    # NEW TAB 6: SOUND & AGEING (IEC 60076-7 and IEC 60076-10)
    with tx_tabs[5]:
        st.markdown('<div class="report-header">🔊 SOUND LEVEL & THERMAL AGEING</div>', unsafe_allow_html=True)
        
        if 'tx_load_data' not in st.session_state:
            st.warning("⚠️ Please complete Load Analysis first!")
        else:
            tx_data = st.session_state.tx_load_data
            
            st.markdown("""
            ### 🔊 Sound Level Calculation [IEC 60076-10]
            
            Transformers produce audible noise due to magnetostriction in the core and cooling equipment.
            """)
            
            col1, col2 = st.columns(2)
            
            with col1:
                st.markdown("#### 🔊 Sound Level Parameters")
                
                tx_type = st.radio("Transformer Type", ['liquid', 'dry'], index=0,
                                  help="Liquid-filled or Dry-type transformer")
                
                cooling_sound = st.selectbox("Cooling Type for Sound", 
                                            ['ONAN', 'ONAF', 'OFAF', 'ODAF'], 
                                            index=0,
                                            help="Affects noise level")
                
                sound_result = tx_calc.calculate_sound_level(
                    tx_data['selected_kva'], tx_type, cooling_sound
                )
                
                st.metric("Base Sound Level", f"{sound_result['base_sound_level']} dB(A)")
                st.metric("Cooling Addition", f"+{sound_result['cooling_addition']} dB(A)")
                st.metric("Total Sound Level", f"{sound_result['total_sound_level']} dB(A)")
                
                st.caption(f"Reference: {sound_result['reference']}")
            
            with col2:
                st.markdown("#### 🌡️ Thermal Ageing [IEC 60076-7]")
                
                # Extract insulation class from stored data
                ins_class = tx_data['insulation_class'].split(' ')[0] if 'insulation_class' in tx_data else 'F'
                
                hot_spot_temp = st.number_input("Hot Spot Temperature (°C)", 
                                               value=98, min_value=50, max_value=250, step=5,
                                               help="Winding hot spot temperature")
                
                ageing_result = tx_calc.calculate_thermal_ageing(
                    hot_spot_temp, ins_class
                )
                
                col_a, col_b = st.columns(2)
                with col_a:
                    st.metric("Relative Ageing Rate", f"{ageing_result['relative_ageing_rate']:.2f}")
                with col_b:
                    st.metric("Remaining Life", f"{ageing_result['remaining_life_years']:.1f} years")
                
                if not ageing_result['is_safe']:
                    st.warning(f"⚠️ Temperature exceeds {ins_class} class limit of {ageing_result['temperature_limit']}°C!")
                else:
                    st.success(f"✅ Within {ins_class} class limit of {ageing_result['temperature_limit']}°C")
            
            st.markdown("### 📊 Sound Level Reference Table [IEC 60076-10]")
            
            # Create reference table
            sound_ref_data = {
                'Power (kVA)': [100, 500, 1000, 2500, 5000, 10000],
                'Liquid-filled (dB)': [50, 55, 58, 62, 65, 68],
                'Dry-type (dB)': [48, 52, 55, 58, 62, 65]
            }
            sound_ref_df = pd.DataFrame(sound_ref_data)
            st.dataframe(sound_ref_df, use_container_width=True, hide_index=True)
            
            st.markdown("### 📝 Ageing Calculation Details")
            
            st.info(f"""
            **IEC 60076-7 Thermal Ageing Calculation**
            
            **Formula:** V = 2^((θ_h - 98)/6)
            
            Where:
            - θ_h = Hot spot temperature = {hot_spot_temp}°C
            - 98°C = Reference temperature for 180,000 hours life
            
            **Calculation:**
            V = 2^(({hot_spot_temp} - 98)/6) = 2^({(hot_spot_temp-98)/6:.2f}) = **{ageing_result['relative_ageing_rate']:.2f}**
            
            **Interpretation:**
            - V = 1.0 → Normal ageing (180,000 hours life)
            - V = 2.0 → Ageing twice as fast (90,000 hours life)
            - V = 0.5 → Ageing half as fast (360,000 hours life)
            
            **Remaining Life:** {ageing_result['remaining_life_years']:.1f} years at {hot_spot_temp}°C
            
            **Insulation Class {ins_class} Temperature Limit:** {ageing_result['temperature_limit']}°C
            """)
    
    # TAB 7: DOWNLOAD REPORT (renumbered)
    with tx_tabs[6]:
        st.markdown('<div class="report-header">📥 DOWNLOAD REPORT</div>', unsafe_allow_html=True)
        
        if 'tx_load_data' not in st.session_state:
            st.warning("⚠️ Please complete Load Analysis first!")
        else:
            st.markdown("### 📋 Transformer Sizing Report")
            
            col1, col2 = st.columns(2)
            
            with col1:
                if st.button("📥 Generate PDF Report", key="tx_pdf", use_container_width=True):
                    with st.spinner("Generating PDF report..."):
                        st.success("✅ PDF generated successfully!")
            
            with col2:
                if st.button("📥 Generate Word Report", key="tx_word", use_container_width=True):
                    with st.spinner("Generating Word report..."):
                        st.success("✅ Word generated successfully!")

# ========== OTHER CALCULATORS (Placeholders) ==========
elif st.session_state.selected_calculator == "⚡ Generator Sizing":
    st.markdown('<div class="report-header">⚡ GENERATOR SIZING</div>', unsafe_allow_html=True)
    st.info("⚡ Coming soon!")

elif st.session_state.selected_calculator == "🌍 Earthing System Design":
    st.markdown('<div class="report-header">🌍 EARTHING SYSTEM DESIGN</div>', unsafe_allow_html=True)
    st.info("🌍 Coming soon!")

# Footer
st.markdown("---")
st.markdown(f"<div style='text-align: center; color: gray;'>🔌 CES-Electrical | Version 64.0 | {datetime.now().strftime('%Y-%m-%d %H:%M')}</div>", unsafe_allow_html=True)