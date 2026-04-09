import streamlit as st
import math
import datetime
import pandas as pd
import base64
from datetime import datetime
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
import os

st.set_page_config(page_title="Professional Engineering Tools", page_icon="🔌", layout="wide")

# ========== HELPER FUNCTIONS FOR FORMATTING ==========
def format_cable_arrangement(arrangement):
    """Convert internal cable arrangement names to nice display format"""
    formats = {
        'bunched_in_air_surface_enclosed': 'Bunched in air surface enclosed',
        'single_layer_wall_floor': 'Single layer wall/floor',
        'single_layer_perforated_tray': 'Single layer perforated tray',
        'single_layer_ladder_cleats': 'Single layer ladder cleats'
    }
    return formats.get(arrangement, arrangement.replace('_', ' ').title())

def format_cable_formation(formation):
    """Convert internal cable formation names to nice display format"""
    formats = {
        'flat': 'Flat',
        'trefoil': 'Trefoil',
        'spaced': 'Spaced'
    }
    return formats.get(formation, formation.title())

def format_cable_type(cable_type):
    """Convert internal cable type names to nice display format"""
    formats = {
        'single_core_non_armoured': 'Single core non-armoured',
        'multi_core_non_armoured': 'Multi core non-armoured',
        'single_core_armoured': 'Single core armoured',
        'multi_core_armoured': 'Multi core armoured'
    }
    return formats.get(cable_type, cable_type.replace('_', ' ').title())

def format_insulation_type(insulation_type):
    """Convert internal insulation type names to nice display format"""
    formats = {
        'XLPE_90': 'XLPE 90°C',
        'PVC_70': 'PVC 70°C'
    }
    return formats.get(insulation_type, insulation_type.replace('_', ' '))

def format_load_type(load_type):
    """Format load type with proper capitalization"""
    return load_type.capitalize()

# ========== CUSTOM CSS ==========
st.markdown("""
<style>
    :root {
        --primary: #1E3A8A;
        --secondary: #00A86B;
        --light-bg: #F0F8FF;
        --card-bg: #FFFFFF;
        --text-dark: #1a1a1a;
        --text-muted: #4a4a4a;
    }
    .report-header {
        background: linear-gradient(135deg, var(--primary) 0%, #3B5BA6 100%);
        color: white;
        padding: 25px;
        border-radius: 12px;
        text-align: center;
        margin-bottom: 25px;
        font-size: 32px !important;
        font-weight: bold;
    }
    .info-box {
        background: linear-gradient(135deg, #E8F4FD 0%, #D6ECFA 100%);
        color: var(--text-dark) !important;
        padding: 15px;
        border-radius: 8px;
        border-left: 5px solid var(--primary);
        margin: 10px 0;
        font-size: 18px !important;
    }
    .info-box * {
        color: var(--text-dark) !important;
    }
    .info-box h4 {
        color: var(--primary) !important;
    }
    .calc-step {
        background-color: #F8F9FA !important;
        padding: 15px;
        border-radius: 8px;
        border-left: 5px solid #00A86B;
        margin: 10px 0;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        font-size: 18px !important;
        color: var(--text-dark) !important;
    }
    .calc-step * {
        color: var(--text-dark) !important;
    }
    .calc-step h4 {
        color: var(--primary) !important;
        margin-top: 0;
        margin-bottom: 10px;
    }
    .calc-step p {
        color: var(--text-dark) !important;
        margin: 5px 0;
    }
    .calc-step b {
        color: var(--primary) !important;
    }
    .result-card {
        background: linear-gradient(135deg, #1E3A8A 0%, #3B5BA6 100%);
        color: white !important;
        padding: 25px;
        border-radius: 12px;
        margin: 20px 0;
    }
    .download-btn {
        display: inline-block;
        padding: 14px 28px;
        margin: 10px;
        color: white !important;
        text-decoration: none;
        border-radius: 8px;
        font-size: 18px !important;
        font-weight: bold;
        text-align: center;
        background: linear-gradient(135deg, #1E3A8A 0%, #3B5BA6 100%);
    }
    .load-type-badge {
        display: inline-block;
        padding: 5px 12px;
        border-radius: 12px;
        font-size: 14px !important;
        font-weight: bold;
    }
    .continuous-badge { 
        background-color: #00A86B; 
        color: white !important; 
    }
    .intermittent-badge { 
        background-color: #FFC107; 
        color: #1a1a1a !important; 
    }
    .standby-badge { 
        background-color: #DC3545; 
        color: white !important; 
    }
    .stTabs [data-baseweb="tab"] {
        font-size: 18px !important;
        font-weight: 600 !important;
    }
    .stDataFrame {
        color: var(--text-dark) !important;
    }
    .stDataFrame table {
        color: var(--text-dark) !important;
    }
    .stDataFrame th {
        background: linear-gradient(135deg, var(--primary) 0%, #3B5BA6 100%) !important;
        color: white !important;
    }
    .stDataFrame td {
        color: var(--text-dark) !important;
        background-color: white !important;
    }
    div[data-testid="stMetricValue"] {
        color: var(--primary) !important;
        font-size: 28px !important;
    }
    div[data-testid="stMetricLabel"] {
        color: var(--text-muted) !important;
    }
    .largest-equipment {
        background: linear-gradient(135deg, #E8F5E9 0%, #C8E6C9 100%);
        padding: 20px;
        border-radius: 10px;
        border-left: 6px solid #00A86B;
        margin: 15px 0;
        color: var(--text-dark) !important;
    }
    .largest-equipment * {
        color: var(--text-dark) !important;
    }
    .largest-equipment h3 {
        color: #006B3C !important;
    }
    .largest-equipment .value {
        font-size: 20px !important;
        font-weight: bold;
        color: #006B3C !important;
    }
    .formula-box {
        background: linear-gradient(135deg, #F8F9FA 0%, #E9ECEF 100%);
        padding: 20px;
        border-radius: 10px;
        border-left: 6px solid var(--secondary);
        margin: 15px 0;
        color: var(--text-dark) !important;
    }
    .formula-box * {
        color: var(--text-dark) !important;
    }
    .formula-box h4 {
        color: var(--primary) !important;
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
    'Schneider Electric': {'MCB': 'Acti9 series', 'MCCB': 'EasyPact EVC series', 'ACB': 'MasterPact MTZ series'},
    'Siemens': {'MCB': '5SY series', 'MCCB': '3VA series', 'ACB': '3WL series'},
    'ABB': {'MCB': 'S200 series', 'MCCB': 'Tmax XT series', 'ACB': 'Emax 2 series'}
}

# ========== LOAD TYPE FACTORS ==========
LOAD_TYPE_FACTORS = {
    'Continuous': {'diversity': 1.0, 'description': 'Continuous (100%) - Full time operation', 'cb_factor': 1.25, 'color': '#00A86B'},
    'Intermittent': {'diversity': 0.3, 'description': 'Intermittent (30%) - Cyclic operation', 'cb_factor': 1.25, 'color': '#FFC107'},
    'Standby': {'diversity': 0.1, 'description': 'Stand-by (10%) - Emergency/backup only', 'cb_factor': 1.25, 'color': '#DC3545'}
}

# ========== TEMPERATURE DERATING FACTORS (Table 4B1 & 4B2 from Sheet1) ==========
TEMPERATURE_FACTORS_AIR = {
    70: {25: 1.03, 30: 1.00, 35: 0.94, 40: 0.87, 45: 0.79, 50: 0.71, 55: 0.61},
    90: {25: 1.02, 30: 1.00, 35: 0.96, 40: 0.91, 45: 0.87, 50: 0.82, 55: 0.76}
}

TEMPERATURE_FACTORS_GROUND = {
    70: {10: 1.10, 15: 1.05, 20: 1.00, 25: 0.95, 30: 0.89, 35: 0.84, 40: 0.77, 45: 0.71},
    90: {10: 1.07, 15: 1.04, 20: 1.00, 25: 0.96, 30: 0.93, 35: 0.89, 40: 0.85, 45: 0.80}
}

# ========== GROUPING FACTORS (Table 4C1 from Sheet1) ==========
GROUPING_FACTORS = {
    'bunched_in_air_surface_enclosed': {1: 1.00, 2: 0.80, 3: 0.70, 4: 0.65, 5: 0.60, 6: 0.57, 7: 0.54, 8: 0.52, 9: 0.50, 12: 0.45, 16: 0.41, 20: 0.38},
    'single_layer_wall_floor': {1: 1.00, 2: 0.85, 3: 0.79, 4: 0.75, 5: 0.73, 6: 0.72, 7: 0.72, 8: 0.71, 9: 0.70, 12: 0.70, 16: 0.70, 20: 0.70},
    'single_layer_perforated_tray': {1: 1.00, 2: 0.88, 3: 0.82, 4: 0.77, 5: 0.75, 6: 0.73, 7: 0.73, 8: 0.72, 9: 0.72, 12: 0.72, 16: 0.72, 20: 0.72},
    'single_layer_ladder_cleats': {1: 1.00, 2: 0.87, 3: 0.82, 4: 0.80, 5: 0.80, 6: 0.79, 7: 0.79, 8: 0.78, 9: 0.78, 12: 0.78, 16: 0.78, 20: 0.78}
}

# ========== SOIL RESISTIVITY FACTORS ==========
SOIL_RESISTIVITY_FACTORS = {0.7: 1.28, 0.8: 1.24, 0.9: 1.19, 1.0: 1.15, 1.2: 1.10, 1.5: 1.00, 2.0: 0.89, 2.5: 0.81, 3.0: 0.75}

# ========== DEPTH FACTORS ==========
DEPTH_FACTORS = {0.5: 1.04, 0.6: 1.02, 0.7: 1.01, 0.8: 1.00, 0.9: 0.99, 1.0: 0.98, 1.25: 0.96, 1.5: 0.95, 1.75: 0.94, 2.0: 0.93}

# ========== CORRECTED 90°C XLPE SINGLE CORE NON-ARMOURED (From Sheet3) ==========
XLPE_90_SINGLE_NON_ARMOURED = {
    # Small cables 1-16mm² - use direct voltage drop values from Sheet3
    1.0: {'ampacity': {'B2': 17.0, 'B34': 15.0, 'C2': 19.0, 'C34': 17.5}, 
          'voltage_drop': {'dc': 46.0, 'ac': 46.0, 'ac34': 40.0}, 'diameter': 8},
    1.5: {'ampacity': {'B2': 23.0, 'B34': 20.0, 'C2': 25.0, 'C34': 23.0}, 
          'voltage_drop': {'dc': 31.0, 'ac': 31.0, 'ac34': 27.0}, 'diameter': 9},
    2.5: {'ampacity': {'B2': 31.0, 'B34': 28.0, 'C2': 34.0, 'C34': 31.0}, 
          'voltage_drop': {'dc': 19.0, 'ac': 19.0, 'ac34': 16.0}, 'diameter': 10},
    4.0: {'ampacity': {'B2': 42.0, 'B34': 37.0, 'C2': 46.0, 'C34': 41.0}, 
          'voltage_drop': {'dc': 12.0, 'ac': 12.0, 'ac34': 10.0}, 'diameter': 11},
    6.0: {'ampacity': {'B2': 54.0, 'B34': 48.0, 'C2': 59.0, 'C34': 54.0}, 
          'voltage_drop': {'dc': 7.9, 'ac': 7.9, 'ac34': 6.8}, 'diameter': 12},
    10.0: {'ampacity': {'B2': 75.0, 'B34': 66.0, 'C2': 81.0, 'C34': 74.0}, 
           'voltage_drop': {'dc': 4.7, 'ac': 4.7, 'ac34': 4.0}, 'diameter': 14},
    16.0: {'ampacity': {'B2': 100.0, 'B34': 88.0, 'C2': 109.0, 'C34': 99.0}, 
           'voltage_drop': {'dc': 2.9, 'ac': 2.9, 'ac34': 2.5}, 'diameter': 16},
    
    # Cables 25mm² and above - use R and X values from Sheet3
    25.0: {'ampacity': {'B2': 133.0, 'B34': 117.0, 'C2': 143.0, 'C34': 130.0, 
                        'F2_flat': 161.0, 'F34_flat': 141.0, 'F34_trefoil': 135.0, 
                        'G2': 182.0, 'G34': 161.0}, 
           'R': 1.85, 
           'X_trefoil': 0.19,
           'X_flat_touching': 0.28,
           'X_spaced': 0.31,
           'diameter': 18},
    
    35.0: {'ampacity': {'B2': 164.0, 'B34': 144.0, 'C2': 176.0, 'C34': 161.0,
                        'F2_flat': 200.0, 'F34_flat': 176.0, 'F34_trefoil': 169.0,
                        'G2': 226.0, 'G34': 201.0},
           'R': 1.35,
           'X_trefoil': 0.18,
           'X_flat_touching': 0.27,
           'X_spaced': 0.29,
           'diameter': 20},
    
    50.0: {'ampacity': {'B2': 198.0, 'B34': 175.0, 'C2': 228.0, 'C34': 209.0,
                        'F2_flat': 242.0, 'F34_flat': 216.0, 'F34_trefoil': 207.0,
                        'G2': 275.0, 'G34': 246.0},
           'R': 0.99,
           'X_trefoil': 0.18,
           'X_flat_touching': 0.27,
           'X_spaced': 0.29,
           'diameter': 22},
    
    70.0: {'ampacity': {'B2': 253.0, 'B34': 222.0, 'C2': 293.0, 'C34': 268.0,
                        'F2_flat': 310.0, 'F34_flat': 279.0, 'F34_trefoil': 268.0,
                        'G2': 353.0, 'G34': 318.0},
           'R': 0.68,
           'X_trefoil': 0.175,
           'X_flat_touching': 0.26,
           'X_spaced': 0.28,
           'diameter': 25},
    
    95.0: {'ampacity': {'B2': 306.0, 'B34': 269.0, 'C2': 355.0, 'C34': 326.0,
                        'F2_flat': 377.0, 'F34_flat': 342.0, 'F34_trefoil': 328.0,
                        'G2': 430.0, 'G34': 389.0},
           'R': 0.49,
           'X_trefoil': 0.17,
           'X_flat_touching': 0.26,
           'X_spaced': 0.27,
           'diameter': 28},
    
    120.0: {'ampacity': {'B2': 354.0, 'B34': 312.0, 'C2': 413.0, 'C34': 379.0,
                         'F2_flat': 437.0, 'F34_flat': 400.0, 'F34_trefoil': 383.0,
                         'G2': 500.0, 'G34': 454.0},
            'R': 0.39,
            'X_trefoil': 0.165,
            'X_flat_touching': 0.25,
            'X_spaced': 0.26,
            'diameter': 30},
    
    150.0: {'ampacity': {'B2': 393.0, 'B34': 342.0, 'C2': 476.0, 'C34': 436.0,
                         'F2_flat': 504.0, 'F34_flat': 464.0, 'F34_trefoil': 444.0,
                         'G2': 577.0, 'G34': 527.0},
            'R': 0.32,
            'X_trefoil': 0.165,
            'X_flat_touching': 0.25,
            'X_spaced': 0.26,
            'diameter': 32},
    
    185.0: {'ampacity': {'B2': 449.0, 'B34': 384.0, 'C2': 545.0, 'C34': 500.0,
                         'F2_flat': 575.0, 'F34_flat': 533.0, 'F34_trefoil': 510.0,
                         'G2': 661.0, 'G34': 605.0},
            'R': 0.25,
            'X_trefoil': 0.165,
            'X_flat_touching': 0.25,
            'X_spaced': 0.26,
            'diameter': 35},
    
    240.0: {'ampacity': {'B2': 528.0, 'B34': 450.0, 'C2': 644.0, 'C34': 590.0,
                         'F2_flat': 679.0, 'F34_flat': 634.0, 'F34_trefoil': 607.0,
                         'G2': 781.0, 'G34': 719.0},
            'R': 0.19,
            'X_trefoil': 0.16,
            'X_flat_touching': 0.25,
            'X_spaced': 0.26,
            'diameter': 38},
    
    300.0: {'ampacity': {'B2': 603.0, 'B34': 514.0, 'C2': 743.0, 'C34': 681.0,
                         'F2_flat': 783.0, 'F34_flat': 736.0, 'F34_trefoil': 703.0,
                         'G2': 902.0, 'G34': 833.0},
            'R': 0.155,
            'X_trefoil': 0.16,
            'X_flat_touching': 0.25,
            'X_spaced': 0.25,
            'diameter': 42},
    
    400.0: {'ampacity': {'B2': 683.0, 'B34': 684.0, 'C2': 868.0, 'C34': 793.0,
                         'F2_flat': 940.0, 'F34_flat': 868.0, 'F34_trefoil': 823.0,
                         'G2': 1058.0, 'G34': 1008.0},
            'R': 0.12,
            'X_trefoil': 0.155,
            'X_flat_touching': 0.24,
            'X_spaced': 0.25,
            'diameter': 45},
    
    500.0: {'ampacity': {'B2': 783.0, 'B34': 666.0, 'C2': 990.0, 'C34': 904.0,
                         'F2_flat': 1083.0, 'F34_flat': 998.0, 'F34_trefoil': 946.0,
                         'G2': 1253.0, 'G34': 1169.0},
            'R': 0.093,
            'X_trefoil': 0.155,
            'X_flat_touching': 0.24,
            'X_spaced': 0.25,
            'diameter': 48},
    
    630.0: {'ampacity': {'B2': 900.0, 'B34': 764.0, 'C2': 1130.0, 'C34': 1033.0,
                         'F2_flat': 1254.0, 'F34_flat': 1151.0, 'F34_trefoil': 1088.0,
                         'G2': 1454.0, 'G34': 1362.0},
            'R': 0.072,
            'X_trefoil': 0.155,
            'X_flat_touching': 0.24,
            'X_spaced': 0.25,
            'diameter': 52},
    
    800.0: {'ampacity': {'C2': 1288.0, 'C34': 1179.0,
                         'F2_flat': 1358.0, 'F34_flat': 1275.0, 'F34_trefoil': 1214.0,
                         'G2': 1581.0, 'G34': 1485.0},
            'R': 0.056,
            'X_trefoil': 0.155,
            'X_flat_touching': 0.23,
            'X_spaced': 0.24,
            'diameter': 60},
    
    1000.0: {'ampacity': {'C2': 1443.0, 'C34': 1323.0,
                          'F2_flat': 1520.0, 'F34_flat': 1436.0, 'F34_trefoil': 1349.0,
                          'G2': 1775.0, 'G34': 1671.0},
             'R': 0.045,
             'X_trefoil': 0.155,
             'X_flat_touching': 0.23,
             'X_spaced': 0.24,
             'diameter': 68},
}

# ========== CORRECTED 70°C PVC SINGLE CORE NON-ARMOURED (From Sheet2) ==========
PVC_70_SINGLE_NON_ARMOURED = {
    1.0: {'ampacity': {'B2': 13.5, 'B34': 12.0, 'C2': 15.5, 'C34': 14.0}, 
          'voltage_drop': {'dc': 2.8, 'ac_single': 2.8, 'ac_three': 2.4}, 'diameter': 8},
    1.5: {'ampacity': {'B2': 17.5, 'B34': 15.5, 'C2': 20.0, 'C34': 18.0}, 
          'voltage_drop': None, 'diameter': 9},
    2.5: {'ampacity': {'B2': 24.0, 'B34': 21.0, 'C2': 27.0, 'C34': 25.0}, 
          'R': 1.75, 'X_trefoil': 0.20, 'X_flat_touching': 0.29, 'X_spaced': 0.33, 'diameter': 10},
    4.0: {'ampacity': {'B2': 32.0, 'B34': 28.0, 'C2': 37.0, 'C34': 33.0}, 
          'R': 1.25, 'X_trefoil': 0.195, 'X_flat_touching': 0.28, 'X_spaced': 0.31, 'diameter': 11},
    6.0: {'ampacity': {'B2': 41.0, 'B34': 36.0, 'C2': 47.0, 'C34': 43.0}, 
          'R': 0.93, 'X_trefoil': 0.19, 'X_flat_touching': 0.28, 'X_spaced': 0.30, 'diameter': 12},
    10.0: {'ampacity': {'B2': 57.0, 'B34': 50.0, 'C2': 65.0, 'C34': 59.0}, 
           'R': 0.63, 'X_trefoil': 0.185, 'X_flat_touching': 0.27, 'X_spaced': 0.29, 'diameter': 14},
    16.0: {'ampacity': {'B2': 76.0, 'B34': 68.0, 'C2': 87.0, 'C34': 79.0}, 
           'R': 0.46, 'X_trefoil': 0.18, 'X_flat_touching': 0.27, 'X_spaced': 0.28, 'diameter': 16},
    25.0: {'ampacity': {'B2': 101.0, 'B34': 89.0, 'C2': 114.0, 'C34': 104.0, 
                        'F2': 131.0, 'F34_flat': 114.0, 'F34_trefoil': 110.0,
                        'G2': 146.0, 'G34': 130.0}, 
           'R': 0.36, 'X_trefoil': 0.175, 'X_flat_touching': 0.26, 'X_spaced': 0.27, 'diameter': 18},
    35.0: {'ampacity': {'B2': 125.0, 'B34': 110.0, 'C2': 141.0, 'C34': 129.0,
                        'F2': 162.0, 'F34_flat': 143.0, 'F34_trefoil': 137.0,
                        'G2': 181.0, 'G34': 162.0}, 
           'R': 0.29, 'X_trefoil': 0.175, 'X_flat_touching': 0.26, 'X_spaced': 0.27, 'diameter': 20},
    50.0: {'ampacity': {'B2': 151.0, 'B34': 134.0, 'C2': 182.0, 'C34': 167.0,
                        'F2': 178.0, 'F34_flat': 174.0, 'F34_trefoil': 167.0,
                        'G2': 219.0, 'G34': 197.0}, 
           'R': 0.23, 'X_trefoil': 0.17, 'X_flat_touching': 0.26, 'X_spaced': 0.27, 'diameter': 22},
    70.0: {'ampacity': {'B2': 192.0, 'B34': 171.0, 'C2': 234.0, 'C34': 214.0,
                        'F2': 251.0, 'F34_flat': 225.0, 'F34_trefoil': 216.0,
                        'G2': 281.0, 'G34': 254.0}, 
           'R': 0.18, 'X_trefoil': 0.165, 'X_flat_touching': 0.25, 'X_spaced': 0.26, 'diameter': 25},
    95.0: {'ampacity': {'B2': 232.0, 'B34': 207.0, 'C2': 284.0, 'C34': 261.0,
                        'F2': 304.0, 'F34_flat': 275.0, 'F34_trefoil': 264.0,
                        'G2': 341.0, 'G34': 311.0}, 
           'R': 0.145, 'X_trefoil': 0.165, 'X_flat_touching': 0.25, 'X_spaced': 0.26, 'diameter': 28},
    120.0: {'ampacity': {'B2': 269.0, 'B34': 239.0, 'C2': 330.0, 'C34': 303.0,
                         'F2': 352.0, 'F34_flat': 321.0, 'F34_trefoil': 308.0,
                         'G2': 396.0, 'G34': 362.0}, 
            'R': 0.105, 'X_trefoil': 0.16, 'X_flat_touching': 0.25, 'X_spaced': 0.26, 'diameter': 30},
    150.0: {'ampacity': {'B2': 300.0, 'B34': 262.0, 'C2': 381.0, 'C34': 349.0,
                         'F2': 406.0, 'F34_flat': 372.0, 'F34_trefoil': 356.0,
                         'G2': 456.0, 'G34': 419.0}, 
            'R': 0.086, 'X_trefoil': 0.155, 'X_flat_touching': 0.24, 'X_spaced': 0.26, 'diameter': 32},
    185.0: {'ampacity': {'B2': 341.0, 'B34': 296.0, 'C2': 436.0, 'C34': 400.0,
                         'F2': 463.0, 'F34_flat': 427.0, 'F34_trefoil': 409.0,
                         'G2': 521.0, 'G34': 480.0}, 
            'R': 0.068, 'X_trefoil': 0.155, 'X_flat_touching': 0.24, 'X_spaced': 0.25, 'diameter': 35},
    240.0: {'ampacity': {'B2': 400.0, 'B34': 346.0, 'C2': 515.0, 'C34': 472.0,
                         'F2': 546.0, 'F34_flat': 507.0, 'F34_trefoil': 485.0,
                         'G2': 615.0, 'G34': 567.0}, 
            'R': 0.053, 'X_trefoil': 0.15, 'X_flat_touching': 0.24, 'X_spaced': 0.25, 'diameter': 38},
    300.0: {'ampacity': {'B2': 458.0, 'B34': 394.0, 'C2': 594.0, 'C34': 545.0,
                         'F2': 629.0, 'F34_flat': 587.0, 'F34_trefoil': 561.0,
                         'G2': 709.0, 'G34': 659.0}, 
            'R': 0.042, 'X_trefoil': 0.15, 'X_flat_touching': 0.24, 'X_spaced': 0.25, 'diameter': 42},
}

# ========== CORRECTED 70°C PVC MULTI-CORE NON-ARMOURED (From Sheet2) ==========
PVC_70_MULTI_NON_ARMOURED = {
    1.0: {'ampacity': {'B2': 13.0, 'B34': 11.5, 'C2': 15.0, 'C34': 13.5, 'F2': 17.0, 'F34': 14.5}, 
          'voltage_drop': 38.0, 'diameter': 10},
    1.5: {'ampacity': {'B2': 16.5, 'B34': 15.0, 'C2': 19.5, 'C34': 17.5, 'F2': 22.0, 'F34': 18.5}, 
          'voltage_drop': 25.0, 'diameter': 11},
    2.5: {'ampacity': {'B2': 23.0, 'B34': 20.0, 'C2': 27.0, 'C34': 24.0, 'F2': 30.0, 'F34': 25.0}, 
          'voltage_drop': 15.0, 'diameter': 12},
    4.0: {'ampacity': {'B2': 30.0, 'B34': 27.0, 'C2': 36.0, 'C34': 32.0, 'F2': 40.0, 'F34': 34.0}, 
          'voltage_drop': 9.5, 'diameter': 14},
    6.0: {'ampacity': {'B2': 38.0, 'B34': 34.0, 'C2': 46.0, 'C34': 41.0, 'F2': 51.0, 'F34': 43.0}, 
          'voltage_drop': 6.4, 'diameter': 16},
    10.0: {'ampacity': {'B2': 52.0, 'B34': 46.0, 'C2': 63.0, 'C34': 57.0, 'F2': 70.0, 'F34': 60.0}, 
           'voltage_drop': 3.8, 'diameter': 18},
    16.0: {'ampacity': {'B2': 69.0, 'B34': 62.0, 'C2': 85.0, 'C34': 76.0, 'F2': 94.0, 'F34': 80.0}, 
           'diameter': 20},
    25.0: {'ampacity': {'B2': 90.0, 'B34': 80.0, 'C2': 112.0, 'C34': 96.0, 'F2': 119.0, 'F34': 101.0}, 
           'R': 1.5, 'X': 0.145, 'diameter': 22},
    35.0: {'ampacity': {'B2': 111.0, 'B34': 99.0, 'C2': 138.0, 'C34': 119.0, 'F2': 148.0, 'F34': 126.0}, 
           'R': 1.1, 'X': 0.145, 'diameter': 25},
    50.0: {'ampacity': {'B2': 133.0, 'B34': 118.0, 'C2': 168.0, 'C34': 144.0, 'F2': 180.0, 'F34': 153.0}, 
           'R': 0.8, 'X': 0.14, 'diameter': 28},
    70.0: {'ampacity': {'B2': 168.0, 'B34': 149.0, 'C2': 213.0, 'C34': 184.0, 'F2': 232.0, 'F34': 196.0}, 
           'R': 0.55, 'X': 0.14, 'diameter': 32},
    95.0: {'ampacity': {'B2': 201.0, 'B34': 179.0, 'C2': 258.0, 'C34': 223.0, 'F2': 282.0, 'F34': 238.0}, 
           'R': 0.41, 'X': 0.135, 'diameter': 35},
    120.0: {'ampacity': {'B2': 232.0, 'B34': 206.0, 'C2': 299.0, 'C34': 259.0, 'F2': 328.0, 'F34': 276.0}, 
            'R': 0.33, 'X': 0.135, 'diameter': 38},
    150.0: {'ampacity': {'B2': 258.0, 'B34': 225.0, 'C2': 344.0, 'C34': 299.0, 'F2': 379.0, 'F34': 319.0}, 
            'R': 0.26, 'X': 0.13, 'diameter': 42},
    185.0: {'ampacity': {'B2': 294.0, 'B34': 255.0, 'C2': 392.0, 'C34': 341.0, 'F2': 434.0, 'F34': 364.0}, 
            'R': 0.21, 'X': 0.13, 'diameter': 45},
    240.0: {'ampacity': {'B2': 344.0, 'B34': 297.0, 'C2': 461.0, 'C34': 403.0, 'F2': 514.0, 'F34': 430.0}, 
            'R': 0.165, 'X': 0.13, 'diameter': 50},
    300.0: {'ampacity': {'B2': 394.0, 'B34': 339.0, 'C2': 530.0, 'C34': 464.0, 'F2': 593.0, 'F34': 497.0}, 
            'R': 0.135, 'X': 0.13, 'diameter': 55},
    400.0: {'ampacity': {'B2': 470.0, 'B34': 402.0, 'C2': 634.0, 'C34': 557.0, 'F2': 715.0, 'F34': 597.0}, 
            'R': 0.1, 'X': 0.125, 'diameter': 60},
}

# ========== CORRECTED 70°C PVC SINGLE CORE ARMOURED (From Sheet2, Table 4D3A) ==========
PVC_70_SINGLE_ARMOURED = {
    50.0: {'ampacity': {'C2': 193.0, 'C34': 179.0, 'F2_touch': 205.0, 'F34_touch_flat': 189.0, 'F34_touch_trefoil': 181.0}, 
           'R': 0.93, 'X_trefoil': 0.19, 'X_flat_touching': 0.26, 'X_spaced': 0.30, 'diameter': 25},
    70.0: {'ampacity': {'C2': 245.0, 'C34': 225.0, 'F2_touch': 259.0, 'F34_touch_flat': 238.0, 'F34_touch_trefoil': 231.0}, 
           'R': 0.63, 'X_trefoil': 0.18, 'X_flat_touching': 0.25, 'X_spaced': 0.29, 'diameter': 28},
    95.0: {'ampacity': {'C2': 296.0, 'C34': 269.0, 'F2_touch': 313.0, 'F34_touch_flat': 285.0, 'F34_touch_trefoil': 280.0}, 
           'R': 0.46, 'X_trefoil': 0.175, 'X_flat_touching': 0.25, 'X_spaced': 0.28, 'diameter': 32},
    120.0: {'ampacity': {'C2': 342.0, 'C34': 309.0, 'F2_touch': 360.0, 'F34_touch_flat': 327.0, 'F34_touch_trefoil': 324.0}, 
            'R': 0.36, 'X_trefoil': 0.17, 'X_flat_touching': 0.24, 'X_spaced': 0.28, 'diameter': 35},
    150.0: {'ampacity': {'C2': 393.0, 'C34': 352.0, 'F2_touch': 413.0, 'F34_touch_flat': 373.0, 'F34_touch_trefoil': 373.0}, 
            'R': 0.29, 'X_trefoil': 0.165, 'X_flat_touching': 0.24, 'X_spaced': 0.27, 'diameter': 38},
    185.0: {'ampacity': {'C2': 447.0, 'C34': 399.0, 'F2_touch': 469.0, 'F34_touch_flat': 422.0, 'F34_touch_trefoil': 425.0}, 
            'R': 0.23, 'X_trefoil': 0.16, 'X_flat_touching': 0.23, 'X_spaced': 0.27, 'diameter': 42},
    240.0: {'ampacity': {'C2': 525.0, 'C34': 465.0, 'F2_touch': 550.0, 'F34_touch_flat': 492.0, 'F34_touch_trefoil': 501.0}, 
            'R': 0.18, 'X_trefoil': 0.16, 'X_flat_touching': 0.23, 'X_spaced': 0.26, 'diameter': 48},
    300.0: {'ampacity': {'C2': 594.0, 'C34': 515.0, 'F2_touch': 624.0, 'F34_touch_flat': 547.0, 'F34_touch_trefoil': 567.0}, 
            'R': 0.145, 'X_trefoil': 0.155, 'X_flat_touching': 0.22, 'X_spaced': 0.26, 'diameter': 52},
    400.0: {'ampacity': {'C2': 687.0, 'C34': 575.0, 'F2_touch': 723.0, 'F34_touch_flat': 618.0, 'F34_touch_trefoil': 657.0}, 
            'R': 0.105, 'X_trefoil': 0.13, 'X_flat_touching': 0.21, 'X_spaced': 0.24, 'diameter': 58},
    500.0: {'ampacity': {'C2': 763.0, 'C34': 622.0, 'F2_touch': 805.0, 'F34_touch_flat': 673.0, 'F34_touch_trefoil': 731.0}, 
            'R': 0.086, 'X_trefoil': 0.145, 'X_flat_touching': 0.20, 'X_spaced': 0.23, 'diameter': 65},
    630.0: {'ampacity': {'C2': 843.0, 'C34': 669.0, 'F2_touch': 891.0, 'F34_touch_flat': 728.0, 'F34_touch_trefoil': 809.0}, 
            'R': 0.068, 'X_trefoil': 0.145, 'X_flat_touching': 0.195, 'X_spaced': 0.22, 'diameter': 72},
    800.0: {'ampacity': {'C2': 919.0, 'C34': 710.0, 'F2_touch': 976.0, 'F34_touch_flat': 777.0, 'F34_touch_trefoil': 886.0}, 
            'R': 0.053, 'X_trefoil': 0.14, 'X_flat_touching': 0.18, 'X_spaced': 0.21, 'diameter': 80},
    1000.0: {'ampacity': {'C2': 975.0, 'C34': 737.0, 'F2_touch': 1041.0, 'F34_touch_flat': 808.0, 'F34_touch_trefoil': 945.0}, 
             'R': 0.042, 'X_trefoil': 0.135, 'X_flat_touching': 0.17, 'X_spaced': 0.19, 'diameter': 88},
}

# ========== CORRECTED 70°C PVC MULTI-CORE ARMOURED (From Sheet2, Table 4D4A) ==========
PVC_70_MULTI_ARMOURED = {
    1.5: {'ampacity': {'C2': 21.0, 'C34': 18.0, 'E2': 22.0, 'E34': 19.0, 'D2': 22.0, 'D34': 18.0}, 
          'voltage_drop': {'dc': 29.0, 'ac': 29.0, 'ac34': 25.0}, 'diameter': 12},
    2.5: {'ampacity': {'C2': 28.0, 'C34': 25.0, 'E2': 31.0, 'E34': 26.0, 'D2': 29.0, 'D34': 24.0}, 
          'voltage_drop': {'dc': 18.0, 'ac': 18.0, 'ac34': 15.0}, 'diameter': 14},
    4.0: {'ampacity': {'C2': 38.0, 'C34': 33.0, 'E2': 41.0, 'E34': 35.0, 'D2': 37.0, 'D34': 30.0}, 
          'voltage_drop': {'dc': 11.0, 'ac': 11.0, 'ac34': 9.5}, 'diameter': 16},
    6.0: {'ampacity': {'C2': 49.0, 'C34': 42.0, 'E2': 53.0, 'E34': 45.0, 'D2': 46.0, 'D34': 38.0}, 
          'voltage_drop': {'dc': 7.3, 'ac': 7.3, 'ac34': 6.4}, 'diameter': 18},
    10.0: {'ampacity': {'C2': 67.0, 'C34': 58.0, 'E2': 72.0, 'E34': 62.0, 'D2': 60.0, 'D34': 50.0}, 
           'voltage_drop': {'dc': 4.4, 'ac': 4.4, 'ac34': 3.8}, 'diameter': 20},
    16.0: {'ampacity': {'C2': 89.0, 'C34': 77.0, 'E2': 97.0, 'E34': 83.0, 'D2': 78.0, 'D34': 64.0}, 
           'voltage_drop': {'dc': 2.8, 'ac': 2.8, 'ac34': 2.4}, 'diameter': 22},
    25.0: {'ampacity': {'C2': 118.0, 'C34': 102.0, 'E2': 128.0, 'E34': 110.0, 'D2': 99.0, 'D34': 82.0}, 
           'R': 1.75, 'X': 0.17, 'diameter': 25},
    35.0: {'ampacity': {'C2': 145.0, 'C34': 125.0, 'E2': 157.0, 'E34': 135.0, 'D2': 119.0, 'D34': 98.0}, 
           'R': 1.25, 'X': 0.165, 'diameter': 28},
    50.0: {'ampacity': {'C2': 175.0, 'C34': 151.0, 'E2': 190.0, 'E34': 163.0, 'D2': 140.0, 'D34': 116.0}, 
           'R': 0.93, 'X': 0.165, 'diameter': 32},
    70.0: {'ampacity': {'C2': 222.0, 'C34': 192.0, 'E2': 241.0, 'E34': 207.0, 'D2': 173.0, 'D34': 143.0}, 
           'R': 0.63, 'X': 0.16, 'diameter': 36},
    95.0: {'ampacity': {'C2': 269.0, 'C34': 231.0, 'E2': 291.0, 'E34': 251.0, 'D2': 204.0, 'D34': 169.0}, 
           'R': 0.46, 'X': 0.155, 'diameter': 40},
    120.0: {'ampacity': {'C2': 310.0, 'C34': 267.0, 'E2': 336.0, 'E34': 290.0, 'D2': 231.0, 'D34': 192.0}, 
            'R': 0.36, 'X': 0.155, 'diameter': 44},
    150.0: {'ampacity': {'C2': 356.0, 'C34': 306.0, 'E2': 386.0, 'E34': 332.0, 'D2': 261.0, 'D34': 217.0}, 
            'R': 0.29, 'X': 0.155, 'diameter': 48},
    185.0: {'ampacity': {'C2': 405.0, 'C34': 348.0, 'E2': 439.0, 'E34': 378.0, 'D2': 292.0, 'D34': 243.0}, 
            'R': 0.23, 'X': 0.15, 'diameter': 52},
    240.0: {'ampacity': {'C2': 476.0, 'C34': 409.0, 'E2': 516.0, 'E34': 445.0, 'D2': 336.0, 'D34': 280.0}, 
            'R': 0.18, 'X': 0.15, 'diameter': 58},
    300.0: {'ampacity': {'C2': 547.0, 'C34': 469.0, 'E2': 592.0, 'E34': 510.0, 'D2': 379.0, 'D34': 316.0}, 
            'R': 0.145, 'X': 0.145, 'diameter': 65},
    400.0: {'ampacity': {'C2': 621.0, 'C34': 540.0, 'E2': 683.0, 'E34': 590.0}, 
            'R': 0.105, 'X': 0.145, 'diameter': 72},
}

# ========== CORRECTED 90°C XLPE MULTI-CORE NON-ARMOURED (From Sheet3) ==========
XLPE_90_MULTI_NON_ARMOURED = {
    1.0: {'ampacity': {'B2': 17.0, 'B34': 15.0, 'C2': 19.0, 'C34': 17.0, 'E2': 21.0, 'E34': 18.0}, 
          'voltage_drop': {'dc': 46.0, 'ac': 46.0, 'ac34': 40.0}, 'diameter': 10},
    1.5: {'ampacity': {'B2': 22.0, 'B34': 19.5, 'C2': 24.0, 'C34': 22.0, 'E2': 26.0, 'E34': 23.0}, 
          'voltage_drop': {'dc': 31.0, 'ac': 31.0, 'ac34': 27.0}, 'diameter': 11},
    2.5: {'ampacity': {'B2': 30.0, 'B34': 26.0, 'C2': 33.0, 'C34': 30.0, 'E2': 36.0, 'E34': 32.0}, 
          'voltage_drop': {'dc': 19.0, 'ac': 19.0, 'ac34': 16.0}, 'diameter': 12},
    4.0: {'ampacity': {'B2': 40.0, 'B34': 35.0, 'C2': 45.0, 'C34': 40.0, 'E2': 49.0, 'E34': 42.0}, 
          'voltage_drop': {'dc': 12.0, 'ac': 12.0, 'ac34': 10.0}, 'diameter': 14},
    6.0: {'ampacity': {'B2': 51.0, 'B34': 44.0, 'C2': 58.0, 'C34': 52.0, 'E2': 63.0, 'E34': 54.0}, 
          'voltage_drop': {'dc': 7.9, 'ac': 7.9, 'ac34': 6.8}, 'diameter': 16},
    10.0: {'ampacity': {'B2': 69.0, 'B34': 60.0, 'C2': 80.0, 'C34': 71.0, 'E2': 86.0, 'E34': 75.0}, 
           'voltage_drop': {'dc': 4.7, 'ac': 4.7, 'ac34': 4.0}, 'diameter': 18},
    16.0: {'ampacity': {'B2': 91.0, 'B34': 80.0, 'C2': 107.0, 'C34': 96.0, 'E2': 115.0, 'E34': 100.0}, 
           'voltage_drop': {'dc': 2.9, 'ac': 2.9, 'ac34': 2.5}, 'diameter': 20},
    25.0: {'ampacity': {'B2': 119.0, 'B34': 105.0, 'C2': 138.0, 'C34': 119.0, 'E2': 149.0, 'E34': 127.0}, 
           'R': 1.85, 'X': 0.16, 'diameter': 22},
    35.0: {'ampacity': {'B2': 146.0, 'B34': 128.0, 'C2': 171.0, 'C34': 147.0, 'E2': 185.0, 'E34': 158.0}, 
           'R': 1.35, 'X': 0.155, 'diameter': 25},
    50.0: {'ampacity': {'B2': 175.0, 'B34': 154.0, 'C2': 209.0, 'C34': 179.0, 'E2': 225.0, 'E34': 192.0}, 
           'R': 0.98, 'X': 0.155, 'diameter': 28},
    70.0: {'ampacity': {'B2': 221.0, 'B34': 194.0, 'C2': 269.0, 'C34': 229.0, 'E2': 289.0, 'E34': 246.0}, 
           'R': 0.67, 'X': 0.15, 'diameter': 32},
    95.0: {'ampacity': {'B2': 265.0, 'B34': 233.0, 'C2': 328.0, 'C34': 278.0, 'E2': 352.0, 'E34': 298.0}, 
           'R': 0.49, 'X': 0.15, 'diameter': 35},
    120.0: {'ampacity': {'B2': 305.0, 'B34': 268.0, 'C2': 382.0, 'C34': 322.0, 'E2': 410.0, 'E34': 346.0}, 
            'R': 0.39, 'X': 0.145, 'diameter': 38},
    150.0: {'ampacity': {'B2': 334.0, 'B34': 300.0, 'C2': 441.0, 'C34': 371.0, 'E2': 473.0, 'E34': 399.0}, 
            'R': 0.31, 'X': 0.145, 'diameter': 42},
    185.0: {'ampacity': {'B2': 384.0, 'B34': 340.0, 'C2': 506.0, 'C34': 424.0, 'E2': 542.0, 'E34': 456.0}, 
            'R': 0.25, 'X': 0.145, 'diameter': 45},
    240.0: {'ampacity': {'B2': 459.0, 'B34': 398.0, 'C2': 599.0, 'C34': 500.0, 'E2': 641.0, 'E34': 538.0}, 
            'R': 0.195, 'X': 0.14, 'diameter': 50},
    300.0: {'ampacity': {'B2': 532.0, 'B34': 455.0, 'C2': 693.0, 'C34': 576.0, 'E2': 741.0, 'E34': 621.0}, 
            'R': 0.155, 'X': 0.14, 'diameter': 55},
    400.0: {'ampacity': {'B2': 625.0, 'B34': 536.0, 'C2': 803.0, 'C34': 667.0, 'E2': 865.0, 'E34': 741.0}, 
            'R': 0.12, 'X': 0.14, 'diameter': 60},
}

# ========== CORRECTED 90°C XLPE SINGLE CORE ARMOURED (From Sheet3) ==========
XLPE_90_SINGLE_ARMOURED = {
    50.0: {'ampacity': {'C2': 237.0, 'C34': 220.0, 'F2_touch': 253.0, 'F34_touch_flat': 232.0, 'F34_touch_trefoil': 222.0}, 
           'R': 0.98, 'X_trefoil': 0.18, 'X_flat_touching': 0.25, 'X_spaced': 0.29, 'diameter': 25},
    70.0: {'ampacity': {'C2': 303.0, 'C34': 277.0, 'F2_touch': 322.0, 'F34_touch_flat': 293.0, 'F34_touch_trefoil': 285.0}, 
           'R': 0.67, 'X_trefoil': 0.17, 'X_flat_touching': 0.25, 'X_spaced': 0.29, 'diameter': 28},
    95.0: {'ampacity': {'C2': 367.0, 'C34': 333.0, 'F2_touch': 389.0, 'F34_touch_flat': 352.0, 'F34_touch_trefoil': 346.0}, 
           'R': 0.49, 'X_trefoil': 0.17, 'X_flat_touching': 0.24, 'X_spaced': 0.28, 'diameter': 32},
    120.0: {'ampacity': {'C2': 425.0, 'C34': 383.0, 'F2_touch': 449.0, 'F34_touch_flat': 405.0, 'F34_touch_trefoil': 402.0}, 
            'R': 0.39, 'X_trefoil': 0.165, 'X_flat_touching': 0.24, 'X_spaced': 0.27, 'diameter': 35},
    150.0: {'ampacity': {'C2': 488.0, 'C34': 437.0, 'F2_touch': 516.0, 'F34_touch_flat': 462.0, 'F34_touch_trefoil': 463.0}, 
            'R': 0.31, 'X_trefoil': 0.16, 'X_flat_touching': 0.23, 'X_spaced': 0.27, 'diameter': 38},
    185.0: {'ampacity': {'C2': 557.0, 'C34': 496.0, 'F2_touch': 587.0, 'F34_touch_flat': 524.0, 'F34_touch_trefoil': 529.0}, 
            'R': 0.25, 'X_trefoil': 0.16, 'X_flat_touching': 0.23, 'X_spaced': 0.26, 'diameter': 42},
    240.0: {'ampacity': {'C2': 656.0, 'C34': 579.0, 'F2_touch': 689.0, 'F34_touch_flat': 612.0, 'F34_touch_trefoil': 625.0}, 
            'R': 0.195, 'X_trefoil': 0.155, 'X_flat_touching': 0.22, 'X_spaced': 0.26, 'diameter': 48},
    300.0: {'ampacity': {'C2': 755.0, 'C34': 662.0, 'F2_touch': 792.0, 'F34_touch_flat': 700.0, 'F34_touch_trefoil': 720.0}, 
            'R': 0.155, 'X_trefoil': 0.15, 'X_flat_touching': 0.22, 'X_spaced': 0.25, 'diameter': 52},
    400.0: {'ampacity': {'C2': 853.0, 'C34': 717.0, 'F2_touch': 899.0, 'F34_touch_flat': 767.0, 'F34_touch_trefoil': 815.0}, 
            'R': 0.115, 'X_trefoil': 0.15, 'X_flat_touching': 0.21, 'X_spaced': 0.24, 'diameter': 58},
    500.0: {'ampacity': {'C2': 962.0, 'C34': 791.0, 'F2_touch': 1016.0, 'F34_touch_flat': 851.0, 'F34_touch_trefoil': 918.0}, 
            'R': 0.093, 'X_trefoil': 0.145, 'X_flat_touching': 0.20, 'X_spaced': 0.24, 'diameter': 65},
    630.0: {'ampacity': {'C2': 1082.0, 'C34': 861.0, 'F2_touch': 1146.0, 'F34_touch_flat': 935.0, 'F34_touch_trefoil': 1027.0}, 
            'R': 0.073, 'X_trefoil': 0.145, 'X_flat_touching': 0.195, 'X_spaced': 0.23, 'diameter': 72},
    800.0: {'ampacity': {'C2': 1170.0, 'C34': 904.0, 'F2_touch': 1246.0, 'F34_touch_flat': 987.0, 'F34_touch_trefoil': 1119.0}, 
            'R': 0.056, 'X_trefoil': 0.14, 'X_flat_touching': 0.18, 'X_spaced': 0.23, 'diameter': 80},
    1000.0: {'ampacity': {'C2': 1261.0, 'C34': 961.0, 'F2_touch': 1345.0, 'F34_touch_flat': 1055.0, 'F34_touch_trefoil': 1214.0}, 
             'R': 0.045, 'X_trefoil': 0.135, 'X_flat_touching': 0.17, 'X_spaced': 0.21, 'diameter': 88},
}

# ========== CORRECTED 90°C XLPE MULTI-CORE ARMOURED (From Sheet3) ==========
XLPE_90_MULTI_ARMOURED = {
    1.5: {'ampacity': {'C2': 27.0, 'C34': 23.0, 'E2': 29.0, 'E34': 25.0, 'D2': 25.0, 'D34': 21.0}, 
          'voltage_drop': {'dc': 31.0, 'ac': 31.0, 'ac34': 27.0}, 'diameter': 12},
    2.5: {'ampacity': {'C2': 36.0, 'C34': 31.0, 'E2': 39.0, 'E34': 33.0, 'D2': 33.0, 'D34': 28.0}, 
          'voltage_drop': {'dc': 19.0, 'ac': 19.0, 'ac34': 16.0}, 'diameter': 14},
    4.0: {'ampacity': {'C2': 49.0, 'C34': 42.0, 'E2': 52.0, 'E34': 44.0, 'D2': 43.0, 'D34': 36.0}, 
          'voltage_drop': {'dc': 12.0, 'ac': 12.0, 'ac34': 10.0}, 'diameter': 16},
    6.0: {'ampacity': {'C2': 62.0, 'C34': 53.0, 'E2': 66.0, 'E34': 56.0, 'D2': 53.0, 'D34': 44.0}, 
          'voltage_drop': {'dc': 7.9, 'ac': 7.9, 'ac34': 6.8}, 'diameter': 18},
    10.0: {'ampacity': {'C2': 85.0, 'C34': 73.0, 'E2': 90.0, 'E34': 78.0, 'D2': 71.0, 'D34': 58.0}, 
           'voltage_drop': {'dc': 4.7, 'ac': 4.7, 'ac34': 4.0}, 'diameter': 20},
    16.0: {'ampacity': {'C2': 110.0, 'C34': 94.0, 'E2': 115.0, 'E34': 99.0, 'D2': 91.0, 'D34': 75.0}, 
           'voltage_drop': {'dc': 2.9, 'ac': 2.9, 'ac34': 2.5}, 'diameter': 22},
    25.0: {'ampacity': {'C2': 146.0, 'C34': 124.0, 'E2': 152.0, 'E34': 131.0, 'D2': 116.0, 'D34': 96.0}, 
           'R': 1.85, 'X': 0.16, 'diameter': 25},
    35.0: {'ampacity': {'C2': 180.0, 'C34': 154.0, 'E2': 188.0, 'E34': 162.0, 'D2': 139.0, 'D34': 115.0}, 
           'R': 1.35, 'X': 0.155, 'diameter': 28},
    50.0: {'ampacity': {'C2': 219.0, 'C34': 187.0, 'E2': 228.0, 'E34': 197.0, 'D2': 164.0, 'D34': 135.0}, 
           'R': 0.98, 'X': 0.155, 'diameter': 32},
    70.0: {'ampacity': {'C2': 279.0, 'C34': 238.0, 'E2': 291.0, 'E34': 251.0, 'D2': 203.0, 'D34': 167.0}, 
           'R': 0.67, 'X': 0.15, 'diameter': 36},
    95.0: {'ampacity': {'C2': 338.0, 'C34': 289.0, 'E2': 354.0, 'E34': 304.0, 'D2': 239.0, 'D34': 197.0}, 
           'R': 0.49, 'X': 0.15, 'diameter': 40},
    120.0: {'ampacity': {'C2': 392.0, 'C34': 335.0, 'E2': 410.0, 'E34': 353.0, 'D2': 271.0, 'D34': 223.0}, 
            'R': 0.39, 'X': 0.145, 'diameter': 44},
    150.0: {'ampacity': {'C2': 451.0, 'C34': 386.0, 'E2': 472.0, 'E34': 406.0, 'D2': 306.0, 'D34': 251.0}, 
            'R': 0.31, 'X': 0.145, 'diameter': 48},
    185.0: {'ampacity': {'C2': 515.0, 'C34': 441.0, 'E2': 539.0, 'E34': 463.0, 'D2': 343.0, 'D34': 281.0}, 
            'R': 0.25, 'X': 0.145, 'diameter': 52},
    240.0: {'ampacity': {'C2': 607.0, 'C34': 520.0, 'E2': 636.0, 'E34': 546.0, 'D2': 395.0, 'D34': 324.0}, 
            'R': 0.195, 'X': 0.14, 'diameter': 58},
    300.0: {'ampacity': {'C2': 698.0, 'C34': 599.0, 'E2': 732.0, 'E34': 628.0, 'D2': 446.0, 'D34': 365.0}, 
            'R': 0.155, 'X': 0.14, 'diameter': 65},
    400.0: {'ampacity': {'C2': 787.0, 'C34': 673.0, 'E2': 847.0, 'E34': 728.0}, 
            'R': 0.12, 'X': 0.14, 'diameter': 72},
}

# ========== COMPLETE CABLE DATABASE STRUCTURE ==========
CABLE_DATABASE = {
    'PVC_70': {
        'insulation_temp': 70,
        'description': 'PVC Insulated, 70°C',
        'single_core_non_armoured': PVC_70_SINGLE_NON_ARMOURED,
        'multi_core_non_armoured': PVC_70_MULTI_NON_ARMOURED,
        'single_core_armoured': PVC_70_SINGLE_ARMOURED,
        'multi_core_armoured': PVC_70_MULTI_ARMOURED
    },
    'XLPE_90': {
        'insulation_temp': 90,
        'description': 'XLPE Insulated, 90°C',
        'single_core_non_armoured': XLPE_90_SINGLE_NON_ARMOURED,
        'multi_core_non_armoured': XLPE_90_MULTI_NON_ARMOURED,
        'single_core_armoured': XLPE_90_SINGLE_ARMOURED,
        'multi_core_armoured': XLPE_90_MULTI_ARMOURED
    }
}

def get_cable_data(insulation_type='XLPE_90', cable_type='single_core_non_armoured'):
    try:
        db = CABLE_DATABASE[insulation_type]
        return db.get(cable_type, {})
    except:
        return {}

def get_temperature_factor(insulation_temp, ambient_temp, installation='air'):
    try:
        if installation in ['buried', 'duct', 'trench', 'ground', 'D']:
            factors = TEMPERATURE_FACTORS_GROUND.get(insulation_temp, TEMPERATURE_FACTORS_GROUND[90])
        else:
            factors = TEMPERATURE_FACTORS_AIR.get(insulation_temp, TEMPERATURE_FACTORS_AIR[90])
        temps = sorted(factors.keys())
        closest_temp = min(temps, key=lambda x: abs(x - ambient_temp))
        return factors[closest_temp]
    except:
        return 1.0

def get_grouping_factor(num_cables, arrangement='bunched_in_air_surface_enclosed'):
    try:
        if arrangement in GROUPING_FACTORS:
            factors = GROUPING_FACTORS[arrangement]
            available = sorted(factors.keys())
            if num_cables in factors:
                return factors[num_cables]
            elif num_cables > max(available):
                return factors[max(available)]
            else:
                closest = min(available, key=lambda x: abs(x - num_cables))
                return factors[closest]
        return 1.0
    except:
        return 1.0

def get_soil_resistivity_factor(resistivity):
    try:
        resistivities = sorted(SOIL_RESISTIVITY_FACTORS.keys())
        closest_res = min(resistivities, key=lambda x: abs(x - resistivity))
        return SOIL_RESISTIVITY_FACTORS[closest_res]
    except:
        return 1.0

def get_depth_factor(depth_m):
    try:
        depths = sorted(DEPTH_FACTORS.keys())
        closest_depth = min(depths, key=lambda x: abs(x - depth_m))
        return DEPTH_FACTORS[closest_depth]
    except:
        return 1.0

# ========== SHORT CIRCUIT CALCULATION (IEC 60949) ==========
def calculate_short_circuit_current(size_mm2, insulation_type, duration_s=1.0, conductor_material='Copper'):
    if conductor_material == 'Copper':
        K = 226
        β = 234.5
    else:
        K = 148
        β = 228
    
    if insulation_type == 'PVC':
        θi = 70
        if size_mm2 <= 300:
            θf = 160
        else:
            θf = 140
    else:
        θi = 90
        θf = 250
    
    first_term = K * size_mm2 / math.sqrt(duration_s)
    log_term = math.log((θf + β) / (θi + β))
    Isc = first_term * math.sqrt(log_term)
    
    return Isc, K, θi, θf

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
    
    def calculate_operating_temperature(self, ambient_temp, load_current, derated_ampacity, insulation_temp):
        if derated_ampacity > 0:
            load_ratio = load_current / derated_ampacity
            temp_rise = (insulation_temp - ambient_temp) * (load_ratio ** 2)
            operating_temp = ambient_temp + temp_rise
        else:
            operating_temp = insulation_temp
        return operating_temp
    
    def calculate_short_circuit(self, size_mm2, insulation_type, ambient_temp, load_current, rated_current, k1, k2, k3, k4, duration_s=1.0, conductor_material='Copper'):
        Isc, K, θi, θf = calculate_short_circuit_current(size_mm2, insulation_type, duration_s, conductor_material)
        insulation_temp = 70 if insulation_type == 'PVC' else 90
        total_k = k1 * k2 * k3 * k4
        derated_ampacity = rated_current * total_k
        operating_temp = self.calculate_operating_temperature(ambient_temp, load_current, derated_ampacity, insulation_temp)
        return Isc, K, operating_temp, θi, θf
    
    def get_derating_factors(self, temp_c, insulation_temp, num_cables, arrangement, installation, soil_resistivity, depth):
        if installation in ['buried', 'duct', 'trench', 'ground', 'D']:
            k1 = get_temperature_factor(insulation_temp, temp_c, 'ground')
        else:
            k1 = get_temperature_factor(insulation_temp, temp_c, 'air')
        
        if arrangement in ['bunched_in_air_surface_enclosed', 'single_layer_wall_floor', 'single_layer_perforated_tray', 'single_layer_ladder_cleats']:
            k2 = get_grouping_factor(num_cables, arrangement)
        else:
            k2 = 1.0
        
        if installation in ['buried', 'duct', 'ground', 'D']:
            k3 = get_soil_resistivity_factor(soil_resistivity)
            k4 = get_depth_factor(depth)
        else:
            k3 = 1.0
            k4 = 1.0
        
        total_k = k1 * k2 * k3 * k4
        factors = {'k1 (Temperature)': k1, 'k2 (Grouping)': k2, 'k3 (Soil Resistivity)': k3, 'k4 (Depth)': k4, 'total': total_k}
        return total_k, factors
    
    def calculate_voltage_drop(self, current, length_m, cable_data, pf, voltage_v, phase='3-phase', installation='C', arrangement='flat'):
        # Check for direct voltage drop values first (for small cables ≤16mm²)
        voltage_drop_mv = cable_data.get('voltage_drop', {})
        if isinstance(voltage_drop_mv, dict) and voltage_drop_mv:
            if phase == '3-phase':
                vd_mv = voltage_drop_mv.get('ac34', voltage_drop_mv.get('ac', 0))
            else:
                vd_mv = voltage_drop_mv.get('ac', voltage_drop_mv.get('ac34', 0))
            
            if vd_mv and vd_mv > 0:
                Vd = vd_mv * current * length_m / 1000
                return Vd, (Vd / voltage_v) * 100
        
        # For larger cables, use R and X values
        r = cable_data.get('R', 0)
        
        # Get X value based on arrangement
        x = 0
        if arrangement == 'trefoil':
            x = cable_data.get('X_trefoil', 0)
        elif arrangement == 'spaced':
            x = cable_data.get('X_spaced', 0)
        else:  # flat or touching
            x = cable_data.get('X_flat_touching', 0)
        
        # For multi-core cables with single X value
        if x == 0:
            x = cable_data.get('X', 0)
        
        if r == 0:
            return 0, 0
        
        phi = math.acos(pf)
        if phase == '3-phase':
            Vd = 1.732 * current * (r * pf + x * math.sin(phi)) * length_m / 1000
        else:
            Vd = 2 * current * (r * pf + x * math.sin(phi)) * length_m / 1000
        
        return Vd, (Vd / voltage_v) * 100
    
    def get_cable_category(self, voltage_v):
        if voltage_v <= 1000:
            return 'LV (0.6/1kV)', 'LV'
        elif voltage_v <= 3300:
            return 'MV (3.3kV)', 'MV_33KV'
        elif voltage_v <= 6600:
            return 'MV (6.6kV)', 'MV_66KV'
        else:
            return 'MV (11kV)', 'MV_11KV'

# ========== AUTOMATIC CABLE SELECTION FUNCTION ==========
def select_cable_automatically(load, cable_db, cable_calc, ambient_temp,
                                insulation_temp, load_current, 
                                load_length, load_pf, load_voltage, load_phase,
                                installation_method, cable_formation, load_cable_type,
                                load_arrangement, load_soil_res, load_depth,
                                load_num_cables):
    
    available_sizes = sorted(cable_db.keys())
    results_list = []
    
    for size in available_sizes:
        cable_data = cable_db[size]
        
        total_k, factors = cable_calc.get_derating_factors(
            ambient_temp, insulation_temp,
            load_num_cables, load_arrangement,
            installation_method,
            load_soil_res, load_depth
        )
        
        ampacity = 0
        if installation_method == 'B':
            if load_phase == '3-phase':
                ampacity = cable_data.get('ampacity', {}).get('B34', 0)
            else:
                ampacity = cable_data.get('ampacity', {}).get('B2', 0)
        elif installation_method == 'C':
            if load_phase == '3-phase':
                ampacity = cable_data.get('ampacity', {}).get('C34', 0)
            else:
                ampacity = cable_data.get('ampacity', {}).get('C2', 0)
        elif installation_method == 'F':
            if cable_formation == 'trefoil':
                ampacity = cable_data.get('ampacity', {}).get('F34_trefoil', 0)
            elif cable_formation == 'spaced':
                ampacity = cable_data.get('ampacity', {}).get('G2', cable_data.get('ampacity', {}).get('G34', 0))
            else:
                if load_phase == '3-phase':
                    ampacity = cable_data.get('ampacity', {}).get('F34_flat', 0)
                else:
                    ampacity = cable_data.get('ampacity', {}).get('F2_flat', 0)
        elif installation_method == 'E':
            if load_phase == '3-phase':
                ampacity = cable_data.get('ampacity', {}).get('E34', 0)
            else:
                ampacity = cable_data.get('ampacity', {}).get('E2', 0)
        elif installation_method == 'D':
            if load_phase == '3-phase':
                ampacity = cable_data.get('ampacity', {}).get('D34', 0)
            else:
                ampacity = cable_data.get('ampacity', {}).get('D2', 0)
        elif installation_method == 'G':
            if load_phase == '3-phase':
                ampacity = cable_data.get('ampacity', {}).get('G34', 0)
            else:
                ampacity = cable_data.get('ampacity', {}).get('G2', 0)
        
        if ampacity == 0:
            ampacity = cable_data.get('ampacity', {}).get('C2', 0)
        
        if ampacity == 0:
            continue
        
        derated = ampacity * total_k
        ampacity_pass = derated >= load_current
        
        vd_v, vd_pct = cable_calc.calculate_voltage_drop(
            load_current, load_length, cable_data,
            load_pf, load_voltage, load_phase,
            installation_method, cable_formation
        )
        
        vd_pass = vd_pct <= 2.5
        
        results_list.append({
            'size': size,
            'ampacity': ampacity,
            'derated': derated,
            'vd_pct': vd_pct,
            'ampacity_pass': ampacity_pass,
            'vd_pass': vd_pass,
            'total_k': total_k
        })
        
        if ampacity_pass and vd_pass:
            return size, cable_data, ampacity, derated, vd_pct, total_k, factors, True, results_list
    
    if available_sizes:
        largest_size = max(available_sizes)
        largest_data = cable_db[largest_size]
        
        total_k, factors = cable_calc.get_derating_factors(
            ambient_temp, insulation_temp,
            load_num_cables, load_arrangement,
            installation_method,
            load_soil_res, load_depth
        )
        
        ampacity = 0
        if installation_method == 'C':
            if load_phase == '3-phase':
                ampacity = largest_data.get('ampacity', {}).get('C34', 0)
            else:
                ampacity = largest_data.get('ampacity', {}).get('C2', 0)
        
        if ampacity == 0:
            ampacity = largest_data.get('ampacity', {}).get('C2', 0)
        
        derated = ampacity * total_k
        vd_v, vd_pct = cable_calc.calculate_voltage_drop(
            load_current, load_length, largest_data,
            load_pf, load_voltage, load_phase,
            installation_method, cable_formation
        )
        
        return largest_size, largest_data, ampacity, derated, vd_pct, total_k, factors, False, results_list
    
    return None, None, 0, 0, 0, 0, {}, False, []

# ========== CIRCUIT BREAKER CALCULATOR CLASS ==========
class CircuitBreakerCalculator:
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
    
    def calculate_cb_size(self, loads_df, design_factor=1.25, manufacturer='Schneider Electric'):
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
            series = MANUFACTURERS[manufacturer][breaker_type]
            
            results.append({
                'Load': load['Load Name'],
                'Power (kW)': load['Power (kW)'],
                'Voltage (V)': load['Voltage (V)'],
                'Phase': load['Phase'],
                'Load Type': load.get('Load Type', 'Continuous'),
                'Current (A)': current,
                'Required CB (A)': required,
                'Selected CB (A)': rating,
                'Breaker Type': breaker_type,
                'Standard': standard,
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
                'design_factor': design_factor,
                'manufacturer': manufacturer,
                'series': series
            })
        
        return results, detailed_reasons
    
    def calculate_main_cb(self, loads_df, voltage=400, pf=0.85, design_factor=1.25):
        total_power = loads_df['Power (kW)'].sum()
        current = total_power * 1000 / (1.732 * voltage * pf)
        rating, required = self.get_standard_rating(current, design_factor)
        breaker_type, standard = self.get_breaker_type(rating)
        
        detailed_reason = f"""
MAIN CIRCUIT BREAKER DETAILED CALCULATION

Step 1: Total load analysis
- Total connected load: {total_power:.2f} kW
- System voltage: {voltage} V (3-phase)
- Power factor: {pf}

Step 2: Total current calculation
Formula: I = P x 1000 / (1.732 x V x PF)
I = {total_power:.2f} x 1000 / (1.732 x {voltage} x {pf})
I = {current:.2f} A

Step 3: Circuit breaker sizing
- Design factor: {design_factor}
- Required rating = {current:.2f} x {design_factor} = {required:.2f} A
- Selected standard rating: {rating} A

Step 4: Breaker type selection
- Based on rating {rating} A -> {breaker_type}
- Application: {BREAKER_TYPES[breaker_type]['application']}

Step 5: Pole selection
- User to select poles based on system requirements

Final selection: {rating} A {breaker_type}
"""
        
        return {
            'total_power': total_power,
            'current': current,
            'required_cb': required,
            'selected_cb': rating,
            'breaker_type': breaker_type,
            'standard': standard,
            'detailed_reason': detailed_reason
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
        
        self.doc.add_heading('1.1 Collection area (Ad)', level=1)
        self.doc.add_paragraph('Formula: Ad = L x W + 2 x (3H) x (L + W) + pi x (3H)^2')
        self.doc.add_paragraph('Reference: IEC 62305-2 Annex A.2.1.1')
        p = self.doc.add_paragraph()
        p.add_run('Result: ').bold = True
        p.add_run(f'Ad = {results["ad"]:.2f} m²')
        
        self.doc.add_heading('1.2 Near strike collection area (Am)', level=1)
        self.doc.add_paragraph('Formula: Am = 2 x 500 x (L + W) + pi x 500^2')
        self.doc.add_paragraph('Reference: IEC 62305-2 Annex A.3')
        p = self.doc.add_paragraph()
        p.add_run('Result: ').bold = True
        p.add_run(f'Am = {results["am"]:.2f} m²')
        
        self.doc.add_heading('1.3 Environmental factor (CD)', level=1)
        self.doc.add_paragraph(f'Selected environment: {inputs.get("environment", "Isolated")}')
        p = self.doc.add_paragraph()
        p.add_run('Result: ').bold = True
        p.add_run(f'CD = {inputs.get("cd", 1)}')
        
        self.doc.add_heading('1.4 Lightning ground flash density (NG)', level=1)
        self.doc.add_paragraph('Formula: NG = 0.1 x Td')
        p = self.doc.add_paragraph()
        p.add_run('Result: ').bold = True
        p.add_run(f'NG = {results.get("ng", 1)} flashes/km²/year')
        
        self.doc.add_heading('1.5 Lightning frequencies', level=1)
        p = self.doc.add_paragraph()
        p.add_run('Nd (Direct): ').bold = True
        p.add_run(f'{results.get("nd", 0):.6f} events/year')
        p = self.doc.add_paragraph()
        p.add_run('Nm (Near): ').bold = True
        p.add_run(f'{results.get("nm", 0):.6f} events/year')
        
        self.doc.add_heading('1.6 Protection level', level=1)
        p = self.doc.add_paragraph()
        p.add_run('Result: ').bold = True
        p.add_run(f'{results.get("lpl", "Class III")}')
        self.doc.add_paragraph(f'Rolling sphere radius: {results.get("sphere", 45)}m')
        
        self.doc.add_page_break()
        self.doc.add_heading('SUMMARY OF RESULTS', level=1)
        table = self.doc.add_table(rows=1, cols=2)
        table.style = 'Light Grid Accent 1'
        hdr_cells = table.rows[0].cells
        hdr_cells[0].text = 'Parameter'
        hdr_cells[1].text = 'Value'
        
        summary_data = [
            ('Collection area (Ad)', f"{results['ad']:.2f} m²"),
            ('Near strike area (Am)', f"{results['am']:.2f} m²"),
            ('Environmental factor (CD)', str(inputs.get('cd', 1))),
            ('Lightning density (NG)', f"{results.get('ng', 1)} flashes/km²/year"),
            ('Direct frequency (Nd)', f"{results.get('nd', 0):.6f} events/year"),
            ('Near frequency (Nm)', f"{results.get('nm', 0):.6f} events/year"),
            ('Protection efficiency', f"{results.get('efficiency', 0):.1%}"),
            ('Protection level', results.get('lpl', 'Class III')),
            ('Rolling sphere radius', f"{results.get('sphere', 45)} m"),
            ('Air terminals required', str(results.get('air_terminals', 4)))
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

# ========== WORD REPORT CLASSES ==========
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
    
    def add_common_parameters(self, ambient_temp):
        heading = self.doc.add_heading('COMMON PARAMETERS FOR ALL LOADS', level=1)
        heading.runs[0].font.color.rgb = RGBColor(0, 51, 102)
        self.doc.add_paragraph(f'Ambient temperature: {ambient_temp}°C')
        self.doc.add_paragraph()
    
    def add_load_details(self, loads_df):
        heading = self.doc.add_heading('LOAD DETAILS', level=1)
        heading.runs[0].font.color.rgb = RGBColor(0, 51, 102)
        
        table = self.doc.add_table(rows=1, cols=7)
        table.style = 'Light Grid Accent 1'
        
        headers = ['Load name', 'Power (kw)', 'Voltage (v)', 'Load type', 'Phase', 'Pf', 'Length (m)']
        for i, header in enumerate(headers):
            table.rows[0].cells[i].text = header
            table.rows[0].cells[i].paragraphs[0].runs[0].bold = True
        
        for idx, load in loads_df.iterrows():
            row = table.add_row().cells
            row[0].text = load['Load Name']
            row[1].text = f"{load['Power (kW)']:.1f}"
            row[2].text = f"{load['Voltage (V)']:.0f}"
            row[3].text = format_load_type(load.get('Load Type', 'Continuous'))
            row[4].text = load['Phase']
            row[5].text = f"{load['Power Factor']:.2f}"
            row[6].text = f"{load['Length (m)']:.0f}"
        
        self.doc.add_paragraph()
    
    def add_cable_results(self, cable_df):
        heading = self.doc.add_heading('CABLE SIZING RESULTS', level=1)
        heading.runs[0].font.color.rgb = RGBColor(0, 51, 102)
        
        table = self.doc.add_table(rows=1, cols=8)
        table.style = 'Light Grid Accent 1'
        
        headers = ['Load name', 'Size', 'Base a', 'Derated a', 'Vd %', 'Sc ka', 'Status', 'Check']
        for i, header in enumerate(headers):
            table.rows[0].cells[i].text = header
            table.rows[0].cells[i].paragraphs[0].runs[0].bold = True
        
        for idx, row in cable_df.iterrows():
            new_row = table.add_row().cells
            new_row[0].text = row['Load Name']
            new_row[1].text = str(row['Size (mm²)'])
            new_row[2].text = str(row['Base Ampacity (A)'])
            new_row[3].text = str(row['Derated Ampacity (A)']).replace(' A', '')
            new_row[4].text = str(row['Voltage Drop (%)']).replace('%', '')
            new_row[5].text = f"{float(row['Short Circuit (kA)']):.2f}" if isinstance(row['Short Circuit (kA)'], (int, float)) else row['Short Circuit (kA)']
            new_row[6].text = row['Status']
            new_row[7].text = row.get('Check', 'N/A')
            
            if row['Status'] == 'PASS':
                new_row[6].paragraphs[0].runs[0].font.color.rgb = RGBColor(0, 128, 0)
            else:
                new_row[6].paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 0, 0)
        
        self.doc.add_paragraph()
    
    def add_detailed_calculations(self, detailed_calcs):
        self.doc.add_page_break()
        heading = self.doc.add_heading('DETAILED CABLE CALCULATIONS', level=1)
        heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
        heading.runs[0].font.color.rgb = RGBColor(0, 51, 102)
        
        for i, calc in enumerate(detailed_calcs):
            if i > 0:
                self.doc.add_paragraph()
            
            self.doc.add_heading(f'Load {i+1}: {calc["load_name"]} ({format_load_type(calc.get("load_type", "Continuous"))})', level=2)
            
            self.doc.add_heading('Step 1: Load current calculation', level=3)
            p = self.doc.add_paragraph()
            p.add_run('Formula: I = P x 1000 / (1.732 x V x PF) for 3-phase').bold = True
            p = self.doc.add_paragraph()
            p.add_run('Calculation: ').bold = True
            p.add_run(f'I = {calc["power"]} x 1000 / (1.732 x {calc["voltage"]} x {calc["pf"]}) = {calc["current"]:.1f} A')
            
            self.doc.add_heading('Step 2: Cable type selection', level=3)
            p = self.doc.add_paragraph()
            p.add_run(f'Based on voltage {calc["voltage"]}V -> {calc["cable_category"]} cables selected')
            
            self.doc.add_heading('Step 3: Derating factors calculation', level=3)
            self.doc.add_paragraph('Derating Factor Formulas:')
            self.doc.add_paragraph('Total derating factor K = k1 × k2 × k3 × k4')
            self.doc.add_paragraph('k1 (Temperature correction): factor based on ambient temperature and insulation type')
            self.doc.add_paragraph('k2 (Grouping correction): factor based on number of circuits and installation arrangement')
            self.doc.add_paragraph('k3 (Soil resistivity correction) (for buried cables): factor based on soil thermal resistivity (K.m/W)')
            self.doc.add_paragraph('k4 (Depth correction) (for buried cables): factor based on burial depth (m)')
            self.doc.add_paragraph('Derated ampacity = Base ampacity × K')
            self.doc.add_paragraph()
            
            self.doc.add_paragraph(f'k1 (Temperature correction) : {calc["k1"]:.3f} - at {calc["ambient_temp"]}°C')
            self.doc.add_paragraph(f'k2 (Grouping)              : {calc["k2"]:.3f} - {format_cable_arrangement(calc["arrangement"])}')
            self.doc.add_paragraph(f'k3 (Soil resistivity)      : {calc["k3"]:.3f}')
            self.doc.add_paragraph(f'k4 (Depth)                 : {calc["k4"]:.3f}')
            p = self.doc.add_paragraph()
            p.add_run('Total K = ').bold = True
            p.add_run(f'{calc["total_k"]:.3f}')
            
            self.doc.add_heading('Step 4: Cable selection (automatic)', level=3)
            p = self.doc.add_paragraph()
            p.add_run('Selected cable: ').bold = True
            p.add_run(f'{calc["size"]} mm² {format_cable_type(calc["cable_type"])} ({format_insulation_type(calc["insulation_type"])})')
            p = self.doc.add_paragraph()
            p.add_run('Base ampacity: ').bold = True
            p.add_run(f'{calc["base_amp"]} A')
            p = self.doc.add_paragraph()
            p.add_run('Derated ampacity: ').bold = True
            p.add_run(f'{calc["derated_amp"]:.1f} A')
            p = self.doc.add_paragraph()
            status = 'PASS' if calc['ampacity_pass'] else 'FAIL'
            p.add_run('Check: ').bold = True
            check = p.add_run(f'{calc["derated_amp"]:.1f} A >= {calc["current"]:.1f} A ? {status}')
            if status == 'PASS':
                check.font.color.rgb = RGBColor(0, 128, 0)
            else:
                check.font.color.rgb = RGBColor(255, 0, 0)
            
            if 'trials' in calc and calc['trials']:
                self.doc.add_heading('Step 4.5: Cable selection trials (all attempted sizes)', level=3)
                
                trial_table = self.doc.add_table(rows=1, cols=6)
                trial_table.style = 'Light Grid Accent 1'
                
                headers = ['Size (mm²)', 'Base Ampacity (A)', 'Derated (A)', 'Vd %', 'Ampacity Check', 'VD Check']
                for j, header in enumerate(headers):
                    trial_table.rows[0].cells[j].text = header
                    trial_table.rows[0].cells[j].paragraphs[0].runs[0].bold = True
                
                for trial in calc['trials']:
                    row = trial_table.add_row().cells
                    row[0].text = str(trial['size'])
                    row[1].text = str(trial.get('ampacity', trial.get('base_amp', 0)))
                    row[2].text = f"{trial['derated']:.1f}"
                    row[3].text = f"{trial['vd_pct']:.2f}"
                    row[4].text = 'PASS' if trial['ampacity_pass'] else 'FAIL'
                    row[5].text = 'PASS' if trial['vd_pass'] else 'FAIL'
                    
                    if trial['ampacity_pass']:
                        row[4].paragraphs[0].runs[0].font.color.rgb = RGBColor(0, 128, 0)
                    else:
                        row[4].paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 0, 0)
                    
                    if trial['vd_pass']:
                        row[5].paragraphs[0].runs[0].font.color.rgb = RGBColor(0, 128, 0)
                    else:
                        row[5].paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 0, 0)
                
                self.doc.add_paragraph()
            
            self.doc.add_heading('Step 5: Voltage drop calculation', level=3)
            
            # Add voltage drop formula
            vd_formula = self.doc.add_paragraph()
            if calc['phase'] == '3-phase':
                vd_formula.add_run('Formula (3-phase): ').bold = True
                vd_formula.add_run('Vd = √3 × I × (R cosφ + X sinφ) × L / 1000')
            else:
                vd_formula.add_run('Formula (1-phase): ').bold = True
                vd_formula.add_run('Vd = 2 × I × (R cosφ + X sinφ) × L / 1000')
            
            # Get R and X values from cable database
            cable_db = get_cable_data(calc['insulation_type'], calc['cable_type'])
            cable_data = cable_db.get(calc['size'], {})
            
            r_value = cable_data.get('R', 0)
            
            if calc['formation'] == 'trefoil':
                x_value = cable_data.get('X_trefoil', 0)
            elif calc['formation'] == 'spaced':
                x_value = cable_data.get('X_spaced', 0)
            else:
                x_value = cable_data.get('X_flat_touching', 0)
            
            # For multi-core cables with single X value
            if x_value == 0:
                x_value = cable_data.get('X', 0)
            
            if r_value == 0:
                if calc['size'] <= 16:
                    r_value = 1.15
                elif calc['size'] <= 35:
                    r_value = 0.73
                elif calc['size'] <= 95:
                    r_value = 0.44
                else:
                    r_value = 0.19
                self.doc.add_paragraph(f'ℹ️ Using typical R = {r_value} Ω/km (actual cable data not available)')
            
            if x_value == 0:
                x_value = 0.08
                self.doc.add_paragraph(f'ℹ️ Using typical X = {x_value} Ω/km (actual cable data not available)')
            
            phi = math.acos(calc['pf'])
            sin_phi = math.sin(phi)
            
            if calc['phase'] == '3-phase':
                vd_calc = 1.732 * calc['current'] * (r_value * calc['pf'] + x_value * sin_phi) * calc['length'] / 1000
                vd_pct_calc = (vd_calc / calc['voltage']) * 100
                
                self.doc.add_paragraph(f'Given: I={calc["current"]:.1f}A, L={calc["length"]:.0f}m, cosφ={calc["pf"]}, V={calc["voltage"]:.0f}V, R={r_value:.4f}Ω/km, X={x_value:.4f}Ω/km')
                self.doc.add_paragraph(f'Step 5.1: sinφ = √(1 - cos²φ) = √(1 - {calc["pf"]:.3f}²) = {sin_phi:.4f}')
                self.doc.add_paragraph(f'Step 5.2: R cosφ = {r_value:.4f} × {calc["pf"]:.3f} = {r_value * calc["pf"]:.4f}')
                self.doc.add_paragraph(f'Step 5.3: X sinφ = {x_value:.4f} × {sin_phi:.4f} = {x_value * sin_phi:.4f}')
                self.doc.add_paragraph(f'Step 5.4: (R cosφ + X sinφ) = {r_value * calc["pf"]:.4f} + {x_value * sin_phi:.4f} = {(r_value * calc["pf"] + x_value * sin_phi):.4f}')
                self.doc.add_paragraph(f'Step 5.5: Vd = 1.732 × {calc["current"]:.1f} × {(r_value * calc["pf"] + x_value * sin_phi):.4f} × {calc["length"]:.0f} / 1000 = {vd_calc:.2f} V')
                self.doc.add_paragraph(f'Step 5.6: Vd% = ({vd_calc:.2f} / {calc["voltage"]:.0f}) × 100 = {vd_pct_calc:.3f}%')
            else:
                vd_calc = 2 * calc['current'] * (r_value * calc['pf'] + x_value * sin_phi) * calc['length'] / 1000
                vd_pct_calc = (vd_calc / calc['voltage']) * 100
                
                self.doc.add_paragraph(f'Given: I={calc["current"]:.1f}A, L={calc["length"]:.0f}m, cosφ={calc["pf"]}, V={calc["voltage"]:.0f}V, R={r_value:.4f}Ω/km, X={x_value:.4f}Ω/km')
                self.doc.add_paragraph(f'Step 5.1: sinφ = √(1 - cos²φ) = √(1 - {calc["pf"]:.3f}²) = {sin_phi:.4f}')
                self.doc.add_paragraph(f'Step 5.2: R cosφ = {r_value:.4f} × {calc["pf"]:.3f} = {r_value * calc["pf"]:.4f}')
                self.doc.add_paragraph(f'Step 5.3: X sinφ = {x_value:.4f} × {sin_phi:.4f} = {x_value * sin_phi:.4f}')
                self.doc.add_paragraph(f'Step 5.4: (R cosφ + X sinφ) = {r_value * calc["pf"]:.4f} + {x_value * sin_phi:.4f} = {(r_value * calc["pf"] + x_value * sin_phi):.4f}')
                self.doc.add_paragraph(f'Step 5.5: Vd = 2 × {calc["current"]:.1f} × {(r_value * calc["pf"] + x_value * sin_phi):.4f} × {calc["length"]:.0f} / 1000 = {vd_calc:.2f} V')
                self.doc.add_paragraph(f'Step 5.6: Vd% = ({vd_calc:.2f} / {calc["voltage"]:.0f}) × 100 = {vd_pct_calc:.3f}%')
            
            self.doc.add_paragraph()
            p = self.doc.add_paragraph()
            p.add_run('Voltage drop for this load: ').bold = True
            p.add_run(f'{calc["vd_pct"]:.3f}%')
            p = self.doc.add_paragraph()
            p.add_run('Maximum allowable voltage drop: ').bold = True
            p.add_run('2.5%')
            p = self.doc.add_paragraph()
            status = 'PASS' if calc['vd_pass'] else 'FAIL'
            p.add_run('Check: ').bold = True
            check = p.add_run(f'{calc["vd_pct"]:.3f}% <= 2.5% ? {status}')
            if status == 'PASS':
                check.font.color.rgb = RGBColor(0, 128, 0)
            else:
                check.font.color.rgb = RGBColor(255, 0, 0)
            
            self.doc.add_heading('Step 6: Short circuit calculation', level=3)
            p = self.doc.add_paragraph()
            p.add_run('Formula: Isc = (K × S / √t) × √(ln((θf + β) / (θi + β)))').bold = True
            p = self.doc.add_paragraph()
            p.add_run('Where: ').bold = True
            p.add_run(f'K = 226 (Copper), β = 234.5, θi = {calc["theta_i"]:.0f}°C, θf = {calc["theta_f"]:.0f}°C')
            p = self.doc.add_paragraph()
            p.add_run('Calculation: ').bold = True
            p.add_run(f'Isc = (226 × {calc["size"]} / √1.0) × √(ln(({calc["theta_f"]:.0f} + 234.5) / ({calc["theta_i"]:.0f} + 234.5)))')
            p = self.doc.add_paragraph()
            p.add_run(f'Isc = {calc["sc"]:.2f} kA')
            
            self.doc.add_heading('Step 7: Circuit breaker sizing', level=3)
            cb_formula = self.doc.add_paragraph()
            cb_formula.add_run('Formula: Icb = Iload × 1.25 (25% safety margin)').bold = True
            cb_example = self.doc.add_paragraph()
            cb_example.add_run('Example: Load current = 40 A → Required CB = 40 × 1.25 = 50 A → Selected = 50 A MCB')
            p = self.doc.add_paragraph()
            p.add_run('For this load: ').bold = True
            p.add_run(f'Icb = {calc["current"]:.1f} × 1.25 = {calc["current"] * 1.25:.1f} A → Selected CB = next standard rating')
            
            self.doc.add_heading('Final status', level=3)
            p = self.doc.add_paragraph()
            if calc['status'] == 'PASS':
                final = p.add_run('PASS')
                final.font.color.rgb = RGBColor(0, 128, 0)
            else:
                final = p.add_run('FAIL')
                final.font.color.rgb = RGBColor(255, 0, 0)
            final.font.size = Pt(14)
            final.font.bold = True
            
            self.doc.add_paragraph('_' * 60)
    
    def add_cb_results(self, cb_results, main_cb, pole_selections, main_pole):
        self.doc.add_page_break()
        heading = self.doc.add_heading('CIRCUIT BREAKER SIZING', level=1)
        heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
        heading.runs[0].font.color.rgb = RGBColor(0, 51, 102)
        
        self.doc.add_heading('Circuit Breaker Sizing Calculation Example', level=2)
        self.doc.add_paragraph('Formula: Icb = Iload × 1.25 (25% safety margin)')
        self.doc.add_paragraph('Example: Load current = 40 A')
        self.doc.add_paragraph('Required CB = 40 × 1.25 = 50 A')
        self.doc.add_paragraph('Selected standard rating = 50 A')
        self.doc.add_paragraph('Breaker type = MCB (≤125A)')
        self.doc.add_paragraph()
        
        self.doc.add_heading('Individual circuit breakers', level=2)
        table = self.doc.add_table(rows=1, cols=7)
        table.style = 'Light Grid Accent 1'
        
        headers = ['Load', 'Power (kw)', 'Current (a)', 'Required (a)', 'Selected (a)', 'Type', 'Poles']
        for i, header in enumerate(headers):
            table.rows[0].cells[i].text = header
            table.rows[0].cells[i].paragraphs[0].runs[0].bold = True
        
        for r in cb_results:
            row = table.add_row().cells
            row[0].text = r['Load']
            row[1].text = f"{r['Power (kW)']:.1f}"
            row[2].text = f"{r['Current (A)']:.1f}"
            row[3].text = f"{r['Required CB (A)']:.1f}"
            row[4].text = str(r['Selected CB (A)'])
            row[5].text = r['Breaker Type']
            row[6].text = pole_selections.get(r['Load'], '3P')
        
        self.doc.add_heading('Main circuit breaker', level=2)
        self.doc.add_paragraph(f'Total power: {main_cb["total_power"]:.1f} kW')
        self.doc.add_paragraph(f'Total current: {main_cb["current"]:.1f} A')
        self.doc.add_paragraph(f'Required rating: {main_cb["required_cb"]:.1f} A')
        p = self.doc.add_paragraph()
        p.add_run(f'Selected main cb: {main_cb["selected_cb"]} A {main_cb["breaker_type"]} {main_pole}').bold = True
    
    def save(self, filename):
        self.doc.save(filename)

class TransformerWordReport:
    def __init__(self):
        self.doc = Document()
        style = self.doc.styles['Normal']
        style.font.name = 'Arial'
        style.font.size = Pt(11)
    
    def add_title(self):
        title = self.doc.add_heading('TRANSFORMER SIZING REPORT', 0)
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        title.runs[0].font.size = Pt(20)
        title.runs[0].font.color.rgb = RGBColor(0, 51, 102)
        self.doc.add_paragraph(f'Date: {datetime.now().strftime("%Y-%m-%d %H:%M")}')
        self.doc.add_paragraph()
    
    def add_load_analysis(self, loads_df):
        heading = self.doc.add_heading('LOAD ANALYSIS', level=1)
        heading.runs[0].font.color.rgb = RGBColor(0, 51, 102)
        
        table = self.doc.add_table(rows=1, cols=6)
        table.style = 'Light Grid Accent 1'
        
        headers = ['Load description', 'Qty', 'Rating (kw)', 'Connected (kw)', 'Diversity', 'P (kw)']
        for i, header in enumerate(headers):
            table.rows[0].cells[i].text = header
            table.rows[0].cells[i].paragraphs[0].runs[0].bold = True
        
        total_p = 0
        for idx, load in loads_df.iterrows():
            load_type_diversity = LOAD_TYPE_FACTORS[load['Load Type']]['diversity']
            connected = load['Rating (kW)'] * load['Quantity']
            p = connected * load_type_diversity
            total_p += p
            
            row = table.add_row().cells
            row[0].text = load['Load Description']
            row[1].text = str(load['Quantity'])
            row[2].text = f"{load['Rating (kW)']:.0f}"
            row[3].text = f"{connected:.0f}"
            row[4].text = f"{load_type_diversity:.1f}"
            row[5].text = f"{p:.1f}"
        
        p_row = table.add_row().cells
        p_row[0].text = 'Total real power (P)'
        p_row[0].paragraphs[0].runs[0].bold = True
        p_row[1].text = ''
        p_row[2].text = ''
        p_row[3].text = ''
        p_row[4].text = ''
        p_row[5].text = f"{total_p:.1f} kW"
        p_row[5].paragraphs[0].runs[0].bold = True
        
        self.doc.add_paragraph()
        return total_p
    
    def add_step_by_step(self, loads_df):
        self.doc.add_page_break()
        heading = self.doc.add_heading('STEP-BY-STEP P, Q, S CALCULATIONS', level=1)
        heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
        heading.runs[0].font.color.rgb = RGBColor(0, 51, 102)
        
        total_p = 0
        total_q = 0
        
        for idx, load in loads_df.iterrows():
            self.doc.add_heading(f'Load {idx+1}: {load["Load Description"]}', level=2)
            
            load_type_diversity = LOAD_TYPE_FACTORS[load['Load Type']]['diversity']
            connected = load['Rating (kW)'] * load['Quantity']
            p = connected * load_type_diversity
            phi = math.acos(load['Power Factor'])
            tan_phi = math.tan(phi)
            q = p * tan_phi
            s = math.sqrt(p**2 + q**2)
            
            self.doc.add_paragraph(f'Step 1 - Connected power: {load["Rating (kW)"]:.0f} kW x {load["Quantity"]} = {connected:.0f} kW')
            self.doc.add_paragraph(f'Step 2 - Demand power (P): {connected:.0f} kW x {load_type_diversity:.1f} = {p:.1f} kW')
            self.doc.add_paragraph(f'Step 3 - Angle φ: acos({load["Power Factor"]}) = {math.degrees(phi):.1f}°')
            self.doc.add_paragraph(f'Step 4 - tan(φ): tan({math.degrees(phi):.1f}°) = {tan_phi:.3f}')
            self.doc.add_paragraph(f'Step 5 - Reactive power (Q): {p:.1f} kW x {tan_phi:.3f} = {q:.1f} kVAR')
            p_step = self.doc.add_paragraph()
            p_step.add_run('Step 6 - Apparent power (S): ').bold = True
            p_step.add_run(f'√({p:.1f}² + {q:.1f}²) = {s:.1f} kVA')
            
            total_p += p
            total_q += q
            
            self.doc.add_paragraph('_' * 50)
        
        return total_p, total_q
    
    def add_largest_equipment(self, loads_df, total_p, total_s, largest_data):
        self.doc.add_page_break()
        heading = self.doc.add_heading('LARGEST EQUIPMENT ANALYSIS', level=1)
        heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
        heading.runs[0].font.color.rgb = RGBColor(0, 51, 102)
        
        if largest_data:
            self.doc.add_heading(f'Largest equipment: {largest_data["load"]["Load Description"]}', level=2)
            self.doc.add_paragraph(f'Connected power: {largest_data["connected"]:.0f} kW')
            self.doc.add_paragraph(f'Demand power (P): {largest_data["p"]:.1f} kW')
            self.doc.add_paragraph(f'Reactive power (Q): {largest_data["q"]:.1f} kVAR')
            self.doc.add_paragraph(f'Apparent power (S): {largest_data["s"]:.1f} kVA')
            
            p_pct = (largest_data["p"] / total_p) * 100 if total_p > 0 else 0
            s_pct = (largest_data["s"] / total_s) * 100 if total_s > 0 else 0
            
            self.doc.add_heading('Impact on total system:', level=3)
            self.doc.add_paragraph(f'• Contributes {p_pct:.1f}% of total real power (P)')
            self.doc.add_paragraph(f'• Contributes {s_pct:.1f}% of total apparent power (S)')
    
    def add_transformer_selection(self, total_p, total_q, future_expansion, selected_kva):
        self.doc.add_page_break()
        heading = self.doc.add_heading('TRANSFORMER SELECTION', level=1)
        heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
        heading.runs[0].font.color.rgb = RGBColor(0, 51, 102)
        
        total_s = math.sqrt(total_p**2 + total_q**2)
        with_future = total_s * (1 + future_expansion/100)
        
        self.doc.add_paragraph(f'Total real power (P) = {total_p:.1f} kW')
        self.doc.add_paragraph(f'Total reactive power (Q) = {total_q:.1f} kVAR')
        p = self.doc.add_paragraph()
        p.add_run(f'Total apparent power (S) = √({total_p:.1f}² + {total_q:.1f}²) = {total_s:.1f} kVA').bold = True
        
        self.doc.add_paragraph()
        self.doc.add_paragraph(f'Future expansion: +{future_expansion}%')
        self.doc.add_paragraph(f'Required with future = {total_s:.1f} x {1 + future_expansion/100:.2f} = {with_future:.1f} kVA')
        self.doc.add_paragraph()
        
        self.doc.add_heading('Standard ratings:', level=3)
        ratings = [50, 100, 160, 200, 250, 315, 400, 500, 630, 800, 1000, 1250, 1600, 2000, 2500, 3150]
        self.doc.add_paragraph(', '.join(str(r) for r in ratings))
        
        final_heading = self.doc.add_heading('', level=2)
        final_heading.add_run(f'Selected transformer: {selected_kva} kVA').bold = True
        final_heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    def save(self, filename):
        self.doc.save(filename)

class SimpleTransformerCalculator:
    def __init__(self):
        self.standard_ratings = [50, 100, 160, 200, 250, 315, 400, 500, 630, 800, 
                                  1000, 1250, 1600, 2000, 2500, 3150, 4000, 5000, 
                                  6300, 8000, 10000, 12500, 16000, 20000, 25000, 
                                  31500, 40000, 50000, 63000]
    
    def calculate_p(self, rating_kw, quantity, diversity):
        return rating_kw * quantity * diversity
    
    def calculate_q(self, p_kw, pf):
        if pf >= 1.0:
            return 0
        phi = math.acos(pf)
        return p_kw * math.tan(phi)
    
    def calculate_s(self, p_kw, q_kvar):
        return math.sqrt(p_kw**2 + q_kvar**2)
    
    def get_standard_rating(self, required_kva):
        for rating in self.standard_ratings:
            if rating >= required_kva:
                return rating
        return self.standard_ratings[-1]
    
    def find_largest_equipment(self, loads):
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

# ========== UNIVERSAL LOAD SHEET ==========
if 'universal_loads' not in st.session_state:
    st.session_state.universal_loads = pd.DataFrame({
        'Load Description': ['LV Motor', 'MV Motor'],
        'Quantity': [1, 1],
        'Rating (kW)': [75, 500],
        'Voltage (V)': [415, 3300],
        'Power Factor': [0.85, 0.85],
        'Load Type': ['Continuous', 'Continuous'],
        'Diversity Factor': [0.8, 0.8]
    })

if 'loads_df' not in st.session_state:
    st.session_state.loads_df = pd.DataFrame(columns=['Load Name', 'Power (kW)', 'Voltage (V)', 'Phase', 'Load Type', 'Power Factor', 'Efficiency', 'Length (m)'])

if 'cable_results_df' not in st.session_state:
    st.session_state.cable_results_df = pd.DataFrame()
if 'detailed_calcs' not in st.session_state:
    st.session_state.detailed_calcs = []
if 'all_derating_factors' not in st.session_state:
    st.session_state.all_derating_factors = {}
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
if 'total_p' not in st.session_state:
    st.session_state.total_p = 0
if 'total_q' not in st.session_state:
    st.session_state.total_q = 0
if 'tx_largest_data' not in st.session_state:
    st.session_state.tx_largest_data = None

# ========== SIDEBAR NAVIGATION ==========
with st.sidebar:
    st.markdown('<div style="background: linear-gradient(135deg, #1E3A8A 0%, #3B5BA6 100%); color: white; padding: 15px; border-radius: 10px; margin-bottom: 20px; text-align: center;"><h2 style="color: white !important; margin: 0;">⚡ CES-Electrical</h2></div>', unsafe_allow_html=True)
    
    if 'selected_calculator' not in st.session_state:
        st.session_state.selected_calculator = "📋 LOAD SHEET"
    
    calculators = [
        "📋 LOAD SHEET",
        "⚡ Lightning Protection",
        "🔌 Cable Sizing",
        "⚙️ Transformer Sizing",
        "🔄 Generator Sizing",
        "⏚ Earthing"
    ]
    
    for calc in calculators:
        if st.button(calc, key=f"nav_{calc}", use_container_width=True):
            st.session_state.selected_calculator = calc
            st.rerun()

# ========== MAIN CONTENT ==========
st.title(f"{st.session_state.selected_calculator} Calculator")

# ========== LOAD SHEET ==========
if st.session_state.selected_calculator == "📋 LOAD SHEET":
    
    st.markdown('<div class="report-header">📋 UNIVERSAL LOAD SHEET</div>', unsafe_allow_html=True)
    
    st.markdown("""
    <div class="info-box">
        <h4>📌 This load sheet is used by all calculators</h4>
        <p>Two example loads provided: LV Motor (415V) and MV Motor (3300V). Load Type determines diversity factor:</p>
        <p><span class="load-type-badge continuous-badge">Continuous</span> = 100% (1.0) | 
           <span class="load-type-badge intermittent-badge">Intermittent</span> = 30% (0.3) | 
           <span class="load-type-badge standby-badge">Standby</span> = 10% (0.1)</p>
    </div>
    """, unsafe_allow_html=True)
    
    with st.expander("📊 tan(acos(PF)) reference table", expanded=False):
        tan_data = {
            'Power factor': [1.0, 0.95, 0.90, 0.85, 0.80, 0.75, 0.70],
            'tan(acos(PF))': [0.00, 0.33, 0.48, 0.62, 0.75, 0.88, 1.02],
            'Example': ['PF=1.0 -> Q=0', 'PF=0.95 -> Q=0.33xP', 'PF=0.90 -> Q=0.48xP', 
                       'PF=0.85 -> Q=0.62xP', 'PF=0.80 -> Q=0.75xP', 'PF=0.75 -> Q=0.88xP', 
                       'PF=0.70 -> Q=1.02xP']
        }
        tan_df = pd.DataFrame(tan_data)
        st.dataframe(tan_df, use_container_width=True, hide_index=True)
        st.caption("Formula: Q = P x tan(acos(PF)) as per document")
    
    col1, col2 = st.columns([1, 1])
    with col1:
        if st.button("➕ Add load", key="add_load_main", use_container_width=True):
            new_row = pd.DataFrame({
                'Load Description': [f'Load {len(st.session_state.universal_loads) + 1}'],
                'Quantity': [1],
                'Rating (kW)': [50.0],
                'Voltage (V)': [415],
                'Power Factor': [0.85],
                'Load Type': ['Continuous'],
                'Diversity Factor': [0.8]
            })
            st.session_state.universal_loads = pd.concat([st.session_state.universal_loads, new_row], ignore_index=True)
            st.rerun()
    
    with col2:
        if st.button("🗑️ Delete last load", key="delete_load_main", use_container_width=True):
            if len(st.session_state.universal_loads) > 1:
                st.session_state.universal_loads = st.session_state.universal_loads[:-1]
                st.rerun()
            else:
                st.warning("At least one row required")
    
    st.markdown("### 📋 Current loads")
    
    for idx, load in st.session_state.universal_loads.iterrows():
        load_type_factor = LOAD_TYPE_FACTORS[load['Load Type']]['diversity']
        badge_color = "continuous-badge" if load['Load Type'] == 'Continuous' else "intermittent-badge" if load['Load Type'] == 'Intermittent' else "standby-badge"
        cable_type = "LV" if load['Voltage (V)'] <= 1000 else "MV"
        
        st.markdown(f"""
        <div class="calc-step">
            <h4>📌 Load {idx+1}: <span style="color: #1E3A8A;">{load['Load Description']}</span></h4>
            <p><b>Quantity:</b> {load['Quantity']} | <b>Rating:</b> {load['Rating (kW)']} kW | <b>Voltage:</b> {load['Voltage (V)']}V ({cable_type})</p>
            <p><b>PF:</b> {load['Power Factor']} | <b>Load type:</b> <span class="load-type-badge {badge_color}">{load['Load Type']} ({load_type_factor*100:.0f}%)</span> | <b>Diversity:</b> {load['Diversity Factor']}</p>
        </div>
        """, unsafe_allow_html=True)
    
    edited_loads = st.data_editor(
        st.session_state.universal_loads,
        num_rows="fixed",
        use_container_width=True,
        column_config={
            "Load Description": st.column_config.TextColumn("Load description", width="medium"),
            "Quantity": st.column_config.NumberColumn("Qty", min_value=1, max_value=100, step=1),
            "Rating (kW)": st.column_config.NumberColumn("Rating (kw)", min_value=0.0, max_value=10000.0, step=1.0),
            "Voltage (V)": st.column_config.NumberColumn("Voltage (v)", min_value=0, max_value=11000, step=100),
            "Power Factor": st.column_config.NumberColumn("Pf", min_value=0.5, max_value=1.0, step=0.05),
            "Load Type": st.column_config.SelectboxColumn("Load type", options=['Continuous', 'Intermittent', 'Standby']),
            "Diversity Factor": st.column_config.NumberColumn("Diversity", min_value=0.0, max_value=1.0, step=0.05)
        }
    )
    st.session_state.universal_loads = edited_loads
    
    total_connected = 0
    total_p = 0
    lv_count = 0
    mv_count = 0
    
    summary_data = []
    for idx, load in st.session_state.universal_loads.iterrows():
        load_type_diversity = LOAD_TYPE_FACTORS[load['Load Type']]['diversity']
        
        connected = load['Rating (kW)'] * load['Quantity']
        p = connected * load_type_diversity
        total_connected += connected
        total_p += p
        
        if load['Voltage (V)'] <= 1000:
            lv_count += 1
        else:
            mv_count += 1
        
        summary_data.append({
            'Load': load['Load Description'],
            'Connected (kw)': f"{connected:.0f}",
            'Demand (kw)': f"{p:.0f} ({load['Load Type']} {load_type_diversity*100:.0f}%)",
            'Pf': f"{load['Power Factor']:.2f}",
            'Type': load['Load Type']
        })
    
    summary_df = pd.DataFrame(summary_data)
    st.dataframe(summary_df, use_container_width=True, hide_index=True)
    
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("Total connected", f"{total_connected:.0f} kW")
    with col2:
        st.metric("Total demand", f"{total_p:.0f} kW")
    with col3:
        st.metric("LV loads", lv_count)
    with col4:
        st.metric("MV loads", mv_count)

# ========== LIGHTNING PROTECTION ==========
elif st.session_state.selected_calculator == "⚡ Lightning Protection":
    
    lp_tabs = st.tabs(["📊 Risk assessment", "🔧 Protection design", "📋 Calculations", "📥 Download report"])
    
    with lp_tabs[0]:
        st.markdown('<div class="report-header">RISK ASSESSMENT (IEC 62305-2)</div>', unsafe_allow_html=True)
        
        structure_type = st.selectbox("Select structure type", 
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
            
            td_days = st.number_input("Thunderstorm days/year", value=10, step=1)
            environment = st.selectbox("Environment", ["Surrounded", "Similar height", "Isolated", "Hilltop"])
        
        with col2:
            st.markdown("### 📊 Environmental factor (CD)")
            cd_values = {"Surrounded": 0.25, "Similar height": 0.5, "Isolated": 1, "Hilltop": 2}
            cd = cd_values[environment]
            
            st.markdown("**IEC 62305-2 Table A.1:**")
            st.markdown("• Surrounded: **0.25** • Similar: **0.5** • Isolated: **1.0** • Hilltop: **2.0**")
            st.success(f"**Selected: {environment} -> CD = {cd}**")
            
            if structure_type == "Column 4-C01":
                c2, c3, c4, c5 = 0.5, 2.0, 3.0, 10.0
            else:
                c2, c3, c4, c5 = 1.0, 3.0, 1.0, 5.0
            
            st.metric("C2 - Type", c2)
            st.metric("C3 - Content", c3)
            st.metric("C4 - Occupancy", c4)
            st.metric("C5 - Consequence", c5)
        
        if st.button("🔧 Calculate risk", type="primary", use_container_width=True):
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
                st.metric("Collection area (Ad)", f"{ad:.0f} m²")
                st.metric("Near strike area (Am)", f"{am:.0f} m²")
            with col_b:
                st.metric("Nd (Direct)", f"{nd:.6f}")
                st.metric("Nm (Near)", f"{nm:.6f}")
            with col_c:
                st.metric("Protection level", lpl)
                st.metric("Efficiency", f"{efficiency:.1%}")
            with col_d:
                st.metric("Rolling sphere", f"{sphere}m")
                st.metric("Air terminals", air_terminals)
            
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
            st.warning("⚠️ Please complete risk assessment first!")
        else:
            results = st.session_state.calc_results
            st.success(f"✅ Designing for: **{results['lpl']}**")
            
            col1, col2 = st.columns(2)
            with col1:
                st.metric("Air terminals", results['air_terminals'])
                st.metric("Rolling sphere", f"{results['sphere']}m")
            with col2:
                if results['lpl'] in ["Class I", "Class II"]:
                    st.metric("Rod diameter", "12.7 mm")
                    st.metric("Down conductor", "58 mm²")
                else:
                    st.metric("Rod diameter", "9.5 mm")
                    st.metric("Down conductor", "29 mm²")
    
    with lp_tabs[2]:
        st.markdown('<div class="report-header">DETAILED CALCULATIONS</div>', unsafe_allow_html=True)
        if not st.session_state.calc_done:
            st.warning("⚠️ Please complete risk assessment first!")
        else:
            results = st.session_state.calc_results
            inputs = st.session_state.input_values
            
            with st.expander("1. Collection area (Ad)", expanded=True):
                st.markdown("**Formula:** Ad = L x W + 2 x (3H) x (L + W) + π x (3H)²")
                st.markdown("**Reference:** IEC 62305-2 Annex A.2.1.1")
                st.markdown(f"**Result:** Ad = **{results['ad']:.2f} m²**")
            
            with st.expander("2. Near strike collection area (Am)", expanded=True):
                st.markdown("**Formula:** Am = 2 x 500 x (L + W) + π x 500²")
                st.markdown("**Reference:** IEC 62305-2 Annex A.3")
                st.markdown(f"**Result:** Am = **{results['am']:.2f} m²**")
            
            with st.expander("3. Environmental factor (CD)"):
                st.markdown(f"**Selected:** {inputs.get('environment', 'Isolated')} -> **{inputs.get('cd', 1)}**")
            
            with st.expander("4. Lightning density (NG)"):
                st.markdown(f"**Result:** NG = **{results.get('ng', 1)} flashes/km²/year**")
            
            with st.expander("5. Lightning frequencies"):
                st.markdown(f"**Nd:** {results.get('nd', 0):.6f} events/year")
                st.markdown(f"**Nm:** {results.get('nm', 0):.6f} events/year")
            
            with st.expander("6. Protection level"):
                st.markdown(f"**Efficiency:** {results.get('efficiency', 0):.1%}")
                st.markdown(f"**Result:** **{results.get('lpl', 'Class III')}**")
    
    with lp_tabs[3]:
        st.markdown('<div class="report-header">DOWNLOAD REPORT</div>', unsafe_allow_html=True)
        if not st.session_state.calc_done:
            st.warning("⚠️ Please complete risk assessment first!")
        else:
            if st.button("📥 Generate word report", key="lp_word", use_container_width=True):
                with st.spinner("Generating word report..."):
                    try:
                        word = LightningWordReport()
                        word.add_calculations(st.session_state.calc_results, st.session_state.input_values)
                        word_path = "temp_lightning.docx"
                        word.save(word_path)
                        with open(word_path, "rb") as f:
                            word_bytes = f.read()
                        b64 = base64.b64encode(word_bytes).decode()
                        filename = f"Lightning_Report_{datetime.now().strftime('%Y%m%d_%H%M')}.docx"
                        os.remove(word_path)
                        st.markdown(f'<a href="data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,{b64}" download="{filename}" class="download-btn">📥 Click here to download word report</a>', unsafe_allow_html=True)
                        st.success("✅ Word generated successfully!")
                    except Exception as e:
                        st.error(f"Error generating word document: {str(e)}")

# ========== CABLE SIZING ==========
elif st.session_state.selected_calculator == "🔌 Cable Sizing":
    
    st.markdown('<div class="report-header">🔌 Cable sizing calculator</div>', unsafe_allow_html=True)
    
    col1, col2 = st.columns(2)
    with col1:
        efficiency_value = st.number_input("Motor efficiency", value=1.0, min_value=0.5, max_value=1.0, step=0.05, format="%.2f")
        st.caption("1.0 = 100%, 0.95 = 95%")
    with col2:
        ambient_temp = st.number_input("Ambient temperature (°C) - common", 
                                      value=30.0, min_value=10.0, max_value=80.0, step=5.0,
                                      key="ambient_temp_global")
        st.info("Same for all cables")
    
    if st.button("📥 Import loads from load sheet", use_container_width=True):
        new_loads = []
        for idx, load in st.session_state.universal_loads.iterrows():
            if load['Voltage (V)'] > 300:
                phase = '3-phase'
            else:
                phase = '1-phase'
            
            load_type_diversity = LOAD_TYPE_FACTORS[load['Load Type']]['diversity']
            
            new_loads.append({
                'Load Name': load['Load Description'],
                'Power (kW)': load['Rating (kW)'] * load['Quantity'] * load_type_diversity,
                'Voltage (V)': load['Voltage (V)'],
                'Phase': phase,
                'Load Type': load['Load Type'],
                'Power Factor': load['Power Factor'],
                'Efficiency': efficiency_value,
                'Length (m)': 50,
                'Insulation Type': 'XLPE_90',
                'Cable Type': 'single_core_non_armoured',
                'Installation Method': 'C',
                'Cables in Group': 3,
                'Cable Arrangement': 'bunched_in_air_surface_enclosed',
                'Cable Formation': 'flat',
                'Soil Resistivity (K.m/W)': 1.5,
                'Burial Depth (m)': 0.8
            })
        st.session_state.loads_df = pd.DataFrame(new_loads)
        st.success("✅ Loads imported successfully!")
        st.rerun()
    
    cable_calc = CableSizingCalculator()
    
    cable_tabs = st.tabs([
        "📥 Loads and derating input", 
        "📊 Derating factors summary", 
        "🔌 Cable selection",
        "⚡ Circuit breakers",
        "📥 Download report"
    ])
    
    with cable_tabs[0]:
        st.markdown("### 📋 Input parameters")
        
        if st.session_state.loads_df.empty:
            st.info("No loads imported yet. Click 'Import loads from load sheet' above.")
        else:
            edited_df = st.data_editor(
                st.session_state.loads_df,
                num_rows="fixed",
                use_container_width=True,
                column_config={
                    "Load Name": st.column_config.TextColumn("Load name", disabled=True),
                    "Power (kW)": st.column_config.NumberColumn("Power (kw)", disabled=True),
                    "Voltage (V)": st.column_config.NumberColumn("Voltage (v)", disabled=True),
                    "Phase": st.column_config.TextColumn("Phase", disabled=True),
                    "Load Type": st.column_config.TextColumn("Load type", disabled=True),
                    "Power Factor": st.column_config.NumberColumn("Pf", disabled=True),
                    "Efficiency": st.column_config.NumberColumn("Efficiency", disabled=True),
                    "Length (m)": st.column_config.NumberColumn("Length (m)", min_value=1.0, max_value=5000.0, step=1.0),
                    "Insulation Type": st.column_config.SelectboxColumn("Insulation type", options=['XLPE_90', 'PVC_70'], format_func=format_insulation_type),
                    "Cable Type": st.column_config.SelectboxColumn("Cable type", options=['single_core_non_armoured', 'multi_core_non_armoured', 'single_core_armoured', 'multi_core_armoured'], format_func=format_cable_type),
                    "Installation Method": st.column_config.SelectboxColumn("Installation method", options=['B', 'C', 'D', 'E', 'F', 'G']),
                    "Cables in Group": st.column_config.NumberColumn("Cables in group", min_value=1, max_value=20, step=1),
                    "Cable Arrangement": st.column_config.SelectboxColumn("Cable arrangement", options=['bunched_in_air_surface_enclosed', 'single_layer_wall_floor', 'single_layer_perforated_tray', 'single_layer_ladder_cleats'], format_func=format_cable_arrangement),
                    "Cable Formation": st.column_config.SelectboxColumn("Cable formation", options=['flat', 'trefoil', 'spaced'], format_func=format_cable_formation),
                    "Soil Resistivity (K.m/W)": st.column_config.NumberColumn("Soil resistivity", min_value=0.5, max_value=3.0, step=0.1),
                    "Burial Depth (m)": st.column_config.NumberColumn("Burial depth (m)", min_value=0.3, max_value=2.0, step=0.1)
                }
            )
            st.session_state.loads_df = edited_df
        
        if st.button("🔧 Calculate with derating factors (auto selection)", type="primary", use_container_width=True):
            with st.spinner("Calculating with automatic cable selection..."):
                cable_results = []
                detailed_calcs = []
                all_factors = {}
                
                for idx, load in st.session_state.loads_df.iterrows():
                    cable_category, _ = cable_calc.get_cable_category(load['Voltage (V)'])
                    
                    insulation_type = load['Insulation Type']
                    insulation_temp = 90 if insulation_type == 'XLPE_90' else 70
                    
                    cable_db = get_cable_data(insulation_type, load['Cable Type'])
                    
                    if not cable_db:
                        st.warning(f"No cable data found for {load['Load Name']}")
                        continue
                    
                    current = cable_calc.calculate_load_current(
                        load['Power (kW)'], load['Voltage (V)'], load['Power Factor'], 
                        load['Efficiency'], load['Phase']
                    )
                    
                    selected_size, cable_data, base_amp, derated_amp, vd_pct, total_k, factors, success, trial_results = select_cable_automatically(
                        load, cable_db, cable_calc, ambient_temp,
                        insulation_temp, current,
                        load['Length (m)'], load['Power Factor'], load['Voltage (V)'], load['Phase'],
                        load['Installation Method'], load['Cable Formation'], load['Cable Type'],
                        load['Cable Arrangement'],
                        load['Soil Resistivity (K.m/W)'],
                        load['Burial Depth (m)'], load['Cables in Group']
                    )
                    
                    if selected_size is None:
                        st.error(f"No suitable cable found for {load['Load Name']}")
                        continue
                    
                    all_factors[load['Load Name']] = factors
                    
                    insulation_short = 'PVC' if insulation_type == 'PVC_70' else 'XLPE'
                    base_ampacity = base_amp
                    
                    total_k_actual, factors_actual = cable_calc.get_derating_factors(
                        ambient_temp, insulation_temp,
                        load['Cables in Group'], load['Cable Arrangement'],
                        load['Installation Method'],
                        load['Soil Resistivity (K.m/W)'], load['Burial Depth (m)']
                    )
                    
                    derated_amp_actual = base_ampacity * total_k_actual
                    
                    isc, k_value, operating_temp, theta_i, theta_f = cable_calc.calculate_short_circuit(
                        selected_size, insulation_short,
                        ambient_temp, current, base_ampacity,
                        factors_actual['k1 (Temperature)'],
                        factors_actual['k2 (Grouping)'],
                        factors_actual['k3 (Soil Resistivity)'],
                        factors_actual['k4 (Depth)'],
                        1.0
                    )
                    
                    status = 'PASS' if success else 'FAIL'
                    
                    cable_results.append({
                        'Load Name': load['Load Name'],
                        'Load Type': load.get('Load Type', 'Continuous'),
                        'Power (kW)': load['Power (kW)'],
                        'Voltage (V)': load['Voltage (V)'],
                        'Phase': load['Phase'],
                        'PF': load['Power Factor'],
                        'Efficiency': f"{load['Efficiency']*100:.0f}%",
                        'Length (m)': load['Length (m)'],
                        'Insulation': format_insulation_type(insulation_type),
                        'Cable Category': cable_category,
                        'Cable Type': format_cable_type(load['Cable Type']),
                        'Size (mm²)': selected_size,
                        'Load Current (A)': f"{current:.1f}",
                        'Base Ampacity (A)': base_ampacity,
                        'Derating Factor K': f"{total_k_actual:.3f}",
                        'Derated Ampacity (A)': f"{derated_amp_actual:.1f} A",
                        'Voltage Drop (%)': f"{vd_pct:.3f}%",
                        'Short Circuit (kA)': f"{isc/1000:.2f}",
                        'K Value': f"{k_value}",
                        'θi (°C)': f"{theta_i:.0f}",
                        'θf (°C)': f"{theta_f:.0f}",
                        'Operating Temp': f"{operating_temp:.1f}°C",
                        'Status': status,
                        'VD Limit': '2.5%',
                        'Check': 'PASS' if (vd_pct <= 2.5 and derated_amp_actual >= current) else 'FAIL'
                    })
                    
                    detailed_calcs.append({
                        'load_name': load['Load Name'],
                        'load_type': load.get('Load Type', 'Continuous'),
                        'power': load['Power (kW)'],
                        'voltage': load['Voltage (V)'],
                        'phase': load['Phase'],
                        'pf': load['Power Factor'],
                        'efficiency': load['Efficiency'],
                        'length': load['Length (m)'],
                        'current': current,
                        'size': selected_size,
                        'insulation_type': insulation_type,
                        'cable_category': cable_category,
                        'cable_type': load['Cable Type'],
                        'base_amp': base_ampacity,
                        'derated_amp': derated_amp_actual,
                        'vd_pct': vd_pct,
                        'sc': isc/1000,
                        'k_value': k_value,
                        'theta_i': theta_i,
                        'theta_f': theta_f,
                        'operating_temp': operating_temp,
                        'k1': factors_actual['k1 (Temperature)'],
                        'k2': factors_actual['k2 (Grouping)'],
                        'k3': factors_actual['k3 (Soil Resistivity)'],
                        'k4': factors_actual['k4 (Depth)'],
                        'total_k': total_k_actual,
                        'ambient_temp': ambient_temp,
                        'arrangement': load['Cable Arrangement'],
                        'formation': load['Cable Formation'],
                        'installation': load['Installation Method'],
                        'soil_res': load['Soil Resistivity (K.m/W)'],
                        'depth': load['Burial Depth (m)'],
                        'num_cables': load['Cables in Group'],
                        'status': status,
                        'vd_pass': vd_pct <= 2.5,
                        'ampacity_pass': derated_amp_actual >= current,
                        'trials': trial_results
                    })
                    
                    if not success:
                        st.warning(f"⚠️ {load['Load Name']}: Even largest cable {selected_size} mm² fails! vd={vd_pct:.2f}% > 2.5%")
                    else:
                        st.success(f"✅ {load['Load Name']}: Selected {selected_size} mm² cable (vd={vd_pct:.2f}%)")
                
                st.session_state.cable_results_df = pd.DataFrame(cable_results)
                st.session_state.detailed_calcs = detailed_calcs
                st.session_state.all_derating_factors = all_factors
                
                cb_calc = CircuitBreakerCalculator()
                manufacturer = 'Schneider Electric'
                cb_results, cb_details = cb_calc.calculate_cb_size(
                    st.session_state.loads_df, 1.25, manufacturer
                )
                main_cb = cb_calc.calculate_main_cb(st.session_state.loads_df, 400, 0.85, 1.25)
                
                st.session_state.cb_results = cb_results
                st.session_state.cb_details = cb_details
                st.session_state.main_cb = main_cb
                
                st.success("✅ Calculations complete with automatic cable selection!")
    
    with cable_tabs[1]:
        st.markdown('<div class="report-header">Derating factors summary</div>', unsafe_allow_html=True)
        
        st.markdown("""
### Derating Factor Formulas

**Total derating factor K = k1 × k2 × k3 × k4**

**k1 (Temperature correction)**
k1 = factor based on ambient temperature and insulation type

**k2 (Grouping correction)**
k2 = factor based on number of circuits and installation arrangement

**k3 (Soil resistivity correction) (for buried cables)**
k3 = factor based on soil thermal resistivity (K.m/W)

**k4 (Depth correction) (for buried cables)**
k4 = factor based on burial depth (m)

**Derated ampacity = Base ampacity × K**
""")
        
        if hasattr(st.session_state, 'all_derating_factors') and st.session_state.all_derating_factors:
            for load_name, factors in st.session_state.all_derating_factors.items():
                with st.expander(f"📊 {load_name} - Derating factors"):
                    st.markdown(f"""
| Factor | Value | Description |
|--------|-------|-------------|
| k1 (Temperature) | {factors['k1 (Temperature)']:.3f} | Based on ambient temperature and insulation type |
| k2 (Grouping) | {factors['k2 (Grouping)']:.3f} | Based on number of circuits and installation arrangement |
| k3 (Soil resistivity) | {factors['k3 (Soil Resistivity)']:.3f} | Based on soil thermal resistivity |
| k4 (Depth) | {factors['k4 (Depth)']:.3f} | Based on burial depth |
| **Total K** | **{factors['total']:.3f}** | **K = k1 × k2 × k3 × k4** |
""")
                    
                    for calc in st.session_state.detailed_calcs:
                        if calc['load_name'] == load_name:
                            st.markdown(f"""
**Installation parameters for {load_name}:**
- Insulation type: {format_insulation_type(calc['insulation_type'])}
- Cable type: {format_cable_type(calc['cable_type'])}
- Installation method: {calc['installation']}
- Cables in group: {calc['num_cables']}
- Cable arrangement: {format_cable_arrangement(calc['arrangement'])}
- Cable formation: {format_cable_formation(calc['formation'])}
- Soil resistivity: {calc['soil_res']} K.m/W
- Burial depth: {calc['depth']} m
""")
        else:
            st.info("👈 Calculate loads first to see derating factors")
    
    with cable_tabs[2]:
        st.markdown('<div class="report-header">🔌 Cable selection results</div>', unsafe_allow_html=True)
        st.markdown("### ⚡ Voltage drop limit: **2.5%**")
        
        if not st.session_state.cable_results_df.empty:
            st.dataframe(st.session_state.cable_results_df, use_container_width=True, hide_index=True)
            st.markdown("### 📋 Detailed calculation")
            
            for calc in st.session_state.detailed_calcs:
                with st.expander(f"🔍 {calc['load_name']} ({format_load_type(calc.get('load_type', 'Continuous'))})"):
                    st.markdown(f"""
**Step 1: Load current**  
I = {calc['power']} x 1000 / (1.732 x {calc['voltage']} x {calc['pf']} x {calc.get('efficiency', 1.0):.2f}) = **{calc['current']:.1f} A**

**Step 2: Cable type selection**  
Voltage {calc['voltage']}V -> {calc['cable_category']}

**Step 3: Derating factors**
- Ambient temperature: {calc['ambient_temp']}°C
- Insulation type: {format_insulation_type(calc['insulation_type'])}
- Cable type: {format_cable_type(calc['cable_type'])}
- Installation method: {calc['installation']}
- Cables in group: {calc['num_cables']}
- Cable arrangement: {format_cable_arrangement(calc['arrangement'])}
- Cable formation: {format_cable_formation(calc['formation'])}
- Soil resistivity: {calc['soil_res']} K.m/W
- Burial depth: {calc['depth']} m

**Calculated factors:**
- k1 (Temperature): {calc['k1']:.3f}
- k2 (Grouping): {calc['k2']:.3f}
- k3 (Soil resistivity): {calc['k3']:.3f}
- k4 (Depth): {calc['k4']:.3f}
- **Total K = {calc['total_k']:.3f}**

**Step 4: Cable selection**  
Selected: {calc['size']} mm² {format_cable_type(calc['cable_type'])} ({format_insulation_type(calc['insulation_type'])})  
Base ampacity: {calc['base_amp']} A  
Derated ampacity = {calc['base_amp']} × {calc['total_k']:.3f} = **{calc['derated_amp']:.1f} A**  
Check: {calc['derated_amp']:.1f} A >= {calc['current']:.1f} A -> **{'PASS' if calc['ampacity_pass'] else 'FAIL'}**

**Step 5: Voltage drop calculation**
""")
                    
                    # Get R and X values from cable database
                    cable_db = get_cable_data(calc['insulation_type'], calc['cable_type'])
                    cable_data = cable_db.get(calc['size'], {})
                    
                    r_value = cable_data.get('R', 0)
                    
                    if calc['formation'] == 'trefoil':
                        x_value = cable_data.get('X_trefoil', 0)
                    elif calc['formation'] == 'spaced':
                        x_value = cable_data.get('X_spaced', 0)
                    else:
                        x_value = cable_data.get('X_flat_touching', 0)
                    
                    # For multi-core cables with single X value
                    if x_value == 0:
                        x_value = cable_data.get('X', 0)
                    
                    if r_value == 0:
                        if calc['size'] <= 16:
                            r_value = 1.15
                        elif calc['size'] <= 35:
                            r_value = 0.73
                        elif calc['size'] <= 95:
                            r_value = 0.44
                        else:
                            r_value = 0.19
                        st.caption(f"ℹ️ Using typical R = {r_value} Ω/km (actual cable data not available)")
                    
                    if x_value == 0:
                        x_value = 0.08
                        st.caption(f"ℹ️ Using typical X = {x_value} Ω/km (actual cable data not available)")
                    
                    phi = math.acos(calc['pf'])
                    sin_phi = math.sin(phi)
                    
                    if calc['phase'] == '3-phase':
                        vd_calc = 1.732 * calc['current'] * (r_value * calc['pf'] + x_value * sin_phi) * calc['length'] / 1000
                        vd_pct_calc = (vd_calc / calc['voltage']) * 100
                        
                        st.markdown(f"""
**Formula (3-phase):** Vd = √3 × I × (R cosφ + X sinφ) × L / 1000

**Given for this load:**
- Load current (I) = {calc['current']:.1f} A
- Cable length (L) = {calc['length']:.0f} m
- Power factor (cosφ) = {calc['pf']}
- Voltage (V) = {calc['voltage']:.0f} V (3-phase)
- Cable resistance (R) = {r_value:.4f} Ω/km
- Cable reactance (X) = {x_value:.4f} Ω/km

**Step 5.1:** sinφ = √(1 - cos²φ) = √(1 - {calc['pf']:.3f}²) = **{sin_phi:.4f}**

**Step 5.2:** R cosφ = {r_value:.4f} × {calc['pf']:.3f} = **{r_value * calc['pf']:.4f}**

**Step 5.3:** X sinφ = {x_value:.4f} × {sin_phi:.4f} = **{x_value * sin_phi:.4f}**

**Step 5.4:** (R cosφ + X sinφ) = {r_value * calc['pf']:.4f} + {x_value * sin_phi:.4f} = **{(r_value * calc['pf'] + x_value * sin_phi):.4f}**

**Step 5.5:** Vd = 1.732 × {calc['current']:.1f} × {(r_value * calc['pf'] + x_value * sin_phi):.4f} × {calc['length']:.0f} / 1000 = **{vd_calc:.2f} V**

**Step 5.6:** Vd% = ({vd_calc:.2f} / {calc['voltage']:.0f}) × 100 = **{vd_pct_calc:.3f}%**
""")
                    else:
                        vd_calc = 2 * calc['current'] * (r_value * calc['pf'] + x_value * sin_phi) * calc['length'] / 1000
                        vd_pct_calc = (vd_calc / calc['voltage']) * 100
                        
                        st.markdown(f"""
**Formula (1-phase):** Vd = 2 × I × (R cosφ + X sinφ) × L / 1000

**Given for this load:**
- Load current (I) = {calc['current']:.1f} A
- Cable length (L) = {calc['length']:.0f} m
- Power factor (cosφ) = {calc['pf']}
- Voltage (V) = {calc['voltage']:.0f} V (1-phase)
- Cable resistance (R) = {r_value:.4f} Ω/km
- Cable reactance (X) = {x_value:.4f} Ω/km

**Step 5.1:** sinφ = √(1 - cos²φ) = √(1 - {calc['pf']:.3f}²) = **{sin_phi:.4f}**

**Step 5.2:** R cosφ = {r_value:.4f} × {calc['pf']:.3f} = **{r_value * calc['pf']:.4f}**

**Step 5.3:** X sinφ = {x_value:.4f} × {sin_phi:.4f} = **{x_value * sin_phi:.4f}**

**Step 5.4:** (R cosφ + X sinφ) = {r_value * calc['pf']:.4f} + {x_value * sin_phi:.4f} = **{(r_value * calc['pf'] + x_value * sin_phi):.4f}**

**Step 5.5:** Vd = 2 × {calc['current']:.1f} × {(r_value * calc['pf'] + x_value * sin_phi):.4f} × {calc['length']:.0f} / 1000 = **{vd_calc:.2f} V**

**Step 5.6:** Vd% = ({vd_calc:.2f} / {calc['voltage']:.0f}) × 100 = **{vd_pct_calc:.3f}%**
""")
                    
                    st.markdown(f"""
**Result for this load:**  
Voltage drop = **{calc['vd_pct']:.3f}%** (Limit: 2.5%)  
Check: {calc['vd_pct']:.3f}% <= 2.5% -> **{'PASS' if calc['vd_pass'] else 'FAIL'}**

**Step 6: Short circuit calculation**  
Isc = **{calc['sc']:.2f} kA**

**Final status: {'PASS' if calc['status'] == 'PASS' else 'FAIL'}**
""")
                    
                    if 'trials' in calc and calc['trials']:
                        st.markdown("**📊 Cable selection trials (all attempted sizes):**")
                        trial_data = []
                        for trial in calc['trials']:
                            trial_data.append({
                                'Size (mm²)': trial['size'],
                                'Base Ampacity (A)': trial.get('ampacity', trial.get('base_amp', 0)),
                                'Derated (A)': f"{trial['derated']:.1f}",
                                'Vd %': f"{trial['vd_pct']:.2f}%",
                                'Ampacity Check': '✓ PASS' if trial['ampacity_pass'] else '✗ FAIL',
                                'VD Check': '✓ PASS' if trial['vd_pass'] else '✗ FAIL'
                            })
                        st.dataframe(pd.DataFrame(trial_data), use_container_width=True, hide_index=True)
        else:
            st.info("👈 Calculate loads first")
    
    with cable_tabs[3]:
        st.markdown('<div class="report-header">⚡ Circuit breaker sizing</div>', unsafe_allow_html=True)
        
        st.markdown("""
### Circuit Breaker Sizing Calculation Example

**Formula:** Icb = Iload × 1.25 (25% safety margin)

**Example:**
- Load current (Iload) = 40 A
- Required CB rating = 40 × 1.25 = 50 A
- Selected standard rating = 50 A
- Breaker type = MCB (since 50A ≤ 125A)

**Breaker types:**
- **MCB** (≤125A): Miniature circuit breaker - For final circuits
- **MCCB** (125A-1600A): Moulded case circuit breaker - For distribution
- **ACB** (≥1600A): Air circuit breaker - For main incomers
""")
        
        if st.session_state.cb_results:
            st.markdown("### ⚡ Individual circuit breakers")
            
            st.markdown("#### 🔧 Pole selection for each load")
            st.info("Select the number of poles based on your system requirements:")
            
            pole_options = {
                '1P': 'Single pole - Phase only (for IT systems or DC)',
                '2P': 'Two pole - Phase + neutral (for single-phase TN/TT systems)',
                '3P': 'Three pole - 3 phases only (for TN-C or 3-wire systems)',
                '4P': 'Four pole - 3 phases + neutral (for TN-S systems)'
            }
            
            pole_selections = {}
            cols = st.columns(min(len(st.session_state.cb_results), 3))
            
            for idx, result in enumerate(st.session_state.cb_results):
                col_idx = idx % 3
                with cols[col_idx]:
                    default_poles = '3P' if result['Phase'] == '3-phase' else '2P'
                    pole_selections[result['Load']] = st.selectbox(
                        f"Poles for {result['Load']}",
                        options=list(pole_options.keys()),
                        format_func=lambda x: f"{x} - {pole_options[x].split('-')[0].strip()}",
                        key=f"pole_{result['Load']}",
                        index=list(pole_options.keys()).index(default_poles)
                    )
                    st.caption(pole_options[pole_selections[result['Load']]])
            
            cb_df = pd.DataFrame([{
                'Load': r['Load'],
                'Power (kW)': r['Power (kW)'],
                'Current (A)': f"{r['Current (A)']:.1f}",
                'Required (A)': f"{r['Required CB (A)']:.1f}",
                'Selected (A)': r['Selected CB (A)'],
                'Type': f"{r['Breaker Type']}",
                'Poles': pole_selections.get(r['Load'], '3P'),
                'Standard': r['Standard']
            } for r in st.session_state.cb_results])
            st.dataframe(cb_df, use_container_width=True, hide_index=True)
            
            st.markdown("### 🔋 Main circuit breaker")
            
            main_pole = st.selectbox(
                "Main circuit breaker poles",
                options=list(pole_options.keys()),
                format_func=lambda x: f"{x} - {pole_options[x]}",
                key="main_pole_select",
                index=2
            )
            
            main = st.session_state.main_cb
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("Total power", f"{main['total_power']:.1f} kW")
            with col2:
                st.metric("Total current", f"{main['current']:.1f} A")
            with col3:
                st.metric("Required cb", f"{main['required_cb']:.1f} A")
            with col4:
                st.metric("Selected cb", f"{main['selected_cb']} A {main['breaker_type']} {main_pole}")
            
            st.markdown("### 📋 Detailed selection calculations")
            
            with st.expander("Main circuit breaker calculation", expanded=True):
                st.markdown(main['detailed_reason'])
                st.markdown(f"""
**Pole selection**
- User selected: {main_pole}
- {pole_options[main_pole]}
""")
            
            st.markdown("### 📋 Individual breaker selection details")
            for detail in st.session_state.cb_details:
                with st.expander(f"Load: {detail['load_name']}"):
                    selected_poles = pole_selections.get(detail['load_name'], '3P')
                    st.markdown(f"""
**Step 1: Load analysis**
- Load type: {detail['phase_desc']}
- Load current: {detail['current']:.2f} A

**Step 2: Rating calculation**
- Design factor: {detail['design_factor']} (25% safety margin)
- Required rating = {detail['current']:.2f} x {detail['design_factor']} = {detail['required']:.2f} A
- Selected standard rating: {detail['selected']} A

**Step 3: Breaker type selection**
- Type: {detail['breaker_type']}
- Application: {BREAKER_TYPES[detail['breaker_type']]['application']}

**Step 4: Pole selection**
- User selected: {selected_poles}
- {pole_options[selected_poles]}

**Step 5: Manufacturer selection**
- Manufacturer: {detail['manufacturer']}
- Series: {detail['series']}

**Final selection: {detail['selected']} A {detail['breaker_type']} {selected_poles}**
""")
        else:
            st.info("👈 Calculate cable sizes first to see circuit breaker results")
    
    with cable_tabs[4]:
        st.markdown('<div class="report-header">📥 Download report</div>', unsafe_allow_html=True)
        
        if not st.session_state.cable_results_df.empty and st.session_state.cb_results:
            if st.button("📥 Generate word report", key="cable_word", use_container_width=True):
                with st.spinner("Generating word with complete detailed calculations..."):
                    try:
                        word = CableWordReport()
                        word.add_title()
                        
                        word.add_common_parameters(ambient_temp)
                        
                        word.add_load_details(st.session_state.loads_df)
                        word.add_cable_results(st.session_state.cable_results_df)
                        
                        if st.session_state.detailed_calcs:
                            word.add_detailed_calculations(st.session_state.detailed_calcs)
                        
                        if st.session_state.cb_results and st.session_state.main_cb:
                            pole_selections = {}
                            for r in st.session_state.cb_results:
                                pole_selections[r['Load']] = '3P' if r['Phase'] == '3-phase' else '2P'
                            word.add_cb_results(st.session_state.cb_results, st.session_state.main_cb, pole_selections, '3P')
                        
                        word_path = "temp_cable_report.docx"
                        word.save(word_path)
                        with open(word_path, "rb") as f:
                            word_bytes = f.read()
                        b64 = base64.b64encode(word_bytes).decode()
                        filename = f"Cable_CB_Report_{datetime.now().strftime('%Y%m%d_%H%M')}.docx"
                        os.remove(word_path)
                        
                        st.markdown(f'<a href="data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,{b64}" download="{filename}" class="download-btn">📥 Click here to download word report</a>', unsafe_allow_html=True)
                        st.success("✅ Word generated successfully!")
                    except Exception as e:
                        st.error(f"Error generating word document: {str(e)}")
        else:
            st.info("👈 Calculate cable sizes first to generate report")

# ========== TRANSFORMER SIZING ==========
elif st.session_state.selected_calculator == "⚙️ Transformer Sizing":
    
    st.markdown('<div class="report-header">⚙️ TRANSFORMER SIZING CALCULATOR</div>', unsafe_allow_html=True)
    
    st.markdown(f"""
    <div class="info-box">
        <h4>📌 Using loads from universal load sheet</h4>
        <p>Total {len(st.session_state.universal_loads)} loads available. Calculations below use these loads.</p>
    </div>
    """, unsafe_allow_html=True)
    
    tx_main_tabs = st.tabs([
        "📊 Load analysis",
        "📈 Largest equipment analysis",
        "📥 Download report"
    ])
    
    tx_calc = SimpleTransformerCalculator()
    
    with tx_main_tabs[0]:
        load_sub_tabs = st.tabs([
            "📋 Step-by-step p, q, s",
            "📊 Summary table"
        ])
        
        with load_sub_tabs[0]:
            st.markdown("### 📋 Step-by-step calculations for each load")
            
            total_p = 0
            total_q = 0
            
            for idx, load in st.session_state.universal_loads.iterrows():
                connected = load['Rating (kW)'] * load['Quantity']
                load_type_diversity = LOAD_TYPE_FACTORS[load['Load Type']]['diversity']
                p = tx_calc.calculate_p(load['Rating (kW)'], load['Quantity'], load_type_diversity)
                phi = math.acos(load['Power Factor'])
                tan_phi = math.tan(phi)
                q = tx_calc.calculate_q(p, load['Power Factor'])
                s = tx_calc.calculate_s(p, q)
                cable_type = "LV" if load['Voltage (V)'] <= 1000 else "MV"
                
                st.markdown(f"""
                <div class="calc-step">
                    <h4>📌 Load {idx+1}: <span style="color: #1E3A8A;">{load['Load Description']}</span> ({load['Load Type']}, {cable_type})</h4>
                    <p><b>Input parameters:</b> Rating = {load['Rating (kW)']:.0f} kW, Quantity = {load['Quantity']}, Pf = {load['Power Factor']}, Load type = {load['Load Type']} ({load_type_diversity*100:.0f}%)</p>
                    <p><b>Step 1 - Connected power:</b> {load['Rating (kW)']:.0f} kW x {load['Quantity']} = <b>{connected:.0f} kW</b></p>
                    <p><b>Step 2 - Demand power (P):</b> {connected:.0f} kW x {load_type_diversity:.1f} = <b>{p:.1f} kW</b></p>
                    <p><b>Step 3 - Angle φ:</b> acos({load['Power Factor']}) = <b>{math.degrees(phi):.1f}°</b></p>
                    <p><b>Step 4 - tan(φ):</b> tan({math.degrees(phi):.1f}°) = <b>{tan_phi:.3f}</b></p>
                    <p><b>Step 5 - Reactive power (Q):</b> {p:.1f} kW x {tan_phi:.3f} = <b>{q:.1f} kVAR</b></p>
                    <p><b>Step 6 - Apparent power (S):</b> √({p:.1f}² + {q:.1f}²) = <b>{s:.1f} kVA</b></p>
                </div>
                """, unsafe_allow_html=True)
                
                total_p += p
                total_q += q
            
            st.session_state.total_p = total_p
            st.session_state.total_q = total_q
        
        with load_sub_tabs[1]:
            st.markdown("### 📊 Load summary table")
            
            summary_data = []
            for idx, load in st.session_state.universal_loads.iterrows():
                load_type_diversity = LOAD_TYPE_FACTORS[load['Load Type']]['diversity']
                connected = load['Rating (kW)'] * load['Quantity']
                p = tx_calc.calculate_p(load['Rating (kW)'], load['Quantity'], load_type_diversity)
                q = tx_calc.calculate_q(p, load['Power Factor'])
                s = tx_calc.calculate_s(p, q)
                cable_type = "LV" if load['Voltage (V)'] <= 1000 else "MV"
                
                summary_data.append({
                    'Load': load['Load Description'],
                    'Type': load['Load Type'],
                    'Cable': cable_type,
                    'Qty': load['Quantity'],
                    'Rating (kW)': load['Rating (kW)'],
                    'Connected (kW)': f"{connected:.0f}",
                    'Load factor': f"{load_type_diversity*100:.0f}%",
                    'P (kW)': f"{p:.1f}",
                    'Pf': load['Power Factor'],
                    'Q (kVAR)': f"{q:.1f}",
                    'S (kVA)': f"{s:.1f}"
                })
            
            summary_df = pd.DataFrame(summary_data)
            st.dataframe(summary_df, use_container_width=True, hide_index=True)
            
            st.markdown("""
            <div class="formula-box">
                <h4>📐 Formulas used:</h4>
                <p><b>Connected power = Rating x Quantity</b></p>
                <p><b>P = Connected power x Load type factor</b> (Real power)</p>
                <p><b>Q = P x tan(acos(PF))</b> (Reactive power)</p>
                <p><b>S = √(P² + Q²)</b> (Apparent power)</p>
                <p><b>Load type factors:</b> Continuous=100%, Intermittent=30%, Standby=10%</p>
            </div>
            """, unsafe_allow_html=True)
    
    with tx_main_tabs[1]:
        st.markdown("### 🏭 Largest equipment analysis")
        
        largest_idx, largest_load, largest_connected = tx_calc.find_largest_equipment(st.session_state.universal_loads)
        
        if largest_load is not None:
            load_type_diversity = LOAD_TYPE_FACTORS[largest_load['Load Type']]['diversity']
            p_largest = tx_calc.calculate_p(largest_load['Rating (kW)'], largest_load['Quantity'], load_type_diversity)
            q_largest = tx_calc.calculate_q(p_largest, largest_load['Power Factor'])
            s_largest = tx_calc.calculate_s(p_largest, q_largest)
            
            total_p = 0
            total_q = 0
            all_loads_data = []
            
            for idx, load in st.session_state.universal_loads.iterrows():
                load_div = LOAD_TYPE_FACTORS[load['Load Type']]['diversity']
                p = tx_calc.calculate_p(load['Rating (kW)'], load['Quantity'], load_div)
                q = tx_calc.calculate_q(p, load['Power Factor'])
                s = tx_calc.calculate_s(p, q)
                
                total_p += p
                total_q += q
                cable_type = "LV" if load['Voltage (V)'] <= 1000 else "MV"
                
                all_loads_data.append({
                    'load': load['Load Description'],
                    'type': load['Load Type'],
                    'cable': cable_type,
                    'p': p,
                    'q': q,
                    's': s,
                    'is_largest': (idx == largest_idx)
                })
            
            total_s = math.sqrt(total_p**2 + total_q**2) if (total_p**2 + total_q**2) > 0 else 0
            
            st.markdown(f"""
            <div class="largest-equipment">
                <h3>🏆 Largest equipment: {largest_load['Load Description']}</h3>
                <table style="width:100%; border-collapse: collapse;">
                    <tr><td style="padding: 10px; font-weight: bold;">Load type: <td style="padding: 10px;"><span class="value">{largest_load['Load Type']} ({load_type_diversity*100:.0f}%)</span></tr>
                    <tr><td style="padding: 10px; font-weight: bold;">Connected power: <td style="padding: 10px;"><span class="value">{largest_connected:.0f} kW</span> ({largest_load['Rating (kW)']:.0f} kW x {largest_load['Quantity']})</span></tr>
                    <tr><td style="padding: 10px; font-weight: bold;">Demand power (P): <td style="padding: 10px;"><span class="value">{p_largest:.1f} kW</span> (after {load_type_diversity*100:.0f}% factor)</span></tr>
                    <tr><td style="padding: 10px; font-weight: bold;">Reactive power (Q): <td style="padding: 10px;"><span class="value">{q_largest:.1f} kVAR</span> (Pf = {largest_load['Power Factor']})</span></tr>
                    <tr><td style="padding: 10px; font-weight: bold;">Apparent power (S): <td style="padding: 10px;"><span class="value">{s_largest:.1f} kVA</span></span></tr>
                </table>
            </div>
            """, unsafe_allow_html=True)
            
            st.markdown("### 📊 Impact on total system")
            
            impact_data = []
            for data in all_loads_data:
                p_pct = (data['p'] / total_p) * 100 if total_p > 0 else 0
                s_pct = (data['s'] / total_s) * 100 if total_s > 0 else 0
                
                impact_data.append({
                    'Load': data['load'],
                    'Type': data['type'],
                    'Cable': data['cable'],
                    'P (kW)': f"{data['p']:.1f}",
                    '% of total P': f"{p_pct:.1f}%",
                    'S (kVA)': f"{data['s']:.1f}",
                    '% of total S': f"{s_pct:.1f}%",
                    'Highlight': '🏆 Largest' if data['is_largest'] else ''
                })
            
            impact_df = pd.DataFrame(impact_data)
            st.dataframe(impact_df, use_container_width=True, hide_index=True)
            
            st.markdown("### 📈 Cumulative effect analysis")
            
            sorted_loads = sorted(all_loads_data, key=lambda x: x['p'], reverse=True)
            
            cumulative_p = 0
            cumulative_data = []
            
            for i, data in enumerate(sorted_loads):
                cumulative_p += data['p']
                p_pct = (cumulative_p / total_p) * 100 if total_p > 0 else 0
                
                cumulative_data.append({
                    'Rank': i+1,
                    'Load': data['load'],
                    'Type': data['type'],
                    'Individual P (kW)': f"{data['p']:.1f}",
                    'Cumulative P (kW)': f"{cumulative_p:.1f}",
                    '% of total': f"{p_pct:.1f}%"
                })
            
            cumulative_df = pd.DataFrame(cumulative_data)
            st.dataframe(cumulative_df, use_container_width=True, hide_index=True)
            
            st.session_state.tx_largest_data = {
                'load': largest_load,
                'connected': largest_connected,
                'p': p_largest,
                'q': q_largest,
                's': s_largest
            }
    
    with tx_main_tabs[2]:
        st.markdown("### ⚙️ Future expansion")
        future_expansion = st.number_input("Future expansion (%)", value=20, min_value=0, max_value=100, step=5)
        
        if 'total_p' in st.session_state and 'total_q' in st.session_state:
            total_p = st.session_state.total_p
            total_q = st.session_state.total_q
            
            total_s = math.sqrt(total_p**2 + total_q**2)
            with_future = total_s * (1 + future_expansion/100)
            selected_kva = tx_calc.get_standard_rating(with_future)
            
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("Total P", f"{total_p:.0f} kW")
            with col2:
                st.metric("Total Q", f"{total_q:.0f} kVAR")
            with col3:
                st.metric("Total S", f"{total_s:.0f} kVA")
            with col4:
                st.metric("With future", f"{with_future:.0f} kVA")
            
            st.markdown(f"""
            <div class="result-card">
                <h3>✅ Final transformer selection</h3>
                <p><b>S = √(P² + Q²) = √({total_p:.0f}² + {total_q:.0f}²) = {total_s:.0f} kVA</b></p>
                <p><b>With {future_expansion}% future = {total_s:.0f} x 1.{future_expansion/100:.0f} = {with_future:.0f} kVA</b></p>
                <p style="font-size: 24px;"><b>Selected: {selected_kva} kVA</b></p>
            </div>
            """, unsafe_allow_html=True)
            
            if st.button("📥 Download word report", key="tx_word", use_container_width=True):
                with st.spinner("Generating word report..."):
                    try:
                        word = TransformerWordReport()
                        word.add_title()
                        total_p_load = word.add_load_analysis(st.session_state.universal_loads)
                        
                        total_p_step, total_q_step = word.add_step_by_step(st.session_state.universal_loads)
                        
                        if 'tx_largest_data' in st.session_state and st.session_state.tx_largest_data:
                            word.add_largest_equipment(st.session_state.universal_loads, total_p, total_s, st.session_state.tx_largest_data)
                        
                        word.add_transformer_selection(total_p_step, total_q_step, future_expansion, selected_kva)
                        
                        word_path = "temp_transformer_report.docx"
                        word.save(word_path)
                        with open(word_path, "rb") as f:
                            word_bytes = f.read()
                        b64 = base64.b64encode(word_bytes).decode()
                        filename = f"Transformer_Report_{datetime.now().strftime('%Y%m%d_%H%M')}.docx"
                        os.remove(word_path)
                        st.markdown(f'<a href="data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,{b64}" download="{filename}" class="download-btn">📥 Click here to download word report</a>', unsafe_allow_html=True)
                        st.success("✅ Word generated successfully!")
                    except Exception as e:
                        st.error(f"Error generating word document: {e}")
        else:
            st.warning("⚠️ Please go to load analysis tab first to calculate totals.")

# ========== GENERATOR SIZING (PLACEHOLDER) ==========
elif st.session_state.selected_calculator == "🔄 Generator Sizing":
    
    st.markdown('<div class="report-header">🔄 GENERATOR SIZING CALCULATOR</div>', unsafe_allow_html=True)
    
    st.markdown("""
    <div class="info-box">
        <h4>🚧 Coming Soon!</h4>
        <p>Generator sizing calculator is currently under development.</p>
        <p>Features to be included:</p>
        <ul>
            <li>Load analysis and kVA calculation</li>
            <li>Motor starting analysis (DOL, Star-Delta, Soft Starter, VFD)</li>
            <li>Fuel consumption calculation</li>
            <li>Generator set selection from standard ratings</li>
            <li>Voltage drop calculation for generator cables</li>
            <li>Automatic transfer switch (ATS) sizing</li>
            <li>Parallel generator operation analysis</li>
        </ul>
        <p style="margin-top: 15px;"><b>Expected release:</b> Coming soon in the next update!</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Show a simple preview of what's coming
    with st.expander("📊 Preview - Generator Sizing Methodology", expanded=False):
        st.markdown("""
        **Generator Sizing Steps:**
        
        1. **Calculate Total Load (kVA)**
           - Sum of all continuous loads
           - Add largest motor starting kVA
        
        2. **Apply Diversity Factors**
           - Based on load types (Continuous, Intermittent, Standby)
        
        3. **Consider Future Expansion**
           - Add 20-30% for future growth
        
        4. **Select Standard Rating**
           - Choose next standard generator size
        
        5. **Verify Starting Requirements**
           - Check voltage dip during motor starting
           - Ensure generator can handle starting current
        
        6. **Fuel Consumption Calculation**
           - At 100%, 75%, 50% load
           - Tank size recommendation
        """)

# ========== EARTHING (PLACEHOLDER) ==========
elif st.session_state.selected_calculator == "⏚ Earthing":
    
    st.markdown('<div class="report-header">⏚ EARTHING CALCULATOR</div>', unsafe_allow_html=True)
    
    st.markdown("""
    <div class="info-box">
        <h4>🚧 Coming Soon!</h4>
        <p>Earthing system calculator is currently under development.</p>
        <p>Features to be included:</p>
        <ul>
            <li>Soil resistivity analysis based on soil type</li>
            <li>Rod, plate, and strip earthing calculations</li>
            <li>Step and touch potential analysis (IEEE Std 80)</li>
            <li>IEC 62305 lightning protection earthing compliance</li>
            <li>Multiple electrode configuration (parallel rods)</li>
            <li>Earth resistance calculation for various electrode types</li>
            <li>Chemical earthing recommendation</li>
        </ul>
        <p style="margin-top: 15px;"><b>Expected release:</b> Coming soon in the next update!</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Show a simple preview of what's coming
    with st.expander("📊 Preview - Earthing Calculation Methodology", expanded=False):
        st.markdown("""
        **Earthing Design Steps (IEEE Std 80 / IEC 62305):**
        
        **1. Soil Resistivity Measurement**
        - Wenner four-pin method
        - Soil types and typical resistivity values
        
        **2. Electrode Selection**
        - Rod electrodes: R = (ρ / 2πL) × ln(4L/d)
        - Plate electrodes: R = ρ / (4 × √(π × A))
        - Strip/ring electrodes: R = (ρ / 2πL) × ln(2L² / (w × h))
        
        **3. Multiple Electrodes**
        - Parallel rod configuration
        - Utilization factor based on spacing
        
        **4. Step & Touch Potential**
        - Step potential: Es = (ρ × I × Ks × Ki) / L
        - Touch potential: Et = (ρ × I × Kt × Ki) / L
        
        **5. IEC 62305 Compliance**
        - Class I: ≤ 1 Ω, min rod length 3m
        - Class II: ≤ 4 Ω, min rod length 2.5m
        - Class III: ≤ 10 Ω, min rod length 2m
        - Class IV: ≤ 20 Ω, min rod length 1.5m
        """)

# Footer
st.markdown("---")
st.markdown(f"<div style='text-align: center; color: gray; font-size: 16px;'>🔌 CES-Electrical | Version 2.0 | {datetime.now().strftime('%Y-%m-%d %H:%M')}</div>", unsafe_allow_html=True)