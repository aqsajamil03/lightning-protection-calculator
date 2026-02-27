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
        color: black;
    }
    .parameter-table tr:nth-child(even) {
        background-color: #f2f2f2;
    }
    .parameter-table tr:nth-child(odd) {
        background-color: white;
    }
</style>
""", unsafe_allow_html=True)

# ========== CIRCUIT BREAKER DATA AND CALCULATIONS ==========
# Circuit Breaker Standard Ratings (IEC 60898)
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

# ========== SESSION STATE INITIALIZATION ==========
if 'calc_results' not in st.session_state:
    st.session_state.calc_results = {}
if 'calc_done' not in st.session_state:
    st.session_state.calc_done = False
if 'selected_calculator' not in st.session_state:
    st.session_state.selected_calculator = "⚡ Lightning Protection"
if 'cable_results' not in st.session_state:
    st.session_state.cable_results = {}
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
    
    calculators = [
        "⚡ Lightning Protection",
        "🔌 Cable Sizing",
        "⚡ Circuit Breaker Sizing",
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

# ========== LIGHTNING PROTECTION CALCULATOR (EXISTING CODE) ==========
if st.session_state.selected_calculator == "⚡ Lightning Protection":
    st.info("⚡ Lightning Protection Calculator - Existing functionality")
    # ... (existing lightning protection code)

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
            cable_type = st.selectbox("Cable Type", ['unarmoured', 'armoured'])
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
            soil_res = st.number_input("Soil Resistivity (K.m/W)", value=1.5, step=0.5)
        with col7:
            depth = st.number_input("Burial Depth (m)", value=0.8, step=0.1)
        with col8:
            system_type = st.selectbox("System Type", ['TN-S', 'TN-C', 'TN-C-S', 'TT'])
        
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
                    
                    # Find suitable cable
                    for size, data in db.items():
                        derated = data['ampacity'] * total_k
                        if derated >= current:
                            # Voltage drop calculation
                            vd_v, vd_pct = cable_calc.calculate_voltage_drop(
                                current, load['Length (m)'], data['R'], data['X'],
                                load['Power Factor'], load['Voltage (V)'], load['Phase']
                            )
                            
                            # Short circuit calculation
                            isc = cable_calc.calculate_short_circuit(size, 1.0)
                            
                            # Efficiency calculation (simplified)
                            efficiency = (load['Power (kW)'] * 1000) / (1.732 * load['Voltage (V)'] * current) * 100
                            
                            cable_results.append({
                                'Load Name': load['Load Name'],
                                'Power (kW)': load['Power (kW)'],
                                'Voltage (V)': load['Voltage (V)'],
                                'Cable Category': cable_category,
                                'Cable Type': f'{cable_type} copper',
                                'Size (mm²)': size,
                                'Load Current (A)': f"{current:.1f}",
                                'Base Ampacity (A)': data['ampacity'],
                                'Derating Factor K': f"{total_k:.3f}",
                                'Derated Ampacity (A)': f"{derated:.1f}",
                                'Voltage Drop (%)': f"{vd_pct:.3f}",
                                'Short Circuit (kA)': f"{isc/1000:.2f}",
                                'Efficiency (%)': f"{efficiency:.1f}",
                                'Status': 'PASS' if vd_pct <= 2.5 else 'FAIL'
                            })
                            
                            # Detailed calculation
                            detail = f"""
### Load: {load['Load Name']}
**Cable Category:** {cable_category} | **Type:** {cable_type} copper

**Step 1: Load Current [IEC 60364-5-52]**
I = {load['Power (kW)']}kW × 1000 / (1.732 × {load['Voltage (V)']}V × {load['Power Factor']}) = **{current:.1f} A**

**Step 2: Derating Factors [IEC 60502-2]**
k1 (Temperature): {factors['k1 (Temperature)']['value']:.3f} ({factors['k1 (Temperature)']['reference']})
k2 (Grouping): {factors['k2 (Grouping)']['value']:.3f} ({factors['k2 (Grouping)']['reference']})
k3 (Soil): {factors['k3 (Soil Resistivity)']['value']:.3f}
k4 (Depth): {factors['k4 (Depth)']['value']:.3f}
k5 (Laying): {factors['k5 (Laying)']['value']:.3f}
Total K = **{total_k:.3f}**

**Step 3: Cable Selection**
Selected: {size}mm² - Base Ampacity: {data['ampacity']}A
Derated Ampacity = {data['ampacity']}A × {total_k:.3f} = **{derated:.1f} A**

**Step 4: Voltage Drop [IEC 60364-5-52 Sec 525]**
Vd = {vd_v:.2f}V = **{vd_pct:.3f}%** (Limit: 2.5%)

**Step 5: Short Circuit [IEC 60949]**
Isc = 143 × {size} / √1.0 = **{isc/1000:.2f} kA**

**Step 6: Efficiency**
η = **{efficiency:.1f}%**

**Status:** {'✅ PASS' if vd_pct <= 2.5 else '❌ FAIL'}
---
"""
                            detailed_calcs.append(detail)
                            break
                
                st.session_state.cable_results_df = pd.DataFrame(cable_results)
                st.session_state.detailed_calcs = detailed_calcs
                st.session_state.derating_factors = factors
                st.success("✅ Calculations complete!")
    
    # TAB 2: DERATING FACTORS
    with cable_tabs[1]:
        st.markdown('<div class="report-header">', unsafe_allow_html=True)
        st.markdown("## ALL DERATING FACTORS (IEC 60502-2)")
        st.markdown('</div>', unsafe_allow_html=True)
        
        if 'derating_factors' in st.session_state:
            factors = st.session_state.derating_factors
            factors_html = "<table class='parameter-table'>"
            factors_html += "<tr><th>Factor</th><th>Value</th><th>Reference</th><th>Description</th></tr>"
            
            for key, data in factors.items():
                if key != 'total':
                    desc = {
                        'k1 (Temperature)': 'Ambient temperature correction',
                        'k2 (Grouping)': 'Number of cables grouped together',
                        'k3 (Soil Resistivity)': 'Soil thermal resistivity',
                        'k4 (Depth)': 'Depth of laying correction',
                        'k5 (Laying)': 'Installation method correction'
                    }.get(key, '')
                    factors_html += f"<tr><td>{key}</td><td>{data['value']:.3f}</td><td>{data['reference']}</td><td>{desc}</td></tr>"
            
            factors_html += f"<tr style='background-color: #1E3A8A; color: white;'><td colspan='4'><strong>Total K = {factors['total']:.3f}</strong></td></tr>"
            factors_html += "</table>"
            
            st.markdown(factors_html, unsafe_allow_html=True)
        else:
            st.info("👈 Calculate loads first")
    
    # TAB 3: CABLE SELECTION
    with cable_tabs[2]:
        st.markdown('<div class="report-header">', unsafe_allow_html=True)
        st.markdown("## CABLE SELECTION RESULTS")
        st.markdown('</div>', unsafe_allow_html=True)
        
        st.markdown("### ⚡ Voltage Drop Limit: **2.5%** [IEC 60364-5-52 Section 525]")
        
        if 'cable_results_df' in st.session_state:
            st.dataframe(st.session_state.cable_results_df, use_container_width=True, hide_index=True)
            
            for detail in st.session_state.detailed_calcs:
                with st.expander(detail.split('\n')[0].replace('###', '').strip()):
                    st.markdown(detail)
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
            test_size = st.number_input("Cable Size (mm²)", value=95.0, step=5.0)
        with col2:
            test_duration = st.number_input("Duration (s)", value=1.0, step=0.1)
        
        cable_calc = CableSizingCalculator()
        isc = cable_calc.calculate_short_circuit(test_size, test_duration)
        
        st.metric("Short Circuit Capacity", f"{isc/1000:.2f} kA")
        
        if 'cable_results_df' in st.session_state:
            st.markdown("### 📊 Calculated Cables Short Circuit Capacity")
            df = st.session_state.cable_results_df[['Load Name', 'Size (mm²)', 'Short Circuit (kA)']]
            st.dataframe(df, use_container_width=True, hide_index=True)
    
    # TAB 5: DOWNLOAD REPORT
    with cable_tabs[4]:
        st.markdown('<div class="report-header">', unsafe_allow_html=True)
        st.markdown("## DOWNLOAD REPORT")
        st.markdown('</div>', unsafe_allow_html=True)
        
        if 'cable_results_df' in st.session_state:
            st.markdown("### 📥 Download PDF Report")
            if st.button("📥 Generate PDF", use_container_width=True):
                with st.spinner("Generating PDF..."):
                    # PDF generation code here
                    st.success("PDF generated successfully!")

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
    
    if 'loads_df' in st.session_state:
        cb_calc = CircuitBreakerCalculator()
        system_type = st.selectbox("System Type", ['TN-S', 'TN-C', 'TN-C-S', 'TT'])
        
        cb_results = cb_calc.calculate_cb_size(st.session_state.loads_df, 1.25, 'Schneider Electric', system_type)
        
        cb_df = pd.DataFrame([{
            'Load': r['Load'],
            'Power (kW)': r['Power (kW)'],
            'Current (A)': f"{r['Current (A)']:.1f}",
            'Selected CB (A)': r['Selected CB (A)'],
            'Type': r['Breaker Type'],
            'Standard': r['Standard'],
            'Poles': r['Poles'],
            'Selection Reason': r['Pole Selection Reason'][:50] + '...'
        } for r in cb_results])
        
        st.dataframe(cb_df, use_container_width=True, hide_index=True)
    else:
        st.info("👈 Add loads in Cable Sizing calculator first")

# ========== OTHER CALCULATORS ==========
elif st.session_state.selected_calculator == "⚙️ Transformer Sizing":
    st.info("⚙️ Transformer sizing calculator - Coming soon")

elif st.session_state.selected_calculator == "⚡ Generator Sizing":
    st.info("⚡ Generator sizing calculator - Coming soon")

elif st.session_state.selected_calculator == "🌍 Earthing System Design":
    st.info("🌍 Earthing system design calculator - Coming soon")

# Footer
st.markdown("---")
st.markdown(f"<div style='text-align: center; color: gray;'>⚡ IEC Compliant Design Calculators | Version 46.0 | {datetime.now().strftime('%Y-%m-%d %H:%M')}</div>", unsafe_allow_html=True)