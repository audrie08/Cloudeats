import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from oauth2client.service_account import ServiceAccountCredentials
from streamlit_gsheets import GSheetsConnection
from streamlit_option_menu import option_menu
from google.oauth2.service_account import Credentials
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta
from streamlit_option_menu import option_menu
import base64
from io import BytesIO
from PIL import Image
import warnings
import json
import pytz
import numpy as np
import time
from google.oauth2 import service_account
from googleapiclient.discovery import build
warnings.filterwarnings('ignore')

# --- PAGE CONFIG ---
st.set_page_config(
    page_title="Commissary Production Schedule",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# --- CREDENTIALS HANDLING KPI DASHBOARD ---
def load_credentials_kpi():
    """Load Google credentials from Streamlit secrets"""
    try:
        # Check if secrets are available
        if "google_credentials1" not in st.secrets:
            st.error("Google credentials not found in secrets")
            return None
            
        # Convert secrets to the format expected by google-auth
        credentials_dict = {
            "type": st.secrets["google_credentials1"]["type"],
            "project_id": st.secrets["google_credentials1"]["project_id"],
            "private_key_id": st.secrets["google_credentials1"]["private_key_id"],
            "private_key": st.secrets["google_credentials1"]["private_key"].replace('\\n', '\n'),
            "client_email": st.secrets["google_credentials1"]["client_email"],
            "client_id": st.secrets["google_credentials1"]["client_id"],
            "auth_uri": st.secrets["google_credentials1"]["auth_uri"],
            "token_uri": st.secrets["google_credentials1"]["token_uri"],
            "auth_provider_x509_cert_url": st.secrets["google_credentials1"]["auth_provider_x509_cert_url"],
            "client_x509_cert_url": st.secrets["google_credentials1"]["client_x509_cert_url"]
        }
        
        scopes = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive"
        ]
        
        credentials = Credentials.from_service_account_info(credentials_dict, scopes=scopes)
        return credentials
        
    except Exception as e:
        st.error(f"Error loading credentials: {str(e)}")
        return None

# --- CREDENTIALS HANDLING PROD SCHED---
def load_credentials_prod():
    """Load Google credentials from Streamlit secrets"""
    try:
        # Convert secrets to the format expected by google-auth
        credentials_dict = {
            "type": st.secrets["google_credentials"]["type"],
            "project_id": st.secrets["google_credentials"]["project_id"],
            "private_key_id": st.secrets["google_credentials"]["private_key_id"],
            "private_key": st.secrets["google_credentials"]["private_key"].replace('\\n', '\n'),
            "client_email": st.secrets["google_credentials"]["client_email"],
            "client_id": st.secrets["google_credentials"]["client_id"],
            "auth_uri": st.secrets["google_credentials"]["auth_uri"],
            "token_uri": st.secrets["google_credentials"]["token_uri"],
            "auth_provider_x509_cert_url": st.secrets["google_credentials"]["auth_provider_x509_cert_url"],
            "client_x509_cert_url": st.secrets["google_credentials"]["client_x509_cert_url"]
        }
        
        scopes = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive"
        ]
        
        credentials = Credentials.from_service_account_info(credentials_dict, scopes=scopes)
        return credentials
        
    except Exception as e:
        st.error(f"Error loading credentials: {str(e)}")
        return None


# --- CUSTOM CSS ---
# Read and encode font (if available)
try:
    with open("TTNorms-Medium.ttf", "rb") as font_file:
        font_base64 = base64.b64encode(font_file.read()).decode()
    font_available = True
except FileNotFoundError:
    font_base64 = ""
    font_available = False

# Enhanced CSS including filter styles and modern navigation
font_face_css = f"""
    @font-face {{
        font-family: 'TT Norms';
        src: url(data:font/ttf;base64,{font_base64}) format('truetype');
        font-weight: 500;
        font-stretch: expanded;
    }}
""" if font_available else ""

st.markdown(f"""
<style>
    {font_face_css}

    /* Hide default Streamlit elements for cleaner look */
    #MainMenu {{visibility: hidden;}}
    footer {{visibility: hidden;}}
    header {{visibility: hidden;}}
    
    /* Adjust main container */
    .block-container {{
        max-width: 1400px;
        padding-left: 1rem;
        padding-right: 1rem;
        padding-top: 0rem;
        padding-bottom: 1rem;
    }}

    /* Background */
    body, .main, .block-container {{
        background-color: #ffffff;
    }}

    /* === MODERN NAVIGATION BAR === */
    .nav-container {{
        background: linear-gradient(135deg, #fffef6 0%, #fefdf0 100%);
        box-shadow: 0 4px 20px rgba(0, 0, 0, 0.1);
        border-bottom: 3px solid #f4d602;
        position: sticky;
        top: 0;
        z-index: 1000;
        margin: -1rem -1rem 1rem -1rem;
        padding: 0 2rem;
    }}
    
    .main-nav {{
        display: flex;
        align-items: center;
        justify-content: space-between;
        padding: 1rem 0;
        max-width: 1200px;
        margin: 0 auto;
    }}
    
    .nav-brand {{
        font-size: 1.5rem;
        font-weight: 700;
        color: #2c3e50;
        text-decoration: none;
        letter-spacing: -0.5px;
        font-family: {'TT Norms' if font_available else 'Segoe UI'}, sans-serif;
    }}
    
    .nav-menu-container {{
        display: flex;
        align-items: center;
        gap: 0.5rem;
    }}

    /* === MAIN HEADER === */
    .main-header {{
        background: #f7d42c;
        padding: 1rem 1rem;
        border-radius: 40px;
        color: #1E2328;
        text-align: center;
        margin-bottom: 1rem;
        box-shadow: 0 8px 16px rgba(0, 0, 0, 0.1);
    }}

    .main-header h1 {{
        font-family: {'TT Norms' if font_available else 'Arial'}, 'Arial', sans-serif;
        font-weight: normal;
        font-size: 2.5em;
        margin: 0;
        letter-spacing: 2px;
    }}

    .main-header p {{
        font-family: {'TT Norms' if font_available else 'Arial'}, 'Arial', sans-serif;
        font-weight: normal;
        margin: 0.5rem 0 0 0;
    }}

    /* === MODERN FILTER CONTAINER === */
    .filter-container {{
        background: #000000;
        backdrop-filter: blur(12px);
        -webkit-backdrop-filter: blur(12px);
        border: 1px solid;
        border-radius: 24px;
        padding: 2.5rem;
        margin: 2rem 0;
        position: relative;
        overflow: hidden;
        transition: all 0.4s cubic-bezier(0.4, 0, 0.2, 1);
    }}

    .filter-container::before {{
        content: '';
        position: absolute;
        top: -50%;
        left: -50%;
        width: 200%;
        height: 200%;
        background: #000000;
        animation: rotateGradient 20s linear infinite;
        pointer-events: none;
        z-index: -1;
    }}

    @keyframes rotateGradient {{
        0% {{ transform: rotate(0deg); }}
        100% {{ transform: rotate(360deg); }}
    }}

    .filter-container:hover {{
        transform: translateY(-4px);
    }}

    /* === ENHANCED FILTER HEADER === */
    .filter-header {{
        text-align: center;
        margin-bottom: 2rem;
        position: relative;
    }}

    .filter-header h3 {{
        font-family: {'TT Norms' if font_available else 'Arial'}, 'Arial', sans-serif;
        font-size: 1.6em;
        margin: 0;
        background: linear-gradient(135deg, #ff8765, #f4d602);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        background-clip: text;
        text-shadow: none;
        font-weight: 600;
    }}

    .filter-header p {{
        margin: 0.5rem 0 0 0;
        font-size: 0.95em;
        color: #000000;
        font-style: italic;
        opacity: 0.8;
    }}

    /* === ENHANCED SELECTBOX STYLING === */
    .stSelectbox > label {{
        font-weight: 700 !important;
        color: #000000 !important;
        font-size: 0.95em !important;
        margin-bottom: 0.8rem !important;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }}

    .stSelectbox > div > div {{
        background: rgba(255, 255, 255, 0.9) !important;
        border: 2px solid !important;
        border-radius: 16px !important;
        transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1) !important;
        box-shadow: 0 4px 12px rgba(237, 174, 73, 0.08) !important;
        backdrop-filter: blur(8px);
    }}

    .stSelectbox > div > div:hover {{
        border-color: rgba(255, 135, 101, 0.6) !important;
        box-shadow: 0 8px 20px rgba(237, 174, 73, 0.15) !important;
        transform: translateY(-1px);
    }}

    .stSelectbox > div > div:focus-within {{
        border-color: #000000 !important;
        box-shadow: 
            0 0 0 4px rgba(255, 135, 101, 0.15) !important,
            0 8px 24px rgba(237, 174, 73, 0.15) !important;
    }}

    /* === KPI CARDS === */
    .kpi-card {{
        border-radius: 16px;
        padding: 2rem;
        box-shadow: 0 8px 32px rgba(0, 0, 0, 0.1);
        border: 1px solid rgba(244, 214, 2, 0.2);
        margin-bottom: 2rem;
        transition: transform 0.3s ease, box-shadow 0.3s ease;
    }}
    
    /* Weekly Production Schedule Card */
    .kpi-card-wps {{
        background: linear-gradient(135deg, ##3c485c 0%, #2a364a 100%);
        color: #2a364a;
    }}
    
    /* Machine Utilization Card */
    .kpi-card-mu {{
        background: linear-gradient(135deg, #3c485c 0%, #2a364a 100%);
        color: #2a364a;
    }}
    
    /* YTD Production Schedule Card */
    .kpi-card-ytd {{
        background: linear-gradient(135deg, ##3c485c 0%, #2a364a 100%);
        color: #2a364a;
    }}
    
    /* KPI Title Styling */
    .kpi-label {{
        color: rgba(255, 255, 255, 0.9);
        font-size: 12px;
        font-weight: 600;
        margin-bottom: 8px;
        text-transform: uppercase;
        letter-spacing: 0.5px;
        font-family: 'TT Norms', 'Segoe UI', sans-serif;
    }}
    
    /* KPI Number Styling */
    .kpi-number {{
        color: #ffffff;
        font-size: 32px;
        font-weight: 700;
        line-height: 1;
        margin: 10px 0;
        font-family: 'TT Norms', 'Segoe UI', sans-serif;
        transition: all 0.3s ease;
    }}
    
    /* Target Styling */
    .kpi-unit {{
        color: rgba(255, 255, 255, 0.7);
        font-size: 14px;
        font-weight: 500;
        font-family: 'TT Norms', 'Segoe UI', sans-serif;
    }}
    
    /* Hover Effects */
    .kpi-card::before {{
        content: '';
        position: absolute;
        top: -2px;
        left: -2px;
        right: -2px;
        bottom: -2px;
        background: linear-gradient(45deg, #ff8765, #f4d602, #ff8765);
        background-size: 400% 400%;
        border-radius: 30px;
        z-index: -1;
        opacity: 0;
        transition: opacity 0.4s ease;
        animation: gradientShift 3s ease infinite;
    }}
    
    @keyframes gradientShift {{
        0%, 100% {{ background-position: 0% 50%; }}
        50% {{ background-position: 100% 50%; }}
    }}
    
    .kpi-card:hover {{
        transform: scale(1.05) translateY(-8px);
        box-shadow: 
            0 25px 50px rgba(244, 214, 2, 0.3),
            0 0 30px rgba(247, 212, 44, 0.2),
            inset 0 1px 0 rgba(255, 255, 255, 0.1);
    }}
    
    .kpi-card:hover::before {{
        opacity: 1;
    }}
    
    .kpi-card:hover .kpi-number {{
        transform: scale(1.1);
    }}

    /* === SKU TABLE === */
    .sku-table {{
        background: #fffef6;
        border-radius: 10px;
        overflow: hidden;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
        border: 1px solid #3b3f46;
    }}
    .table-header {{
        background: #1e2323;
        color: #f4d602;
        padding: 1rem;
        font-weight: bold;
        text-align: center;
    }}
    .sku-row {{
        padding: 1rem;
        border-bottom: 1px solid #3b3f46;
        transition: background-color 0.2s;
    }}
    .sku-row:hover {{
        background-color: rgba(244, 214, 2, 0.15);
    }}
    .sku-row:last-child {{
        border-bottom: none;
    }}

    /* === SIDEBAR (for fallback/secondary pages) === */
    [data-testid="stSidebar"] > div:first-child {{
        display: flex;
        flex-direction: column;
        height: 100vh;
        justify-content: flex-start;
        background-color: #fffef6;
        color: #f4d602;
    }}
    .css-1d391kg {{
        width: 280px !important;
        min-width: 280px !important;
    }}

    [data-testid="stSidebar"] h3 {{
        font-family: {'TT Norms' if font_available else 'Arial'}, 'Arial', sans-serif;
    }}

    /* === DASHBOARD CARDS === */
    .dashboard-card {{
        background: linear-gradient(135deg, #ffffff, #f8f9fa);
        border-radius: 16px;
        padding: 2rem;
        box-shadow: 0 8px 32px rgba(0, 0, 0, 0.1);
        border: 1px solid rgba(244, 214, 2, 0.2);
        margin-bottom: 2rem;
        transition: transform 0.3s ease, box-shadow 0.3s ease;
    }}
    
    .dashboard-card:hover {{
        transform: translateY(-4px);
        box-shadow: 0 12px 48px rgba(0, 0, 0, 0.15);
    }}
    
    .dashboard-card h3 {{
        color: #2c3e50;
        font-family: {'TT Norms' if font_available else 'Arial'}, sans-serif;
        font-size: 1.4rem;
        font-weight: 600;
        margin-bottom: 1rem;
        border-bottom: 2px solid #f4d602;
        padding-bottom: 0.5rem;
    }}

    /* === RESPONSIVE DESIGN === */
    @media (max-width: 768px) {{
        .nav-container {{
            padding: 0 1rem;
        }}
        
        .main-nav {{
            flex-direction: column;
            gap: 1rem;
        }}
        
        .main-header h1 {{
            font-size: 2em;
        }}
        
        .kpi-number {{
            font-size: 1.8em;
        }}
        
        .kpi-card {{
            min-height: 140px;
        }}
        
        .dashboard-card {{
            padding: 1.5rem;
        }}
    }}
</style>
""", unsafe_allow_html=True)


# --- STATION MAPPING ---
STATIONS = {
    'All Stations': {'color': '#34495e'},
    'Hot Kitchen': {'color': '#e74c3c'},
    'Cold Sauce': {'color': '#f39c12'},
    'Fabrication': {'color': '#3498db'},
    'Pastry': {'color': '#9b59b6'}
}

STATION_RANGES = {
    'Hot Kitchen': [(8, 36), (38, 72)],  # Hot Sauce + Hot Savory
    'Cold Sauce': [(74, 113)],
    'Fabrication': [(115, 128), (130, 152)],  # Poultry + Meats
    'Pastry': [(154, 176)]
}

# FIXED Column mappings with proper key names
COLUMNS = {
    'sku': 1,                    # Column B
    'batch_qty': 2,              # Column C
    'kg_per_hr': 3,              # Column D
    'man_hr': 4,                 # Column E
    'hrs_per_batch': 5,          # Column F
    'std_manpower': 7,           # Column H
    'frequency': 9,
    'batches_start': 10,          # Column I (18Aug)
    'batches_end': 16,           # Column O (24Aug)
    'volume_start': 19,          # Column R (18Aug volume)
    'volume_end': 25,            # Column X (24Aug volume)
    'hours_start': 27,           # Column Z (18Aug hours)
    'hours_end': 33,             # Column AF (24Aug hours)
    'manpower_start': 43,        # Column AP (18Aug manpower)
    'manpower_end': 49,          # Column AV (24Aug manpower)
    'overtime_start': 51,        # Column AX (18Aug individual OT)
    'overtime_end': 57,          # Column BD (24Aug individual OT)
    'overtime_percentage_start': 67, # Column BN (18Aug overall %)
    'overtime_percentage_end': 73,   # Column BS (24Aug overall %)
    'data_start_row': 6,         # Row 7 (0-based index 6) - where actual SKU data starts
    'header_row': 5              # Row 6 (0-based index 5) - where date headers are
}

# Machine Utilization Column mappings
MACHINE_COLUMNS = {
    'machine': 1,                # Column B (machines in rows 7-22)
    'rated_capacity': 2,           # Column C (kg/hr)
    'ideal_run_time': 3,         # Column D (ideal machine run time rate)
    'working_hours': 4,          # Column E 
    'run_time': 5,              # Column F
    'qty': 6,                   # Column G
    'available_hrs': 7,         # Column H
    'gwa': 8,
    'needed_hrs_start': 10,        # Column K (run hrs data)
    'needed_hrs_end': 16,          # Column Q
    'remaining_hrs_start': 18,  # Column S (available hours data)
    'remaining_hrs_end': 24,    # Column Y
    'machine_needed_start': 26, # Column AA (machine needed data)
    'machine_needed_end': 32,   # Column AG
    'capacity_utilization_start': 34,
    'capacity_utilization_end': 40,
    'machine_start_row': 6,     # Row 7 (0-based index 6)
    'machine_end_row': 23,      # Row 22 (0-based index 21)
    'header_row': 1             # Row 2 (0-based index 1)
}

# Constants for YTD Production data structure
YTD_COLUMNS = {
    'subrecipe': 1,          # Column B
    'batch_qty': 2,          # Column C  
    'kg_per_mhr': 3,         # Column D
    'mhr_per_kg': 4,         # Column E
    'hrs_per_run': 5,        # Column F
    'working_hrs': 6,        # Column G
    'std_manpower': 7,       # Column H
    'data_start': 8,         # Column I (Jan Week 1)
    'data_end': 372           # Approximate end column (Dec Week 53)
}

# Updated station ranges based on your specification
STATION_RANGES = {
    'Hot Kitchen': [(8, 36), (38, 72)],  # Hot Sauce + Hot Savory
    'Cold Sauce': [(74, 113)],
    'Fabrication': [(115, 128), (130, 152)],  # Poultry + Meats
    'Pastry': [(154, 176)]
}

# Key production summary rows
PRODUCTION_SUMMARY_ROWS = {
    'total_stations': 5,        # Row 6 (0-indexed as 5)
    'hot_kitchen_sauce': 6,     # Row 7
    'hot_kitchen_savory': 36,   # Row 37
    'cold_sauce': 72,           # Row 73
    'fab_poultry': 113,         # Row 114
    'fab_meats': 128,           # Row 129
    'pastry': 152               # Row 153
}


# --- UTILITY FUNCTIONS ---
def safe_float_convert(value):
    """Safely convert a value to float, handling various edge cases - ONLY POSITIVE VALUES"""
    if not value:
        return 0.0
    
    # Convert to string and clean up
    str_val = str(value).strip()
    
    # Handle empty strings
    if not str_val or str_val == '':
        return 0.0
    
    # Handle explicit zeros
    if str_val == '0':
        return 0.0
    
    # Handle negative signs with no number or malformed decimals
    if str_val in ['-', '-.', '- .', '-.0', '- .0']:
        return 0.0
    
    # Try to clean up the string
    str_val = str_val.replace(' ', '')  # Remove spaces
    
    # Handle cases like '-.' or '-0'
    if str_val.startswith('-') and len(str_val) <= 2:
        return 0.0
    
    # NEW: Check if the string starts with a negative sign - treat as 0
    if str_val.startswith('-'):
        return 0.0
    
    try:
        result = float(str_val)
        # NEW: If the result is negative, return 0
        return max(0.0, result)
    except ValueError:
        # If conversion fails, return 0
        return 0.0

# --- Helper functions for safe summing ---
def safe_sum(values):
    """Safely sum a list, ignoring None or non-numeric values."""
    return sum(v for v in values if isinstance(v, (int, float)) and v is not None)

def safe_sum_for_day(values, index):
    """Safely get a value for a specific day, ignoring errors."""
    if index < len(values):
        v = values[index]
        return v if isinstance(v, (int, float)) and v is not None else 0
    return 0


# --- LOAD KPI DASHBOARD ---
@st.cache_data(ttl=60)
def load_kpi_data(sheet_index=3):
    """Load KPI data from Google Sheets (sheet index 0) with last modified time"""
    credentials = load_credentials_kpi()
    if not credentials:
        return pd.DataFrame(), pd.DataFrame(), None
    
    try:
        gc = gspread.authorize(credentials)
        spreadsheet_id = "12ScL8L6Se7jRTqM2nL3hboxQkc8MLhEp7hEDlGUPKZg"
        sh = gc.open_by_key(spreadsheet_id)
        
        # Get the spreadsheet metadata for last modified time - USE CORRECT METHOD
        # Try different ways to get the modified time
        last_modified_time = None
        
        # Method 1: Try using the drive API for file metadata
        try:
            drive_service = build('drive', 'v3', credentials=credentials)
            file_metadata = drive_service.files().get(fileId=spreadsheet_id, fields='modifiedTime').execute()
            last_modified_time = file_metadata.get('modifiedTime')
        except Exception as e:
            st.write(f"❌ Drive API method failed: {e}")
        
        # Method 2: Try the spreadsheet metadata (fallback)
        if last_modified_time is None:
            try:
                spreadsheet_metadata = sh.fetch_sheet_metadata()
                last_modified_time = spreadsheet_metadata.get('properties', {}).get('modifiedTime')
            except Exception as e:
                st.write(f"❌ Sheets metadata method failed: {e}")
        
        # Method 3: If still None, use current time as fallback
        if last_modified_time is None:
            last_modified_time = datetime.now().isoformat() + 'Z'
        
        worksheet = sh.get_worksheet(sheet_index)
        data = worksheet.get_all_values()

        if len(data) < 4:
            st.warning("Not enough data in the spreadsheet")
            return pd.DataFrame(), pd.DataFrame(), last_modified_time
            
        # Extract headers (row 2), targets (row 3), and data rows (from row 4)
        headers = data[1]  # Row 2 (0-indexed as row 1)
        targets = data[2]  # Row 3 (0-indexed as row 2)
        ytd = data[3]      # Row 4 (0-indexed as row 2)
        data_rows = data[4:60]  # From row 5 onwards
        
        # Create DataFrame
        df = pd.DataFrame(data_rows, columns=headers)
        df = df.replace('', pd.NA).dropna(how='all')  # Remove completely empty rows
        
        # Create targets DataFrame
        targets_df = pd.DataFrame([targets], columns=headers)
        
        return df, targets_df, last_modified_time
        
    except Exception as e:
        st.error(f"Error loading KPI data: {str(e)}")
        return pd.DataFrame(), pd.DataFrame(), None

# --- UTILITY FUNCTIONS ---
def safe_float(value, default=0.0):
    """Safely convert string to float"""
    try:
        if value == '' or value is None or pd.isna(value):
            return default
        # Remove % and other characters
        clean_value = str(value).replace('%', '').replace(',', '').replace('₱', '').strip()
        return float(clean_value)
    except (ValueError, TypeError):
        return default

def format_kpi_value(value, kpi_type):
    """Format KPI values based on type - return N/A for empty values"""
    try:
        # Check if value is empty, None, or NaN
        if value == '' or value is None or pd.isna(value) or str(value).strip() == '':
            return "N/A"
            
        num_value = safe_float(value)
        
        # Check if the numeric value is 0 (which might indicate empty data)
        if num_value == 0 and str(value).strip() in ['', '0', '0.0', '0.00']:
            return "N/A"
            
        if kpi_type in ['percentage', 'yield', 'efficiency', 'quality', 'attendance', 'ot']:
            return f"{num_value:.2f}%"
        elif kpi_type in ['currency', 'labor_cost']:
            return f"₱{num_value:,.2f}"
        elif kpi_type in ['count', 'manpower']:
            return f"{int(num_value):,}"
        else:
            return f"{num_value:.2f}"
    except:
        return "N/A"

def get_kpi_color(current, target, kpi_type):
    """Determine color based on KPI performance vs target"""
    try:
        # Return gray for empty/NA values
        if current == '' or current is None or pd.isna(current) or str(current).strip() == '':
            return "#64748b"  # Gray for empty data
            
        current_val = safe_float(current)
        target_str = str(target).strip()
        
        # Handle symbolic targets (>, <)
        if target_str.startswith('>'):
            # Target is "greater than" value
            target_val = safe_float(target_str[1:])
            if kpi_type in ['spoilage', 'variance', 'labor_cost', 'overtime']:
                # For cost-based KPIs with > target: current should be LESS than target
                if current_val < target_val:
                    return "#22c55e"  # Green (good)
                else:
                    return "#ef4444"  # Red (bad)
            else:
                # For performance KPIs with > target: current should be GREATER than target
                if current_val > target_val:
                    return "#22c55e"  # Green (good)
                else:
                    return "#ef4444"  # Red (bad)
                    
        elif target_str.startswith('<'):
            # Target is "less than" value
            target_val = safe_float(target_str[1:])
            if kpi_type in ['spoilage', 'variance', 'labor_cost', 'overtime']:
                # For cost-based KPIs with < target: current should be LESS than target
                if current_val < target_val:
                    return "#22c55e"  # Green (good)
                else:
                    return "#ef4444"  # Red (bad)
            else:
                # For performance KPIs with < target: current should be LESS than target
                if current_val < target_val:
                    return "#22c55e"  # Green (good)
                else:
                    return "#ef4444"  # Red (bad)
                    
        else:
            # Regular numeric target (no symbols)
            target_val = safe_float(target_str)
            if target_val == 0:
                return "#4f7dbd"  # Gray for no target
                
            # For cost-based KPIs, lower is better
            if kpi_type in ['spoilage', 'variance', 'labor_cost', 'overtime']:
                if current_val <= target_val:
                    return "#22c55e"  # Green (good)
                else:
                    return "#ef4444"  # Red (bad)
            else:
                # For performance KPIs, higher is better
                if current_val >= target_val:
                    return "#22c55e"  # Green (good)
                else:
                    return "#ef4444"  # Red (bad)
                    
    except:
        return "#64748b"  # Gray for errors

# --- DASHBOARD COMPONENTS ---
def create_kpi_card(title, value, target, kpi_type, size="small"):
    """Create a modern KPI card with hover effects"""
    # Check if value is empty and show N/A - FIXED: Handle pd.NA properly
    if value is None or (isinstance(value, (int, float, str)) and value == '') or pd.isna(value):
        formatted_value = "N/A"
        color = "#64748b"  # Gray for empty data
    else:
        # Also handle string representations of empty values
        try:
            if str(value).strip() == '':
                formatted_value = "N/A"
                color = "#64748b"
            else:
                formatted_value = format_kpi_value(value, kpi_type)
                color = get_kpi_color(value, target, kpi_type)
        except:
            formatted_value = "N/A"
            color = "#64748b"
    
    # Preserve the < and > symbols from the target
    target_str = str(target)
    if target_str.startswith(('<', '>')):
        symbol = target_str[0]
        numeric_part = target_str[1:]
        try:
            numeric_value = float(numeric_part.strip('%'))
            if kpi_type == "percentage":
                formatted_numeric = f"{numeric_value:.1f}%"
            elif kpi_type == "currency":
                formatted_numeric = f"₱{numeric_value:,.2f}" if kpi_type == "currency" else f"${numeric_value:,.2f}"
            else:
                formatted_numeric = f"{numeric_value:,.0f}"
            formatted_target = f"{symbol}{formatted_numeric}"
        except ValueError:
            formatted_target = target_str
    else:
        formatted_target = format_kpi_value(target, kpi_type)
        
    if size == "large":
        card_height = "200px"
        title_size = "18px"
        value_size = "50px"
    else:
        card_height = "140px"
        title_size = "12px"
        value_size = "32px"

    # Check if target is empty and conditionally include target line
    target_html = ""
    if target and str(target).strip() and str(target).strip().lower() not in ['', 'nan', 'none', '0', '0.0', '0.00']:
        target_html = f"""<div style="
            color: #acb4bf;
            font-size: 14px;
            font-weight: 500;
        ">Target: {formatted_target}</div>"""
    
    card_html = f"""
    <div class="kpi-card" style="height: {card_height};">
        <div style="
            color: #94a3b8;
            font-size: {title_size};
            font-weight: 600;
            margin-bottom: 8px;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        ">{title}</div>
        <div class="kpi-number" style="
            color: {color};
            font-size: {value_size};
            font-weight: 700;
            line-height: 1;
            margin: 10px 0;
        ">{formatted_value}</div>
        {target_html}
    </div>
    """
    return card_html

def create_volume_chart(kpi_data, week_column):
    """Create a modern volume chart using Plotly"""
    try:
        import plotly.graph_objects as go
        import plotly.express as px
        from plotly.subplots import make_subplots
        
        # Find the volume column (assuming it's one of the first few columns after week)
        volume_column = None
        for col in kpi_data.columns[2:6]:  # Check columns 2-5 for volume data
            if any(term in str(col).lower() for term in ['volume', 'vol', 'production', 'output']):
                volume_column = col
                break
        
        # If not found by name, use the first data column after week
        if volume_column is None:
            volume_column = kpi_data.columns[2] if len(kpi_data.columns) > 2 else kpi_data.columns[1]
        
        # Prepare data
        chart_data = []
        for _, row in kpi_data.iterrows():
            week = str(row[week_column]).strip()
            volume = safe_float(row[volume_column]) if volume_column in row.index else 0
            
            # Skip empty weeks or zero volumes
            if week and week.lower() not in ['', 'nan', 'none'] and volume > 0:
                chart_data.append({
                    'Week': week,
                    'Volume': volume
                })
        
        if not chart_data:
            st.warning("No volume data available for charting.")
            return
        
        # Convert to DataFrame for easier handling
        chart_df = pd.DataFrame(chart_data)
        
        # Create the modern chart
        fig = go.Figure()
        
        # Add the main bar chart with gradient colors
        fig.add_trace(go.Bar(
            x=chart_df['Week'],
            y=chart_df['Volume'],
            name='Volume',
            marker=dict(
                color=chart_df['Volume'],
                colorscale=[
                    [0, '#1e293b'],      # Dark slate
                    [0.3, '#3b82f6'],    # Blue
                    [0.6, '#06b6d4'],    # Cyan  
                    [1, '#10b981']       # Emerald
                ],
                line=dict(color='rgba(255,255,255,0.2)', width=1),
                opacity=0.9
            ),
            text=[f'{v:.1f}' for v in chart_df['Volume']],
            textposition='outside',
            textfont=dict(
                color='black',
                size=10,
                family='Segoe UI'
            ),
            hovertemplate='<b>%{x}</b><br>Volume: %{y:.1f}<br><extra></extra>',
            hoverlabel=dict(
                bgcolor='rgba(248, 246, 240, 0.95)',
                bordercolor='rgba(160, 174, 192, 0.4)',
                font=dict(color='#4a5568', family='Segoe UI')
            )
        ))
        

        
        # Update layout with modern styling
        fig.update_layout(
            title=dict(
                text='<b>Weekly Volume Performance</b>',
                x=0.5,
                xanchor='center',
                font=dict(
                    size=24,
                    color='#2c3e50',
                    family='TT Norms'
                )
            ),
            xaxis=dict(
                title=dict(text='Week', font=dict(size=14, color='#5a6c7d', family='Segoe UI')),
                tickfont=dict(size=11, color='#4a5568', family='Segoe UI'),
                gridcolor='rgba(160, 174, 192, 0.3)',
                zeroline=False,
                tickangle=-45
            ),
            yaxis=dict(
                title=dict(text='Volume', font=dict(size=14, color='#5a6c7d', family='Segoe UI')),
                tickfont=dict(size=11, color='#4a5568', family='Segoe UI'),
                gridcolor='rgba(160, 174, 192, 0.3)',
                zeroline=False
            ),
            plot_bgcolor='#ffffff',
            paper_bgcolor='#ffffff',
            font=dict(family='Segoe UI'),
            showlegend=True,
            legend=dict(
                x=0.02,
                y=0.98,
                bgcolor='rgba(248, 246, 240, 0.9)',
                bordercolor='rgba(160, 174, 192, 0.4)',
                borderwidth=1,
                font=dict(color='#4a5568', size=11)
            ),
            margin=dict(l=60, r=40, t=80, b=100),
            height=450,
            hovermode='x unified'
        )
        
        # Add subtle animations
        fig.update_traces(
            marker=dict(
                line=dict(width=1.5)
            )
        )
        
        # Display the chart
        st.plotly_chart(fig, use_container_width=True, config={
            'displayModeBar': False,
            'staticPlot': False
        })
        
        # Custom CSS for rounded corners
        st.markdown("""
            <style>
            .js-plotly-plot .plotly .modebar {
                display: none;
            }
            .js-plotly-plot .plotly {
                border-radius: 25px !important;
                overflow: hidden;
            }
            .js-plotly-plot .plotly .main-svg {
                border-radius: 25px !important;
            }
            </style>
        """, unsafe_allow_html=True)
        
    except ImportError:
        st.error("Plotly is required for charts. Please install it: pip install plotly")
    except Exception as e:
        st.error(f"Error creating volume chart: {str(e)}")


def create_multi_kpi_chart(kpi_data, week_column):
    """Create a multi-KPI line chart with dropdown selector"""
    try:
        import plotly.graph_objects as go
        
        # Define KPI mappings (display name -> column index, format type)
        kpi_mappings = {
            "Volume": (2, "count"),
            "Production Plan Performance": (3, "percentage"),
            "Capacity Utilization": (4, "percentage"),
            "Production Plan Compliance": (5, "percentage"),
            "Spoilage": (6, "percentage"),
            "Variance": (7, "percentage"),
            "Yield": (8, "percentage"),
            "Attendance": (9, "percentage"),
            "Overtime": (10, "percentage"),
            "Labor cost per kilo (₱)": (11, "currency"),
            "Availability": (12, "percentage"),
            "Efficiency": (13, "percentage"),
            "Quality": (14, "percentage"),
            "OEE": (15, "percentage"),
            "Man-hr": (16, "count"),
            "KGMH": (17, "count"),
            "Manpower": (18, "count"),
            "Revenue": (19, "currency"),
            "Spoilage Cost%": (20, "percentage"),
            "Spoilage vs Revenue": (21, "percentage")
        }
        
        # KPI selector
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            selected_kpi = st.selectbox(
                "Select KPI to View",
                list(kpi_mappings.keys()),
                index=0,
                key="kpi_selector"
            )
        
        # Get selected KPI details
        column_index, format_type = kpi_mappings[selected_kpi]
        
        # Prepare data for the selected KPI
        chart_data = []
        for _, row in kpi_data.iterrows():
            week = str(row[week_column]).strip()
            if column_index < len(row):
                value = safe_float(row.iloc[column_index]) if row.iloc[column_index] != '' else None
            else:
                value = None
            
            # Skip empty weeks but include weeks with zero values for KPIs
            if week and week.lower() not in ['', 'nan', 'none'] and value is not None:
                chart_data.append({
                    'Week': week,
                    'Value': value
                })
        
        if not chart_data:
            st.warning(f"No {selected_kpi} data available for charting.")
            return
        
        # Convert to DataFrame
        chart_df = pd.DataFrame(chart_data)
        
        # Create the line chart
        fig = go.Figure()
        
        # Add the line trace with markers
        fig.add_trace(go.Scatter(
            x=chart_df['Week'],
            y=chart_df['Value'],
            mode='lines+markers',
            name=selected_kpi,
            line=dict(
                color='#f59e0b',  # Golden color for the line
                width=3
            ),
            marker=dict(
                color='#f59e0b',
                size=8,
                line=dict(color='#d97706', width=2)
            ),
            hovertemplate=f'<b>%{{x}}</b><br>{selected_kpi}: %{{y:.1f}}<br><extra></extra>',
            hoverlabel=dict(
                bgcolor='rgba(248, 246, 240, 0.95)',
                bordercolor='rgba(160, 174, 192, 0.4)',
                font=dict(color='#4a5568', family='Segoe UI')
            )
        ))
        
        # Format y-axis title based on KPI type
        if format_type == "percentage":
            y_title = f"{selected_kpi} (%)"
        elif format_type == "currency":
            y_title = f"{selected_kpi} (₱)"
        else:
            y_title = selected_kpi
        
        # Update layout
        fig.update_layout(
            title=dict(
                text=f'<b>{selected_kpi} Trend</b>',
                x=0.5,
                xanchor='center',
                font=dict(
                    size=20,
                    color='#2c3e50',
                    family='TT Norms'
                )
            ),
            xaxis=dict(
                title=dict(text='Week', font=dict(size=14, color='#5a6c7d', family='Segoe UI')),
                tickfont=dict(size=11, color='#4a5568', family='Segoe UI'),
                gridcolor='rgba(160, 174, 192, 0.3)',
                zeroline=False,
                tickangle=-45
            ),
            yaxis=dict(
                title=dict(text=y_title, font=dict(size=14, color='#5a6c7d', family='Segoe UI')),
                tickfont=dict(size=11, color='#4a5568', family='Segoe UI'),
                gridcolor='rgba(160, 174, 192, 0.3)',
                zeroline=False
            ),
            plot_bgcolor='#ffffff',
            paper_bgcolor='#ffffff',
            font=dict(family='Segoe UI'),
            showlegend=False,
            margin=dict(l=60, r=40, t=80, b=100),
            height=450,
            hovermode='x unified'
        )
        
        # Display the chart
        st.plotly_chart(fig, use_container_width=True, config={
            'displayModeBar': False,
            'staticPlot': False
        })
        
        # Custom CSS for rounded corners
        st.markdown("""
            <style>
            .js-plotly-plot .plotly .modebar {
                display: none;
            }
            .js-plotly-plot .plotly {
                border-radius: 25px !important;
                overflow: hidden;
            }
            .js-plotly-plot .plotly .main-svg {
                border-radius: 25px !important;
            }
            </style>
        """, unsafe_allow_html=True)
        
    except ImportError:
        st.error("Plotly is required for charts. Please install it: pip install plotly")
    except Exception as e:
        st.error(f"Error creating multi-KPI chart: {str(e)}")


def display_volume_section():
    """Display the volume chart section in your KPI dashboard"""
    try:
        # Load data
        kpi_data, targets_data, last_modified_time = load_kpi_data()
        
        if kpi_data.empty:
            st.warning("No data available for volume chart.")
            return
        
        # Find week column (same logic as your main dashboard)
        week_column = None
        for col in kpi_data.columns:
            col_lower = str(col).lower()
            if any(term in col_lower for term in ['week', 'wk']):
                week_column = col
                break
        
        if week_column is None:
            week_column = kpi_data.columns[1] if len(kpi_data.columns) > 1 else kpi_data.columns[0]
        
        # Add section header with modern styling
        st.markdown("<br><br>", unsafe_allow_html=True)
        
        # Create and display the volume chart
        create_volume_chart(kpi_data, week_column)
        
    except Exception as e:
        st.error(f"Error displaying volume section: {str(e)}")


def create_multi_kpi_chart(kpi_data, week_column):
    """Create a multi-KPI line chart with dropdown selector"""
    try:
        import plotly.graph_objects as go
        
        # Define KPI mappings (display name -> column index, format type)
        kpi_mappings = {
            "Volume": (2, "count"),
            "Production Plan Performance": (3, "percentage"),
            "Capacity Utilization": (4, "percentage"),
            "Production Plan Compliance": (5, "percentage"),
            "Spoilage": (6, "percentage"),
            "Variance": (7, "percentage"),
            "Yield": (8, "percentage"),
            "Attendance": (9, "percentage"),
            "Overtime": (10, "percentage"),
            "Labor cost per kilo (₱)": (11, "currency"),
            "Availability": (12, "percentage"),
            "Efficiency": (13, "percentage"),
            "Quality": (14, "percentage"),
            "OEE": (15, "percentage"),
            "Man-hr": (16, "count"),
            "KGMH": (17, "count"),
            "Manpower": (18, "count"),
            "Revenue": (19, "currency"),
            "Spoilage Cost%": (20, "percentage"),
            "Spoilage vs Revenue": (21, "percentage")
        }
        
        # KPI selector
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            selected_kpi = st.selectbox(
                "Select KPI to View",
                list(kpi_mappings.keys()),
                index=0,
                key="kpi_selector"
            )
        
        # Get selected KPI details
        column_index, format_type = kpi_mappings[selected_kpi]
        
        # Prepare data for the selected KPI
        chart_data = []
        for _, row in kpi_data.iterrows():
            week = str(row[week_column]).strip()
            
            # Safe value extraction
            if column_index < len(row):
                raw_value = row.iloc[column_index]
                # Handle pandas NA values properly
                if pd.isna(raw_value) or raw_value == '' or raw_value is None:
                    value = None
                else:
                    value = safe_float(raw_value)
            else:
                value = None
            
            # Skip empty weeks but include weeks with zero values for KPIs
            week_valid = week and str(week).lower() not in ['', 'nan', 'none']
            value_valid = value is not None and not pd.isna(value)
            
            if week_valid and value_valid:
                chart_data.append({
                    'Week': week,
                    'Value': value
                })
        
        if not chart_data:
            st.warning(f"No {selected_kpi} data available for charting.")
            return
        
        # Convert to DataFrame
        chart_df = pd.DataFrame(chart_data)
        
        # Create the line chart
        fig = go.Figure()
        
        # Add the line trace with markers
        fig.add_trace(go.Scatter(
            x=chart_df['Week'],
            y=chart_df['Value'],
            mode='lines+markers',
            name=selected_kpi,
            line=dict(
                color='#f59e0b',  # Golden color for the line
                width=3
            ),
            marker=dict(
                color='#f59e0b',
                size=8,
                line=dict(color='#d97706', width=2)
            ),
            hovertemplate=f'<b>%{{x}}</b><br>{selected_kpi}: %{{y:.1f}}<br><extra></extra>',
            hoverlabel=dict(
                bgcolor='rgba(248, 246, 240, 0.95)',
                bordercolor='rgba(160, 174, 192, 0.4)',
                font=dict(color='#4a5568', family='Segoe UI')
            )
        ))
        
        # Format y-axis title based on KPI type
        if format_type == "percentage":
            y_title = f"{selected_kpi} (%)"
        elif format_type == "currency":
            y_title = f"{selected_kpi} (₱)"
        else:
            y_title = selected_kpi
        
        # Update layout
        fig.update_layout(
            title=dict(
                text=f'<b>{selected_kpi} Trend</b>',
                x=0.5,
                xanchor='center',
                font=dict(
                    size=20,
                    color='#2c3e50',
                    family='TT Norms'
                )
            ),
            xaxis=dict(
                title=dict(text='Week', font=dict(size=14, color='#5a6c7d', family='Segoe UI')),
                tickfont=dict(size=11, color='#4a5568', family='Segoe UI'),
                gridcolor='rgba(160, 174, 192, 0.3)',
                zeroline=False,
                tickangle=-45
            ),
            yaxis=dict(
                title=dict(text=y_title, font=dict(size=14, color='#5a6c7d', family='Segoe UI')),
                tickfont=dict(size=11, color='#4a5568', family='Segoe UI'),
                gridcolor='rgba(160, 174, 192, 0.3)',
                zeroline=False
            ),
            plot_bgcolor='#ffffff',
            paper_bgcolor='#ffffff',
            font=dict(family='Segoe UI'),
            showlegend=False,
            margin=dict(l=60, r=40, t=80, b=100),
            height=450,
            hovermode='x unified'
        )
        
        # Display the chart
        st.plotly_chart(fig, use_container_width=True, config={
            'displayModeBar': False,
            'staticPlot': False
        })
        
        # Custom CSS for rounded corners
        st.markdown("""
            <style>
            .js-plotly-plot .plotly .modebar {
                display: none;
            }
            .js-plotly-plot .plotly {
                border-radius: 25px !important;
                overflow: hidden;
            }
            .js-plotly-plot .plotly .main-svg {
                border-radius: 25px !important;
            }
            </style>
        """, unsafe_allow_html=True)
        
    except ImportError:
        st.error("Plotly is required for charts. Please install it: pip install plotly")
    except Exception as e:
        st.error(f"Error creating multi-KPI chart: {str(e)}")


def display_volume_section():
    """Display the volume chart section in your KPI dashboard"""
    try:
        # Load data
        kpi_data, targets_data, last_modified_time = load_kpi_data()
        
        if kpi_data.empty:
            st.warning("No data available for volume chart.")
            return
        
        # Find week column (same logic as your main dashboard)
        week_column = None
        for col in kpi_data.columns:
            col_lower = str(col).lower()
            if any(term in col_lower for term in ['week', 'wk']):
                week_column = col
                break
        
        if week_column is None:
            week_column = kpi_data.columns[1] if len(kpi_data.columns) > 1 else kpi_data.columns[0]
        
        # Add section header with modern styling
        st.markdown("<br><br>", unsafe_allow_html=True)
        
        # Create and display the volume chart
        create_volume_chart(kpi_data, week_column)
        
    except Exception as e:
        st.error(f"Error displaying volume section: {str(e)}")


def create_kpi_scatter_chart(kpi_data, week_column):
    """Create a scatter plot with dual KPI selection for X and Y axes"""
    try:
        import plotly.graph_objects as go
        
        # Define KPI mappings (display name -> column index, format type)
        kpi_mappings = {
            "Volume": (2, "count"),
            "Production Plan Performance": (3, "percentage"),
            "Capacity Utilization": (4, "percentage"),
            "Production Plan Compliance": (5, "percentage"),
            "Spoilage": (6, "percentage"),
            "Variance": (7, "percentage"),
            "Yield": (8, "percentage"),
            "Attendance": (9, "percentage"),
            "Overtime": (10, "percentage"),
            "Labor cost per kilo (₱)": (11, "currency"),
            "Availability": (12, "percentage"),
            "Efficiency": (13, "percentage"),
            "Quality": (14, "percentage"),
            "OEE": (15, "percentage"),
            "Man-hr": (16, "count"),
            "KGMH": (17, "count"),
            "Manpower": (18, "count"),
            "Revenue": (19, "currency"),
            "Spoilage Cost%": (20, "percentage"),
            "Spoilage vs Revenue": (21, "percentage")
        }
        
        # KPI selectors for X and Y axes
        col1, col2 = st.columns(2)
        with col1:
            x_kpi = st.selectbox(
                "X-Axis KPI",
                list(kpi_mappings.keys()),
                index=10,  # Default to "Labor cost per kilo (₱)"
                key="x_kpi_selector"
            )
        with col2:
            y_kpi = st.selectbox(
                "Y-Axis KPI",
                list(kpi_mappings.keys()),
                index=0,   # Default to "Volume"
                key="y_kpi_selector"
            )
        
        # Get selected KPI details
        x_column_index, x_format_type = kpi_mappings[x_kpi]
        y_column_index, y_format_type = kpi_mappings[y_kpi]
        
        # Prepare data for the scatter plot
        scatter_data = []
        for _, row in kpi_data.iterrows():
            week = str(row[week_column]).strip()
            
            # Extract X value
            if x_column_index < len(row):
                raw_x = row.iloc[x_column_index]
                if pd.isna(raw_x) or raw_x == '' or raw_x is None:
                    x_value = None
                else:
                    x_value = safe_float(raw_x)
            else:
                x_value = None
            
            # Extract Y value
            if y_column_index < len(row):
                raw_y = row.iloc[y_column_index]
                if pd.isna(raw_y) or raw_y == '' or raw_y is None:
                    y_value = None
                else:
                    y_value = safe_float(raw_y)
            else:
                y_value = None
            
            # Only include points where both X and Y have values
            week_valid = week and str(week).lower() not in ['', 'nan', 'none']
            x_valid = x_value is not None and not pd.isna(x_value)
            y_valid = y_value is not None and not pd.isna(y_value)
            
            if week_valid and x_valid and y_valid:
                scatter_data.append({
                    'Week': week,
                    'X': x_value,
                    'Y': y_value
                })
        
        if not scatter_data:
            st.warning(f"No data available for {x_kpi} vs {y_kpi} scatter plot.")
            return
        
        # Convert to DataFrame
        scatter_df = pd.DataFrame(scatter_data)
        
        # Create the scatter plot
        fig = go.Figure()
        
        # Add scatter trace
        fig.add_trace(go.Scatter(
            x=scatter_df['X'],
            y=scatter_df['Y'],
            mode='markers',
            name=f'{y_kpi} vs {x_kpi}',
            marker=dict(
                color='#3b82f6',  # Blue color for markers
                size=10,
                line=dict(color='#1e40af', width=2),
                opacity=0.8
            ),
            text=scatter_df['Week'],
            hovertemplate=f'<b>%{{text}}</b><br>{x_kpi}: %{{x:.1f}}<br>{y_kpi}: %{{y:.1f}}<br><extra></extra>',
            hoverlabel=dict(
                bgcolor='rgba(248, 246, 240, 0.95)',
                bordercolor='rgba(160, 174, 192, 0.4)',
                font=dict(color='#4a5568', family='TT Norms')
            )
        ))
        
        # Format axis titles based on KPI types
        if x_format_type == "percentage":
            x_title = f"{x_kpi} (%)"
        elif x_format_type == "currency":
            x_title = f"{x_kpi} (₱)"
        else:
            x_title = x_kpi
            
        if y_format_type == "percentage":
            y_title = f"{y_kpi} (%)"
        elif y_format_type == "currency":
            y_title = f"{y_kpi} (₱)"
        else:
            y_title = y_kpi
        
        # Update layout
        fig.update_layout(
            title=dict(
                text=f'<b>{y_kpi} vs {x_kpi}</b>',
                x=0.5,
                xanchor='center',
                font=dict(
                    size=20,
                    color='#2c3e50',
                    family='TT Norms'
                )
            ),
            xaxis=dict(
                title=dict(text=x_title, font=dict(size=14, color='#5a6c7d', family='Segoe UI')),
                tickfont=dict(size=11, color='#4a5568', family='Segoe UI'),
                gridcolor='rgba(160, 174, 192, 0.3)',
                zeroline=False
            ),
            yaxis=dict(
                title=dict(text=y_title, font=dict(size=14, color='#5a6c7d', family='Segoe UI')),
                tickfont=dict(size=11, color='#4a5568', family='Segoe UI'),
                gridcolor='rgba(160, 174, 192, 0.3)',
                zeroline=False
            ),
            plot_bgcolor='#ffffff',
            paper_bgcolor='#ffffff',
            font=dict(family='Segoe UI'),
            showlegend=False,
            margin=dict(l=60, r=40, t=80, b=60),
            height=450,
            hovermode='closest'
        )
        
        # Display the chart
        st.plotly_chart(fig, use_container_width=True, config={
            'displayModeBar': False,
            'staticPlot': False
        })
        
        # Custom CSS for rounded corners
        st.markdown("""
            <style>
            .js-plotly-plot .plotly .modebar {
                display: none;
            }
            .js-plotly-plot .plotly {
                border-radius: 25px !important;
                overflow: hidden;
            }
            .js-plotly-plot .plotly .main-svg {
                border-radius: 25px !important;
            }
            </style>
        """, unsafe_allow_html=True)
        
    except ImportError:
        st.error("Plotly is required for charts. Please install it: pip install plotly")
    except Exception as e:
        st.error(f"Error creating scatter chart: {str(e)}")


def display_multi_kpi_section():
    """Display the multi-KPI chart section with line and scatter charts side by side"""
    try:
        # Load data
        kpi_data, targets_data, last_modified_time = load_kpi_data()
        
        if kpi_data.empty:
            st.warning("No data available for KPI charts.")
            return
        
        # Find week column
        week_column = None
        for col in kpi_data.columns:
            col_lower = str(col).lower()
            if any(term in col_lower for term in ['week', 'wk']):
                week_column = col
                break
        
        if week_column is None:
            week_column = kpi_data.columns[1] if len(kpi_data.columns) > 1 else kpi_data.columns[0]
        
        # Create side-by-side layout
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("""
                <div style="text-align: center; margin: 20px 0 15px 0;">
                    <h3 style="color: #2c3e50; font-size: 18px; font-weight: 600;">
                        Trend Analysis
                    </h3>
                </div>
            """, unsafe_allow_html=True)
            create_multi_kpi_chart(kpi_data, week_column)
        
        with col2:
            st.markdown("""
                <div style="text-align: center; margin: 20px 0 15px 0;">
                    <h3 style="color: #2c3e50; font-size: 18px; font-weight: 600;">
                        Correlation Analysis
                    </h3>
                </div>
            """, unsafe_allow_html=True)
            create_kpi_scatter_chart(kpi_data, week_column)
        
    except Exception as e:
        st.error(f"Error displaying multi-KPI section: {str(e)}")
        
def display_kpi_dashboard():
    """Display the main KPI dashboard"""
    # Get Philippines timezone
    try:
        ph_timezone = pytz.timezone('Asia/Manila')
    except Exception as e:
        st.write(f"❌ Timezone error: {e}")
        ph_timezone = None
    
    st.markdown("""
        <style>
        .kpi-card {
            background: linear-gradient(135deg, #1e293b 0%, #334155 100%);
            border: 1px solid #475569;
            border-radius: 27px;
            padding: 20px;
            display: flex;
            flex-direction: column;
            justify-content: space-between;
            align-items: flex-start;
            box-shadow: 0 10px 25px rgba(0,0,0,0.2);
            transition: all 0.3s ease;
            position: relative;
            overflow: hidden;
            font-family: 'Segoe UI', sans-serif;
            text-align: left;
        }
        
        .kpi-card::before {
            content: '';
            position: absolute;
            top: 0;
            left: 0;
            right: 0;
            bottom: 0;
            background: linear-gradient(45deg, rgba(255, 255, 255, 0.1), rgba(255, 255, 255, 0.05));
            opacity: 0;
            transition: opacity 0.3s ease;
            z-index: 1;
        }
        
        .kpi-card:hover {
            transform: scale(1.05) translateY(-8px) rotateY(5deg);
            box-shadow: 
                0 25px 50px rgba(255, 255, 255, 0.2),
                0 0 30px rgba(255, 255, 255, 0.1),
                inset 0 1px 0 rgba(255, 255, 255, 0.05);
        }
        
        .kpi-card:hover::before {
            opacity: 1;
        }
        
        .kpi-number {
            transition: transform 0.3s ease;
        }
        
        .kpi-card:hover .kpi-number {
            transform: scale(1.1);
        }
    
        /* Your existing CSS styles below - KEEP THESE */
        .kpi-container {
            background: #0f172a;
            padding: 20px;
            border-radius: 30px;
            margin: 10px 0;
        }
        .dashboard-title {
            color: #000000;
            text-align: center;
            font-size: 32px;
            font-weight: 700;
            margin-bottom: 30px;
            text-transform: uppercase;
            letter-spacing: 2px;
            font-family: {'TT Norms' if font_available else 'Segoe UI'}, sans-serif;
        }
        .last-updated {
            color: #94a3b8;
            text-align: center;
            font-size: 14px;
            font-weight: 500;
            margin-bottom: 20px;
            font-style: italic;
        }
        .main {
            background-color: #0f172a;
        }
        .stSelectbox > div > div {
            background-color: #000000;
            border: 1px solid #475569;
            border-radius: 8px;
        }
        .stSelectbox label {
            color: #000000 !important;
            font-weight: 600;
        }
        .stSelectbox div[data-baseweb="select"] > div {
            color: #000000 !important;
        }
        div[role="listbox"] {
            background-color: #1e293b !important;
            color: #000000 !important;
        }
        div[role="option"] {
            color: #000000 !important;
            background-color: #1e293b !important;
        }
        div[role="option"]:hover {
            background-color: #334155 !important;
        }
        </style>
        """, unsafe_allow_html=True)
    
    try:
        # Load data WITH last modified time
        kpi_data, targets_data, last_modified_time = load_kpi_data()
        
        if kpi_data.empty:
            st.error("No KPI data available. Please check if the spreadsheet is accessible and contains data.")
            return

        
        # Try to find the week column
        week_column = None
        for col in kpi_data.columns:
            col_lower = str(col).lower()
            if any(term in col_lower for term in ['week', 'wk']):
                week_column = col
                break
        
        if week_column is None:
            for col in kpi_data.columns:
                sample_values = kpi_data[col].dropna().head(5).astype(str)
                if any(any(term in val.lower() for term in ['wk', 'week']) for val in sample_values):
                    week_column = col
                    break
        
        if week_column is None:
            week_column = kpi_data.columns[1] if len(kpi_data.columns) > 1 else kpi_data.columns[0]
        
        # Get available weeks - FIXED: Create week_options here
        weeks = kpi_data[week_column].dropna().unique()
        weeks = [str(w).strip() for w in weeks if str(w).strip() != '']
        week_options = weeks  # This was missing!
        
        if not weeks:
            st.error(f"No week data available in column '{week_column}'.")
            return
        
        # Find default week (most recent with data) - FIXED: Create current_week here
        default_week_index = None
        for i in range(len(kpi_data) - 1, -1, -1):
            if kpi_data.iloc[i, 2:22].notna().any():
                default_week_index = i
                break
        
        if default_week_index is not None:
            current_week = kpi_data.iloc[default_week_index][week_column]
            current_week = str(current_week).strip()
            if current_week in weeks:
                default_index = weeks.index(current_week)
            else:
                default_index = len(weeks) - 1
                current_week = weeks[default_index] if weeks else ""
        else:
            default_index = len(weeks) - 1
            current_week = weeks[default_index] if weeks else ""
        
        # Week selection - FIXED: Now using the correctly defined variables
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            selected_week = st.selectbox(
                "Select Week",
                options=week_options,
                index=week_options.index(current_week) if current_week in week_options else 0,
                key="kpi_week_selector",  # Unique key for KPI dashboard
                help="Select week to view KPI data"
            )
        
        # Filter data for selected week
        week_data = kpi_data[kpi_data[week_column] == selected_week]
        if week_data.empty:
            st.warning(f"No data available for {selected_week}")
            return
            
        week_row = week_data.iloc[0]
        target_row = targets_data.iloc[0] if not targets_data.empty else pd.Series()
        
        # Dashboard title
        st.markdown(f'<div class="dashboard-title">Key Performance Metrics - {selected_week}</div>', unsafe_allow_html=True)
        
        # Format the last modified time in Philippines time
        if last_modified_time:
            try:
                
                # Convert to datetime object
                dt = pd.to_datetime(last_modified_time)

                # Adjust to Philippines time
                if ph_timezone:
                    if dt.tzinfo is not None:
                        dt = dt.astimezone(ph_timezone)
                    else:
                        dt = dt.replace(tzinfo=pytz.UTC).astimezone(ph_timezone)
                
                formatted_time = dt.strftime("%b %d, %Y %I:%M %p")
                
            except Exception as e:
                formatted_time = str(last_modified_time)
        else:
            formatted_time = "Unknown time"
        
        # Display the actual spreadsheet last modified time
        st.markdown(f'<div class="last-updated">Last updated: {formatted_time}</div>', unsafe_allow_html=True)

# ---- 
        
        # Top KPI metrics row
        st.markdown("### Key Performance Indicators")
        cols = st.columns(6)
        
        kpi_metrics = [
            ("Spoilage vs Revenue", week_row.iloc[21] if len(week_row) > 21 else '', 
             target_row.iloc[21] if len(target_row) > 21 else '', "percentage"),
            ("Variance", week_row.iloc[7] if len(week_row) > 7 else '', 
             target_row.iloc[7] if len(target_row) > 7 else '', "percentage"),
            ("Yield", week_row.iloc[8] if len(week_row) > 8 else '', 
             target_row.iloc[8] if len(target_row) > 8 else '', "percentage"),
            ("Availability", week_row.iloc[12] if len(week_row) > 12 else '', 
             target_row.iloc[12] if len(target_row) > 12 else '', "percentage"),
            ("Efficiency", week_row.iloc[13] if len(week_row) > 13 else '', 
             target_row.iloc[13] if len(target_row) > 13 else '', "percentage"),
            ("Quality", week_row.iloc[14] if len(week_row) > 14 else '', 
             target_row.iloc[14] if len(target_row) > 14 else '', "percentage"),
        ]
        
        for i, (title, value, target, kpi_type) in enumerate(kpi_metrics):
            with cols[i]:
                st.markdown(create_kpi_card(title, value, target, kpi_type), unsafe_allow_html=True)
        
        st.markdown("<br>", unsafe_allow_html=True)
        
        # Second row of KPIs
        cols2 = st.columns(6)
        
        kpi_metrics_2 = [
            ("Man-hr", week_row.iloc[16] if len(week_row) > 16 else '', 
             target_row.iloc[16] if len(target_row) > 16 else '', "count"),
            ("Attendance", week_row.iloc[9] if len(week_row) > 9 else '', 
             target_row.iloc[9] if len(target_row) > 9 else '', "percentage"),
            ("Overtime %", week_row.iloc[10] if len(week_row) > 10 else '', 
             target_row.iloc[10] if len(target_row) > 10 else '', "percentage"),
            ("Labor Cost/kg", week_row.iloc[11] if len(week_row) > 11 else '', 
             target_row.iloc[11] if len(target_row) > 11 else '', "currency"),
            ("KGMH", week_row.iloc[17] if len(week_row) > 17 else '', 
             target_row.iloc[17] if len(target_row) > 17 else '', "count"),
            ("Manpower", week_row.iloc[18] if len(week_row) > 18 else '', 
             target_row.iloc[18] if len(target_row) > 18 else '', "count"),
        ]
        
        for i, (title, value, target, kpi_type) in enumerate(kpi_metrics_2):
            with cols2[i]:
                st.markdown(create_kpi_card(title, value, target, kpi_type), unsafe_allow_html=True)
        
        st.markdown("<br><br>", unsafe_allow_html=True)
        
        # Big KPI cards
        st.markdown("### Core Performance Metrics")
        big_cols = st.columns(4)
        
        big_kpis = [
            ("Capacity Utilization", week_row.iloc[4] if len(week_row) > 4 else '', 
             target_row.iloc[4] if len(target_row) > 4 else '', "percentage"),
            ("Production Plan Performance", week_row.iloc[3] if len(week_row) > 3 else '', 
             target_row.iloc[3] if len(target_row) > 3 else '', "percentage"),
            ("OEE", week_row.iloc[15] if len(week_row) > 15 else '', 
             target_row.iloc[15] if len(target_row) > 15 else '', "percentage"),
            ("Production Plan Compliance", week_row.iloc[5] if len(week_row) > 5 else '', 
             target_row.iloc[5] if len(target_row) > 5 else '', "percentage"),
        ]
        
        for i, (title, value, target, kpi_type) in enumerate(big_kpis):
            with big_cols[i]:
                st.markdown(create_kpi_card(title, value, target, kpi_type, size="large"), unsafe_allow_html=True)

        # Volume Chart Section
        display_volume_section()
        display_multi_kpi_section()

        # Simple Footer
        st.markdown("---")
        st.markdown("""
            <div style="text-align: center; padding: 20px 0; color: #666;">
                <p style="margin: 0; font-size: 14px;">© 2025 Commissary KPI Dashboard</p>
                <p style="margin: 5px 0 0 0; font-size: 12px; color: #888;">KPI Dashboard v1</p>
            </div>
        """, unsafe_allow_html=True)
        
    except Exception as e:
        st.error(f"Error displaying KPI dashboard: {str(e)}")
        st.exception(e)  # Show full exception details for debugging

# --- DATA LOADER FUNCTION ---
@st.cache_data(ttl=60)
def load_production_data(sheet_index=1):
    """Load production data from Google Sheets"""
    credentials = load_credentials_prod()
    if not credentials:
        return pd.DataFrame()
    
    try:
        # Use gspread instead of googleapiclient for consistency with working version
        import gspread
        gc = gspread.authorize(credentials)

        spreadsheet_id = "1PxdGZDltF2OWj5b6A3ncd7a1O4H-1ARjiZRBH0kcYrI"
        sh = gc.open_by_key(spreadsheet_id)
        worksheet = sh.get_worksheet(sheet_index)
        data = worksheet.get_all_values()

        df = pd.DataFrame(data)
        df = df.fillna('')
        
        return df
        
    except Exception as e:
        st.error(f"Error loading data: {str(e)}")
        return pd.DataFrame()

def update_week_dropdown(worksheet, selected_week):
    """Update the week dropdown selection in the spreadsheet"""
    try:
        # Adjust this cell reference to match where your week dropdown is located
        worksheet.update('H1', selected_week)  # Adjust cell reference
        
        # Clear cache to force reload of updated data
        load_production_data.clear()
        
        return True
    except Exception as e:
        st.error(f"Error updating spreadsheet: {str(e)}")
        return False

class ProductionDataExtractor:
    def __init__(self, df):
        self.df = df

    def get_unique_skus_by_station(self, station_filter="All Stations"):
        """Get list of unique SKU names filtered by station"""
        skus = self.get_all_skus(station_filter=station_filter)
        return sorted(list(set([sku['sku'] for sku in skus if sku['sku']])))        

    def get_week_info(self):
        """Extract week dates for title with improved error handling"""
        try:
            # Try multiple strategies to find date information
            dates = []
            
            # Strategy 1: Look in row 1 (index 1) for dates in batches columns
            for col in range(COLUMNS['batches_start'], COLUMNS['batches_end'] + 1):
                if col < len(self.df.columns):
                    date_val = str(self.df.iloc[1, col]).strip()
                    if date_val and date_val != '' and date_val != 'nan':
                        dates.append(date_val)
            
            # Strategy 2: If no dates found, try row 0
            if not dates:
                for col in range(COLUMNS['batches_start'], COLUMNS['batches_end'] + 1):
                    if col < len(self.df.columns):
                        date_val = str(self.df.iloc[0, col]).strip()
                        if date_val and date_val != '' and date_val != 'nan':
                            dates.append(date_val)
            
            # Strategy 3: If still no dates, look for any date-like patterns in first few rows
            if not dates:
                for row_idx in range(min(5, len(self.df))):
                    for col in range(COLUMNS['batches_start'], COLUMNS['batches_end'] + 1):
                        if col < len(self.df.columns):
                            cell_val = str(self.df.iloc[row_idx, col]).strip()
                            # Look for date patterns (containing /, -, or numbers)
                            if cell_val and any(char in cell_val for char in ['/', '-']) and any(char.isdigit() for char in cell_val):
                                dates.append(cell_val)
                                if len(dates) >= 7:  # Stop after finding enough dates
                                    break
                    if len(dates) >= 7:
                        break
            
            if len(dates) >= 2:
                start_date = dates[0]
                end_date = dates[-1]
                return start_date, end_date
            elif len(dates) == 1:
                # If only one date found, use it as both start and end
                return dates[0], dates[0]
            else:
                # Fallback: use current date
                current_date = datetime.now().strftime("%m/%d/%Y")
                return current_date, current_date
                
        except Exception as e:
            # Fallback to current date
            current_date = datetime.now().strftime("%m/%d/%Y")
            return current_date, current_date
    
    def get_days_of_week(self):
        """Extract days of week from row 4 (index 3)"""
        try:
            days = []
            # Try row 4 first (index 3) as shown in your screenshot
            for col in range(COLUMNS['batches_start'], COLUMNS['batches_end'] + 1):
                if col < len(self.df.columns):
                    day_val = str(self.df.iloc[2, col]).strip()  # Changed from index 2 to 3
                    if day_val and day_val != '':
                        days.append(day_val)
            
            # If no days found in row 4, try row 3 (index 2) as fallback
            if not days:
                for col in range(COLUMNS['batches_start'], COLUMNS['batches_end'] + 1):
                    if col < len(self.df.columns):
                        day_val = str(self.df.iloc[2, col]).strip()
                        if day_val and day_val != '':
                            days.append(day_val)
            
            return days
        except Exception as e:
            return []
        
    def get_week_number(self):
        """Extract week number from column H, row 1"""
        try:
            # Column H is index 7 (0-based), Row 1 is index 0 (0-based)
            week_number = str(self.df.iloc[0, 7]).strip()
            if week_number and week_number != '' and week_number != 'nan':
                return week_number
            else:
                return "Week N/A"
        except Exception as e:
            return "Week N/A"
    
    def get_all_skus(self, station_filter="All Stations", sku_filter="All SKUs", day_filter="All Days"):
        """Get all SKUs based on filters"""
        skus = []
        
        # Determine which ranges to process
        if station_filter == "All Stations":
            ranges_to_process = []
            for station_ranges in STATION_RANGES.values():
                ranges_to_process.extend(station_ranges)
        else:
            ranges_to_process = STATION_RANGES.get(station_filter, [])
        
        # Extract SKUs from ranges
        for start_row, end_row in ranges_to_process:
            for row_idx in range(start_row - 1, end_row):
                if row_idx < len(self.df):
                    sku_data = self.extract_sku_data(row_idx)
                    if sku_data and sku_data['sku']:
                        # Determine station
                        sku_data['station'] = self.determine_station(row_idx + 1)
                        skus.append(sku_data)
        
        # Apply SKU filter
        if sku_filter != "All SKUs":
            skus = [sku for sku in skus if sku['sku'] == sku_filter]
        
        # Apply day filter (filter based on which days have production)
        if day_filter != "Current Week":
            days = self.get_days_of_week()
            if day_filter in days:
                day_index = days.index(day_filter)
                filtered_skus = []
                for sku in skus:
                    if (day_index < len(sku['daily_batches']) and 
                        sku['daily_batches'][day_index] and 
                        sku['daily_batches'][day_index] != '' and 
                        sku['daily_batches'][day_index] != '0'):
                        filtered_skus.append(sku)
                skus = filtered_skus
        
        return skus
    
    def determine_station(self, row_number):
        """Determine which station a row belongs to"""
        for station, ranges in STATION_RANGES.items():
            for start_row, end_row in ranges:
                if start_row <= row_number <= end_row:
                    return station
        return "Unknown"
    
    def extract_sku_data(self, row_idx):
        """Extract all data for a SKU row"""
        try:
            row = self.df.iloc[row_idx]
            
            def safe_value(col_idx, default=''):
                try:
                    return str(row[col_idx]).strip() if col_idx < len(row) else default
                except:
                    return default
            
            sku_name = safe_value(COLUMNS['sku'])
            if not sku_name:
                return None
            
            sku_data = {
                'sku': sku_name,
                'batch_qty': safe_value(COLUMNS['batch_qty']),
                'daily_batches': [],
                'daily_volume': [],
                'daily_hours': [],
                'daily_manpower': [],
                'overtime': [],
                'overtime_percentage': [],           
            }
            
            # Extract daily data
            for col in range(COLUMNS['batches_start'], COLUMNS['batches_end'] + 1):
                sku_data['daily_batches'].append(safe_value(col))
            
            for col in range(COLUMNS['volume_start'], COLUMNS['volume_end'] + 1):
                sku_data['daily_volume'].append(safe_value(col))
            
            for col in range(COLUMNS['hours_start'], COLUMNS['hours_end'] + 1):
                sku_data['daily_hours'].append(safe_value(col))
            
            # Extract manpower from columns AP to AV (41-47)
            for col in range(COLUMNS['manpower_start'], COLUMNS['manpower_end'] + 1):
                sku_data['daily_manpower'].append(safe_value(col))

            # Extract overtime per person (AX-BD, columns 49-55)
            for col in range(COLUMNS['overtime_start'], COLUMNS['overtime_end'] + 1):
                sku_data['overtime'].append(safe_value(col))

            # Extract overtime percentage
            for col in range(COLUMNS['overtime_percentage_start'], COLUMNS['overtime_percentage_end'] + 1):
                sku_data['overtime_percentage'].append(safe_value(col))

            return sku_data
            
        except Exception as e:
            return None
        
    def get_overtime_percentage(self):
        """Get overtime percentage from row 6 (index 5), columns BP to BV (60 to 66)"""
        try:
            overtime_percentages = []
            row_idx = COLUMNS['header_row']  # Row 6 (index 5)
            
            # Use the CORRECT column range from your COLUMNS mapping
            for col in range(COLUMNS['overtime_percentage_start'], COLUMNS['overtime_percentage_end'] + 1):
                if row_idx < len(self.df) and col < len(self.df.columns):
                    value = str(self.df.iloc[row_idx, col]).strip()
                    overtime_percentages.append(value)
                else:
                    overtime_percentages.append('')
            
            return overtime_percentages
        except Exception as e:
            st.error(f"Error extracting overtime percentage: {str(e)}")
            return []
    
    def get_unique_skus(self):
        """Get list of unique SKU names"""
        skus = self.get_all_skus()
        return sorted(list(set([sku['sku'] for sku in skus if sku['sku']])))

def calculate_totals(skus, extractor=None, day_filter="Current Week", days=None):
    """Calculate KPI totals with day-specific filtering for manpower and overtime percentage"""
    def safe_sum_daily_values(values_list, attr):
        total = 0
        for sku in values_list:
            daily_values = sku.get(attr, [])
            for val in daily_values:
                total += safe_float_convert(val)
        return total
    
    def safe_sum_daily_values_for_day(values_list, attr, day_index):
        """Sum values for a specific day index"""
        total = 0
        for sku in values_list:
            daily_values = sku.get(attr, [])
            if day_index < len(daily_values):
                total += safe_float_convert(daily_values[day_index])
        return total
    
    # Standard calculations (unchanged)
    total_batches = safe_sum_daily_values(skus, 'daily_batches')
    total_volume = safe_sum_daily_values(skus, 'daily_volume')  
    total_hours = safe_sum_daily_values(skus, 'daily_hours')
    
    # FIXED: Manpower calculation based on day filter
    if day_filter == "Current Week":
        # Sum all manpower for the week
        total_manpower = safe_sum_daily_values(skus, 'daily_manpower')
    else:
        # Sum manpower for specific day only
        if days and day_filter in days:
            day_index = days.index(day_filter)
            total_manpower = safe_sum_daily_values_for_day(skus, 'daily_manpower', day_index)
        else:
            total_manpower = 0.0

    # FIXED: Overtime percentage calculation based on day filter
    overtime_percentage = 0.0
    if extractor:
        overtime_percentages = extractor.get_overtime_percentage()
        
        if day_filter == "Current Week":
            # Average of all valid percentages for the week (excluding empty/zero values)
            valid_percentages = []
            for p in overtime_percentages:
                # Clean the percentage value first
                clean_p = str(p).replace('%', '').strip()
                converted_value = safe_float_convert(clean_p)
                if converted_value > 0:
                    valid_percentages.append(converted_value)
            
            if valid_percentages:
                overtime_percentage = sum(valid_percentages) / len(valid_percentages)
        else:
            # Get percentage for specific day
            if days and day_filter in days:
                day_index = days.index(day_filter)
                if day_index < len(overtime_percentages):
                    # Clean the percentage value before conversion
                    raw_value = overtime_percentages[day_index]
                    clean_value = str(raw_value).replace('%', '').strip()
                    overtime_percentage = safe_float_convert(clean_value)
    
    return total_batches, total_volume, total_hours, total_manpower, overtime_percentage

def render_sku_table(skus, day_filter="Current Week", days=None):
    """Render the main SKU table with day-specific manpower calculation and color-coded station pills"""
    if not skus:
        st.warning("No SKUs match the current filters.")
        return
    
    st.markdown("<div style='margin:20px 0;'></div>", unsafe_allow_html=True)

    st.markdown("### Production List")
    
    # Station color mapping for pills
    station_colors = {
        'Hot Kitchen': "#f26556",
        'Cold Sauce': "#7dbfea", 
        'Fabrication': "#febc51",
        'Pastry': "#ba85cf",
        'Unknown': "#94abad"
    }
    
    # Prepare table data
    table_data = []
    for sku in skus:
        def safe_sum(values):
            total = 0
            for v in values:
                total += safe_float_convert(v)
            return total
        
        def safe_sum_for_day(values, day_index):
            if day_index < len(values):
                return safe_float_convert(values[day_index])
            return 0.0
        
        sku_overtime_per_person = safe_sum(sku.get('overtime', []))
        
        # Manpower calc depending on filter
        if day_filter == "Current Week":
            sku_manpower = safe_sum(sku.get('daily_manpower', []))
        else:
            if days and day_filter in days:
                day_index = days.index(day_filter)
                sku_manpower = safe_sum_for_day(sku.get('daily_manpower', []), day_index)
            else:
                sku_manpower = 0.0
        
        station = sku.get('station', 'Unknown')
        station_color = station_colors.get(station, station_colors['Unknown'])
        
        # --- Improved HTML pill with better escaping ---
        station_pill = f"""<span style="background-color: {station_color}; color: white; padding: 4px 12px; border-radius: 20px; font-size: 11px; font-weight: 600; display: inline-block; text-align: center; box-shadow: 0 2px 4px rgba(0,0,0,0.15); text-transform: uppercase; letter-spacing: 0.5px; white-space: nowrap;">{station}</span>"""
        
        table_data.append({
            'Station': station_pill,
            'SKU': sku['sku'],
            'Batches (no.)': f"{safe_sum(sku.get('daily_batches', [])):,.0f}",
            'Volume (kg)': f"{safe_sum(sku.get('daily_volume', [])):,.1f}",
            'Hours (hr)': f"{safe_sum(sku.get('daily_hours', [])):,.1f}",
            'Manpower (count)': f"{sku_manpower:,.0f}",
            'Overtime per Person (hrs)': f"{sku_overtime_per_person:,.0f}"
        })

    # Build DataFrame
    df_display = pd.DataFrame(table_data)
    
    # Add CSS styling first
    st.markdown("""
    <style>
    .scrollable-table-container {
        max-height: 600px;
        overflow-y: auto;
        overflow-x: auto;
        border-radius: 10px;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
        margin: 20px 0;
    }
    .station-table {
        width: 100%;
        border-collapse: collapse;
        font-size: 14px;
        background: white;
        font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'Roboto', sans-serif;
        margin: 0;
    }
    .station-table th {
        background: #1e2323;
        color: #f4d602;
        font-weight: bold;
        padding: 12px 8px;
        text-align: center;
        border-bottom: 2px solid #3b3f46;
        position: sticky;
        top: 0;
        z-index: 10;
    }
    .station-table td {
        padding: 12px 8px;
        border-bottom: 1px solid #e0e0e0;
        vertical-align: middle;
        text-align: center;
        font-weight: 500;
    }
    .station-table tr:hover {
        background-color: rgba(244, 214, 2, 0.1);
        transition: background-color 0.2s ease;
    }
    .station-table tr:last-child td {
        border-bottom: none;
    }
    /* Set equal widths for numeric columns (Batches to Overtime per Person) */
    .station-table th:nth-child(3),
    .station-table th:nth-child(4),
    .station-table th:nth-child(5),
    .station-table th:nth-child(6),
    .station-table th:nth-child(7),
    .station-table td:nth-child(3),
    .station-table td:nth-child(4),
    .station-table td:nth-child(5),
    .station-table td:nth-child(6),
    .station-table td:nth-child(7) {
        width: 10%;
        min-width: 100px;
    }
    /* Station column width */
    .station-table td:first-child,
    .station-table th:first-child {
        min-width: 120px;
        width: 18%;
    }
    /* SKU column width */
    .station-table td:nth-child(2),
    .station-table th:nth-child(2) {
        width: 30%;
        min-width: 150px;
    }
    </style>
    """, unsafe_allow_html=True)
    
    # Use the formatted_df as-is (categories are already the index)
    html_table = df_display.to_html(
        escape=False, 
        index=False, 
        classes='station-table',
        table_id='sku-table'
    )
    
    # Wrap table in scrollable container
    scrollable_html = f"""
    <div class="scrollable-table-container">
        {html_table}
    </div>
    """
    
    st.markdown(scrollable_html, unsafe_allow_html=True)

class YTDProductionExtractor:
    def __init__(self, df):
        self.df = df
        self.setup_station_mappings()
    
    def setup_station_mappings(self):
        """Define station row mappings based on your specification"""
        self.station_mappings = {
            "All Stations": {"row": 6, "name": "All Stations"},
            "Hot Kitchen Sauce": {"row": 7, "name": "Hot Kitchen Sauce"},
            "Hot Kitchen Savory": {"row": 37, "name": "Hot Kitchen Savory"}, 
            "Cold Sauce": {"row": 73, "name": "Cold Sauce"},
            "Fabrication Poultry": {"row": 114, "name": "Fabrication Poultry"},
            "Fabrication Meats": {"row": 129, "name": "Fabrication Meats"},
            "Pastry": {"row": 153, "name": "Pastry"}
        }
    
    def get_available_weeks(self):
        """Extract unique weeks from row 2 (week numbers)"""
        try:
            # Get week numbers from row 2, columns I onwards
            week_row = self.df.iloc[1, YTD_COLUMNS['data_start']:]
            
            seen_weeks = set()
            available_weeks = []
            current_week = None
            week_start_col = None
            
            for col_idx, week_num in enumerate(week_row):
                actual_col_idx = col_idx + YTD_COLUMNS['data_start']
                
                if pd.notna(week_num) and str(week_num).strip() != '':
                    try:
                        week_number = int(float(week_num))
                        
                        # If this is a new week number
                        if week_number != current_week:
                            # Save the previous week if it exists
                            if current_week is not None and current_week not in seen_weeks:
                                available_weeks.append({
                                    'week_number': current_week,
                                    'start_column': week_start_col,
                                    'end_column': actual_col_idx - 1
                                })
                                seen_weeks.add(current_week)
                            
                            # Start tracking the new week
                            current_week = week_number
                            week_start_col = actual_col_idx
                    except (ValueError, TypeError):
                        continue
                
                # Check if we've reached the end of data
                if actual_col_idx >= len(self.df.columns) - 1 or col_idx >= len(week_row) - 1:
                    # Save the last week if it hasn't been added yet
                    if current_week is not None and current_week not in seen_weeks:
                        available_weeks.append({
                            'week_number': current_week,
                            'start_column': week_start_col,
                            'end_column': actual_col_idx
                        })
                    break
            
            return sorted(available_weeks, key=lambda x: x['week_number'])
        except Exception as e:
            st.error(f"Error extracting weeks: {e}")
            return []
    
    def get_week_days(self, week_number):
        """Get the days for a specific week from row 3 (dates) and row 4 (day names)"""
        try:
            available_weeks = self.get_available_weeks()
            week_info = next((w for w in available_weeks if w['week_number'] == week_number), None)
            
            if not week_info:
                return []
            
            week_days = []
            for col_idx in range(week_info['start_column'], week_info['end_column'] + 1):
                if col_idx < len(self.df.columns):
                    # Get the date from row 3 (0-indexed as 2)
                    date_value = self.df.iloc[2, col_idx]
                    # Get the day name from row 4 (0-indexed as 3)
                    day_name_value = self.df.iloc[3, col_idx] if len(self.df) > 3 else None
                    
                    if pd.notna(date_value):
                        try:
                            # Try to parse the date
                            if isinstance(date_value, str):
                                # Handle different date formats
                                if '/' in date_value:
                                    date_obj = datetime.strptime(date_value, '%m/%d')
                                    date_obj = date_obj.replace(year=datetime.now().year)
                                elif '-' in date_value:
                                    date_obj = datetime.strptime(date_value, '%m-%d')
                                    date_obj = date_obj.replace(year=datetime.now().year)
                                else:
                                    date_obj = pd.to_datetime(date_value)
                            else:
                                date_obj = pd.to_datetime(date_value)
                            
                            # Use the day name from row 4 if available, otherwise calculate from date
                            if pd.notna(day_name_value) and str(day_name_value).strip() != '':
                                day_name = str(day_name_value).strip()
                            else:
                                day_name = date_obj.strftime('%A')
                            
                            week_days.append({
                                'column_index': col_idx,
                                'date': date_obj,
                                'day_name': day_name,
                                'formatted_date': date_obj.strftime('%b %d')
                            })
                        except:
                            # If date parsing fails, use raw values
                            day_name = f"Day {len(week_days) + 1}"
                            if pd.notna(day_name_value) and str(day_name_value).strip() != '':
                                day_name = str(day_name_value).strip()
                            
                            week_days.append({
                                'column_index': col_idx,
                                'date': date_value,
                                'day_name': day_name,
                                'formatted_date': str(date_value)
                            })
            
            return week_days
        except Exception as e:
            st.error(f"Error extracting week days: {e}")
            return []
    
    def get_station_skus(self, station_name):
        """Get all SKUs for a specific station from the station ranges"""
        try:
            STATION_RANGES = {
                'Hot Kitchen': [(8, 36), (38, 72)],
                'Cold Sauce': [(74, 113)],
                'Fabrication': [(115, 128), (130, 152)],
                'Pastry': [(154, 176)]
            }
            
            skus = []
            
            # Get the appropriate ranges for the station
            if station_name in STATION_RANGES:
                ranges = STATION_RANGES[station_name]
            else:
                # Handle individual stations
                station_to_range = {
                    "Hot Kitchen Sauce": [(8, 36)],
                    "Hot Kitchen Savory": [(38, 72)],
                    "Fabrication Poultry": [(115, 128)],
                    "Fabrication Meats": [(130, 152)]
                }
                ranges = station_to_range.get(station_name, [])
            
            # Extract SKUs from the ranges
            for start_row, end_row in ranges:
                start_idx = start_row - 1  # Convert to 0-based
                end_idx = end_row - 1
                
                for idx in range(start_idx, min(end_idx + 1, len(self.df))):
                    subrecipe = self.df.iloc[idx, YTD_COLUMNS['subrecipe']]
                    if pd.notna(subrecipe) and str(subrecipe).strip() != '':
                        skus.append(subrecipe)
            
            return skus
        
        except Exception as e:
            st.error(f"Error getting station SKUs: {e}")
            return []
    
    def get_production_totals(self):
        """Calculate total SKUs and total batches from the YTD data"""
        try:
            total_skus = 0
            total_batches = 0
           
            STATION_RANGES = {
                'Hot Kitchen': [(8, 36), (38, 72)],
                'Cold Sauce': [(74, 113)],
                'Fabrication': [(115, 128), (130, 152)],
                'Pastry': [(154, 176)]
            }
           
            # Count SKUs and batches from all station ranges
            for station_name, ranges in STATION_RANGES.items():
                for start_row, end_row in ranges:
                    start_idx = start_row - 1
                    end_idx = end_row - 1
                   
                    for idx in range(start_idx, min(end_idx + 1, len(self.df))):
                        subrecipe = self.df.iloc[idx, YTD_COLUMNS['subrecipe']]
                        if pd.notna(subrecipe) and str(subrecipe).strip() != '':
                            total_skus += 1
                           
                            # Sum production data from columns I to NI
                            start_col = YTD_COLUMNS['data_start']
                            end_col = min(YTD_COLUMNS['data_end'], len(self.df.columns))
                            
                            for col_idx in range(start_col, end_col + 1):
                                if col_idx < len(self.df.columns):
                                    value = self.df.iloc[idx, col_idx]
                                    try:
                                        batch_count = float(value) if value else 0
                                    except (ValueError, TypeError):
                                        batch_count = 0
                                    total_batches += batch_count
           
            return total_skus, total_batches
           
        except Exception as e:
            st.error(f"Error calculating production totals: {e}")
            return 0, 0


    
    def get_station_production_summary(self):
        """Create a production summary by station"""
        try:
            production_data = []
            
            # Define station ranges
            station_ranges = {
                "Hot Kitchen Sauce": (7, 37),
                "Hot Kitchen Savory": (37, 73),
                "Cold Sauce": (73, 114),
                "Fabrication Poultry": (114, 129),
                "Fabrication Meats": (129, 153),
                "Pastry": (153, None)
            }
            
            for station, (start_row, end_row) in station_ranges.items():
                station_skus = 0
                station_batches = 0
                
                # Convert to 0-based indexing
                start_idx = start_row - 1
                end_idx = (end_row - 1) if end_row else len(self.df)
                
                for idx in range(start_idx, min(end_idx, len(self.df))):
                    try:
                        subrecipe = self.df.iloc[idx, YTD_COLUMNS['subrecipe']]
                        
                        if pd.notna(subrecipe) and str(subrecipe).strip() != '':
                            station_skus += 1
                            
                            # Sum all production values for this SKU
                            production_values = self.df.iloc[idx, YTD_COLUMNS['data_start']:min(YTD_COLUMNS['data_end'], len(self.df.columns))]
                            sku_batches = sum(pd.to_numeric(val, errors='coerce') or 0 for val in production_values)
                            station_batches += sku_batches
                    except:
                        continue
                
                if station_skus > 0:
                    production_data.append({
                        'Station': station,
                        'Total SKUs': station_skus,
                        'Total Batches': station_batches,
                        'Avg Batches per SKU': station_batches / station_skus if station_skus > 0 else 0
                    })
            
            return pd.DataFrame(production_data)
        except Exception as e:
            st.error(f"Error creating production summary: {e}")
            return pd.DataFrame()
    
    def get_all_stations(self):
        """Get list of all available stations (no week_number parameter needed)"""
        return list(self.station_mappings.keys())
    
    def get_filtered_production_data(self, selected_week=None, selected_day=None,
                                   selected_station=None, selected_sku=None):
        """Get filtered production data based on user selections"""
        try:
            STATION_RANGES = {
                'Hot Kitchen': [(8, 36), (38, 72)],
                'Cold Sauce': [(74, 113)],
                'Fabrication': [(115, 128), (130, 152)],
                'Pastry': [(154, 176)]
            }
           
            production_data = []
           
            # Determine which station ranges to process
            stations_to_process = []
            if selected_station and selected_station != "All Stations":
                # Map station names to their ranges
                station_to_main = {
                    "Hot Kitchen Sauce": "Hot Kitchen",
                    "Hot Kitchen Savory": "Hot Kitchen", 
                    "Cold Sauce": "Cold Sauce",
                    "Fabrication Poultry": "Fabrication",
                    "Fabrication Meats": "Fabrication",
                    "Pastry": "Pastry"
                }
                if selected_station in station_to_main:
                    stations_to_process = [station_to_main[selected_station]]
            else:
                stations_to_process = list(STATION_RANGES.keys())
           
            # Process each station
            for station_name in stations_to_process:
                ranges = STATION_RANGES[station_name]
               
                for start_row, end_row in ranges:
                    start_idx = start_row - 1  # Convert to 0-based
                    end_idx = end_row - 1
                   
                    for idx in range(start_idx, min(end_idx + 1, len(self.df))):
                        subrecipe = self.df.iloc[idx, YTD_COLUMNS['subrecipe']]
                       
                        # Apply SKU filter
                        if selected_sku and selected_sku != "All SKUs" and subrecipe != selected_sku:
                            continue
                       
                        if pd.notna(subrecipe) and str(subrecipe).strip() != '':
                            total_batches = 0
                           
                            if selected_week:
                                # Get specific week data
                                week_days = self.get_week_days(selected_week)
                                
                                for day_info in week_days:
                                    # If a specific day is selected, only process that day
                                    if selected_day and selected_day != "All Days":
                                        # Compare the day info with the selected day
                                        day_display = f"{day_info['day_name']} ({day_info['formatted_date']})"
                                        if day_display != selected_day:
                                            continue
                                    
                                    col_idx = day_info['column_index']
                                    if col_idx < len(self.df.columns):
                                        value = self.df.iloc[idx, col_idx]
                                        try:
                                            batch_count = float(value) if value else 0
                                        except (ValueError, TypeError):
                                            batch_count = 0
                                        total_batches += batch_count
                            else:
                                # Get all data (columns I to NI)
                                start_col = YTD_COLUMNS['data_start']
                                end_col = min(YTD_COLUMNS['data_end'], len(self.df.columns))
                                for col_idx in range(start_col, end_col + 1):
                                    if col_idx < len(self.df.columns):
                                        value = self.df.iloc[idx, col_idx]
                                        try:
                                            batch_count = float(value) if value else 0
                                        except (ValueError, TypeError):
                                            batch_count = 0
                                        total_batches += batch_count
                           
                            # Only add if we have production data
                            if total_batches > 0:
                                production_data.append({
                                    'Station': station_name,
                                    'SKU': subrecipe,
                                    'Batches': total_batches
                                })
           
            return pd.DataFrame(production_data)
           
        except Exception as e:
            st.error(f"Error filtering production data: {e}")
            import traceback
            st.error(f"Traceback: {traceback.format_exc()}")
            return pd.DataFrame()

            
# --- MACHINE UTILIZATION EXTRACTOR ---
class MachineUtilizationExtractor:
    def __init__(self, df):
        self.df = df

    def calculate_totals(self, machines, day_index=None):
        """Calculate totals for all machines, optionally for a specific day"""
        total_machines = 0
        total_needed_hrs = 0
        total_remaining_hrs = 0  # This will now only include positive values
        total_machine_needed = 0
        capacity_utilization_values = []
        
        for machine in machines:
            # Sum quantities
            total_machines += machine.get('qty', 1)
            
            if day_index is None:
                # Weekly totals - ensure remaining hours are non-negative
                total_needed_hrs += safe_sum(machine.get('daily_needed_hrs', []))
                
                # FIXED: Only sum positive remaining hours
                remaining_hrs = machine.get('daily_remaining_hrs', [])
                total_remaining_hrs += sum(max(0, hr) if hr is not None else 0 for hr in remaining_hrs)
                
                total_machine_needed += safe_sum_positive_only(machine.get('daily_machine_needed', []))
                
                # For capacity utilization
                daily_capacities = machine.get('daily_capacity_utilization', [])
                valid_capacities = [val for val in daily_capacities if val is not None and val > 0]
                capacity_utilization_values.extend(valid_capacities)
                
            else:
                # Specific day totals - ensure remaining hours are non-negative
                total_needed_hrs += safe_sum_for_day(machine.get('daily_needed_hrs', []), day_index)
                
                # FIXED: Only use positive remaining hours for specific day
                remaining_hrs = machine.get('daily_remaining_hrs', [])
                if day_index < len(remaining_hrs):
                    day_remaining = remaining_hrs[day_index]
                    total_remaining_hrs += max(0, day_remaining) if day_remaining is not None else 0
                else:
                    total_remaining_hrs += 0
                
                total_machine_needed += safe_positive_for_day(machine.get('daily_machine_needed', []), day_index)
                
                # For specific day capacity utilization
                day_capacity = safe_value_for_day(machine.get('daily_capacity_utilization', []), day_index)
                if day_capacity > 0:
                    capacity_utilization_values.append(day_capacity)
        
        # Calculate average capacity utilization
        total_capacity_utilization = safe_average(capacity_utilization_values)
        
        return (total_machines, total_needed_hrs, total_remaining_hrs, 
                total_machine_needed, total_capacity_utilization)
    
    def get_machine_data(self):
        """Extract machine data with proper capacity utilization handling"""
        machines = []
        
        for row_idx in range(MACHINE_COLUMNS['machine_start_row'], MACHINE_COLUMNS['machine_end_row'] + 1):
            machine_name = self.df.iloc[row_idx, MACHINE_COLUMNS['machine']]
            
            if pd.isna(machine_name) or machine_name == "":
                continue
                
            # Get daily capacity utilization data (columns 34-40, which is AI-AO)
            daily_capacity_utilization = []
            for col_idx in range(MACHINE_COLUMNS['capacity_utilization_start'], 
                               MACHINE_COLUMNS['capacity_utilization_end'] + 1):
                try:
                    value = self.df.iloc[row_idx, col_idx]
                    if pd.isna(value) or value == "":
                        daily_capacity_utilization.append(0)
                    else:
                        # Handle different formats: percentage strings, decimals, or numbers
                        if isinstance(value, str):
                            # Remove % sign if present and convert
                            clean_value = value.replace('%', '').strip()
                            try:
                                num_value = float(clean_value)
                                daily_capacity_utilization.append(num_value)
                            except ValueError:
                                daily_capacity_utilization.append(0)
                        elif isinstance(value, (int, float)):
                            # If it's a decimal between 0-1, convert to percentage
                            if 0 <= value <= 1:
                                daily_capacity_utilization.append(value * 100)
                            else:
                                # Already a percentage or larger number
                                daily_capacity_utilization.append(float(value))
                        else:
                            daily_capacity_utilization.append(0)
                except (IndexError, ValueError, TypeError):
                    daily_capacity_utilization.append(0)
            
            # Get daily needed hours data
            daily_needed_hrs = []
            for col_idx in range(MACHINE_COLUMNS['needed_hrs_start'], MACHINE_COLUMNS['needed_hrs_end'] + 1):
                try:
                    value = self.df.iloc[row_idx, col_idx]
                    daily_needed_hrs.append(float(value) if not pd.isna(value) else 0)
                except (IndexError, ValueError, TypeError):
                    daily_needed_hrs.append(0)
            
            # Get daily remaining hours data
            daily_remaining_hrs = []
            for col_idx in range(MACHINE_COLUMNS['remaining_hrs_start'], MACHINE_COLUMNS['remaining_hrs_end'] + 1):
                try:
                    value = self.df.iloc[row_idx, col_idx]
                    daily_remaining_hrs.append(float(value) if not pd.isna(value) else 0)
                except (IndexError, ValueError, TypeError):
                    daily_remaining_hrs.append(0)
            
            # Get daily machine needed data
            daily_machine_needed = []
            for col_idx in range(MACHINE_COLUMNS['machine_needed_start'], MACHINE_COLUMNS['machine_needed_end'] + 1):
                try:
                    value = self.df.iloc[row_idx, col_idx]
                    daily_machine_needed.append(float(value) if not pd.isna(value) else 0)
                except (IndexError, ValueError, TypeError):
                    daily_machine_needed.append(0)
            
            # Get other machine properties
            try:
                rated_capacity = float(self.df.iloc[row_idx, MACHINE_COLUMNS['rated_capacity']]) if not pd.isna(self.df.iloc[row_idx, MACHINE_COLUMNS['rated_capacity']]) else 0
                qty = int(self.df.iloc[row_idx, MACHINE_COLUMNS['qty']]) if not pd.isna(self.df.iloc[row_idx, MACHINE_COLUMNS['qty']]) else 1
            except (ValueError, TypeError):
                rated_capacity = 0
                qty = 1
            
            machine_data = {
                'machine': str(machine_name),
                'rated_capacity': rated_capacity,
                'qty': qty,
                'daily_needed_hrs': daily_needed_hrs,
                'daily_remaining_hrs': daily_remaining_hrs,
                'daily_machine_needed': daily_machine_needed,
                'daily_capacity_utilization': daily_capacity_utilization
            }
            
            # Debug print to check if data is being extracted
            print(f"Machine: {machine_name}")
            print(f"Capacity Utilization Data: {daily_capacity_utilization}")
            
            machines.append(machine_data)
        
        return machines

def safe_sum(values):
    """Safely sum a list of values, handling None and empty lists"""
    if not values:
        return 0
    return sum(value or 0 for value in values)

def safe_sum_positive_only(values):
    """Safely sum only positive values, treating negative values as 0"""
    if not values:
        return 0
    return sum(max(0, value) if value is not None else 0 for value in values)

def safe_average(values):
    """Safely calculate average of a list of values, handling None and empty lists"""
    if not values:
        return 0
    # Filter out None/zero values for percentage calculations
    valid_values = [value for value in values if value is not None and value != 0]
    if not valid_values:
        return 0
    return sum(valid_values) / len(valid_values)

def safe_sum_for_day(values, day_index):
    """Safely get a value for a specific day, handling index errors"""
    try:
        return values[day_index] or 0
    except (IndexError, TypeError):
        return 0

def safe_value_for_day(values, day_index):
    """Safely get a value for a specific day without summing (for percentages)"""
    try:
        return values[day_index] or 0
    except (IndexError, TypeError):
        return 0

def safe_positive_for_day(values, day_index):
    """Safely get a positive value for a specific day, return 0 for negative values"""
    try:
        value = values[day_index]
        return max(0, value) if value is not None else 0  # Return 0 if negative or None
    except (IndexError, TypeError):
        return 0

def render_machine_table(machines, day_filter="Current Week", day_options=None):
    """Render the machine utilization table with non-negative remaining hours"""
    if not machines:
        st.warning("No machines match the current filters.")
        return
    
    st.markdown("### Machine Utilization Details")
    
    # Prepare table data
    table_data = []
    for machine in machines:
        if day_filter == "Current Week":
            needed_hrs = safe_sum(machine.get('daily_needed_hrs', []))
            
            # FIXED: Only sum positive remaining hours
            remaining_hrs_list = machine.get('daily_remaining_hrs', [])
            remaining_hrs = sum(max(0, hr) if hr is not None else 0 for hr in remaining_hrs_list)
            
            machine_needed = safe_sum_positive_only(machine.get('daily_machine_needed', []))
            capacity_utilization = safe_average(machine.get('daily_capacity_utilization', []))
        else:
            if day_options and day_filter in day_options:
                day_index = day_options.index(day_filter)
                needed_hrs = safe_sum_for_day(machine.get('daily_needed_hrs', []), day_index)
                
                # FIXED: Only use positive remaining hours for specific day
                remaining_hrs_list = machine.get('daily_remaining_hrs', [])
                if day_index < len(remaining_hrs_list):
                    day_remaining = remaining_hrs_list[day_index]
                    remaining_hrs = max(0, day_remaining) if day_remaining is not None else 0
                else:
                    remaining_hrs = 0
                
                machine_needed = safe_positive_for_day(machine.get('daily_machine_needed', []), day_index)
                capacity_utilization = safe_value_for_day(machine.get('daily_capacity_utilization', []), day_index)
            else:
                needed_hrs, remaining_hrs, machine_needed, capacity_utilization = 0, 0, 0, 0

        table_data.append({
            'Machine': machine["machine"],
            'Rated Capacity (kg/hr)': f"{machine.get('rated_capacity', 0):,.0f}",
            'Qty (no.)': f"{machine.get('qty', 1):,.0f}",
            'Needed Run Hours (hrs)': f"{needed_hrs:,.1f}",
            'Remaining Available Hours (hrs)': f"{max(0, remaining_hrs):,.1f}",
            'Additional Machines Needed (no.)': f"{machine_needed:,.0f}",
            'Capacity Utilization %': f"{capacity_utilization:,.1f}%",
        })
    
    # Display as DataFrame
    df_display = pd.DataFrame(table_data)

    # Add CSS styling
    st.markdown("""
    <style>
    .scrollable-machine-container {
        max-height: 600px;
        overflow-y: auto;
        overflow-x: auto;
        border-radius: 10px;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
        margin: 20px 0;
    }
    .machine-table {
        width: 100%;
        border-collapse: collapse;
        font-size: 14px;
        background: white;
        font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'Roboto', sans-serif;
        margin: 0;
    }
    .machine-table th {
        background: #1e2323;
        color: #f4d602;
        font-weight: bold;
        padding: 12px 8px;
        text-align: center;
        border-bottom: 2px solid #3b3f46;
        position: sticky;
        top: 0;
        z-index: 10;
    }
    .machine-table td {
        padding: 12px 8px;
        border-bottom: 1px solid #e0e0e0;
        vertical-align: middle;
        text-align: center;
        font-weight: 500;
    }
    .machine-table tr:hover {
        background-color: rgba(244, 214, 2, 0.1);
        transition: background-color 0.2s ease;
    }
    .machine-table tr:last-child td {
        border-bottom: none;
    }
    /* Set equal widths for numeric columns */
    .machine-table th:nth-child(2),
    .machine-table th:nth-child(3),
    .machine-table th:nth-child(4),
    .machine-table th:nth-child(5),
    .machine-table th:nth-child(6),
    .machine-table th:nth-child(7),
    .machine-table td:nth-child(2),
    .machine-table td:nth-child(3),
    .machine-table td:nth-child(4),
    .machine-table td:nth-child(5),
    .machine-table td:nth-child(6),
    .machine-table td:nth-child(7) {
        width: 14%;
        min-width: 100px;
    }
    /* Machine column width */
    .machine-table td:first-child,
    .machine-table th:first-child {
        width: 16%;
        min-width: 150px;
    }
    </style>
    """, unsafe_allow_html=True)
    
    # Render as HTML table
    html_table = df_display.to_html(
        escape=False, 
        index=False, 
        classes='machine-table',
        table_id='machine-table'
    )
    
    # Wrap table in scrollable container
    scrollable_html = f"""
    <div class="scrollable-machine-container">
        {html_table}
    </div>
    """
    
    st.markdown(scrollable_html, unsafe_allow_html=True)

def logo_to_base64(img):
    """Convert PIL image to base64 string"""
    buffer = BytesIO()
    img.save(buffer, format="PNG")
    return base64.b64encode(buffer.getvalue()).decode()


def create_navigation():
    """Create a modern, professional navigation header with logo"""
    
    try:
        # Load and convert logo to base64
        from PIL import Image
        logo_img = Image.open("cloudeats.png")
        logo_base64 = logo_to_base64(logo_img)
        
        # Modern navigation with logo
        st.markdown(f"""
        <style>
        .modern-nav-container {{
            background: linear-gradient(135deg, #000000 0%, #2f1d38 100%);
            padding: 0.75rem 2rem;
            margin: -1rem -1rem 2rem -1rem;
            position: sticky;
            top: 0;
            z-index: 1000;
            box-shadow: 0 4px 20px rgba(0,0,0,0.1);
            border-bottom: 3px solid rgba(255,255,255,0.1);
        }}
        
        .nav-brand {{
            display: flex;
            align-items: center;
            gap: 18px;
            margin-bottom: 0;
        }}
        
        .brand-logo {{
            width: 70px;   /* Increased size */
            height: auto;  /* Keep aspect ratio */
        }}
        
        .brand-logo img {{
            width: 100%;
            height: auto;
            object-fit: contain;
            display: block;
        }}
        
        .brand-text {{
            font-family: 'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif;
            font-size: 28px;
            font-weight: 700;
            color: #ffffff;
            letter-spacing: -0.5px;
            text-shadow: 0 2px 4px rgba(0,0,0,0.3);
            line-height: 1.2;
        }}
        </style>
        
        <div class="modern-nav-container">
            <div class="nav-brand">
                <div class="brand-logo">
                    <img src="data:image/png;base64,{logo_base64}" alt="Bites To Go Logo">
                </div>
                <div class="brand-text">Bites To Go - Commissary Dashboard</div>
            </div>
        </div>
        """, unsafe_allow_html=True)
        
    except FileNotFoundError:
        # Fallback if logo file is not found
        st.markdown("""
        <style>
        .modern-nav-container {
            background: linear-gradient(135deg, #000000 0%, #2f1d38 100%);
            padding: 0.75rem 2rem;
            margin: -1rem -1rem 2rem -1rem;
            position: sticky;
            top: 0;
            z-index: 1000;
            box-shadow: 0 4px 20px rgba(0,0,0,0.1);
            border-bottom: 3px solid rgba(255,255,255,0.1);
        }
        
        .nav-brand {
            display: flex;
            align-items: center;
            gap: 18px;
            margin-bottom: 0;
        }
        
        .brand-icon {
            width: 55px;
            height: 55px;
            background: linear-gradient(135deg, #ffd700, #ffed4a);
            border-radius: 8px;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 24px;
            font-weight: bold;
            color: #333;
            box-shadow: 0 4px 15px rgba(255,215,0,0.3);
        }
        
        .brand-text {
            font-family: 'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif;
            font-size: 28px;
            font-weight: 700;
            color: #ffffff;
            letter-spacing: -0.5px;
            text-shadow: 0 2px 4px rgba(0,0,0,0.3);
        }
        </style>
        
        <div class="modern-nav-container">
            <div class="nav-brand">
                <div class="brand-icon">🍽️</div>
                <div class="brand-text">Bites To Go</div>
            </div>
        </div>
        """, unsafe_allow_html=True)
        
        st.warning("Logo file 'cloudeats.png' not found. Using fallback icon.")

class SummaryDataExtractor:
    """Class to extract and process summary data from the Google Sheets"""
    
    def __init__(self, sheet_client, spreadsheet_id):
        self.client = sheet_client
        self.spreadsheet_id = spreadsheet_id
        self.sheet = None
        self._initialize_sheet()
    
    def _initialize_sheet(self):
        """Initialize the first sheet (index 0) of the spreadsheet"""
        try:
            spreadsheet = self.client.open_by_key(self.spreadsheet_id)
            self.sheet = spreadsheet.get_worksheet(0)  # Get first sheet (index 0)
        except Exception as e:
            st.error(f"Failed to access sheet: {e}")
    
    def update_week_dropdown(self, week_number):
        """Update the week dropdown in cell C1 of the sheet"""
        try:
            if self.sheet:
                self.sheet.update('C1', week_number)
                st.success(f"Updated sheet to week {week_number}")
                return True
        except Exception as e:
            st.error(f"Failed to update week in sheet: {e}")
            return False
        return False
    
    def extract_summary_data(self):
        """Extract summary data from the sheet after week update"""
        try:
            if not self.sheet:
                return None, None, None
            
            # Get all data from the sheet
            all_data = self.sheet.get_all_values()
            
            # Extract header row (B3:L3) - adjust for 0-indexing
            header_row = all_data[2][1:12]  # Row 3, columns B to L
            
            # Extract data rows (B4:L10) - rows 4-10, columns B to L
            data_rows = []
            for i in range(3, 10):  # Rows 4-10 (0-indexed: 3-9)
                if i < len(all_data):
                    row = all_data[i][1:12]  # Columns B to L
                    data_rows.append(row)
            
            # Create DataFrame
            df = pd.DataFrame(data_rows, columns=header_row)
            
            # Set category names as index
            categories = ['Batches', 'Volume', 'Total Run Mhrs', 'Total Manpower Required', 
                         'Total OT Manhrs', '%OT', 'Capacity Utilization']
            df.index = categories
            
            # Extract staff metrics from column N
            staff_metrics = {}
            if len(all_data) > 3 and len(all_data[3]) > 13:  # N4
                staff_metrics['Total Staff Count'] = all_data[3][13]
            if len(all_data) > 5 and len(all_data[5]) > 13:  # N6
                staff_metrics['Production Staff'] = all_data[5][13]
            if len(all_data) > 7 and len(all_data[7]) > 13:  # N8
                staff_metrics['Support Staff'] = all_data[7][13]
            
            # Get current week from C1
            current_week = self.sheet.acell('C1').value
            
            return df, staff_metrics, current_week
            
        except Exception as e:
            st.error(f"Failed to extract data: {e}")
            return None, None, None
    
    def get_date_headers(self, week_number):
        """Generate date headers for the selected week"""
        try:
            # Calculate start date of the week (assuming week 1 starts on a specific date)
            # You may need to adjust this based on your year start date
            year = datetime.now().year
            start_of_year = datetime(year, 1, 1)
            
            # Find the first Monday of the year
            days_ahead = 0 - start_of_year.weekday()  # Monday is 0
            if days_ahead < 0:
                days_ahead += 7
            first_monday = start_of_year + timedelta(days=days_ahead)
            
            # Calculate the start date of the selected week
            week_start = first_monday + timedelta(weeks=week_number - 1)
            
            # Generate 7 dates for the week
            dates = []
            for i in range(7):
                date = week_start + timedelta(days=i)
                dates.append(date.strftime('%d%b'))  # Format as 11Aug, 12Aug, etc.
            
            return dates
        except:
            # Fallback to generic labels if date calculation fails
            return ['Day1', 'Day2', 'Day3', 'Day4', 'Day5', 'Day6', 'Day7']
    
            
@st.cache_resource
def init_google_sheets():
    """Initialize Google Sheets connection"""
    try:
        credentials = load_credentials_prod()
        if credentials:
            client = gspread.authorize(credentials)
            return client
    except Exception as e:
        st.error(f"Failed to connect to Google Sheets: {e}")
        return None

def create_metric_cards(staff_metrics):
    """Create metric cards for staff information"""
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.metric(
            label="📊 Total Staff Count",
            value=staff_metrics.get('Total Staff Count', 'N/A'),
            delta=None
        )
    
    with col2:
        st.metric(
            label="👷 Production Staff",
            value=staff_metrics.get('Production Staff', 'N/A'),
            delta=None
        )
    
    with col3:
        st.metric(
            label="🔧 Support Staff",
            value=staff_metrics.get('Support Staff', 'N/A'),
            delta=None
        )

def format_dataframe(df):
    """Apply formatting to the DataFrame for better display"""
    # Create a copy to avoid modifying the original
    formatted_df = df.copy()
    
    # Apply number formatting for specific rows
    for idx in formatted_df.index:
        for col in formatted_df.columns:
            try:
                value = formatted_df.loc[idx, col]
                if value and str(value).replace('.', '').replace(',', '').isdigit():
                    # Format numbers with commas
                    formatted_df.loc[idx, col] = f"{float(value):,.0f}" if '.' not in str(value) else f"{float(value):,.2f}"
                elif '%' in str(value) or idx in ['%OT', 'Capacity Utilization']:
                    # Keep percentage formatting
                    pass
            except:
                continue
    
    return formatted_df



# --- Updated Summary Page - DataFrame Only ---
def summary_page():
    """Summary page showing weekly production data as DataFrame"""
    
    st.markdown("""
    <div class="main-header">
        <h1><b>Production Summary</b></h1>
        <p><b>Weekly Production Summary</b></p>
    </div>
    """, unsafe_allow_html=True)
    
    # Initialize Google Sheets connection
    sheets_client = init_google_sheets()
    
    if not sheets_client:
        st.error("❌ Unable to connect to Google Sheets. Please check your credentials.")
        return
    
    # Initialize data extractor (you'll need to provide your spreadsheet ID)
    SPREADSHEET_ID = "1PxdGZDltF2OWj5b6A3ncd7a1O4H-1ARjiZRBH0kcYrI"  # Replace with actual spreadsheet ID
    
    try:
        extractor = SummaryDataExtractor(sheets_client, SPREADSHEET_ID)
    except Exception as e:
        st.error(f"Failed to initialize data extractor: {e}")
        return
    
   # Center the controls container
    st.markdown("""
    <div style="display: flex; justify-content: center; margin: 20px 0;">
        <div style="width: 100%; max-width: 600px;">
    """, unsafe_allow_html=True)
    
    # Week selection and update button - perfectly aligned
    col1, col2 = st.columns([2, 1])
    
    with col1:
        selected_week = st.selectbox(
            "Select Week:",
            options=list(range(1, 54)),  # Weeks 1-53
            index=None,
            key="week_selector"
        )
    
    with col2:
        # Add custom CSS for perfect button alignment
        st.markdown("""
        <style>
        /* Reset any margins on the button container */
        .stButton {
            margin-top: 0 !important;
            padding-top: 0 !important;
        }
        
        .stButton > button {
            background-color: #6c757d !important;
            color: white !important;
            border: none !important;
            border-radius: 20px !important;
            padding: 0.625rem 1rem !important;
            font-weight: 500 !important;
            margin-top: -0.45rem !important;
            width: 100% !important;
            height: -0.45rem !important;
            transition: all 0.3s ease !important;
            display: flex !important;
            align-items: center !important;
            justify-content: center !important;
        }
        .stButton > button:hover {
            background-color: #5a6268 !important;
            transform: translateY(-1px) !important;
            box-shadow: 0 4px 8px rgba(0, 0, 0, 0.2) !important;
        }
        .stButton > button:focus {
            box-shadow: 0 0 0 3px rgba(108, 117, 125, 0.25) !important;
        }
        
        /* Ensure selectbox has consistent styling */
        .stSelectbox > div > div > div {
            height: 2.75rem !important;
        }
        </style>
        """, unsafe_allow_html=True)
        
        # Add some spacing to match the selectbox label
        st.markdown("<br>", unsafe_allow_html=True)
        
        if st.button("🔄 Update Data", key="update_button"):
            with st.spinner("Updating spreadsheet and fetching data..."):
                # Update the week in the spreadsheet
                success = extractor.update_week_dropdown(selected_week)
                if success:
                    # Small delay to allow spreadsheet to update
                    import time
                    time.sleep(2)
                    st.rerun()  # Refresh the page to show updated data
    
    # Close the centered container
    st.markdown("</div></div>", unsafe_allow_html=True)
        
    # Extract and display data
    with st.spinner("Loading data..."):
        df, staff_metrics, current_week = extractor.extract_summary_data()
    
    if df is not None and staff_metrics is not None:
        # Modern current week display
        st.markdown(f"""
            <div style="position: relative; z-index: 1;">
                <h3 style="
                    color: black;
                    margin: 0;
                    font-size: 1.5rem;
                    font-weight: 700;
                    text-shadow: 0 2px 4px rgba(0, 0, 0, 0.3);
                    display: flex;
                    align-items: center;
                    justify-content: center;
                    gap: 12px;
                ">
                    Week 
                    <span style="
                        color: white;
                        background: #334155;
                        padding: 5px 15px;
                        border-radius: 15px;
                        font-size: 1.8rem;
                        font-weight: 900;
                        margin-left: 5px;
                        border: 2px solid rgba(255, 255, 255, 0.3);
                        backdrop-filter: blur(5px);
                    ">{current_week}</span>
                </h3>
            </div>
            """, unsafe_allow_html=True)
                
        # Display main data table
        st.subheader("Weekly Production Data")
        
        # Format the DataFrame for better display
        formatted_df = format_dataframe(df)
        
        # Add CSS styling
        st.markdown("""
        <style>
        .scrollable-table-container {
            max-height: 600px;
            overflow-y: auto;
            overflow-x: auto;
            border-radius: 10px;
            box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
            margin: 20px 0;
        }
        .station-table {
            width: 100%;
            border-collapse: collapse;
            font-size: 14px;
            background: white;
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'Roboto', sans-serif;
            margin: 0;
        }
        .station-table th {
            background: #1e2323;
            color: #f4d602;
            font-weight: bold;
            padding: 12px 8px;
            text-align: center;
            border-bottom: 2px solid #3b3f46;
            position: sticky;
            top: 0;
            z-index: 10;
        }
        .station-table td {
            padding: 12px 8px;
            border-bottom: 1px solid #e0e0e0;
            vertical-align: middle;
            text-align: center;
            font-weight: 500;
        }
        .station-table tr:hover {
            background-color: rgba(244, 214, 2, 0.1);
            transition: background-color 0.2s ease;
        }
        .station-table tr:last-child td {
            border-bottom: none;
        }
        /* Set equal widths for numeric columns */
        .station-table th:nth-child(n+2),
        .station-table td:nth-child(n+2) {
            width: 12%;
            min-width: 100px;
        }
        /* First column (category names) */
        .station-table td:first-child,
        .station-table th:first-child {
            min-width: 180px;
            width: 22%;
            text-align: left;
            padding-left: 15px;
        }
        
        /* Center the controls container */
        .controls-container {
            display: flex;
            justify-content: center;
            align-items: flex-end;
            gap: 15px;
            margin: 20px 0;
            padding: 20px;
            background: #f8f9fa;
            border-radius: 12px;
            box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
        }
        
        /* Ensure selectbox and button are aligned */
        .stSelectbox > div > div {
            margin-bottom: 0 !important;
        }
        </style>
        """, unsafe_allow_html=True)
        
        # Reset index to make categories a column
        if formatted_df.index.name is not None or not any('category' in col.lower() for col in formatted_df.columns):
            display_df = formatted_df.reset_index()
            # Remove duplicate category columns if they exist
            category_cols = [col for col in display_df.columns if 'category' in col.lower()]
            if len(category_cols) > 1:
                # Keep only the first category column
                cols_to_drop = category_cols[1:]
                display_df = display_df.drop(columns=cols_to_drop)
        else:
            display_df = formatted_df.copy()
        
        # Render as HTML table with pills in scrollable container
        html_table = display_df.to_html(
            escape=False, 
            index=False, 
            classes='station-table',
            table_id='production-table'
        )
        
        # Wrap table in scrollable container
        scrollable_html = f"""
        <div class="scrollable-table-container">
            {html_table}
        </div>
        """
        
        st.markdown(scrollable_html, unsafe_allow_html=True)

        # Display staff metrics as cards
        st.subheader("👥 Staff Details")
        create_metric_cards(staff_metrics)
        
    else:
        st.error("❌ Failed to load data from the spreadsheet.")
        st.info("Please check your spreadsheet connection and data format.")
        
def weekly_prod_schedule():

    st.markdown("""
    <div class="main-header">
        <h1><b>Commissary Production Schedule</b></h1>
        <p><b>Weekly Production Schedule Management</b></p>
    </div>
    """, unsafe_allow_html=True)
    
    # Load data
    with st.spinner("Loading production schedule..."):
        df = load_production_data()
    
    if df is None:
        st.stop()
    
    # Initialize extractor
    extractor = ProductionDataExtractor(df)
    
    # Get week info, week number, and days
    start_date, end_date = extractor.get_week_info()
    week_number = extractor.get_week_number()
    days = extractor.get_days_of_week()

    # Filters with dynamic SKU dropdown and sheet selection
    col1, col2, col3, col4 = st.columns(4)

    with col1:
        # Week number display (read-only)
        st.selectbox(
            "Week Number",
            options=[week_number],
            index=0,
            key="week_display",
            disabled=True,
            help="Current week number from the selected spreadsheet"
        )

    with col2:
        station_filter = st.selectbox(
            "Filter per Station",
            options=list(STATIONS.keys()),
            index=0,
            key="station_filter"
        )

    with col3:
        # Get SKUs filtered by selected station
        if station_filter == "All Stations":
            unique_skus = extractor.get_unique_skus()
        else:
            unique_skus = extractor.get_unique_skus_by_station(station_filter)
        
        sku_options = ["All SKUs"] + unique_skus
        sku_filter = st.selectbox(
            "Filter per SKU",
            options=sku_options,
            index=0,
            key="sku_filter"
        )

    with col4:
        day_options = ["Current Week"] + days
        day_filter = st.selectbox(
            "Filter per Day",
            options=day_options,
            index=0,
            key="day_filter"
        )

    # Get filtered SKUs
    filtered_skus = extractor.get_all_skus(station_filter, sku_filter, day_filter)
    
    # FIXED: Calculate totals with day filter and days parameters
    total_batches, total_volume, total_hours, total_total_manpower, overtime_percentage = calculate_totals(
        filtered_skus, extractor, day_filter, days
    )
    
    st.markdown("<div style='margin:20px 0;'></div>", unsafe_allow_html=True)

    # 
    st.markdown("### Summary")
    
    col1, col2, col3, col4, col5, col6 = st.columns(6)
    
    with col1:
        st.markdown(f"""
        <div class="kpi-card kpi-card-wps">
            <div class="kpi-label">Total SKUs</div>
            <div class="kpi-number">{len(filtered_skus)}</div>
            <div class="kpi-unit">(no.)</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        st.markdown(f"""
        <div class="kpi-card kpi-card-wps">
            <div class="kpi-label">Total Batches</div>
            <div class="kpi-number">{total_batches:,.0f}</div>
            <div class="kpi-unit">(no.)</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col3:
        st.markdown(f"""
        <div class="kpi-card kpi-card-wps">
            <div class="kpi-label">Total Volume</div>
            <div class="kpi-number">{total_volume:,.0f}</div>
            <div class="kpi-unit">(kg)</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col4:
        st.markdown(f"""
        <div class="kpi-card kpi-card-wps">
            <div class="kpi-label">Total Hours</div>
            <div class="kpi-number">{total_hours:,.0f}</div>
            <div class="kpi-unit">(hrs)</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col5:
        st.markdown(f"""
        <div class="kpi-card kpi-card-wps">
            <div class="kpi-label">Total Manpower</div>
            <div class="kpi-number">{total_total_manpower:,.0f}</div>
            <div class="kpi-unit">(count)</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col6:
        st.markdown(f"""
        <div class="kpi-card kpi-card-wps">
            <div class="kpi-label">Overtime</div>
            <div class="kpi-number">{overtime_percentage:.1f}%</div>
            <div class="kpi-unit">%</div>
        </div>
        """, unsafe_allow_html=True)

    # FIXED: SKU Table with day filter and days parameters
    render_sku_table(filtered_skus, day_filter, days)

    # Simple Footer
    st.markdown("---")
    st.markdown("""
        <div style="text-align: center; padding: 20px 0; color: #666;">
            <p style="margin: 0; font-size: 14px;">© 2025 Weekly Production Schedule</p>
        </div>
    """, unsafe_allow_html=True)

def machine_utilization():

    # --- Header ---
    st.markdown("""
    <div class="main-header">
        <h1><b>Machine Utilization Dashboard</b></h1>
        <p><b>Track machine usage vs capacity</b></p>
    </div>
    """, unsafe_allow_html=True)

    # --- Load Data ---
    df_machine = load_production_data(sheet_index=2)  
    extractor = MachineUtilizationExtractor(df_machine)
    machines = extractor.get_machine_data()

    # --- Filters (day & machine) ---
    col1, col2 = st.columns(2)

    with col1:
        machine_options = ["All Machines"] + [m['machine'] for m in machines]
        machine_filter = st.selectbox("Filter per Machine", options=machine_options, index=0)

    with col2:
        # Row 2 = weekdays
        weekdays = extractor.df.iloc[2, MACHINE_COLUMNS['needed_hrs_start']:MACHINE_COLUMNS['needed_hrs_end']+1].tolist()
        # Row 3 = dates
        dates = extractor.df.iloc[3, MACHINE_COLUMNS['needed_hrs_start']:MACHINE_COLUMNS['needed_hrs_end']+1].tolist()
        # Combine weekday + date like "Mon (11 Aug)"
        day_labels = [f"{wd} ({dt})" for wd, dt in zip(weekdays, dates)]
        day_options = ["Current Week"] + day_labels
        day_filter = st.selectbox("Filter per Day", options=day_options, index=0)

    # --- Apply filters ---
    if machine_filter != "All Machines":
        machines = [m for m in machines if m['machine'] == machine_filter]

    # --- Totals for KPI cards ---
    if day_filter == "Current Week":
        totals = extractor.calculate_totals(machines)
    else:
        day_index = day_options.index(day_filter) - 1
        totals = extractor.calculate_totals(machines, day_index=day_index)
    
    # Now unpack 5 values instead of 4
    total_machines, total_needed_hrs, total_remaining_hrs, total_machine_needed, total_capacity_utilization = totals

    # Adjust "Total Machines" display
    if machine_filter != "All Machines" and machines:
        # If filtered by machine → show only that machine's qty
        total_machines = machines[0].get("qty", 1)

    # --- KPI Cards ---
    st.markdown("### Summary")
    
    colA, colB, colC, colD, colE = st.columns([0.9, 1, 1.2, 1.2, 1])
    
    # Use consistent naming with your totals
    total_machines, total_needed_hrs, total_remaining_hrs, total_machine_needed, total_capacity_utilization = totals
    
    with colA:
        st.markdown(f"""
        <div class="kpi-card kpi-card-mu">
            <div class="kpi-label">Total Machines</div>
            <div class="kpi-number">{total_machines:,.0f}</div>
            <div class="kpi-unit">(no.)</div>
        </div>
        """, unsafe_allow_html=True)
    
    with colB:
        st.markdown(f"""
        <div class="kpi-card kpi-card-mu">
            <div class="kpi-label">Needed Run Hours</div>
            <div class="kpi-number">{total_needed_hrs:,.0f}</div>
            <div class="kpi-unit">(hrs)</div>
        </div>
        """, unsafe_allow_html=True)
    
    with colC:
        st.markdown(f"""
        <div class="kpi-card kpi-card-mu">
            <div class="kpi-label">Remaining Available Hours</div>
            <div class="kpi-number">{total_remaining_hrs:,.0f}</div>
            <div class="kpi-unit">(hrs)</div>
        </div>
        """, unsafe_allow_html=True)
    
    with colD:
        st.markdown(f"""
        <div class="kpi-card kpi-card-mu">
            <div class="kpi-label">Additional Machines Needed</div>
            <div class="kpi-number">{total_machine_needed:,.0f}</div>
            <div class="kpi-unit">(no.)</div>
        </div>
        """, unsafe_allow_html=True)
    
    with colE:
        st.markdown(f"""
        <div class="kpi-card kpi-card-mu">
            <div class="kpi-label">Capacity Utilization</div>
            <div class="kpi-number">{total_capacity_utilization:,.0f}%</div>
            <div class="kpi-unit">(%)</div>
        </div>
        """, unsafe_allow_html=True)

    # --- Call the render function ---
    render_machine_table(machines, day_filter, day_options)

    # Simple Footer
    st.markdown("---")
    st.markdown("""
        <div style="text-align: center; padding: 20px 0; color: #666;">
            <p style="margin: 0; font-size: 14px;">© 2025 Machine Utilization</p>
        </div>
    """, unsafe_allow_html=True)

def ytd_production():
    """YTD Production Schedule Page"""
   
    # --- Header ---
    st.markdown("""
    <div class="main-header">
        <h1><b>YTD Production Schedule</b></h1>
        <p><b>Comprehensive Production Schedule – 2025</b></p>
    </div>
    """, unsafe_allow_html=True)
   
    # --- Load Data ---
    try:
        # Load YTD Production data from sheet index 6
        df_ytd = load_production_data(sheet_index=6)
        
        extractor = YTDProductionExtractor(df_ytd)
       
        # --- Single Row of Filters --  
        col1, col2, col3, col4 = st.columns(4)
       
        with col1:
            # Week Selection
            available_weeks = extractor.get_available_weeks()
            week_options = ["All Weeks"] + [f"Week {week['week_number']}" for week in available_weeks]
            selected_week_display = st.selectbox("Select Week", options=week_options, index=0)
           
            # Get actual week number
            if selected_week_display == "All Weeks":
                selected_week = None
            else:
                week_numbers = [week['week_number'] for week in available_weeks]
                selected_week = week_numbers[week_options.index(selected_week_display) - 1]
       
        with col2:
            # Day Selection
            selected_day_filter = None
            selected_day_display = "All Days"
            
            if selected_week:
                week_days = extractor.get_week_days(selected_week)
                day_options = ["All Days"] + [f"{day['day_name']} ({day['formatted_date']})" for day in week_days]
                selected_day_display = st.selectbox("Select Day", options=day_options, index=0)
                
                # Store the selected day for filtering
                if selected_day_display != "All Days":
                    selected_day_filter = selected_day_display
            else:
                selected_day_display = st.selectbox("Select Day", options=["All Days"], index=0, disabled=True)
       
        with col3:
            # Station Selection
            all_stations = extractor.get_all_stations()
            selected_station = st.selectbox("Select Station", options=all_stations, index=0)
       
        with col4:
            # SKU Selection
            if selected_station and selected_station != "All Stations":
                station_skus = extractor.get_station_skus(selected_station)
                sku_options = ["All SKUs"] + station_skus
                selected_sku = st.selectbox("Select SKU", options=sku_options, index=0)
            else:
                selected_sku = st.selectbox("Select SKU", options=["All SKUs"], index=0, disabled=True)
       
        # --- Get Production Data for KPIs ---
        # Get filtered production data for KPI calculations
        production_df = extractor.get_filtered_production_data(
            selected_week=selected_week,
            selected_day=selected_day_display if selected_day_display != "All Days" else None,
            selected_station=selected_station if selected_station != "All Stations" else None,
            selected_sku=selected_sku if selected_sku != "All SKUs" else None
        )
        
        # Calculate filtered totals for KPI cards
        if not production_df.empty:
            filtered_skus = production_df['SKU'].nunique()
            filtered_batches = production_df['Batches'].sum()
        else:
            filtered_skus = 0
            filtered_batches = 0
        
        # Get overall totals (without filters) for comparison
        total_skus, total_batches = extractor.get_production_totals()
       
        # --- KPI Cards ---
        st.markdown("### Production Summary")
        
        col_kpi1, col_kpi2 = st.columns(2)
        
        with col_kpi1:
            st.markdown(f"""
            <div class="kpi-card kpi-card-ytd">
                <div class="kpi-label">Total SKUs</div>
                <div class="kpi-number">{filtered_skus:,.0f}</div>
                <div class="kpi-unit">(no.)</div>
            </div>
            """, unsafe_allow_html=True)
        
        with col_kpi2:
            st.markdown(f"""
            <div class="kpi-card kpi-card-ytd">
                <div class="kpi-label">Total Batches</div>
                <div class="kpi-number">{filtered_batches:,.0f}</div>
                <div class="kpi-unit">(no.)</div>
            </div>
            """, unsafe_allow_html=True)
       
        # --- Production Data Table ---
        st.markdown("### Production Data")
       
        if not production_df.empty:
            # Format batches column
            production_df['Batches'] = production_df['Batches'].apply(lambda x: f"{x:,.0f}")
           
            # Display the dataframe
            st.dataframe(production_df, width='stretch', hide_index=True)
       
        else:
            st.warning("No production data matches the selected filters")
            # Show empty dataframe structure
            empty_df = pd.DataFrame(columns=['Station', 'SKU', 'Batches'])
            st.dataframe(empty_df, width='stretch', hide_index=True)

    except Exception as e:
        st.error(f"Error loading YTD Production data: {str(e)}")

    # Simple Footer
    st.markdown("---")
    st.markdown("""
        <div style="text-align: center; padding: 20px 0; color: #666;">
            <p style="margin: 0; font-size: 14px;">© 2025 YTD Production Schedule</p>
        </div>
    """, unsafe_allow_html=True)
        
def main():
    """Main application function - CORRECTED VERSION"""
    
    # Create modern navigation header
    create_navigation()
    
    # Initialize session state for navigation FIRST
    if 'main_tab' not in st.session_state:
        st.session_state.main_tab = "KPI Dashboard"
    if 'sub_tab' not in st.session_state:
        st.session_state.sub_tab = "Summary"  # Default to Summary page
    
    # Main navigation with smaller, centered buttons
    main_page_selection = option_menu(
        menu_title=None,
        options=["KPI Dashboard", "Production Details"],
        icons=["house-fill", "clipboard-data-fill"],
        default_index=0 if st.session_state.main_tab == "KPI Dashboard" else 1,
        orientation="horizontal",
        key="main_navigation",
        styles={
            "container": {
                "max-width": "350px",
                "text-align": "center",
                "border-radius": "20px", 
                "color": "#ffffff",
            },
            "icon": {
                "color": "#ffe712",
                "font-size": "12px",
                "margin-right": "4px"
            },
            "nav-link": {
                "font-family": "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
                "font-size": "11px",
                "font-weight": "500",
                "text-align": "center",
                "color": "#543559",
                "margin": "0.1rem",
                "padding": "0.4rem 0.6rem",
                "border-radius": "17px",
                "transition": "all 0.3s ease",
                "border": "1px solid transparent",
                "background": "rgba(248, 249, 250, 0.5)",
                "white-space": "nowrap"
            },
            "nav-link-selected": {
                "background": "linear-gradient(135deg, #495057 0%, #6c757d 100%)",
                "color": "#ffffff",
                "font-weight": "600",
                "box-shadow": "0 4px 15px rgba(73, 80, 87, 0.3)",
                "border": "2px solid rgba(255,255,255,0.1)",
                "transform": "translateY(-1px)"
            }
        }
    )
    
    # Store the main page selection in session state
    st.session_state.main_tab = main_page_selection
    
    # Show sub-navigation only if Production Details is selected
    if main_page_selection == "Production Details":
        
        # UPDATED: Added Summary to the sub-navigation options
        sub_page_selection = option_menu(
            menu_title=None,
            options=["Summary", "Weekly Production Schedule", "Machine Utilization", "YTD Production Schedule"],
            icons=["bar-chart-fill", "calendar-week-fill", "gear-fill", "graph-up"],
            default_index=["Summary", "Weekly Production Schedule", "Machine Utilization", "YTD Production Schedule"].index(st.session_state.sub_tab),
            orientation="horizontal",
            key="sub_navigation",
            styles={
                "container": {
                    "max-width": "800px",  # Increased width for 4 options
                    "text-align": "center",
                    "border-radius": "20px",
                    "color": "#ffffff"
                },
                "icon": {
                    "color": "#ffe712",
                    "font-size": "12px",
                    "margin-right": "4px"
                },
                "nav-link": {
                    "font-family": "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
                    "font-size": "11px",
                    "font-weight": "500",
                    "text-align": "center",
                    "color": "#543559",
                    "margin": "0.1rem",
                    "padding": "0.4rem 0.6rem",
                    "border-radius": "17px",
                    "transition": "all 0.3s ease",
                    "border": "1px solid transparent",
                    "background": "rgba(248, 249, 250, 0.5)",
                    "white-space": "nowrap"
                },
                "nav-link-selected": {
                    "background": "linear-gradient(135deg, #495057 0%, #6c757d 100%)",
                    "color": "#ffffff",
                    "font-weight": "600",
                    "box-shadow": "0 4px 15px rgba(73, 80, 87, 0.3)",
                    "border": "2px solid rgba(255,255,255,0.1)",
                    "transform": "translateY(-1px)"
                }
            }
        )

        # Store the sub page selection in session state
        st.session_state.sub_tab = sub_page_selection
    
    # Display the appropriate content based on navigation
    if st.session_state.main_tab == "KPI Dashboard":
        # FIXED: Call the existing main_page() function instead of display_kpi_dashboard()
        display_kpi_dashboard()
    else:
        # UPDATED: Added Summary page routing
        if st.session_state.sub_tab == "Summary":
            summary_page()
        elif st.session_state.sub_tab == "Weekly Production Schedule":
            weekly_prod_schedule()
        elif st.session_state.sub_tab == "Machine Utilization":
            machine_utilization()
        elif st.session_state.sub_tab == "YTD Production Schedule":
            ytd_production()      
            
if __name__ == "__main__":
    main()
