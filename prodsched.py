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
from google.oauth2 import service_account
from googleapiclient.discovery import build
warnings.filterwarnings('ignore')

# --- PAGE CONFIG ---
st.set_page_config(
    page_title="Commissary Production Scheduler",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# --- CREDENTIALS HANDLING ---
def load_credentials():
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
        background: #ff8765;
        color: #f0ebe4;
        padding: 1rem;
        border-radius: 25px;
        text-align: center;
        margin: 0rem;
        box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1);
        transition: transform 0.2s;
        display: flex;
        flex-direction: column;
        justify-content: center;
        align-items: center;
        min-height: 170px;
        position: relative;
    }}
    
    .kpi-card::before {{
        content: '';
        position: absolute;
        top: -2px;
        left: -2px;
        right: -2px;
        bottom: -2px;
        background: #272c2f;
        background-size: 400% 400%;
        border-radius: 27px;
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
        transform: scale(1.05) translateY(-8px) rotateY(5deg);
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

    .kpi-number {{
        font-family: {'TT Norms' if font_available else 'Arial'}, 'Arial', sans-serif;
        font-size: 2.2em;
        font-weight: 700;
        margin-bottom: 0.9rem;
        transition: all 0.3s ease;
    }}

    .kpi-label {{
        font-size: 0.8em;
        font-weight: 620;
        margin-bottom: 0.9rem;
        opacity: 1;
    }}

    .kpi-unit {{
        font-size: 0.9em;
        font-weight: 450;
        opacity: 1;
        margin-bottom: 0.1rem;
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
    
    # Render as HTML table with pills in scrollable container
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
            "Combined Hot Kitchen": {"rows": [7, 37], "name": "Combined Hot Kitchen"},  # Sum of sauce + savory
            "Cold Sauce": {"row": 73, "name": "Cold Sauce"},
            "Fab Poultry": {"row": 114, "name": "Fab Poultry"},
            "Fab Meats": {"row": 129, "name": "Fab Meats"},
            "Pastry": {"row": 153, "name": "Pastry"}
        }
    
    def get_available_weeks(self):
        """Extract available weeks from row 2 (week numbers)"""
        try:
            # Get week numbers from row 2, columns I onwards (column index 8+)
            week_row = self.df.iloc[1, 8:]  # Row 2 (0-indexed as 1), from column I
            
            # Find non-empty week numbers
            available_weeks = []
            for col_idx, week_num in enumerate(week_row):
                if pd.notna(week_num) and str(week_num).strip() != '':
                    try:
                        week_number = int(float(week_num))
                        available_weeks.append({
                            'week_number': week_number,
                            'column_index': col_idx + 8  # Adjust for actual column position
                        })
                    except (ValueError, TypeError):
                        continue
            
            return available_weeks
        except Exception as e:
            st.error(f"Error extracting weeks: {e}")
            return []
    
    def get_week_days(self, week_number):
        """Get the days for a specific week from row 3"""
        try:
            available_weeks = self.get_available_weeks()
            week_columns = []
            
            # Find all columns for the selected week
            for week_info in available_weeks:
                if week_info['week_number'] == week_number:
                    col_idx = week_info['column_index']
                    
                    # Get the date from row 3 (0-indexed as 2)
                    date_value = self.df.iloc[2, col_idx]
                    
                    if pd.notna(date_value):
                        # Try to parse the date
                        try:
                            if isinstance(date_value, str):
                                # Handle different date formats
                                if '/' in date_value:
                                    date_obj = datetime.strptime(date_value, '%m/%d')
                                    # Add current year
                                    date_obj = date_obj.replace(year=datetime.now().year)
                                elif '-' in date_value:
                                    date_obj = datetime.strptime(date_value, '%m-%d')
                                    date_obj = date_obj.replace(year=datetime.now().year)
                                else:
                                    # Try parsing as is
                                    date_obj = pd.to_datetime(date_value)
                            else:
                                date_obj = pd.to_datetime(date_value)
                            
                            week_columns.append({
                                'column_index': col_idx,
                                'date': date_obj,
                                'day_name': date_obj.strftime('%A'),
                                'formatted_date': date_obj.strftime('%b %d')
                            })
                        except:
                            # If date parsing fails, use raw value
                            week_columns.append({
                                'column_index': col_idx,
                                'date': date_value,
                                'day_name': f"Day {len(week_columns) + 1}",
                                'formatted_date': str(date_value)
                            })
            
            return sorted(week_columns, key=lambda x: x['column_index'])
        except Exception as e:
            st.error(f"Error extracting week days: {e}")
            return []
    
    def get_station_data(self, station_name, week_number):
        """Extract production data for a specific station and week"""
        try:
            if station_name not in self.station_mappings:
                return {}
            
            station_config = self.station_mappings[station_name]
            week_days = self.get_week_days(week_number)
            
            if not week_days:
                return {}
            
            # Get basic recipe data (columns B-H)
            station_data = {
                'station_name': station_name,
                'week_number': week_number,
                'days_data': {},
                'recipe_info': {}
            }
            
            # Handle combined stations (Hot Kitchen = Sauce + Savory)
            if station_name == "Combined Hot Kitchen":
                # Sum data from both rows 7 and 37
                for day_info in week_days:
                    col_idx = day_info['column_index']
                    
                    # Get values from both rows
                    sauce_value = self.df.iloc[6, col_idx] if col_idx < len(self.df.columns) else 0  # Row 7 (0-indexed as 6)
                    savory_value = self.df.iloc[36, col_idx] if col_idx < len(self.df.columns) else 0  # Row 37 (0-indexed as 36)
                    
                    # Convert to numeric, handling NaN
                    sauce_val = pd.to_numeric(sauce_value, errors='coerce') or 0
                    savory_val = pd.to_numeric(savory_value, errors='coerce') or 0
                    
                    combined_value = sauce_val + savory_val
                    
                    station_data['days_data'][day_info['formatted_date']] = {
                        'value': combined_value,
                        'day_name': day_info['day_name'],
                        'column_index': col_idx
                    }
            else:
                # Single station data
                station_row = station_config['row'] - 1  # Convert to 0-based index
                
                # Extract recipe information (columns B-H)
                try:
                    station_data['recipe_info'] = {
                        'subrecipe': self.df.iloc[station_row, 1] if len(self.df.columns) > 1 else '',
                        'batch_qty': self.df.iloc[station_row, 2] if len(self.df.columns) > 2 else 0,
                        'kg_per_mhr': self.df.iloc[station_row, 3] if len(self.df.columns) > 3 else 0,
                        'mhr_per_kg': self.df.iloc[station_row, 4] if len(self.df.columns) > 4 else 0,
                        'hrs_per_run': self.df.iloc[station_row, 5] if len(self.df.columns) > 5 else 0,
                        'working_hrs': self.df.iloc[station_row, 6] if len(self.df.columns) > 6 else 0,
                        'std_manpower': self.df.iloc[station_row, 7] if len(self.df.columns) > 7 else 0
                    }
                except:
                    station_data['recipe_info'] = {}
                
                # Extract daily production data
                for day_info in week_days:
                    col_idx = day_info['column_index']
                    
                    if col_idx < len(self.df.columns):
                        value = self.df.iloc[station_row, col_idx]
                        numeric_value = pd.to_numeric(value, errors='coerce') or 0
                        
                        station_data['days_data'][day_info['formatted_date']] = {
                            'value': numeric_value,
                            'day_name': day_info['day_name'],
                            'column_index': col_idx
                        }
            
            return station_data
        except Exception as e:
            st.error(f"Error extracting station data: {e}")
            return {}
    
    def get_all_stations_summary(self, week_number):
        """Get summary data for all stations for a specific week"""
        try:
            summary_data = {}
            
            for station_name in self.station_mappings.keys():
                if station_name != "All Stations":  # Skip the summary row for individual calculations
                    station_data = self.get_station_data(station_name, week_number)
                    if station_data and station_data.get('days_data'):
                        summary_data[station_name] = station_data
            
            return summary_data
        except Exception as e:
            st.error(f"Error creating stations summary: {e}")
            return {}

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
                <div class="brand-text">Bites To Go</div>
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


def main_page():
    """Main Page Content - Your existing dashboard"""
    # Main Header (your existing style)
    st.markdown(f"""
    <div class="main-header">
        <h1>Main Dashboard</h1>
    </div>
    """, unsafe_allow_html=True)

def weekly_prod_schedule():

    st.markdown("""
    <div class="main-header">
        <h1><b>Commissary Production Scheduler</b></h1>
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
        <div class="kpi-card">
            <div class="kpi-number">{len(filtered_skus)}</div>
            <div class="kpi-title">Total SKUs</div>
            <div class="kpi-unit">(no.)</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-number">{total_batches:,.0f}</div>
            <div class="kpi-title">Total Batches</div>
            <div class="kpi-unit">(no.)</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col3:
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-number">{total_volume:,.0f}</div>
            <div class="kpi-title">Total Volume</div>
            <div class="kpi-unit">(kg)</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col4:
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-number">{total_hours:,.0f}</div>
            <div class="kpi-title">Total Hours</div>
            <div class="kpi-unit">(hrs)</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col5:
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-number">{total_total_manpower:,.0f}</div>
            <div class="kpi-title">Total Manpower</div>
            <div class="kpi-unit">(count)</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col6:
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-number">{overtime_percentage:.1f}%</div>
            <div class="kpi-title">Overtime</div>
            <div class="kpi-unit">%</div>
        </div>
        """, unsafe_allow_html=True)

    # FIXED: SKU Table with day filter and days parameters
    render_sku_table(filtered_skus, day_filter, days)

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
        <div class="kpi-card">
            <div class="kpi-number">{total_machines:,.0f}</div>
            <div class="kpi-title">Total Machines</div>
            <div class="kpi-unit">(no.)</div>
        </div>
        """, unsafe_allow_html=True)
    
    with colB:
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-number">{total_needed_hrs:,.0f}</div>
            <div class="kpi-title">Needed Run Hours</div>
            <div class="kpi-unit">(hrs)</div>
        </div>
        """, unsafe_allow_html=True)
    
    with colC:
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-number">{total_remaining_hrs:,.0f}</div>
            <div class="kpi-title">Remaining Available Hours</div>
            <div class="kpi-unit">(hrs)</div>
        </div>
        """, unsafe_allow_html=True)
    
    with colD:
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-number">{total_machine_needed:,.0f}</div>
            <div class="kpi-title">Additional Machines Needed</div>
            <div class="kpi-unit">(no.)</div>
        </div>
        """, unsafe_allow_html=True)
    
    with colE:
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-number">{total_capacity_utilization:,.0f}%</div>
            <div class="kpi-title">Capacity Utilization</div>
            <div class="kpi-unit">(%)</div>
        </div>
        """, unsafe_allow_html=True)

    # --- Call the render function ---
    render_machine_table(machines, day_filter, day_options)


def ytd_production():
    """YTD Production Schedule Page"""
    
    # --- Header ---
    st.markdown("""
    <div class="main-header">
        <h1><b>📈 YTD Production Schedule</b></h1>
        <p><b>Year-to-date production metrics, trends, and comprehensive analytics</b></p>
    </div>
    """, unsafe_allow_html=True)
    
    # --- Load Data ---
    try:
        df_ytd = load_production_data(sheet_index=6)
        extractor = YTDProductionExtractor(df_ytd)
        
        # --- Week and Day Selection Filters ---
        st.markdown("### 📅 Time Period Selection")
        
        col_time1, col_time2 = st.columns(2)
        
        with col_time1:
            # Get available weeks
            available_weeks = extractor.get_available_weeks()
            if available_weeks:
                week_options = [f"Week {week['week_number']}" for week in available_weeks]
                week_numbers = [week['week_number'] for week in available_weeks]
                
                selected_week_display = st.selectbox(
                    "Select Week", 
                    options=week_options,
                    index=0,
                    help="Choose a week to view detailed daily production data"
                )
                
                # Get the actual week number
                selected_week = week_numbers[week_options.index(selected_week_display)]
            else:
                st.warning("No week data available")
                selected_week = 1
        
        with col_time2:
            # Get days for selected week
            if available_weeks:
                week_days = extractor.get_week_days(selected_week)
                if week_days:
                    day_options = ["All Days"] + [f"{day['day_name']} ({day['formatted_date']})" for day in week_days]
                    selected_day = st.selectbox(
                        "Select Day", 
                        options=day_options,
                        index=0,
                        help="Choose a specific day or view all days in the week"
                    )
                else:
                    st.warning("No day data available for selected week")
                    selected_day = "All Days"
            else:
                selected_day = "All Days"
        
        # --- Get Production Totals ---
        total_skus, total_batches = extractor.get_production_totals()
        
        # --- KPI Cards ---
        st.markdown("### 📊 Production Summary")
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown(f"""
            <div class="kpi-card">
                <div class="kpi-number">{total_skus:,.0f}</div>
                <div class="kpi-title">Total SKUs</div>
                <div class="kpi-unit">(subrecipes)</div>
            </div>
            """, unsafe_allow_html=True)
        
        with col2:
            st.markdown(f"""
            <div class="kpi-card">
                <div class="kpi-number">{total_batches:,.0f}</div>
                <div class="kpi-title">Total Batches</div>
                <div class="kpi-unit">(units)</div>
            </div>
            """, unsafe_allow_html=True)
        
        # --- Weekly Station Summary ---
        if available_weeks:
            st.markdown(f"### 🏭 Station Production Summary - Week {selected_week}")
            
            # Get all stations data for the selected week
            stations_summary = extractor.get_all_stations_summary(selected_week)
            
            if stations_summary:
                # Create summary table
                summary_data = []
                for station_name, station_data in stations_summary.items():
                    week_total = station_data.get('week_total', 0)
                    daily_avg = week_total / len(station_data['days_data']) if station_data['days_data'] else 0
                    
                    summary_data.append({
                        'Station': station_name,
                        'Week Total': f"{week_total:,.0f}",
                        'Daily Average': f"{daily_avg:.1f}",
                        'Days Active': len([d for d in station_data['days_data'].values() if d['value'] > 0])
                    })
                
                summary_df = pd.DataFrame(summary_data)
                st.dataframe(summary_df, use_container_width=True, hide_index=True)
                
                # --- Detailed Day View ---
                if selected_day != "All Days":
                    st.markdown(f"### 📋 Detailed Production - {selected_day}")
                    
                    # Extract day info from selection
                    day_info = None
                    for day in week_days:
                        if f"{day['day_name']} ({day['formatted_date']})" == selected_day:
                            day_info = day
                            break
                    
                    if day_info:
                        detailed_data = []
                        for station_name, station_data in stations_summary.items():
                            day_key = day_info['formatted_date']
                            if day_key in station_data['days_data']:
                                day_production = station_data['days_data'][day_key]['value']
                                
                                # Add recipe info if available
                                recipe_info = station_data.get('recipe_info', {})
                                
                                detailed_data.append({
                                    'Station': station_name,
                                    'Production': f"{day_production:,.0f}",
                                    'Batch Qty': recipe_info.get('batch_qty', 'N/A'),
                                    'Kg per MHr': recipe_info.get('kg_per_mhr', 'N/A'),
                                    'Working Hrs': recipe_info.get('working_hrs', 'N/A'),
                                    'Std Manpower': recipe_info.get('std_manpower', 'N/A')
                                })
                        
                        detailed_df = pd.DataFrame(detailed_data)
                        st.dataframe(detailed_df, use_container_width=True, hide_index=True)
        
        # --- Production List DataFrame ---
        st.markdown("### 📋 Production List by Station")
        
        production_df = extractor.get_production_list()
        
        if not production_df.empty:
            # Add filters for the production list
            col_filter1, col_filter2 = st.columns(2)
            
            with col_filter1:
                station_options = ["All Stations"] + sorted(production_df['Station'].unique().tolist())
                station_filter = st.selectbox("Filter by Station", options=station_options, index=0)
            
            with col_filter2:
                # Sort options
                sort_options = ["Station (A-Z)", "Station (Z-A)", "SKUs (High-Low)", "SKUs (Low-High)", "Batches (High-Low)", "Batches (Low-High)"]
                sort_filter = st.selectbox("Sort by", options=sort_options, index=0)
            
            # Apply station filter
            filtered_df = production_df.copy()
            if station_filter != "All Stations":
                filtered_df = filtered_df[filtered_df['Station'] == station_filter]
            
            # Apply sorting
            if sort_filter == "Station (A-Z)":
                filtered_df = filtered_df.sort_values('Station')
            elif sort_filter == "Station (Z-A)":
                filtered_df = filtered_df.sort_values('Station', ascending=False)
            elif sort_filter == "SKUs (High-Low)":
                filtered_df = filtered_df.sort_values('Total SKUs', ascending=False)
            elif sort_filter == "SKUs (Low-High)":
                filtered_df = filtered_df.sort_values('Total SKUs')
            elif sort_filter == "Batches (High-Low)":
                filtered_df = filtered_df.sort_values('Batches per SKU', ascending=False)
            elif sort_filter == "Batches (Low-High)":
                filtered_df = filtered_df.sort_values('Batches per SKU')
            
            # Display the filtered and sorted dataframe
            st.dataframe(filtered_df, use_container_width=True, hide_index=True)
            
            # Summary stats for filtered data
            col_stat1, col_stat2, col_stat3 = st.columns(3)
            with col_stat1:
                st.metric("Stations", len(filtered_df))
            with col_stat2:
                st.metric("Total SKUs", f"{filtered_df['Total SKUs'].sum():,.0f}")
            with col_stat3:
                st.metric("Total Batches", f"{filtered_df['Batches per SKU'].sum():,.0f}")
        else:
            st.warning("No production list data available")
        
        # --- Station Performance Chart ---
        if not production_df.empty:
            st.markdown("### 📊 Station Performance Visualization")
            
            chart_type = st.selectbox(
                "Chart Type", 
                options=["Bar Chart - SKUs", "Bar Chart - Batches", "Pie Chart - SKUs", "Pie Chart - Batches"],
                index=0
            )
            
            if chart_type == "Bar Chart - SKUs":
                st.bar_chart(production_df.set_index('Station')['Total SKUs'])
            elif chart_type == "Bar Chart - Batches":
                st.bar_chart(production_df.set_index('Station')['Batches per SKU'])
            elif chart_type == "Pie Chart - SKUs":
                fig_pie = px.pie(production_df, values='Total SKUs', names='Station', 
                               title="SKU Distribution by Station")
                st.plotly_chart(fig_pie, use_container_width=True)
            elif chart_type == "Pie Chart - Batches":
                fig_pie = px.pie(production_df, values='Batches per SKU', names='Station', 
                               title="Batch Distribution by Station")
                st.plotly_chart(fig_pie, use_container_width=True)
    
    except Exception as e:
        st.error(f"Error loading YTD Production data: {str(e)}")
        st.markdown("""
        <div class="dashboard-card">
            <h3>📈 YTD Production Analysis</h3>
            <p>Year-to-date production metrics, trends, and comprehensive analytics.</p>
            <p><em>Please ensure the production data file is available and properly formatted.</em></p>
            <ul>
                <li>Sheet index 6 should contain YTD Production Schedule</li>
                <li>Row 2: Week numbers (starting from column I)</li>
                <li>Row 3: Corresponding dates for each day</li>
                <li>Columns B-H: Recipe information (subrecipe, batch qty, etc.)</li>
                <li>Key station rows: 6 (All), 7 (Hot Kitchen Sauce), 37 (Hot Kitchen Savory), 73 (Cold Sauce), 114 (Fab Poultry), 129 (Fab Meats), 153 (Pastry)</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)
            
def main():
    """Main application function"""
    
    # Create modern navigation header
    create_navigation()
    
    # Initialize session state for navigation
    if 'main_tab' not in st.session_state:
        st.session_state.main_tab = "Main Page"
    if 'sub_tab' not in st.session_state:
        st.session_state.sub_tab = "Weekly Production Schedule"
    
    # Main navigation with smaller, centered buttons
    main_page_selection = option_menu(
        menu_title=None,
        options=["Main Page", "Dashboard Details"],
        icons=["house-fill", "clipboard-data-fill"],
        default_index=0 if st.session_state.main_tab == "Main Page" else 1,
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
    
    # Show sub-navigation only if Dashboard Details is selected
    if main_page_selection == "Dashboard Details":
        
        sub_page_selection = option_menu(
            menu_title=None,
            options=["Weekly Production Schedule", "Machine Utilization", "YTD Production Schedule"],
            icons=["calendar-week-fill", "gear-fill", "graph-up"],
            default_index=["Weekly Production Schedule", "Machine Utilization", "YTD Production Schedule"].index(st.session_state.sub_tab),
            orientation="horizontal",
            key="sub_navigation",
            styles={
                "container": {
                    "max-width": "600px",
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
    if st.session_state.main_tab == "Main Page":
        main_page()
    else:
        if st.session_state.sub_tab == "Weekly Production Schedule":
            weekly_prod_schedule()
        elif st.session_state.sub_tab == "Machine Utilization":
            machine_utilization()
        elif st.session_state.sub_tab == "YTD Production Schedule":
            ytd_production()      
            
if __name__ == "__main__":
    main()
