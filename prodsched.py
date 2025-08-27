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

# --- Convert Logo to Base64 ---
def logo_to_base64(img):
    buffer = BytesIO()
    img.save(buffer, format="PNG")
    return base64.b64encode(buffer.getvalue()).decode()

# --- Load and Encode Logo ---
try:
    logo = Image.open("cloudeats.png")
    logo_base64 = logo_to_base64(logo)
    
    # --- Sidebar Logo without Rounded Corners ---
    st.sidebar.markdown(
        f"""
        <img src="data:image/png;base64,{logo_base64}" 
             style="width: 320px; height: auto; border-radius: 0px; display: block; margin: 0 auto;" />
        """,
        unsafe_allow_html=True
    )
except FileNotFoundError:
    st.sidebar.markdown(
        """
        <div style="text-align: center; padding: 20px; background: #f0f0f0; border-radius: 10px; margin-bottom: 20px;">
            <h3 style="color: #2c3e50; margin: 0;">CloudEats</h3>
            <p style="color: #7f8c8d; margin: 5px 0 0 0; font-size: 14px;">Production Scheduler</p>
        </div>
        """,
        unsafe_allow_html=True
    )

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
        padding-top: 0rem;
        padding-bottom: 1rem;
        padding-left: 1rem;
        padding-right: 1rem;
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

# --- DATA LOADER FUNCTION ---
@st.cache_data(ttl=60)
def load_production_data(sheet_index=1):
    """Load production data from Google Sheets"""
    credentials = load_credentials()
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
    
def create_navigation():
    """Create the modern navigation bar"""
    nav_html = """
    <div class="nav-container">
        <div class="main-nav">
            <div class="nav-brand">ProductionPro</div>
            <div class="nav-menu-container">
                <!-- Navigation menu will be inserted here -->
            </div>
        </div>
    </div>
    """
    st.markdown(nav_html, unsafe_allow_html=True)

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
    st.markdown("""
    <div class="dashboard-card">
        <h3>📈 YTD Production Analysis</h3>
        <p>Year-to-date production metrics, trends, and comprehensive analytics.</p>
    </div>
    """, unsafe_allow_html=True)
            
def main():
    """Main application function"""
    
    # Create navigation bar
    create_navigation()
    
    # Navigation menu using option_menu (modern horizontal navigation)
    with st.container():
        # Center the navigation menu
        col1, col2, col3 = st.columns([1, 2, 1])
        
        with col2:
            page = option_menu(
                menu_title=None,
                options=["Main Page", "Weekly Production Schedule", "Machine Utilization", "YTD Production Schedule"],
                icons=["house-fill", "calendar-week-fill", "gear-fill", "graph-up"],
                default_index=0,
                orientation="horizontal",
                key="main_navigation",
                styles={
                    "container": {
                        "padding": "0rem",
                        "background-color": "transparent",
                        "border-radius": "12px",
                        "margin": "-1.5rem 0 2rem 0",
                        "box-shadow": "none"
                    },
                    "icon": {
                        "color": "#f4d602",
                        "font-size": "16px",
                        "margin-right": "8px"
                    },
                    "nav-link": {
                        f"font-family": f"{'TT Norms' if font_available else 'Segoe UI'}, sans-serif",
                        "font-size": "14px",
                        "font-weight": "500",
                        "text-align": "center",
                        "color": "#2c3e50",
                        "margin": "0 0.25rem",
                        "padding": "0.875rem 1.5rem",
                        "border-radius": "10px",
                        "transition": "all 0.3s ease",
                        "border": "1px solid rgba(44, 62, 80, 0.1)",
                        "background": "rgba(255, 255, 255, 0.8)",
                        "backdrop-filter": "blur(8px)"
                    },
                    "nav-link-selected": {
                        "background": "linear-gradient(135deg, #f4d602, #f7e842)",
                        "color": "#000000",
                        "font-weight": "600",
                        "box-shadow": "0 4px 12px rgba(244, 214, 2, 0.4)",
                        "border": "1px solid rgba(244, 214, 2, 0.6)",
                        "transform": "translateY(-2px)"
                    }
                }
            )
    
    # Page content based on selection
    if page == "Main Page":
        main_page()

    elif page == "Weekly Production Schedule":
        weekly_prod_schedule()
    
    elif page == "Machine Utilization":
        machine_utilization()
    
    elif page == "YTD Production Schedule":
        ytd_production()
   
            
            
if __name__ == "__main__":
    main()
