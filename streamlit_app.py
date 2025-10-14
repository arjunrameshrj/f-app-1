import streamlit as st
import pandas as pd
import plotly.express as px
import os
import hashlib
from datetime import datetime

# Streamlit page configuration
st.set_page_config(page_title="Warranty Conversion Dashboard", layout="wide", initial_sidebar_state="expanded")

# Custom CSS for professional styling with centered metrics table and bold colored headings
st.markdown("""
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Poppins:wght@400;500;600;700&family=Inter:wght@400;500&display=swap');

        body {
            font-family: 'Poppins', 'Inter', sans-serif;
        }
        .stApp {
            background: linear-gradient(135deg, #dbeafe 0%, #f9fafb 100%);
            padding: 20px;
        }
        .main-header {
            color: #3730a3;
            font-size: 2.8em;
            font-weight: 700;
            text-align: center;
            margin-bottom: 30px;
            text-shadow: 0 2px 4px rgba(0, 0, 0, 0.05);
        }
        .subheader {
            color: #3730a3;
            font-size: 1.6em;
            font-weight: 600;
            margin-top: 25px;
            margin-bottom: 15px;
        }
        .stButton>button {
            background: linear-gradient(90deg, #3730a3, #4338ca);
            color: white;
            border-radius: 10px;
            padding: 12px 24px;
            font-weight: 500;
            border: none;
            box-shadow: 0 2px 6px rgba(0, 0, 0, 0.1);
            transition: all 0.3s ease;
        }
        .stButton>button:hover {
            background: linear-gradient(90deg, #06b6d4, #22d3ee);
            box-shadow: 0 4px 8px rgba(0, 0, 0, 0.15);
            transform: translateY(-1px);
        }
        .stTextInput>div>input {
            border-radius: 8px;
            border: 1px solid #d1d5db;
            padding: 12px;
            background-color: #ffffff;
            transition: border-color 0.3s ease;
        }
        .stTextInput>div>input:focus {
            border-color: #06b6d4;
            box-shadow: 0 0 0 3px rgba(6, 182, 212, 0.1);
        }
        .stSelectbox>div>div>select {
            border-radius: 8px;
            border: 1px solid #d1d5db;
            padding: 12px;
            background-color: #ffffff;
            color: #1f2937;
            transition: border-color 0.3s ease;
        }
        .stSelectbox>div>div>select:focus {
            border-color: #06b6d4;
            box-shadow: 0 0 0 3px rgba(6, 182, 212, 0.1);
        }
        .stCheckbox label {
            color: #1f2937;
            font-weight: 500;
        }
        .sidebar .sidebar-content {
            background: #ffffff;
            border-right: 2px solid transparent;
            border-image: linear-gradient(to bottom, #3730a3, #06b6d4) 1;
            padding: 25px;
            box-shadow: 0 4px 12px rgba(0, 0, 0, 0.05);
        }
        .sidebar .stForm {
            background-color: #f3f4f6;
            padding: 20px;
            border-radius: 12px;
            margin-bottom: 20px;
            box-shadow: inset 0 1px 3px rgba(0, 0, 0, 0.05);
        }
        .stDataFrame {
            border-radius: 12px;
            overflow: hidden;
            background-color: #ffffff;
            box-shadow: 0 4px 12px rgba(0, 0, 0, 0.08);
            padding: 10px;
            text-align: center;
        }
        .stDataFrame table {
            margin-left: auto;
            margin-right: auto;
            width: 100%;
        }
        .stDataFrame th {
            background-color: #3730a3 !important;
            color: white !important;
            font-weight: 600 !important;
            text-align: center !important;
        }
        .stDataFrame td {
            text-align: center !important;
        }
        .stPlotlyChart {
            border-radius: 12px;
            box-shadow: 0 4px 12px rgba(0, 0, 0, 0.08);
            background-color: #ffffff;
            padding: 10px;
        }
        .stSuccess, .stWarning, .stInfo, .stError {
            border-radius: 10px;
            padding: 15px;
            margin-bottom: 15px;
            box-shadow: 0 2px 6px rgba(0, 0, 0, 0.05);
        }
        .stExpander {
            border-radius: 10px;
            background-color: #ffffff;
            box-shadow: 0 2px 6px rgba(0, 0, 0, 0.05);
            margin-bottom: 15px;
        }
        .stExpander>summary {
            background: linear-gradient(to right, #f3f4f6, #ffffff);
            padding: 12px;
            font-weight: 500;
            color: #3730a3;
        }
    </style>
""", unsafe_allow_html=True)

# Create data directory if it doesn't exist
DATA_DIR = "data"
if not os.path.exists(DATA_DIR):
    os.makedirs(DATA_DIR)

JULY_FILE_PATH = os.path.join(DATA_DIR, "july_data.csv")
AUGUST_FILE_PATH = os.path.join(DATA_DIR, "august_data.csv")
SEPTEMBER_FILE_PATH = os.path.join(DATA_DIR, "september_data.csv")
USER_CREDENTIALS_FILE = os.path.join(DATA_DIR, "user_credentials.txt")

# Function to hash passwords
def hash_password(password):
    return hashlib.sha256(password.encode()).hexdigest()

# Initialize user credentials (admin user for file upload access)
def initialize_credentials():
    if not os.path.exists(USER_CREDENTIALS_FILE):
        with open(USER_CREDENTIALS_FILE, 'w') as f:
            f.write(f"admin:{hash_password('admin123')}\n")

initialize_credentials()

# Load credentials
def load_credentials():
    credentials = {}
    if os.path.exists(USER_CREDENTIALS_FILE):
        with open(USER_CREDENTIALS_FILE, 'r') as f:
            for line in f:
                username, hashed_pwd = line.strip().split(':')
                credentials[username] = hashed_pwd
    return credentials

# Authentication function
def authenticate(username, password):
    credentials = load_credentials()
    return username in credentials and credentials[username] == hash_password(password)

# Function to map item categories to replacement warranty categories
def map_to_replacement_category(item_category):
    # Fan categories
    fan_categories = ['CEILING FAN', 'PEDESTAL FAN', 'RECHARGABLE FAN', 'TABLE FAN', 'TOWER FAN', 'WALL FAN']
    # Steamer categories
    steamer_categories = ['GARMENTS STEAMER', 'STEAMER']
    
    if any(fan in item_category.upper() for fan in fan_categories):
        return 'FAN'
    elif 'MIXER GRINDER' in item_category.upper():
        return 'MIXER GRINDER'
    elif 'IRON BOX' in item_category.upper():
        return 'IRON BOX'
    elif 'ELECTRIC KETTLE' in item_category.upper():
        return 'ELECTRIC KETTLE'
    elif 'OTG' in item_category.upper():
        return 'OTG'
    elif any(steamer in item_category.upper() for steamer in steamer_categories):
        return 'STEAMER'
    elif 'INDUCTION' in item_category.upper():
        return 'INDUCTION COOKER'
    else:
        return item_category  # Return original if not a replacement category

# Session state initialization
if 'authenticated' not in st.session_state:
    st.session_state.authenticated = False
if 'show_upload_form' not in st.session_state:
    st.session_state.show_upload_form = False
if 'user_role' not in st.session_state:
    st.session_state.user_role = "non-admin"

# Main dashboard
st.markdown('<h1 class="main-header">📊 Warranty Conversion Analysis Dashboard</h1>', unsafe_allow_html=True)

# Load data function
required_columns = ['Item Category', 'BDM', 'RBM', 'Store', 'Staff Name', 'TotalSoldPrice', 'WarrantyPrice', 'TotalCount', 'WarrantyCount']

@st.cache_data
def load_data(file, month, file_path=None):
    try:
        if isinstance(file, str):  # File path
            df = pd.read_csv(file)
        elif file.name.endswith('.csv'):  # Uploaded CSV
            df = pd.read_csv(file)
        else:  # Uploaded Excel
            df = pd.read_excel(file, engine='openpyxl')
        
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            st.error(f"Missing columns in {month} file: {', '.join(missing_columns)}")
            return None
        
        numeric_cols = ['TotalSoldPrice', 'WarrantyPrice', 'TotalCount', 'WarrantyCount']
        for col in numeric_cols:
            df[col] = pd.to_numeric(df[col], errors='coerce')
        
        if df[numeric_cols].isna().any().any():
            st.warning(f"Missing or invalid values in numeric columns for {month}. Filling with 0.")
            df[numeric_cols] = df[numeric_cols].fillna(0)
        
        # Add replacement category column
        df['Replacement Category'] = df['Item Category'].apply(map_to_replacement_category)
        
        df['Conversion% (Count)'] = (df['WarrantyCount'] / df['TotalCount'] * 100).round(2)
        df['Conversion% (Price)'] = (df['WarrantyPrice'] / df['TotalSoldPrice'] * 100).where(df['TotalSoldPrice'] > 0, 0).round(2)
        df['AHSP'] = (df['WarrantyPrice'] / df['WarrantyCount']).where(df['WarrantyCount'] > 0, 0).round(2)
        df['Month'] = month
        
        if file_path and not isinstance(file, str):
            if file.name.endswith('.csv'):
                with open(file_path, 'wb') as f:
                    f.write(file.getvalue())
            else:
                df.to_csv(file_path, index=False)
            st.success(f"{month} data saved successfully and available to all users.")
        
        return df
    except Exception as e:
        st.error(f"Error loading {month} file: {str(e)}")
        return None

# Load saved data or handle new uploads
df = None

# Sidebar content
with st.sidebar:
    st.markdown('<h2 style="color: #3730a3; font-weight: 600;">🔍 Dashboard Controls</h2>', unsafe_allow_html=True)
    st.markdown('<hr style="border: 1px solid #e5e7eb; margin: 10px 0;">', unsafe_allow_html=True)
    
    # File upload form
    st.markdown('<h3 style="color: #3730a3; font-weight: 500;">📁 Upload Data Files</h3>', unsafe_allow_html=True)
    with st.form("upload_form"):
        col1, col2, col3 = st.columns(3)
        with col1:
            july_file = st.file_uploader("July Data", type=["csv", "xlsx", "xls"], key="july")
        with col2:
            august_file = st.file_uploader("August Data", type=["csv", "xlsx", "xls"], key="august")
        with col3:
            september_file = st.file_uploader("September Data", type=["csv", "xlsx", "xls"], key="september")
        upload_password = st.text_input("Upload Password", type="password", placeholder="Enter password to upload")
        submit_upload = st.form_submit_button("Upload", type="primary")
        if submit_upload:
            if authenticate("admin", upload_password):
                july_df = load_data(july_file, "July", JULY_FILE_PATH) if july_file else None
                august_df = load_data(august_file, "August", AUGUST_FILE_PATH) if august_file else None
                september_df = load_data(september_file, "September", SEPTEMBER_FILE_PATH) if september_file else None
                df_list = [df for df in [july_df, august_df, september_df] if df is not None]
                if df_list:
                    df = pd.concat(df_list, ignore_index=True)
                else:
                    st.warning("Please select at least one file to upload.")
            else:
                st.error("Invalid password. Upload failed.")

    # Admin login
    st.markdown('<hr style="border: 1px solid #e5e7eb; margin: 20px 0;">', unsafe_allow_html=True)
    if st.session_state.user_role == "admin" and st.session_state.authenticated:
        st.button("Logout", on_click=lambda: st.session_state.update(authenticated=False, show_upload_form=False, user_role="non-admin"))
    else:
        st.button("Admin Login", on_click=lambda: st.session_state.update(show_upload_form=True))
        if st.session_state.show_upload_form and not st.session_state.authenticated:
            with st.form("login_form"):
                st.markdown('<h3 style="color: #3730a3; font-weight: 500;">🔐 Admin Login</h3>', unsafe_allow_html=True)
                username = st.text_input("Username", placeholder="Enter your username")
                password = st.text_input("Password", type="password", placeholder="Enter your password")
                submit_login = st.form_submit_button("Login", type="primary")
                if submit_login:
                    if authenticate(username, password):
                        st.session_state.authenticated = True
                        st.session_state.user_role = "admin"
                        st.success("Logged in successfully!")
                    else:
                        st.error("Invalid username or password.")

    # Sidebar filters
    st.markdown('<hr style="border: 1px solid #e5e7eb; margin: 20px 0;">', unsafe_allow_html=True)
    st.markdown('<h3 style="color: #3730a3; font-weight: 500;">⚙️ Filters</h3>', unsafe_allow_html=True)
    
    # Replacement warranty filter
    replacement_filter = st.checkbox("Show Replacement Warranty Categories Only")
    
    # Speaker category filter
    speaker_filter = st.checkbox("Show Speaker Categories Only")
    
    # Month filter
    month_options = ['All', 'July', 'August', 'September']
    selected_month = st.selectbox("Month", month_options, index=0)

# Load saved data if no new upload
if df is None:
    july_df = load_data(JULY_FILE_PATH, "July") if os.path.exists(JULY_FILE_PATH) else None
    august_df = load_data(AUGUST_FILE_PATH, "August") if os.path.exists(AUGUST_FILE_PATH) else None
    september_df = load_data(SEPTEMBER_FILE_PATH, "September") if os.path.exists(SEPTEMBER_FILE_PATH) else None
    df_list = [df for df in [july_df, august_df, september_df] if df is not None]
    if df_list:
        df = pd.concat(df_list, ignore_index=True)
        st.info("Loaded saved July, August, and/or September data.")

# If no uploaded or saved files, use sample data with speaker categories
if df is None:
    data = {
        'Item Category': ['CEILING FAN', 'PEDESTAL FAN', 'MIXER GRINDER', 'IRON BOX', 'ELECTRIC KETTLE', 'OTG', 'GARMENTS STEAMER', 'INDUCTION COOKER', 'SOUND BAR', 'PARTY SPEAKER', 'BLUETOOTH SPEAKER', 'HOME THEATRE'] * 3,
        'BDM': ['BDM1'] * 36,
        'RBM': ['MAHESH'] * 12 + ['RENJITH'] * 12 + ['JOHN'] * 12,
        'Store': ['Palakkad FUTURE', 'Store B'] * 6 + ['Kannur FUTURE', 'Store C'] * 6 + ['Kochi FUTURE', 'Store D'] * 6,
        'Staff Name': ['Staff1', 'Staff2'] * 18,
        'TotalSoldPrice': [48239177/12, 48239177/12, 48239177/12, 48239177/12, 48239177/12, 48239177/12, 48239177/12, 48239177/12, 48239177/12, 48239177/12, 48239177/12, 48239177/12] * 3,
        'WarrantyPrice': [300619/12, 300619/12, 300619/12, 300619/12, 300619/12, 300619/12, 300619/12, 300619/12, 300619/12, 300619/12, 300619/12, 300619/12] * 3,
        'TotalCount': [5286/12, 5286/12, 5286/12, 5286/12, 5286/12, 5286/12, 5286/12, 5286/12, 5286/12, 5286/12, 5286/12, 5286/12] * 3,
        'WarrantyCount': [483/12, 483/12, 483/12, 483/12, 483/12, 483/12, 483/12, 483/12, 483/12, 483/12, 483/12, 483/12] * 3,
        'Month': ['July'] * 12 + ['August'] * 12 + ['September'] * 12
    }
    df = pd.DataFrame(data)
    df['Replacement Category'] = df['Item Category'].apply(map_to_replacement_category)
    df['Conversion% (Count)'] = (df['WarrantyCount'] / df['TotalCount'] * 100).round(2)
    df['Conversion% (Price)'] = (df['WarrantyPrice'] / df['TotalSoldPrice'] * 100).round(2)
    df['AHSP'] = (df['WarrantyPrice'] / df['WarrantyCount']).where(df['WarrantyCount'] > 0, 0).round(2)
    st.warning("Using sample data for July, August, and September with speaker categories.")

# Ensure Month column is categorical
df['Month'] = pd.Categorical(df['Month'], categories=['July', 'August', 'September'], ordered=True)

# Apply replacement or speaker filter if selected
if replacement_filter:
    replacement_categories = ['FAN', 'MIXER GRINDER', 'IRON BOX', 'ELECTRIC KETTLE', 'OTG', 'STEAMER', 'INDUCTION COOKER']
    df = df[df['Replacement Category'].isin(replacement_categories)]
    category_column = 'Replacement Category'
elif speaker_filter:
    speaker_categories = ['SOUND BAR', 'PARTY SPEAKER', 'BLUETOOTH SPEAKER', 'HOME THEATRE']
    df = df[df['Item Category'].isin(speaker_categories)]
    category_column = 'Item Category'
else:
    category_column = 'Item Category'

# Define filters after df is loaded
with st.sidebar:
    bdm_options = ['All'] + sorted(df['BDM'].unique().tolist())
    rbm_options = ['All'] + sorted(df['RBM'].unique().tolist())
    store_options = ['All'] + sorted(df['Store'].unique().tolist())
    category_options = ['All'] + sorted(df[category_column].unique().tolist())

    selected_bdm = st.selectbox("BDM", bdm_options, index=0)
    selected_rbm = st.selectbox("RBM", rbm_options, index=0)
    selected_store = st.selectbox("Store", store_options, index=0)
    selected_category = st.selectbox(category_column, category_options, index=0)

    if selected_rbm != 'All':
        staff_options = ['All'] + sorted(df[df['RBM'] == selected_rbm]['Staff Name'].unique().tolist())
    else:
        staff_options = ['All'] + sorted(df['Staff Name'].unique().tolist())
    selected_staff = st.selectbox("Staff", staff_options, index=0)

    future_filter = st.checkbox("Show only FUTURE stores")
    decreased_rbm_filter = st.checkbox("Show only RBMs with decreased count conversion")

# Apply filters
filtered_df = df.copy()
if selected_month != 'All':
    filtered_df = filtered_df[filtered_df['Month'] == selected_month]
if selected_bdm != 'All':
    filtered_df = filtered_df[filtered_df['BDM'] == selected_bdm]
if selected_rbm != 'All':
    filtered_df = filtered_df[filtered_df['RBM'] == selected_rbm]
if selected_store != 'All':
    filtered_df = filtered_df[filtered_df['Store'] == selected_store]
if selected_category != 'All':
    filtered_df = filtered_df[filtered_df[category_column] == selected_category]
if selected_staff != 'All':
    filtered_df = filtered_df[filtered_df['Staff Name'] == selected_staff]
if future_filter:
    filtered_df = filtered_df[filtered_df['Store'].str.contains('FUTURE', case=True)]

if filtered_df.empty:
    st.warning("No data matches your filters. Adjust your selection.")
    st.stop()

# Main dashboard
st.markdown('<h2 class="subheader">📈 Performance Comparison: July vs August vs September</h2>', unsafe_allow_html=True)

# KPI comparison
st.markdown('<h3 class="subheader">Warranty Total Count and KPIs</h3>', unsafe_allow_html=True)
col1, col2, col3, col4, col5 = st.columns(5)

# Calculate metrics from the raw data (not averages)
july_data = filtered_df[filtered_df['Month'] == 'July']
august_data = filtered_df[filtered_df['Month'] == 'August']
september_data = filtered_df[filtered_df['Month'] == 'September']

# July metrics
july_total_warranty = july_data['WarrantyPrice'].sum()
july_total_units = july_data['TotalCount'].sum()
july_total_warranty_units = july_data['WarrantyCount'].sum()
july_total_sales = july_data['TotalSoldPrice'].sum()
july_value_conversion = (july_total_warranty / july_total_sales * 100) if july_total_sales > 0 else 0
july_count_conversion = (july_total_warranty_units / july_total_units * 100) if july_total_units > 0 else 0
july_ahsp = (july_total_warranty / july_total_warranty_units) if july_total_warranty_units > 0 else 0

# August metrics
august_total_warranty = august_data['WarrantyPrice'].sum()
august_total_units = august_data['TotalCount'].sum()
august_total_warranty_units = august_data['WarrantyCount'].sum()
august_total_sales = august_data['TotalSoldPrice'].sum()
august_value_conversion = (august_total_warranty / august_total_sales * 100) if august_total_sales > 0 else 0
august_count_conversion = (august_total_warranty_units / august_total_units * 100) if august_total_units > 0 else 0
august_ahsp = (august_total_warranty / august_total_warranty_units) if august_total_warranty_units > 0 else 0

# September metrics
september_total_warranty = september_data['WarrantyPrice'].sum()
september_total_units = september_data['TotalCount'].sum()
september_total_warranty_units = september_data['WarrantyCount'].sum()
september_total_sales = september_data['TotalSoldPrice'].sum()
september_value_conversion = (september_total_warranty / september_total_sales * 100) if september_total_sales > 0 else 0
september_count_conversion = (september_total_warranty_units / september_total_units * 100) if september_total_units > 0 else 0
september_ahsp = (september_total_warranty / september_total_warranty_units) if september_total_warranty_units > 0 else 0

with col1:
    st.metric("Warranty Sales (July)", f"₹{july_total_warranty:,.0f}")
    st.metric("Warranty Sales (August)", f"₹{august_total_warranty:,.0f}", 
              delta=f"{((august_total_warranty - july_total_warranty) / july_total_warranty * 100):.2f}%" if july_total_warranty > 0 else "N/A")
    st.metric("Warranty Sales (September)", f"₹{september_total_warranty:,.0f}", 
              delta=f"{((september_total_warranty - august_total_warranty) / august_total_warranty * 100):.2f}%" if august_total_warranty > 0 else "N/A")

with col2:
    st.metric("Count Conversion (July)", f"{july_count_conversion:.2f}%")
    st.metric("Count Conversion (August)", f"{august_count_conversion:.2f}%", 
              delta=f"{august_count_conversion - july_count_conversion:.2f}%")
    st.metric("Count Conversion (September)", f"{september_count_conversion:.2f}%", 
              delta=f"{september_count_conversion - august_count_conversion:.2f}%")

with col3:
    st.metric("Value Conversion (July)", f"{july_value_conversion:.2f}%")
    st.metric("Value Conversion (August)", f"{august_value_conversion:.2f}%", 
              delta=f"{august_value_conversion - july_value_conversion:.2f}%")
    st.metric("Value Conversion (September)", f"{september_value_conversion:.2f}%", 
              delta=f"{september_value_conversion - august_value_conversion:.2f}%")

with col4:
    st.metric("AHSP (July)", f"₹{july_ahsp:,.2f}")
    st.metric("AHSP (August)", f"₹{august_ahsp:,.2f}", 
              delta=f"₹{august_ahsp - july_ahsp:,.2f}" if july_ahsp > 0 else "N/A")
    st.metric("AHSP (September)", f"₹{september_ahsp:,.2f}", 
              delta=f"₹{september_ahsp - august_ahsp:,.2f}" if august_ahsp > 0 else "N/A")

with col5:
    st.metric("Warranty Unit Count (July)", f"{july_total_warranty_units:,.0f}")
    st.metric("Warranty Unit Count (August)", f"{august_total_warranty_units:,.0f}", 
              delta=f"{((august_total_warranty_units - july_total_warranty_units) / july_total_warranty_units * 100):.2f}%" if july_total_warranty_units > 0 else "N/A")
    st.metric("Warranty Unit Count (September)", f"{september_total_warranty_units:,.0f}", 
              delta=f"{((september_total_warranty_units - august_total_warranty_units) / august_total_warranty_units * 100):.2f}%" if august_total_warranty_units > 0 else "N/A")

# Store Performance Comparison
st.markdown('<h3 class="subheader">🏬 Store Performance Comparison</h3>', unsafe_allow_html=True)

# First calculate store-level aggregates
store_summary = filtered_df.groupby(['Store', 'Month']).agg({
    'TotalSoldPrice': 'sum',
    'WarrantyPrice': 'sum',
    'TotalCount': 'sum',
    'WarrantyCount': 'sum'
}).reset_index()

# Calculate metrics for each store-month combination
store_summary['Conversion% (Count)'] = (store_summary['WarrantyCount'] / store_summary['TotalCount'] * 100).round(2)
store_summary['Conversion% (Price)'] = (store_summary['WarrantyPrice'] / store_summary['TotalSoldPrice'] * 100).round(2)
store_summary['AHSP'] = (store_summary['WarrantyPrice'] / store_summary['WarrantyCount']).where(store_summary['WarrantyCount'] > 0, 0).round(2)

# Create pivot table for comparison
store_conv_pivot = store_summary.pivot_table(index='Store', columns='Month', values=['Conversion% (Count)', 'Conversion% (Price)', 'AHSP'], aggfunc='first').fillna(0)
store_conv_pivot.columns = [f"{col[1]} {col[0]}" for col in store_conv_pivot.columns]

# Ensure all required columns exist
for col in ['July Conversion% (Count)', 'August Conversion% (Count)', 'September Conversion% (Count)', 
            'July Conversion% (Price)', 'August Conversion% (Price)', 'September Conversion% (Price)', 
            'July AHSP', 'August AHSP', 'September AHSP']:
    if col not in store_conv_pivot.columns:
        store_conv_pivot[col] = 0

# Calculate changes
store_conv_pivot['Count Conv Change Jul-Aug (%)'] = (store_conv_pivot['August Conversion% (Count)'] - store_conv_pivot['July Conversion% (Count)']).round(2)
store_conv_pivot['Count Conv Change Aug-Sep (%)'] = (store_conv_pivot['September Conversion% (Count)'] - store_conv_pivot['August Conversion% (Count)']).round(2)
store_conv_pivot['Value Conv Change Jul-Aug (%)'] = (store_conv_pivot['August Conversion% (Price)'] - store_conv_pivot['July Conversion% (Price)']).round(2)
store_conv_pivot['Value Conv Change Aug-Sep (%)'] = (store_conv_pivot['September Conversion% (Price)'] - store_conv_pivot['August Conversion% (Price)']).round(2)
store_conv_pivot['AHSP Change Jul-Aug (%)'] = ((store_conv_pivot['August AHSP'] - store_conv_pivot['July AHSP']) / store_conv_pivot['July AHSP'] * 100).where(store_conv_pivot['July AHSP'] > 0, 0).round(2)
store_conv_pivot['AHSP Change Aug-Sep (%)'] = ((store_conv_pivot['September AHSP'] - store_conv_pivot['August AHSP']) / store_conv_pivot['August AHSP'] * 100).where(store_conv_pivot['August AHSP'] > 0, 0).round(2)

# Sort by September count conversion
store_conv_pivot = store_conv_pivot.sort_values('September Conversion% (Count)', ascending=False)

# Calculate TOTAL row using the same method as KPIs (sum of all, then calculate metrics)
total_july = store_summary[store_summary['Month'] == 'July'][['TotalSoldPrice', 'WarrantyPrice', 'TotalCount', 'WarrantyCount']].sum()
total_august = store_summary[store_summary['Month'] == 'August'][['TotalSoldPrice', 'WarrantyPrice', 'TotalCount', 'WarrantyCount']].sum()
total_september = store_summary[store_summary['Month'] == 'September'][['TotalSoldPrice', 'WarrantyPrice', 'TotalCount', 'WarrantyCount']].sum()

total_row = pd.DataFrame({
    'Store': ['Total'],
    'July Conversion% (Count)': [(total_july['WarrantyCount'] / total_july['TotalCount'] * 100).round(2) if total_july['TotalCount'] > 0 else 0],
    'August Conversion% (Count)': [(total_august['WarrantyCount'] / total_august['TotalCount'] * 100).round(2) if total_august['TotalCount'] > 0 else 0],
    'September Conversion% (Count)': [(total_september['WarrantyCount'] / total_september['TotalCount'] * 100).round(2) if total_september['TotalCount'] > 0 else 0],
    'July Conversion% (Price)': [(total_july['WarrantyPrice'] / total_july['TotalSoldPrice'] * 100).round(2) if total_july['TotalSoldPrice'] > 0 else 0],
    'August Conversion% (Price)': [(total_august['WarrantyPrice'] / total_august['TotalSoldPrice'] * 100).round(2) if total_august['TotalSoldPrice'] > 0 else 0],
    'September Conversion% (Price)': [(total_september['WarrantyPrice'] / total_september['TotalSoldPrice'] * 100).round(2) if total_september['TotalSoldPrice'] > 0 else 0],
    'July AHSP': [(total_july['WarrantyPrice'] / total_july['WarrantyCount']).round(2) if total_july['WarrantyCount'] > 0 else 0],
    'August AHSP': [(total_august['WarrantyPrice'] / total_august['WarrantyCount']).round(2) if total_august['WarrantyCount'] > 0 else 0],
    'September AHSP': [(total_september['WarrantyPrice'] / total_september['WarrantyCount']).round(2) if total_september['WarrantyCount'] > 0 else 0],
    'Count Conv Change Jul-Aug (%)': [
        (total_august['WarrantyCount'] / total_august['TotalCount'] * 100 - total_july['WarrantyCount'] / total_july['TotalCount'] * 100).round(2) 
        if total_august['TotalCount'] > 0 and total_july['TotalCount'] > 0 else 0
    ],
    'Count Conv Change Aug-Sep (%)': [
        (total_september['WarrantyCount'] / total_september['TotalCount'] * 100 - total_august['WarrantyCount'] / total_august['TotalCount'] * 100).round(2) 
        if total_september['TotalCount'] > 0 and total_august['TotalCount'] > 0 else 0
    ],
    'Value Conv Change Jul-Aug (%)': [
        (total_august['WarrantyPrice'] / total_august['TotalSoldPrice'] * 100 - total_july['WarrantyPrice'] / total_july['TotalSoldPrice'] * 100).round(2) 
        if total_august['TotalSoldPrice'] > 0 and total_july['TotalSoldPrice'] > 0 else 0
    ],
    'Value Conv Change Aug-Sep (%)': [
        (total_september['WarrantyPrice'] / total_september['TotalSoldPrice'] * 100 - total_august['WarrantyPrice'] / total_august['TotalSoldPrice'] * 100).round(2) 
        if total_september['TotalSoldPrice'] > 0 and total_august['TotalSoldPrice'] > 0 else 0
    ],
    'AHSP Change Jul-Aug (%)': [
        ((total_august['WarrantyPrice'] / total_august['WarrantyCount'] - total_july['WarrantyPrice'] / total_july['WarrantyCount']) / 
         (total_july['WarrantyPrice'] / total_july['WarrantyCount']) * 100).round(2) 
        if total_july['WarrantyCount'] > 0 and total_august['WarrantyCount'] > 0 else 0
    ],
    'AHSP Change Aug-Sep (%)': [
        ((total_september['WarrantyPrice'] / total_september['WarrantyCount'] - total_august['WarrantyPrice'] / total_august['WarrantyCount']) / 
         (total_august['WarrantyPrice'] / total_august['WarrantyCount']) * 100).round(2) 
        if total_august['WarrantyCount'] > 0 and total_september['WarrantyCount'] > 0 else 0
    ]
})

# Prepare display dataframe
store_display = store_conv_pivot.reset_index()
store_display = store_display[['Store', 'July Conversion% (Count)', 'August Conversion% (Count)', 'September Conversion% (Count)', 
                              'July Conversion% (Price)', 'August Conversion% (Price)', 'September Conversion% (Price)', 
                              'July AHSP', 'August AHSP', 'September AHSP', 
                              'Count Conv Change Jul-Aug (%)', 'Count Conv Change Aug-Sep (%)', 
                              'Value Conv Change Jul-Aug (%)', 'Value Conv Change Aug-Sep (%)', 
                              'AHSP Change Jul-Aug (%)', 'AHSP Change Aug-Sep (%)']]
store_display = pd.concat([store_display, total_row], ignore_index=True)
store_display['Total Change (%)'] = store_display[['Count Conv Change Jul-Aug (%)', 'Count Conv Change Aug-Sep (%)', 
                                                  'Value Conv Change Jul-Aug (%)', 'Value Conv Change Aug-Sep (%)', 
                                                  'AHSP Change Jul-Aug (%)', 'AHSP Change Aug-Sep (%)']].mean(axis=1).round(2)
store_display.columns = ['Store', 'July Count Conv (%)', 'August Count Conv (%)', 'September Count Conv (%)', 
                        'July Value Conv (%)', 'August Value Conv (%)', 'September Value Conv (%)', 
                        'July AHSP (₹)', 'August AHSP (₹)', 'September AHSP (₹)', 
                        'Count Conv Change Jul-Aug (%)', 'Count Conv Change Aug-Sep (%)', 
                        'Value Conv Change Jul-Aug (%)', 'Value Conv Change Aug-Sep (%)', 
                        'AHSP Change Jul-Aug (%)', 'AHSP Change Aug-Sep (%)', 'Total Change (%)']
store_display = store_display.fillna(0)

def highlight_change(val):
    color = 'green' if val > 0 else 'red' if val < 0 else 'black'
    return f'color: {color}'

st.dataframe(store_display.style.format({
    'July Count Conv (%)': '{:.2f}%',
    'August Count Conv (%)': '{:.2f}%',
    'September Count Conv (%)': '{:.2f}%',
    'July Value Conv (%)': '{:.2f}%',
    'August Value Conv (%)': '{:.2f}%',
    'September Value Conv (%)': '{:.2f}%',
    'July AHSP (₹)': '₹{:.2f}',
    'August AHSP (₹)': '₹{:.2f}',
    'September AHSP (₹)': '₹{:.2f}',
    'Count Conv Change Jul-Aug (%)': '{:.2f}%',
    'Count Conv Change Aug-Sep (%)': '{:.2f}%',
    'Value Conv Change Jul-Aug (%)': '{:.2f}%',
    'Value Conv Change Aug-Sep (%)': '{:.2f}%',
    'AHSP Change Jul-Aug (%)': '{:.2f}%',
    'AHSP Change Aug-Sep (%)': '{:.2f}%',
    'Total Change (%)': '{:.2f}%'
}).applymap(highlight_change, subset=['Count Conv Change Jul-Aug (%)', 'Count Conv Change Aug-Sep (%)', 
                                     'Value Conv Change Jul-Aug (%)', 'Value Conv Change Aug-Sep (%)', 
                                     'AHSP Change Jul-Aug (%)', 'AHSP Change Aug-Sep (%)', 'Total Change (%)'])
.set_properties(**{'font-weight': 'bold'}, subset=pd.IndexSlice[store_display.index[-1], :]), 
use_container_width=True)

# Visualization
fig_store = px.bar(store_summary, 
                   x='Store', 
                   y='Conversion% (Count)', 
                   color='Month', 
                   barmode='group', 
                   title='Store Count Conversion: July vs August vs September',
                   template='plotly_white',
                   color_discrete_sequence=['#3730a3', '#06b6d4', '#10b981'])
fig_store.update_layout(
    plot_bgcolor='rgba(0,0,0,0)',
    paper_bgcolor='rgba(0,0,0,0)',
    font=dict(family="Poppins, Inter, sans-serif", size=12, color="#1f2937"),
    showlegend=True,
    xaxis=dict(showgrid=False, tickangle=45),
    yaxis=dict(showgrid=True, gridcolor='#e5e7eb')
)
st.plotly_chart(fig_store, use_container_width=True)

# RBM Performance Comparison
st.markdown('<h3 class="subheader">👥 RBM Performance Comparison</h3>', unsafe_allow_html=True)

# First calculate RBM-level aggregates
rbm_summary = filtered_df.groupby(['RBM', 'Month']).agg({
    'TotalSoldPrice': 'sum',
    'WarrantyPrice': 'sum',
    'TotalCount': 'sum',
    'WarrantyCount': 'sum'
}).reset_index()

# Calculate metrics for each RBM-month combination
rbm_summary['Conversion% (Count)'] = (rbm_summary['WarrantyCount'] / rbm_summary['TotalCount'] * 100).round(2)
rbm_summary['Conversion% (Price)'] = (rbm_summary['WarrantyPrice'] / rbm_summary['TotalSoldPrice'] * 100).round(2)
rbm_summary['AHSP'] = (rbm_summary['WarrantyPrice'] / rbm_summary['WarrantyCount']).where(rbm_summary['WarrantyCount'] > 0, 0).round(2)

# Create pivot table for comparison
rbm_pivot = rbm_summary.pivot_table(index='RBM', columns='Month', values=['Conversion% (Count)', 'Conversion% (Price)', 'AHSP'], aggfunc='first').fillna(0)
rbm_pivot.columns = [f"{col[1]} {col[0]}" for col in rbm_pivot.columns]

# Ensure all required columns exist
for col in ['July Conversion% (Count)', 'August Conversion% (Count)', 'September Conversion% (Count)', 
            'July Conversion% (Price)', 'August Conversion% (Price)', 'September Conversion% (Price)', 
            'July AHSP', 'August AHSP', 'September AHSP']:
    if col not in rbm_pivot.columns:
        rbm_pivot[col] = 0

# Calculate changes
rbm_pivot['Count Conv Change Jul-Aug (%)'] = (rbm_pivot['August Conversion% (Count)'] - rbm_pivot['July Conversion% (Count)']).round(2)
rbm_pivot['Count Conv Change Aug-Sep (%)'] = (rbm_pivot['September Conversion% (Count)'] - rbm_pivot['August Conversion% (Count)']).round(2)
rbm_pivot['Value Conv Change Jul-Aug (%)'] = (rbm_pivot['August Conversion% (Price)'] - rbm_pivot['July Conversion% (Price)']).round(2)
rbm_pivot['Value Conv Change Aug-Sep (%)'] = (rbm_pivot['September Conversion% (Price)'] - rbm_pivot['August Conversion% (Price)']).round(2)
rbm_pivot['AHSP Change Jul-Aug (%)'] = ((rbm_pivot['August AHSP'] - rbm_pivot['July AHSP']) / rbm_pivot['July AHSP'] * 100).where(rbm_pivot['July AHSP'] > 0, 0).round(2)
rbm_pivot['AHSP Change Aug-Sep (%)'] = ((rbm_pivot['September AHSP'] - rbm_pivot['August AHSP']) / rbm_pivot['August AHSP'] * 100).where(rbm_pivot['August AHSP'] > 0, 0).round(2)
rbm_pivot = rbm_pivot.sort_values('September Conversion% (Count)', ascending=False)

if decreased_rbm_filter:
    rbm_pivot = rbm_pivot[(rbm_pivot['Count Conv Change Jul-Aug (%)'] < 0) | (rbm_pivot['Count Conv Change Aug-Sep (%)'] < 0)]
    if rbm_pivot.empty:
        st.warning("No RBMs with decreased count conversion match the filters.")

# Prepare display dataframe
rbm_display = rbm_pivot.reset_index()
rbm_display = rbm_display[['RBM', 'July Conversion% (Count)', 'August Conversion% (Count)', 'September Conversion% (Count)', 
                          'July Conversion% (Price)', 'August Conversion% (Price)', 'September Conversion% (Price)', 
                          'July AHSP', 'August AHSP', 'September AHSP', 
                          'Count Conv Change Jul-Aug (%)', 'Count Conv Change Aug-Sep (%)', 
                          'Value Conv Change Jul-Aug (%)', 'Value Conv Change Aug-Sep (%)', 
                          'AHSP Change Jul-Aug (%)', 'AHSP Change Aug-Sep (%)']]
rbm_display.columns = ['RBM', 'July Count Conv (%)', 'August Count Conv (%)', 'September Count Conv (%)', 
                       'July Value Conv (%)', 'August Value Conv (%)', 'September Value Conv (%)', 
                       'July AHSP (₹)', 'August AHSP (₹)', 'September AHSP (₹)', 
                       'Count Conv Change Jul-Aug (%)', 'Count Conv Change Aug-Sep (%)', 
                       'Value Conv Change Jul-Aug (%)', 'Value Conv Change Aug-Sep (%)', 
                       'AHSP Change Jul-Aug (%)', 'AHSP Change Aug-Sep (%)']
rbm_display = rbm_display.fillna(0)

st.dataframe(rbm_display.style.format({
    'July Count Conv (%)': '{:.2f}%',
    'August Count Conv (%)': '{:.2f}%',
    'September Count Conv (%)': '{:.2f}%',
    'July Value Conv (%)': '{:.2f}%',
    'August Value Conv (%)': '{:.2f}%',
    'September Value Conv (%)': '{:.2f}%',
    'July AHSP (₹)': '₹{:.2f}',
    'August AHSP (₹)': '₹{:.2f}',
    'September AHSP (₹)': '₹{:.2f}',
    'Count Conv Change Jul-Aug (%)': '{:.2f}%',
    'Count Conv Change Aug-Sep (%)': '{:.2f}%',
    'Value Conv Change Jul-Aug (%)': '{:.2f}%',
    'Value Conv Change Aug-Sep (%)': '{:.2f}%',
    'AHSP Change Jul-Aug (%)': '{:.2f}%',
    'AHSP Change Aug-Sep (%)': '{:.2f}%'
}).applymap(highlight_change, subset=['Count Conv Change Jul-Aug (%)', 'Count Conv Change Aug-Sep (%)', 
                                     'Value Conv Change Jul-Aug (%)', 'Value Conv Change Aug-Sep (%)', 
                                     'AHSP Change Jul-Aug (%)', 'AHSP Change Aug-Sep (%)']), 
use_container_width=True)

# Visualization
fig_rbm = px.bar(rbm_summary, 
                 x='RBM', 
                 y='Conversion% (Count)', 
                 color='Month', 
                 barmode='group', 
                 title='RBM Count Conversion: July vs August vs September',
                 template='plotly_white',
                 color_discrete_sequence=['#3730a3', '#06b6d4', '#10b981'])
fig_rbm.update_layout(
    plot_bgcolor='rgba(0,0,0,0)',
    paper_bgcolor='rgba(0,0,0,0)',
    font=dict(family="Poppins, Inter, sans-serif", size=12, color="#1f2937"),
    showlegend=True,
    xaxis=dict(showgrid=False, tickangle=45),
    yaxis=dict(showgrid=True, gridcolor='#e5e7eb')
)
st.plotly_chart(fig_rbm, use_container_width=True)

# Category Performance Comparison
st.markdown(f'<h3 class="subheader">📦 {category_column} Performance Comparison</h3>', unsafe_allow_html=True)

# First calculate category-level aggregates
category_summary = filtered_df.groupby([category_column, 'Month']).agg({
    'TotalSoldPrice': 'sum',
    'WarrantyPrice': 'sum',
    'TotalCount': 'sum',
    'WarrantyCount': 'sum'
}).reset_index()

# Calculate metrics for each category-month combination
category_summary['Conversion% (Count)'] = (category_summary['WarrantyCount'] / category_summary['TotalCount'] * 100).round(2)
category_summary['Conversion% (Price)'] = (category_summary['WarrantyPrice'] / category_summary['TotalSoldPrice'] * 100).round(2)
category_summary['AHSP'] = (category_summary['WarrantyPrice'] / category_summary['WarrantyCount']).where(category_summary['WarrantyCount'] > 0, 0).round(2)

if not category_summary.empty:
    # Create pivot table for comparison
    category_pivot = category_summary.pivot_table(index=category_column, columns='Month', values=['Conversion% (Count)', 'Conversion% (Price)', 'AHSP'], aggfunc='first').fillna(0)
    category_pivot.columns = [f"{col[1]} {col[0]}" for col in category_pivot.columns]

    # Ensure all required columns exist
    for col in ['July Conversion% (Count)', 'August Conversion% (Count)', 'September Conversion% (Count)', 
                'July Conversion% (Price)', 'August Conversion% (Price)', 'September Conversion% (Price)', 
                'July AHSP', 'August AHSP', 'September AHSP']:
        if col not in category_pivot.columns:
            category_pivot[col] = 0

    # Calculate changes
    category_pivot['Count Conv Change Jul-Aug (%)'] = (category_pivot['August Conversion% (Count)'] - category_pivot['July Conversion% (Count)']).round(2)
    category_pivot['Count Conv Change Aug-Sep (%)'] = (category_pivot['September Conversion% (Count)'] - category_pivot['August Conversion% (Count)']).round(2)
    category_pivot['Value Conv Change Jul-Aug (%)'] = (category_pivot['August Conversion% (Price)'] - category_pivot['July Conversion% (Price)']).round(2)
    category_pivot['Value Conv Change Aug-Sep (%)'] = (category_pivot['September Conversion% (Price)'] - category_pivot['August Conversion% (Price)']).round(2)
    category_pivot['AHSP Change Jul-Aug (%)'] = ((category_pivot['August AHSP'] - category_pivot['July AHSP']) / category_pivot['July AHSP'] * 100).where(category_pivot['July AHSP'] > 0, 0).round(2)
    category_pivot['AHSP Change Aug-Sep (%)'] = ((category_pivot['September AHSP'] - category_pivot['August AHSP']) / category_pivot['August AHSP'] * 100).where(category_pivot['August AHSP'] > 0, 0).round(2)
    category_pivot = category_pivot.sort_values('September Conversion% (Count)', ascending=False)

    # Prepare display dataframe
    category_display = category_pivot.reset_index()
    category_display = category_display[[category_column, 'July Conversion% (Count)', 'August Conversion% (Count)', 'September Conversion% (Count)', 
                                       'July Conversion% (Price)', 'August Conversion% (Price)', 'September Conversion% (Price)', 
                                       'July AHSP', 'August AHSP', 'September AHSP', 
                                       'Count Conv Change Jul-Aug (%)', 'Count Conv Change Aug-Sep (%)', 
                                       'Value Conv Change Jul-Aug (%)', 'Value Conv Change Aug-Sep (%)', 
                                       'AHSP Change Jul-Aug (%)', 'AHSP Change Aug-Sep (%)']]
    category_display.columns = [category_column, 'July Count Conv (%)', 'August Count Conv (%)', 'September Count Conv (%)', 
                               'July Value Conv (%)', 'August Value Conv (%)', 'September Value Conv (%)', 
                               'July AHSP (₹)', 'August AHSP (₹)', 'September AHSP (₹)', 
                               'Count Conv Change Jul-Aug (%)', 'Count Conv Change Aug-Sep (%)', 
                               'Value Conv Change Jul-Aug (%)', 'Value Conv Change Aug-Sep (%)', 
                               'AHSP Change Jul-Aug (%)', 'AHSP Change Aug-Sep (%)']
    category_display = category_display.fillna(0)

    st.dataframe(category_display.style.format({
        'July Count Conv (%)': '{:.2f}%',
        'August Count Conv (%)': '{:.2f}%',
        'September Count Conv (%)': '{:.2f}%',
        'July Value Conv (%)': '{:.2f}%',
        'August Value Conv (%)': '{:.2f}%',
        'September Value Conv (%)': '{:.2f}%',
        'July AHSP (₹)': '₹{:.2f}',
        'August AHSP (₹)': '₹{:.2f}',
        'September AHSP (₹)': '₹{:.2f}',
        'Count Conv Change Jul-Aug (%)': '{:.2f}%',
        'Count Conv Change Aug-Sep (%)': '{:.2f}%',
        'Value Conv Change Jul-Aug (%)': '{:.2f}%',
        'Value Conv Change Aug-Sep (%)': '{:.2f}%',
        'AHSP Change Jul-Aug (%)': '{:.2f}%',
        'AHSP Change Aug-Sep (%)': '{:.2f}%'
    }).applymap(highlight_change, subset=['Count Conv Change Jul-Aug (%)', 'Count Conv Change Aug-Sep (%)', 
                                         'Value Conv Change Jul-Aug (%)', 'Value Conv Change Aug-Sep (%)', 
                                         'AHSP Change Jul-Aug (%)', 'AHSP Change Aug-Sep (%)']), 
    use_container_width=True)

    # Visualization
    fig_category = px.bar(category_summary, 
                          x=category_column, 
                          y='Conversion% (Count)', 
                          color='Month', 
                          barmode='group', 
                          title=f'{category_column} Count Conversion: July vs August vs September',
                          template='plotly_white',
                          color_discrete_sequence=['#3730a3', '#06b6d4', '#10b981'])
    fig_category.update_layout(
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        font=dict(family="Poppins, Inter, sans-serif", size=12, color="#1f2937"),
        showlegend=True,
        xaxis=dict(showgrid=False, tickangle=45),
        yaxis=dict(showgrid=True, gridcolor='#e5e7eb')
    )
    st.plotly_chart(fig_category, use_container_width=True)
else:
    st.warning(f"No {category_column.lower()} data available with current filters.")
    category_pivot = pd.DataFrame()

# Insights
st.markdown('<h2 class="subheader">💡 Insights & Recommendations</h2>', unsafe_allow_html=True)

# Define targets
TARGET_COUNT_CONVERSION = 15.0
TARGET_VALUE_CONVERSION = 2.0
TARGET_AHSP = 1000.0  # Example target, adjust as needed

# Overall performance change and target comparison
with st.expander("Overall Performance Change & Target Comparison", expanded=True):
    if september_count_conversion >= TARGET_COUNT_CONVERSION:
        st.success(f"September Count Conversion ({september_count_conversion:.2f}%) meets or exceeds target ({TARGET_COUNT_CONVERSION}%).")
    else:
        st.warning(f"September Count Conversion ({september_count_conversion:.2f}%) is below target ({TARGET_COUNT_CONVERSION}%). Gap: {TARGET_COUNT_CONVERSION - september_count_conversion:.2f}%.")
    if september_count_conversion > august_count_conversion:
        st.success(f"Count Conversion improved from {august_count_conversion:.2f}% in August to {september_count_conversion:.2f}% in September.")
    elif september_count_conversion > july_count_conversion:
        st.info(f"Count Conversion improved from {july_count_conversion:.2f}% in July to {september_count_conversion:.2f}% in September, but monitor August's dip ({august_count_conversion:.2f}%).")
    else:
        st.warning(f"Count Conversion declined or remained flat from {july_count_conversion:.2f}% (July) and {august_count_conversion:.2f}% (August) to {september_count_conversion:.2f}% in September.")

    if september_value_conversion >= TARGET_VALUE_CONVERSION:
        st.success(f"September Value Conversion ({september_value_conversion:.2f}%) meets or exceeds target ({TARGET_VALUE_CONVERSION}%).")
    else:
        st.warning(f"September Value Conversion ({september_value_conversion:.2f}%) is below target ({TARGET_VALUE_CONVERSION}%). Gap: {TARGET_VALUE_CONVERSION - september_value_conversion:.2f}%.")
    if september_value_conversion > august_value_conversion:
        st.success(f"Value Conversion improved from {august_value_conversion:.2f}% in August to {september_value_conversion:.2f}% in September.")
    elif september_value_conversion > july_value_conversion:
        st.info(f"Value Conversion improved from {july_value_conversion:.2f}% in July to {september_value_conversion:.2f}% in September, but monitor August's dip ({august_value_conversion:.2f}%).")
    else:
        st.warning(f"Value Conversion declined or remained flat from {july_value_conversion:.2f}% (July) and {august_value_conversion:.2f}% (August) to {september_value_conversion:.2f}% in September.")

    if september_ahsp >= TARGET_AHSP:
        st.success(f"September AHSP (₹{september_ahsp:,.2f}) meets or exceeds target (₹{TARGET_AHSP:,.2f}).")
    else:
        st.warning(f"September AHSP (₹{september_ahsp:,.2f}) is below target (₹{TARGET_AHSP:,.2f}). Gap: ₹{TARGET_AHSP - september_ahsp:,.2f}.")
    if september_ahsp > august_ahsp:
        st.success(f"AHSP improved from ₹{august_ahsp:,.2f} in August to ₹{september_ahsp:,.2f} in September.")
    elif september_ahsp > july_ahsp:
        st.info(f"AHSP improved from ₹{july_ahsp:,.2f} in July to ₹{september_ahsp:,.2f} in September, but monitor August's dip (₹{august_ahsp:,.2f}).")
    else:
        st.warning(f"AHSP declined or remained flat from ₹{july_ahsp:,.2f} (July) and ₹{august_ahsp:,.2f} (August) to ₹{september_ahsp:,.2f} in September.")

# Store-Level Warranty Sales Analysis
with st.expander("Store-Level Warranty Sales Analysis", expanded=True):
    store_warranty = filtered_df.groupby(['Store', 'Month'])['WarrantyPrice'].sum().reset_index()
    store_pivot = store_warranty.pivot_table(index='Store', columns='Month', values='WarrantyPrice', aggfunc='first').fillna(0)
    store_pivot['Change Jul-Aug (₹)'] = store_pivot['August'] - store_pivot['July']
    store_pivot['Change Aug-Sep (₹)'] = store_pivot['September'] - store_pivot['August']
    store_pivot['Change Jul-Aug (%)'] = ((store_pivot['August'] - store_pivot['July']) / store_pivot['July'] * 100).where(store_pivot['July'] > 0, 0).round(2)
    store_pivot['Change Aug-Sep (%)'] = ((store_pivot['September'] - store_pivot['August']) / store_pivot['August'] * 100).where(store_pivot['August'] > 0, 0).round(2)
    store_pivot = store_pivot.sort_values('Change Aug-Sep (%)', ascending=False).reset_index()

    if not store_pivot.empty:
        st.write("Warranty sales comparison between July, August, and September for each store:")
        display_df = store_pivot[['Store', 'July', 'August', 'September', 'Change Jul-Aug (₹)', 'Change Aug-Sep (₹)', 'Change Jul-Aug (%)', 'Change Aug-Sep (%)']]
        display_df.columns = ['Store', 'July Warranty Sales (₹)', 'August Warranty Sales (₹)', 'September Warranty Sales (₹)', 'Change Jul-Aug (₹)', 'Change Aug-Sep (₹)', 'Change Jul-Aug (%)', 'Change Aug-Sep (%)']
        st.dataframe(display_df.style.format({
            'July Warranty Sales (₹)': '₹{:.0f}',
            'August Warranty Sales (₹)': '₹{:.0f}',
            'September Warranty Sales (₹)': '₹{:.0f}',
            'Change Jul-Aug (₹)': '₹{:.0f}',
            'Change Aug-Sep (₹)': '₹{:.0f}',
            'Change Jul-Aug (%)': '{:.2f}%',
            'Change Aug-Sep (%)': '{:.2f}%'
        }).applymap(highlight_change, subset=['Change Jul-Aug (₹)', 'Change Aug-Sep (₹)', 'Change Jul-Aug (%)', 'Change Aug-Sep (%)']), use_container_width=True)

        st.write("**Reasons for Warranty Sales Decreases:**")
        decreased_stores = store_pivot[(store_pivot['Change Jul-Aug (%)'] < 0) | (store_pivot['Change Aug-Sep (%)'] < 0)]
        if not decreased_stores.empty:
            category_warranty = filtered_df.groupby(['Store', 'Month', category_column])['WarrantyPrice'].sum().reset_index()
            category_pivot_warranty = category_warranty.pivot_table(index=['Store', category_column], columns='Month', values='WarrantyPrice', aggfunc='first').fillna(0)
            category_pivot_warranty['Change Jul-Aug (₹)'] = category_pivot_warranty['August'] - category_pivot_warranty['July']
            category_pivot_warranty['Change Aug-Sep (₹)'] = category_pivot_warranty['September'] - category_pivot_warranty['August']
            
            store_metrics = filtered_df.groupby(['Store', 'Month']).agg({
                'TotalSoldPrice': 'sum',
                'WarrantyCount': 'sum',
                'TotalCount': 'sum',
                'AHSP': 'mean'
            }).reset_index()
            metrics_pivot = store_metrics.pivot_table(index='Store', columns='Month', values=['TotalSoldPrice', 'WarrantyCount', 'TotalCount', 'AHSP'], aggfunc='first').fillna(0)
            metrics_pivot.columns = [f"{col[1]} {col[0]}" for col in metrics_pivot.columns]
            
            for _, row in decreased_stores.iterrows():
                store = row['Store']
                st.write(f"**{store}:**")
                reasons = []
                
                store_cat = category_pivot_warranty.loc[store] if store in category_pivot_warranty.index.get_level_values(0) else pd.DataFrame()
                if not store_cat.empty:
                    decreased_cats = store_cat[(store_cat['Change Jul-Aug (₹)'] < 0) | (store_cat['Change Aug-Sep (₹)'] < 0)]
                    for cat, cat_row in decreased_cats.iterrows():
                        if cat_row['Change Jul-Aug (₹)'] < 0:
                            reasons.append(f"{cat} warranty sales decreased by ₹{abs(cat_row['Change Jul-Aug (₹)']):,.0f} from July to August.")
                        if cat_row['Change Aug-Sep (₹)'] < 0:
                            reasons.append(f"{cat} warranty sales decreased by ₹{abs(cat_row['Change Aug-Sep (₹)']):,.0f} from August to September.")
                
                if store in metrics_pivot.index:
                    july_total_sales = metrics_pivot.loc[store, 'July TotalSoldPrice']
                    august_total_sales = metrics_pivot.loc[store, 'August TotalSoldPrice']
                    september_total_sales = metrics_pivot.loc[store, 'September TotalSoldPrice']
                    if august_total_sales > july_total_sales:
                        reasons.append(f"Total sales increased (₹{august_total_sales:,.0f} vs. ₹{july_total_sales:,.0f}) from July to August, potentially diluting warranty sales.")
                    if september_total_sales > august_total_sales:
                        reasons.append(f"Total sales increased (₹{september_total_sales:,.0f} vs. ₹{august_total_sales:,.0f}) from August to September, potentially diluting warranty sales.")
                    
                    july_warranty_count = metrics_pivot.loc[store, 'July WarrantyCount']
                    august_warranty_count = metrics_pivot.loc[store, 'August WarrantyCount']
                    september_warranty_count = metrics_pivot.loc[store, 'September WarrantyCount']
                    if august_warranty_count < july_warranty_count:
                        reasons.append(f"Fewer warranty units sold ({august_warranty_count:.0f} vs. {july_warranty_count:.0f}) from July to August.")
                    if september_warranty_count < august_warranty_count:
                        reasons.append(f"Fewer warranty units sold ({september_warranty_count:.0f} vs. {august_warranty_count:.0f}) from August to September.")
                    
                    july_ahsp = metrics_pivot.loc[store, 'July AHSP']
                    august_ahsp = metrics_pivot.loc[store, 'August AHSP']
                    september_ahsp = metrics_pivot.loc[store, 'September AHSP']
                    if august_ahsp < july_ahsp:
                        reasons.append(f"Lower average warranty selling price (₹{august_ahsp:,.2f} vs. ₹{july_ahsp:,.2f}) from July to August.")
                    if september_ahsp < august_ahsp:
                        reasons.append(f"Lower average warranty selling price (₹{september_ahsp:,.2f} vs. ₹{august_ahsp:,.2f}) from August to September.")
                
                if reasons:
                    for reason in reasons:
                        st.write(f"- {reason}")
                    st.write("**Recommendations:**")
                    st.write("- Review sales strategies for underperforming product categories.")
                    st.write("- Enhance staff training on warranty benefits.")
                    st.write("- Introduce targeted promotions for warranty products.")
                else:
                    st.write("- No specific reasons identified; further analysis needed.")
        else:
            st.info("No stores experienced a warranty sales decrease with current filters.")
    else:
        st.info("No warranty sales data available for stores with current filters.")

# Significant Changes
with st.expander("Significant Changes", expanded=True):
    significant_stores = store_conv_pivot[(abs(store_conv_pivot['Count Conv Change Jul-Aug (%)']) > 2) | (abs(store_conv_pivot['Count Conv Change Aug-Sep (%)']) > 2)]
    if not significant_stores.empty:
        st.write("**Stores with Significant Count Conversion Changes:**")
        for store in significant_stores.index:
            change_jul_aug = float(store_conv_pivot.loc[store, 'Count Conv Change Jul-Aug (%)'])
            change_aug_sep = float(store_conv_pivot.loc[store, 'Count Conv Change Aug-Sep (%)'])
            if abs(change_jul_aug) > 2:
                st.write(f"- {store}: {change_jul_aug:.2f}% {'increase' if change_jul_aug > 0 else 'decrease'} from July to August")
            if abs(change_aug_sep) > 2:
                st.write(f"- {store}: {change_aug_sep:.2f}% {'increase' if change_aug_sep > 0 else 'decrease'} from August to September")

    significant_rbms = rbm_pivot[(abs(rbm_pivot['Count Conv Change Jul-Aug (%)']) > 2) | (abs(rbm_pivot['Count Conv Change Aug-Sep (%)']) > 2)]
    if not significant_rbms.empty:
        st.write("**RBMs with Significant Count Conversion Changes:**")
        for rbm in significant_rbms.index:
            change_jul_aug = float(rbm_pivot.loc[rbm, 'Count Conv Change Jul-Aug (%)'])
            change_aug_sep = float(rbm_pivot.loc[rbm, 'Count Conv Change Aug-Sep (%)'])
            if abs(change_jul_aug) > 2:
                st.write(f"- {rbm}: {change_jul_aug:.2f}% {'increase' if change_jul_aug > 0 else 'decrease'} from July to August")
            if abs(change_aug_sep) > 2:
                st.write(f"- {rbm}: {change_aug_sep:.2f}% {'increase' if change_aug_sep > 0 else 'decrease'} from August to September")

    if not category_pivot.empty:
        significant_categories = category_pivot[(abs(category_pivot['Count Conv Change Jul-Aug (%)']) > 2) | (abs(category_pivot['Count Conv Change Aug-Sep (%)']) > 2)]
        if not significant_categories.empty:
            st.write(f"**{category_column} with Significant Count Conversion Changes:**")
            for category in significant_categories.index:
                change_jul_aug = float(category_pivot.loc[category, 'Count Conv Change Jul-Aug (%)'])
                change_aug_sep = float(category_pivot.loc[category, 'Count Conv Change Aug-Sep (%)'])
                if abs(change_jul_aug) > 2:
                    st.write(f"- {category}: {change_jul_aug:.2f}% {'increase' if change_jul_aug > 0 else 'decrease'} from July to August")
                if abs(change_aug_sep) > 2:
                    st.write(f"- {category}: {change_aug_sep:.2f}% {'increase' if change_aug_sep > 0 else 'decrease'} from August to September")
    else:
        st.info("No significant item category changes with current filters.")

# Improvement Opportunities
with st.expander("Improvement Opportunities (September)", expanded=True):
    low_count_performers = store_summary[(store_summary['Month'] == 'September') & (store_summary['Conversion% (Count)'] < TARGET_COUNT_CONVERSION)]
    low_value_performers = store_summary[(store_summary['Month'] == 'September') & (store_summary['Conversion% (Price)'] < TARGET_VALUE_CONVERSION)]
    low_ahsp_performers = store_summary[(store_summary['Month'] == 'September') & (store_summary['AHSP'] < TARGET_AHSP)]
    
    if not low_count_performers.empty:
        st.write(f"These stores in September have below-target count conversion (target: {TARGET_COUNT_CONVERSION}%):")
        for _, row in low_count_performers.iterrows():
            st.write(f"- {row['Store']}: {row['Conversion% (Count)']:.2f}%")
    else:
        st.success(f"All stores meet or exceed the count conversion target ({TARGET_COUNT_CONVERSION}%) in September.")

    if not low_value_performers.empty:
        st.write(f"These stores in September have below-target value conversion (target: {TARGET_VALUE_CONVERSION}%):")
        for _, row in low_value_performers.iterrows():
            st.write(f"- {row['Store']}: {row['Conversion% (Price)']:.2f}%")
    else:
        st.success(f"All stores meet or exceed the value conversion target ({TARGET_VALUE_CONVERSION}%) in September.")

    if not low_ahsp_performers.empty:
        st.write(f"These stores in September have below-target AHSP (target: ₹{TARGET_AHSP:,.2f}):")
        for _, row in low_ahsp_performers.iterrows():
            st.write(f"- {row['Store']}: ₹{row['AHSP']:,.2f}")
    else:
        st.success(f"All stores meet or exceed the AHSP target (₹{TARGET_AHSP:,.2f}) in September.")

    if not low_count_performers.empty or not low_value_performers.empty or not low_ahsp_performers.empty:
        st.write("**Recommendations:**")
        st.write("1. Provide additional training on warranty benefits to improve conversion rates.")
        st.write("2. Create targeted promotions for high-value warranty products.")
        st.write("3. Review staffing and sales strategies in underperforming stores.")
        st.write("4. Implement incentives for selling higher-value warranty packages.")
    else:
        st.write("**Recommendations:**")
        st.write("1. Maintain current strategies to sustain performance.")
        st.write("2. Explore opportunities to exceed targets through innovative promotions.")

# FUTURE stores analysis
if future_filter or any('FUTURE' in store for store in filtered_df['Store'].unique()):
    with st.expander("FUTURE Stores Analysis", expanded=True):
        future_stores = filtered_df[filtered_df['Store'].str.contains('FUTURE', case=True)]
        if not future_stores.empty:
            future_summary = future_stores.groupby('Month').agg({
                'TotalSoldPrice': 'sum',
                'WarrantyPrice': 'sum',
                'TotalCount': 'sum',
                'WarrantyCount': 'sum',
                'AHSP': 'mean'
            }).reset_index()
            future_summary['Conversion% (Count)'] = (future_summary['WarrantyCount'] / future_summary['TotalCount'] * 100).round(2)
            future_summary['Conversion% (Price)'] = (future_summary['WarrantyPrice'] / future_summary['TotalSoldPrice'] * 100).round(2)
            
            july_future = future_summary[future_summary['Month'] == 'July']
            august_future = future_summary[future_summary['Month'] == 'August']
            september_future = future_summary[future_summary['Month'] == 'September']
            
            july_future_count = july_future['Conversion% (Count)'].iloc[0] if not july_future.empty else 0
            august_future_count = august_future['Conversion% (Count)'].iloc[0] if not august_future.empty else 0
            september_future_count = september_future['Conversion% (Count)'].iloc[0] if not september_future.empty else 0
            july_future_value = july_future['Conversion% (Price)'].iloc[0] if not july_future.empty else 0
            august_future_value = august_future['Conversion% (Price)'].iloc[0] if not august_future.empty else 0
            september_future_value = september_future['Conversion% (Price)'].iloc[0] if not september_future.empty else 0
            july_future_ahsp = july_future['AHSP'].iloc[0] if not july_future.empty else 0
            august_future_ahsp = august_future['AHSP'].iloc[0] if not august_future.empty else 0
            september_future_ahsp = september_future['AHSP'].iloc[0] if not september_future.empty else 0
            
            st.write(f"Average count conversion in FUTURE stores (July): {july_future_count:.2f}%")
            st.write(f"Average count conversion in FUTURE stores (August): {august_future_count:.2f}%")
            st.write(f"Average count conversion in FUTURE stores (September): {september_future_count:.2f}%")
            st.write(f"Average value conversion in FUTURE stores (July): {july_future_value:.2f}%")
            st.write(f"Average value conversion in FUTURE stores (August): {august_future_value:.2f}%")
            st.write(f"Average value conversion in FUTURE stores (September): {september_future_value:.2f}%")
            st.write(f"Average AHSP in FUTURE stores (July): ₹{july_future_ahsp:,.2f}")
            st.write(f"Average AHSP in FUTURE stores (August): ₹{august_future_ahsp:,.2f}")
            st.write(f"Average AHSP in FUTURE stores (September): ₹{september_future_ahsp:,.2f}")
            
            if september_future_count < TARGET_COUNT_CONVERSION or september_future_value < TARGET_VALUE_CONVERSION or september_future_ahsp < TARGET_AHSP:
                st.warning("FUTURE stores are below target in September.")
                st.write("**Recommendations:**")
                st.write("- Conduct store-specific training to boost conversion rates.")
                st.write("- Analyze customer demographics for targeted promotions.")
                st.write("- Review warranty pricing strategy to improve AHSP.")
            else:
                st.success("FUTURE stores meet or exceed targets in September!")
        else:
            st.info("No FUTURE stores in current selection")

# Replacement Warranty Categories Analysis
if replacement_filter:
    with st.expander("Replacement Warranty Categories Deep Dive", expanded=True):
        st.markdown('<h3 class="subheader">🔄 Replacement Warranty Category Performance</h3>', unsafe_allow_html=True)
        
        # Calculate warranty sales by replacement category
        replacement_sales = filtered_df.groupby(['Replacement Category', 'Month']).agg({
            'TotalSoldPrice': 'sum',
            'WarrantyPrice': 'sum',
            'TotalCount': 'sum',
            'WarrantyCount': 'sum'
        }).reset_index()
        
        # Calculate metrics
        replacement_sales['Conversion% (Count)'] = (replacement_sales['WarrantyCount'] / replacement_sales['TotalCount'] * 100).round(2)
        replacement_sales['Conversion% (Price)'] = (replacement_sales['WarrantyPrice'] / replacement_sales['TotalSoldPrice'] * 100).round(2)
        replacement_sales['AHSP'] = (replacement_sales['WarrantyPrice'] / replacement_sales['WarrantyCount']).where(replacement_sales['WarrantyCount'] > 0, 0).round(2)
        
        # Pivot for comparison
        replacement_pivot = replacement_sales.pivot_table(index='Replacement Category', columns='Month', 
                                                         values=['WarrantyPrice', 'Conversion% (Count)', 'Conversion% (Price)', 'AHSP'], 
                                                         aggfunc='first').fillna(0)
        replacement_pivot.columns = [f"{col[1]} {col[0]}" for col in replacement_pivot.columns]
        
        # Calculate changes
        replacement_pivot['Warranty Sales Change Jul-Aug (₹)'] = replacement_pivot['August WarrantyPrice'] - replacement_pivot['July WarrantyPrice']
        replacement_pivot['Warranty Sales Change Aug-Sep (₹)'] = replacement_pivot['September WarrantyPrice'] - replacement_pivot['August WarrantyPrice']
        replacement_pivot['Count Conversion Change Jul-Aug (%)'] = replacement_pivot['August Conversion% (Count)'] - replacement_pivot['July Conversion% (Count)']
        replacement_pivot['Count Conversion Change Aug-Sep (%)'] = replacement_pivot['September Conversion% (Count)'] - replacement_pivot['August Conversion% (Count)']
        replacement_pivot['Value Conversion Change Jul-Aug (%)'] = replacement_pivot['August Conversion% (Price)'] - replacement_pivot['July Conversion% (Price)']
        replacement_pivot['Value Conversion Change Aug-Sep (%)'] = replacement_pivot['September Conversion% (Price)'] - replacement_pivot['August Conversion% (Price)']
        replacement_pivot['AHSP Change Jul-Aug (₹)'] = replacement_pivot['August AHSP'] - replacement_pivot['July AHSP']
        replacement_pivot['AHSP Change Aug-Sep (₹)'] = replacement_pivot['September AHSP'] - replacement_pivot['August AHSP']
        
        # Prepare display
        display_cols = ['July WarrantyPrice', 'August WarrantyPrice', 'September WarrantyPrice', 'Warranty Sales Change Jul-Aug (₹)', 'Warranty Sales Change Aug-Sep (₹)',
                       'July Conversion% (Count)', 'August Conversion% (Count)', 'September Conversion% (Count)', 'Count Conversion Change Jul-Aug (%)', 'Count Conversion Change Aug-Sep (%)',
                       'July Conversion% (Price)', 'August Conversion% (Price)', 'September Conversion% (Price)', 'Value Conversion Change Jul-Aug (%)', 'Value Conversion Change Aug-Sep (%)',
                       'July AHSP', 'August AHSP', 'September AHSP', 'AHSP Change Jul-Aug (₹)', 'AHSP Change Aug-Sep (₹)']
        
        display_df = replacement_pivot[display_cols].reset_index()
        display_df.columns = ['Replacement Category', 'July Warranty (₹)', 'August Warranty (₹)', 'September Warranty (₹)', 'Warranty Change Jul-Aug (₹)', 'Warranty Change Aug-Sep (₹)',
                             'July Count Conv (%)', 'August Count Conv (%)', 'September Count Conv (%)', 'Count Conv Change Jul-Aug (%)', 'Count Conv Change Aug-Sep (%)',
                             'July Value Conv (%)', 'August Value Conv (%)', 'September Value Conv (%)', 'Value Conv Change Jul-Aug (%)', 'Value Conv Change Aug-Sep (%)',
                             'July AHSP (₹)', 'August AHSP (₹)', 'September AHSP (₹)', 'AHSP Change Jul-Aug (₹)', 'AHSP Change Aug-Sep (₹)']
        
        def highlight_replacement_change(val, column):
            if column in ['Warranty Change Jul-Aug (₹)', 'Warranty Change Aug-Sep (₹)', 'Count Conv Change Jul-Aug (%)', 'Count Conv Change Aug-Sep (%)', 'Value Conv Change Jul-Aug (%)', 'Value Conv Change Aug-Sep (%)', 'AHSP Change Jul-Aug (₹)', 'AHSP Change Aug-Sep (₹)']:
                return 'color: green' if val > 0 else 'color: red' if val < 0 else 'color: black'
            return ''
        
        styled_df = display_df.style.format({
            'July Warranty (₹)': '₹{:.0f}',
            'August Warranty (₹)': '₹{:.0f}',
            'September Warranty (₹)': '₹{:.0f}',
            'Warranty Change Jul-Aug (₹)': '₹{:.0f}',
            'Warranty Change Aug-Sep (₹)': '₹{:.0f}',
            'July Count Conv (%)': '{:.2f}%',
            'August Count Conv (%)': '{:.2f}%',
            'September Count Conv (%)': '{:.2f}%',
            'Count Conv Change Jul-Aug (%)': '{:.2f}%',
            'Count Conv Change Aug-Sep (%)': '{:.2f}%',
            'July Value Conv (%)': '{:.2f}%',
            'August Value Conv (%)': '{:.2f}%',
            'September Value Conv (%)': '{:.2f}%',
            'Value Conv Change Jul-Aug (%)': '{:.2f}%',
            'Value Conv Change Aug-Sep (%)': '{:.2f}%',
            'July AHSP (₹)': '₹{:.2f}',
            'August AHSP (₹)': '₹{:.2f}',
            'September AHSP (₹)': '₹{:.2f}',
            'AHSP Change Jul-Aug (₹)': '₹{:.2f}',
            'AHSP Change Aug-Sep (₹)': '₹{:.2f}'
        })
        
        for col in ['Warranty Change Jul-Aug (₹)', 'Warranty Change Aug-Sep (₹)', 'Count Conv Change Jul-Aug (%)', 'Count Conv Change Aug-Sep (%)', 'Value Conv Change Jul-Aug (%)', 'Value Conv Change Aug-Sep (%)', 'AHSP Change Jul-Aug (₹)', 'AHSP Change Aug-Sep (₹)']:
            styled_df = styled_df.apply(lambda x: [highlight_replacement_change(val, col) for val in x], subset=[col], axis=0)
        
        st.dataframe(styled_df, use_container_width=True)
        
        # Visualization
        fig_replacement = px.bar(replacement_sales, 
                                x='Replacement Category', 
                                y='Conversion% (Count)', 
                                color='Month', 
                                barmode='group', 
                                title='Replacement Category Count Conversion: July vs August vs September',
                                template='plotly_white',
                                color_discrete_sequence=['#3730a3', '#06b6d4', '#10b981'])
        fig_replacement.update_layout(
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)',
            font=dict(family="Poppins, Inter, sans-serif", size=12, color="#1f2937"),
            showlegend=True,
            xaxis=dict(showgrid=False, tickangle=45),
            yaxis=dict(showgrid=True, gridcolor='#e5e7eb')
        )
        st.plotly_chart(fig_replacement, use_container_width=True)

# Speaker Categories Analysis
if speaker_filter:
    with st.expander("Speaker Categories Deep Dive", expanded=True):
        st.markdown('<h3 class="subheader">🔊 Speaker Category Performance</h3>', unsafe_allow_html=True)
        
        # Calculate warranty sales by speaker category
        speaker_categories = ['SOUND BAR', 'PARTY SPEAKER', 'BLUETOOTH SPEAKER', 'HOME THEATRE']
        speaker_sales = filtered_df[filtered_df['Item Category'].isin(speaker_categories)].groupby(['Item Category', 'Month']).agg({
            'TotalSoldPrice': 'sum',
            'WarrantyPrice': 'sum',
            'TotalCount': 'sum',
            'WarrantyCount': 'sum'
        }).reset_index()
        
        # Calculate metrics
        speaker_sales['Conversion% (Count)'] = (speaker_sales['WarrantyCount'] / speaker_sales['TotalCount'] * 100).round(2)
        speaker_sales['Conversion% (Price)'] = (speaker_sales['WarrantyPrice'] / speaker_sales['TotalSoldPrice'] * 100).round(2)
        speaker_sales['AHSP'] = (speaker_sales['WarrantyPrice'] / speaker_sales['WarrantyCount']).where(speaker_sales['WarrantyCount'] > 0, 0).round(2)
        
        # Pivot for comparison
        speaker_pivot = speaker_sales.pivot_table(index='Item Category', columns='Month', 
                                                 values=['WarrantyPrice', 'Conversion% (Count)', 'Conversion% (Price)', 'AHSP'], 
                                                 aggfunc='first').fillna(0)
        speaker_pivot.columns = [f"{col[1]} {col[0]}" for col in speaker_pivot.columns]
        
        # Ensure all required columns exist
        for col in ['July WarrantyPrice', 'August WarrantyPrice', 'September WarrantyPrice', 'July Conversion% (Count)', 'August Conversion% (Count)', 'September Conversion% (Count)', 
                    'July Conversion% (Price)', 'August Conversion% (Price)', 'September Conversion% (Price)', 'July AHSP', 'August AHSP', 'September AHSP']:
            if col not in speaker_pivot.columns:
                speaker_pivot[col] = 0
        
        # Calculate changes
        speaker_pivot['Warranty Sales Change Jul-Aug (₹)'] = speaker_pivot['August WarrantyPrice'] - speaker_pivot['July WarrantyPrice']
        speaker_pivot['Warranty Sales Change Aug-Sep (₹)'] = speaker_pivot['September WarrantyPrice'] - speaker_pivot['August WarrantyPrice']
        speaker_pivot['Count Conversion Change Jul-Aug (%)'] = speaker_pivot['August Conversion% (Count)'] - speaker_pivot['July Conversion% (Count)']
        speaker_pivot['Count Conversion Change Aug-Sep (%)'] = speaker_pivot['September Conversion% (Count)'] - speaker_pivot['August Conversion% (Count)']
        speaker_pivot['Value Conversion Change Jul-Aug (%)'] = speaker_pivot['August Conversion% (Price)'] - speaker_pivot['July Conversion% (Price)']
        speaker_pivot['Value Conversion Change Aug-Sep (%)'] = speaker_pivot['September Conversion% (Price)'] - speaker_pivot['August Conversion% (Price)']
        speaker_pivot['AHSP Change Jul-Aug (₹)'] = speaker_pivot['August AHSP'] - speaker_pivot['July AHSP']
        speaker_pivot['AHSP Change Aug-Sep (₹)'] = speaker_pivot['September AHSP'] - speaker_pivot['August AHSP']
        
        # Prepare display
        display_cols = ['July WarrantyPrice', 'August WarrantyPrice', 'September WarrantyPrice', 'Warranty Sales Change Jul-Aug (₹)', 'Warranty Sales Change Aug-Sep (₹)',
                       'July Conversion% (Count)', 'August Conversion% (Count)', 'September Conversion% (Count)', 'Count Conversion Change Jul-Aug (%)', 'Count Conversion Change Aug-Sep (%)',
                       'July Conversion% (Price)', 'August Conversion% (Price)', 'September Conversion% (Price)', 'Value Conversion Change Jul-Aug (%)', 'Value Conversion Change Aug-Sep (%)',
                       'July AHSP', 'August AHSP', 'September AHSP', 'AHSP Change Jul-Aug (₹)', 'AHSP Change Aug-Sep (₹)']
        
        display_df = speaker_pivot[display_cols].reset_index()
        display_df.columns = ['Speaker Category', 'July Warranty (₹)', 'August Warranty (₹)', 'September Warranty (₹)', 'Warranty Change Jul-Aug (₹)', 'Warranty Change Aug-Sep (₹)',
                             'July Count Conv (%)', 'August Count Conv (%)', 'September Count Conv (%)', 'Count Conv Change Jul-Aug (%)', 'Count Conv Change Aug-Sep (%)',
                             'July Value Conv (%)', 'August Value Conv (%)', 'September Value Conv (%)', 'Value Conv Change Jul-Aug (%)', 'Value Conv Change Aug-Sep (%)',
                             'July AHSP (₹)', 'August AHSP (₹)', 'September AHSP (₹)', 'AHSP Change Jul-Aug (₹)', 'AHSP Change Aug-Sep (₹)']
        
        st.dataframe(display_df.style.format({
            'July Warranty (₹)': '₹{:.0f}',
            'August Warranty (₹)': '₹{:.0f}',
            'September Warranty (₹)': '₹{:.0f}',
            'Warranty Change Jul-Aug (₹)': '₹{:.0f}',
            'Warranty Change Aug-Sep (₹)': '₹{:.0f}',
            'July Count Conv (%)': '{:.2f}%',
            'August Count Conv (%)': '{:.2f}%',
            'September Count Conv (%)': '{:.2f}%',
            'Count Conv Change Jul-Aug (%)': '{:.2f}%',
            'Count Conv Change Aug-Sep (%)': '{:.2f}%',
            'July Value Conv (%)': '{:.2f}%',
            'August Value Conv (%)': '{:.2f}%',
            'September Value Conv (%)': '{:.2f}%',
            'Value Conv Change Jul-Aug (%)': '{:.2f}%',
            'Value Conv Change Aug-Sep (%)': '{:.2f}%',
            'July AHSP (₹)': '₹{:.2f}',
            'August AHSP (₹)': '₹{:.2f}',
            'September AHSP (₹)': '₹{:.2f}',
            'AHSP Change Jul-Aug (₹)': '₹{:.2f}',
            'AHSP Change Aug-Sep (₹)': '₹{:.2f}'
        }).applymap(highlight_change, subset=['Warranty Change Jul-Aug (₹)', 'Warranty Change Aug-Sep (₹)', 'Count Conv Change Jul-Aug (%)', 'Count Conv Change Aug-Sep (%)', 'Value Conv Change Jul-Aug (%)', 'Value Conv Change Aug-Sep (%)', 'AHSP Change Jul-Aug (₹)', 'AHSP Change Aug-Sep (₹)']), 
        use_container_width=True)
        
        # Visualization
        fig_speaker = px.bar(speaker_sales, 
                             x='Item Category', 
                             y='Conversion% (Count)', 
                             color='Month', 
                             barmode='group', 
                             title='Speaker Category Count Conversion: July vs August vs September',
                             template='plotly_white',
                             color_discrete_sequence=['#3730a3', '#06b6d4', '#10b981'])
        fig_speaker.update_layout(
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)',
            font=dict(family="Poppins, Inter, sans-serif", size=12, color="#1f2937"),
            showlegend=True,
            xaxis=dict(showgrid=False, tickangle=45),
            yaxis=dict(showgrid=True, gridcolor='#e5e7eb')
        )
        st.plotly_chart(fig_speaker, use_container_width=True)

# Correlation Analysis
st.markdown('<h3 class="subheader">🔗 Correlation Analysis (September)</h3>', unsafe_allow_html=True)
corr_matrix = filtered_df[filtered_df['Month'] == 'September'][['TotalSoldPrice', 'WarrantyPrice', 'TotalCount', 'WarrantyCount', 'AHSP']].corr()
fig_corr = px.imshow(
    corr_matrix,
    text_auto=True,
    title="Correlation Matrix of Key Metrics (September)",
    template='plotly_white',
    color_continuous_scale='Blues',
    zmin=-1,
    zmax=1
)
fig_corr.update_layout(
    plot_bgcolor='rgba(0,0,0,0)',
    paper_bgcolor='rgba(0,0,0,0)',
    font=dict(family="Poppins, Inter, sans-serif", size=12, color="#1f2937"),
    showlegend=True
)
st.plotly_chart(fig_corr, use_container_width=True)
