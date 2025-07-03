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

MAY_FILE_PATH = os.path.join(DATA_DIR, "may_data.csv")
JUNE_FILE_PATH = os.path.join(DATA_DIR, "june_data.csv")
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
        col1, col2 = st.columns(2)
        with col1:
            may_file = st.file_uploader("May Data", type=["csv", "xlsx", "xls"], key="may")
        with col2:
            june_file = st.file_uploader("June Data", type=["csv", "xlsx", "xls"], key="june")
        upload_password = st.text_input("Upload Password", type="password", placeholder="Enter password to upload")
        submit_upload = st.form_submit_button("Upload", type="primary")
        if submit_upload:
            if authenticate("admin", upload_password):
                may_df = load_data(may_file, "May", MAY_FILE_PATH) if may_file else None
                june_df = load_data(june_file, "June", JUNE_FILE_PATH) if june_file else None
                if may_df is not None or june_df is not None:
                    df_list = [df for df in [may_df, june_df] if df is not None]
                    if df_list:
                        df = pd.concat(df_list, ignore_index=True)
            else:
                st.error("Invalid password. Upload failed.")
            if not may_file and not june_file:
                st.warning("Please select at least one file to upload.")

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
    
    # Month filter
    month_options = ['All', 'May', 'June']
    selected_month = st.selectbox("Month", month_options, index=0)

# Load saved data if no new upload
if df is None:
    may_df = load_data(MAY_FILE_PATH, "May") if os.path.exists(MAY_FILE_PATH) else None
    june_df = load_data(JUNE_FILE_PATH, "June") if os.path.exists(JUNE_FILE_PATH) else None
    df_list = [df for df in [may_df, june_df] if df is not None]
    if df_list:
        df = pd.concat(df_list, ignore_index=True)
        st.info("Loaded saved May and/or June data.")

# If no uploaded or saved files, use sample data with replacement categories
if df is None:
    data = {
        'Item Category': ['CEILING FAN', 'PEDESTAL FAN', 'MIXER GRINDER', 'IRON BOX', 'ELECTRIC KETTLE', 'OTG', 'GARMENTS STEAMER', 'INDUCTION COOKER'] * 3,
        'BDM': ['BDM1'] * 24,
        'RBM': ['MAHESH'] * 12 + ['RENJITH'] * 12,
        'Store': ['Palakkad FUTURE', 'Store B'] * 6 + ['Kannur FUTURE', 'Store C'] * 6,
        'Staff Name': ['Staff1', 'Staff2'] * 12,
        'TotalSoldPrice': [48239177/8, 48239177/8, 48239177/8, 48239177/8, 48239177/8, 48239177/8, 48239177/8, 48239177/8] * 3,
        'WarrantyPrice': [300619/8, 300619/8, 300619/8, 300619/8, 300619/8, 300619/8, 300619/8, 300619/8] * 3,
        'TotalCount': [5286/8, 5286/8, 5286/8, 5286/8, 5286/8, 5286/8, 5286/8, 5286/8] * 3,
        'WarrantyCount': [483/8, 483/8, 483/8, 483/8, 483/8, 483/8, 483/8, 483/8] * 3,
        'Month': ['May'] * 8 + ['June'] * 8 + ['May'] * 8
    }
    df = pd.DataFrame(data)
    df['Replacement Category'] = df['Item Category'].apply(map_to_replacement_category)
    df['Conversion% (Count)'] = (df['WarrantyCount'] / df['TotalCount'] * 100).round(2)
    df['Conversion% (Price)'] = (df['WarrantyPrice'] / df['TotalSoldPrice'] * 100).round(2)
    df['AHSP'] = (df['WarrantyPrice'] / df['WarrantyCount']).where(df['WarrantyCount'] > 0, 0).round(2)
    st.warning("Using sample data for May and June with replacement warranty categories.")

# Ensure Month column is categorical
df['Month'] = pd.Categorical(df['Month'], categories=['May', 'June'], ordered=True)

# Apply replacement warranty filter if selected
if replacement_filter:
    replacement_categories = ['FAN', 'MIXER GRINDER', 'IRON BOX', 'ELECTRIC KETTLE', 'OTG', 'STEAMER', 'INDUCTION COOKER']
    df = df[df['Replacement Category'].isin(replacement_categories)]
    category_column = 'Replacement Category'
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
st.markdown('<h2 class="subheader">📈 Performance Comparison: May vs June</h2>', unsafe_allow_html=True)

# KPI comparison
st.markdown('<h3 class="subheader">Overall KPIs</h3>', unsafe_allow_html=True)
col1, col2, col3, col4 = st.columns(4)

# Calculate metrics from the raw data (not averages)
may_data = filtered_df[filtered_df['Month'] == 'May']
june_data = filtered_df[filtered_df['Month'] == 'June']

may_total_warranty = may_data['WarrantyPrice'].sum()
may_total_units = may_data['TotalCount'].sum()
may_total_warranty_units = may_data['WarrantyCount'].sum()
may_total_sales = may_data['TotalSoldPrice'].sum()
may_value_conversion = (may_total_warranty / may_total_sales * 100) if may_total_sales > 0 else 0
may_count_conversion = (may_total_warranty_units / may_total_units * 100) if may_total_units > 0 else 0
may_ahsp = (may_total_warranty / may_total_warranty_units) if may_total_warranty_units > 0 else 0

june_total_warranty = june_data['WarrantyPrice'].sum()
june_total_units = june_data['TotalCount'].sum()
june_total_warranty_units = june_data['WarrantyCount'].sum()
june_total_sales = june_data['TotalSoldPrice'].sum()
june_value_conversion = (june_total_warranty / june_total_sales * 100) if june_total_sales > 0 else 0
june_count_conversion = (june_total_warranty_units / june_total_units * 100) if june_total_units > 0 else 0
june_ahsp = (june_total_warranty / june_total_warranty_units) if june_total_warranty_units > 0 else 0

with col1:
    st.metric("Warranty Sales (May)", f"₹{may_total_warranty:,.0f}")
    st.metric("Warranty Sales (June)", f"₹{june_total_warranty:,.0f}", 
              delta=f"{((june_total_warranty - may_total_warranty) / may_total_warranty * 100):.2f}%" if may_total_warranty > 0 else "N/A")

with col2:
    st.metric("Count Conversion (May)", f"{may_count_conversion:.2f}%")
    st.metric("Count Conversion (June)", f"{june_count_conversion:.2f}%", 
              delta=f"{june_count_conversion - may_count_conversion:.2f}%")

with col3:
    st.metric("Value Conversion (May)", f"{may_value_conversion:.2f}%")
    st.metric("Value Conversion (June)", f"{june_value_conversion:.2f}%", 
              delta=f"{june_value_conversion - may_value_conversion:.2f}%")

with col4:
    st.metric("AHSP (May)", f"₹{may_ahsp:,.2f}")
    st.metric("AHSP (June)", f"₹{june_ahsp:,.2f}", 
              delta=f"₹{june_ahsp - may_ahsp:,.2f}" if may_ahsp > 0 else "N/A")

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
for col in ['May Conversion% (Count)', 'June Conversion% (Count)', 'May Conversion% (Price)', 'June Conversion% (Price)', 'May AHSP', 'June AHSP']:
    if col not in store_conv_pivot.columns:
        store_conv_pivot[col] = 0

# Calculate changes
store_conv_pivot['Count Conversion Change (%)'] = (store_conv_pivot['June Conversion% (Count)'] - store_conv_pivot['May Conversion% (Count)']).round(2)
store_conv_pivot['Value Conversion Change (%)'] = (store_conv_pivot['June Conversion% (Price)'] - store_conv_pivot['May Conversion% (Price)']).round(2)
store_conv_pivot['AHSP Change (%)'] = ((store_conv_pivot['June AHSP'] - store_conv_pivot['May AHSP']) / store_conv_pivot['May AHSP'] * 100).where(store_conv_pivot['May AHSP'] > 0, 0).round(2)

# Sort by count conversion change
store_conv_pivot = store_conv_pivot.sort_values('Count Conversion Change (%)', ascending=False)

# Calculate TOTAL row using the same method as KPIs (sum of all, then calculate metrics)
total_may = store_summary[store_summary['Month'] == 'May'][['TotalSoldPrice', 'WarrantyPrice', 'TotalCount', 'WarrantyCount']].sum()
total_june = store_summary[store_summary['Month'] == 'June'][['TotalSoldPrice', 'WarrantyPrice', 'TotalCount', 'WarrantyCount']].sum()

total_row = pd.DataFrame({
    'Store': ['Total'],
    'May Conversion% (Count)': [(total_may['WarrantyCount'] / total_may['TotalCount'] * 100).round(2) if total_may['TotalCount'] > 0 else 0],
    'June Conversion% (Count)': [(total_june['WarrantyCount'] / total_june['TotalCount'] * 100).round(2) if total_june['TotalCount'] > 0 else 0],
    'May Conversion% (Price)': [(total_may['WarrantyPrice'] / total_may['TotalSoldPrice'] * 100).round(2) if total_may['TotalSoldPrice'] > 0 else 0],
    'June Conversion% (Price)': [(total_june['WarrantyPrice'] / total_june['TotalSoldPrice'] * 100).round(2) if total_june['TotalSoldPrice'] > 0 else 0],
    'May AHSP': [(total_may['WarrantyPrice'] / total_may['WarrantyCount']).round(2) if total_may['WarrantyCount'] > 0 else 0],
    'June AHSP': [(total_june['WarrantyPrice'] / total_june['WarrantyCount']).round(2) if total_june['WarrantyCount'] > 0 else 0],
    'Count Conversion Change (%)': [
        (total_june['WarrantyCount'] / total_june['TotalCount'] * 100 - total_may['WarrantyCount'] / total_may['TotalCount'] * 100).round(2) 
        if total_june['TotalCount'] > 0 and total_may['TotalCount'] > 0 else 0
    ],
    'Value Conversion Change (%)': [
        (total_june['WarrantyPrice'] / total_june['TotalSoldPrice'] * 100 - total_may['WarrantyPrice'] / total_may['TotalSoldPrice'] * 100).round(2) 
        if total_june['TotalSoldPrice'] > 0 and total_may['TotalSoldPrice'] > 0 else 0
    ],
    'AHSP Change (%)': [
        ((total_june['WarrantyPrice'] / total_june['WarrantyCount'] - total_may['WarrantyPrice'] / total_may['WarrantyCount']) / 
         (total_may['WarrantyPrice'] / total_may['WarrantyCount']) * 100).round(2) 
        if total_may['WarrantyCount'] > 0 and total_june['WarrantyCount'] > 0 else 0
    ]
})

# Prepare display dataframe
store_display = store_conv_pivot.reset_index()
store_display = store_display[['Store', 'May Conversion% (Count)', 'June Conversion% (Count)', 'May Conversion% (Price)', 'June Conversion% (Price)', 'May AHSP', 'June AHSP', 'Count Conversion Change (%)', 'Value Conversion Change (%)', 'AHSP Change (%)']]
store_display = pd.concat([store_display, total_row], ignore_index=True)
store_display['Total Change (%)'] = store_display[['Count Conversion Change (%)', 'Value Conversion Change (%)', 'AHSP Change (%)']].mean(axis=1).round(2)
store_display.columns = ['Store', 'May Count Conv (%)', 'June Count Conv (%)', 'May Value Conv (%)', 'June Value Conv (%)', 'May AHSP (₹)', 'June AHSP (₹)', 'Count Conv Change (%)', 'Value Conv Change (%)', 'AHSP Change (%)', 'Total Change (%)']
store_display = store_display.fillna(0)

def highlight_change(val):
    color = 'green' if val > 0 else 'red' if val < 0 else 'black'
    return f'color: {color}'

st.dataframe(store_display.style.format({
    'May Count Conv (%)': '{:.2f}%',
    'June Count Conv (%)': '{:.2f}%',
    'May Value Conv (%)': '{:.2f}%',
    'June Value Conv (%)': '{:.2f}%',
    'May AHSP (₹)': '₹{:.2f}',
    'June AHSP (₹)': '₹{:.2f}',
    'Count Conv Change (%)': '{:.2f}%',
    'Value Conv Change (%)': '{:.2f}%',
    'AHSP Change (%)': '{:.2f}%',
    'Total Change (%)': '{:.2f}%'
}).applymap(highlight_change, subset=['Count Conv Change (%)', 'Value Conv Change (%)', 'AHSP Change (%)', 'Total Change (%)'])
.set_properties(**{'font-weight': 'bold'}, subset=pd.IndexSlice[store_display.index[-1], :]), 
use_container_width=True)

# Visualization
fig_store = px.bar(store_summary, 
                   x='Store', 
                   y='Conversion% (Count)', 
                   color='Month', 
                   barmode='group', 
                   title='Store Count Conversion: May vs June',
                   template='plotly_white',
                   color_discrete_sequence=['#3730a3', '#06b6d4'])
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
for col in ['May Conversion% (Count)', 'June Conversion% (Count)', 'May Conversion% (Price)', 'June Conversion% (Price)', 'May AHSP', 'June AHSP']:
    if col not in rbm_pivot.columns:
        rbm_pivot[col] = 0

# Calculate changes
rbm_pivot['Count Conversion Change (%)'] = (rbm_pivot['June Conversion% (Count)'] - rbm_pivot['May Conversion% (Count)']).round(2)
rbm_pivot['Value Conversion Change (%)'] = (rbm_pivot['June Conversion% (Price)'] - rbm_pivot['May Conversion% (Price)']).round(2)
rbm_pivot['AHSP Change (%)'] = ((rbm_pivot['June AHSP'] - rbm_pivot['May AHSP']) / rbm_pivot['May AHSP'] * 100).where(rbm_pivot['May AHSP'] > 0, 0).round(2)
rbm_pivot = rbm_pivot.sort_values('Count Conversion Change (%)', ascending=False)

if decreased_rbm_filter:
    rbm_pivot = rbm_pivot[rbm_pivot['Count Conversion Change (%)'] < 0]
    if rbm_pivot.empty:
        st.warning("No RBMs with decreased count conversion match the filters.")

# Prepare display dataframe
rbm_display = rbm_pivot.reset_index()
rbm_display = rbm_display[['RBM', 'May Conversion% (Count)', 'June Conversion% (Count)', 'May Conversion% (Price)', 'June Conversion% (Price)', 'May AHSP', 'June AHSP', 'Count Conversion Change (%)', 'Value Conversion Change (%)', 'AHSP Change (%)']]
rbm_display.columns = ['RBM', 'May Count Conv (%)', 'June Count Conv (%)', 'May Value Conv (%)', 'June Value Conv (%)', 'May AHSP (₹)', 'June AHSP (₹)', 'Count Conv Change (%)', 'Value Conv Change (%)', 'AHSP Change (%)']
rbm_display = rbm_display.fillna(0)

st.dataframe(rbm_display.style.format({
    'May Count Conv (%)': '{:.2f}%',
    'June Count Conv (%)': '{:.2f}%',
    'May Value Conv (%)': '{:.2f}%',
    'June Value Conv (%)': '{:.2f}%',
    'May AHSP (₹)': '₹{:.2f}',
    'June AHSP (₹)': '₹{:.2f}',
    'Count Conv Change (%)': '{:.2f}%',
    'Value Conv Change (%)': '{:.2f}%',
    'AHSP Change (%)': '{:.2f}%'
}).applymap(highlight_change, subset=['Count Conv Change (%)', 'Value Conv Change (%)', 'AHSP Change (%)']), use_container_width=True)

# Visualization
fig_rbm = px.bar(rbm_summary, 
                 x='RBM', 
                 y='Conversion% (Count)', 
                 color='Month', 
                 barmode='group', 
                 title='RBM Count Conversion: May vs June',
                 template='plotly_white',
                 color_discrete_sequence=['#3730a3', '#06b6d4'])
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
    for col in ['May Conversion% (Count)', 'June Conversion% (Count)', 'May Conversion% (Price)', 'June Conversion% (Price)', 'May AHSP', 'June AHSP']:
        if col not in category_pivot.columns:
            category_pivot[col] = 0

    # Calculate changes
    category_pivot['Count Conversion Change (%)'] = (category_pivot['June Conversion% (Count)'] - category_pivot['May Conversion% (Count)']).round(2)
    category_pivot['Value Conversion Change (%)'] = (category_pivot['June Conversion% (Price)'] - category_pivot['May Conversion% (Price)']).round(2)
    category_pivot['AHSP Change (%)'] = ((category_pivot['June AHSP'] - category_pivot['May AHSP']) / category_pivot['May AHSP'] * 100).where(category_pivot['May AHSP'] > 0, 0).round(2)
    category_pivot = category_pivot.sort_values('Count Conversion Change (%)', ascending=False)

    # Prepare display dataframe
    category_display = category_pivot.reset_index()
    category_display = category_display[[category_column, 'May Conversion% (Count)', 'June Conversion% (Count)', 'May Conversion% (Price)', 'June Conversion% (Price)', 'May AHSP', 'June AHSP', 'Count Conversion Change (%)', 'Value Conversion Change (%)', 'AHSP Change (%)']]
    category_display.columns = [category_column, 'May Count Conv (%)', 'June Count Conv (%)', 'May Value Conv (%)', 'June Value Conv (%)', 'May AHSP (₹)', 'June AHSP (₹)', 'Count Conv Change (%)', 'Value Conv Change (%)', 'AHSP Change (%)']
    category_display = category_display.fillna(0)

    st.dataframe(category_display.style.format({
        'May Count Conv (%)': '{:.2f}%',
        'June Count Conv (%)': '{:.2f}%',
        'May Value Conv (%)': '{:.2f}%',
        'June Value Conv (%)': '{:.2f}%',
        'May AHSP (₹)': '₹{:.2f}',
        'June AHSP (₹)': '₹{:.2f}',
        'Count Conv Change (%)': '{:.2f}%',
        'Value Conv Change (%)': '{:.2f}%',
        'AHSP Change (%)': '{:.2f}%'
    }).applymap(highlight_change, subset=['Count Conv Change (%)', 'Value Conv Change (%)', 'AHSP Change (%)']), use_container_width=True)

    # Visualization
    fig_category = px.bar(category_summary, 
                          x=category_column, 
                          y='Conversion% (Count)', 
                          color='Month', 
                          barmode='group', 
                          title=f'{category_column} Count Conversion: May vs June',
                          template='plotly_white',
                          color_discrete_sequence=['#3730a3', '#06b6d4'])
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
    if june_count_conversion >= TARGET_COUNT_CONVERSION:
        st.success(f"June Count Conversion ({june_count_conversion:.2f}%) meets or exceeds target ({TARGET_COUNT_CONVERSION}%).")
    else:
        st.warning(f"June Count Conversion ({june_count_conversion:.2f}%) is below target ({TARGET_COUNT_CONVERSION}%). Gap: {TARGET_COUNT_CONVERSION - june_count_conversion:.2f}%.")
    if june_count_conversion > may_count_conversion:
        st.success(f"Count Conversion improved from {may_count_conversion:.2f}% in May to {june_count_conversion:.2f}% in June.")
    else:
        st.warning(f"Count Conversion declined from {may_count_conversion:.2f}% in May to {june_count_conversion:.2f}% in June.")

    if june_value_conversion >= TARGET_VALUE_CONVERSION:
        st.success(f"June Value Conversion ({june_value_conversion:.2f}%) meets or exceeds target ({TARGET_VALUE_CONVERSION}%).")
    else:
        st.warning(f"June Value Conversion ({june_value_conversion:.2f}%) is below target ({TARGET_VALUE_CONVERSION}%). Gap: {TARGET_VALUE_CONVERSION - june_value_conversion:.2f}%.")
    if june_value_conversion > may_value_conversion:
        st.success(f"Value Conversion improved from {may_value_conversion:.2f}% in May to {june_value_conversion:.2f}% in June.")
    else:
        st.warning(f"Value Conversion declined from {may_value_conversion:.2f}% in May to {june_value_conversion:.2f}% in June.")

    if june_ahsp >= TARGET_AHSP:
        st.success(f"June AHSP (₹{june_ahsp:,.2f}) meets or exceeds target (₹{TARGET_AHSP:,.2f}).")
    else:
        st.warning(f"June AHSP (₹{june_ahsp:,.2f}) is below target (₹{TARGET_AHSP:,.2f}). Gap: ₹{TARGET_AHSP - june_ahsp:,.2f}.")
    if june_ahsp > may_ahsp:
        st.success(f"AHSP improved from ₹{may_ahsp:,.2f} in May to ₹{june_ahsp:,.2f} in June.")
    else:
        st.warning(f"AHSP declined from ₹{may_ahsp:,.2f} in May to ₹{june_ahsp:,.2f} in June.")

# Store-Level Warranty Sales Analysis
with st.expander("Store-Level Warranty Sales Analysis", expanded=True):
    store_warranty = filtered_df.groupby(['Store', 'Month'])['WarrantyPrice'].sum().reset_index()
    store_pivot = store_warranty.pivot_table(index='Store', columns='Month', values='WarrantyPrice', aggfunc='first').fillna(0)
    store_pivot['Change (₹)'] = store_pivot['June'] - store_pivot['May']
    store_pivot['Change (%)'] = ((store_pivot['June'] - store_pivot['May']) / store_pivot['May'] * 100).where(store_pivot['May'] > 0, 0).round(2)
    store_pivot = store_pivot.sort_values('Change (%)', ascending=False).reset_index()

    if not store_pivot.empty:
        st.write("Warranty sales comparison between May and June for each store:")
        display_df = store_pivot[['Store', 'May', 'June', 'Change (₹)', 'Change (%)']]
        display_df.columns = ['Store', 'May Warranty Sales (₹)', 'June Warranty Sales (₹)', 'Change (₹)', 'Change (%)']
        st.dataframe(display_df.style.format({
            'May Warranty Sales (₹)': '₹{:.0f}',
            'June Warranty Sales (₹)': '₹{:.0f}',
            'Change (₹)': '₹{:.0f}',
            'Change (%)': '{:.2f}%'
        }).applymap(highlight_change, subset=['Change (₹)', 'Change (%)']), use_container_width=True)

        st.write("**Reasons for Warranty Sales Decreases:**")
        decreased_stores = store_pivot[store_pivot['Change (%)'] < 0]
        if not decreased_stores.empty:
            category_warranty = filtered_df.groupby(['Store', 'Month', category_column])['WarrantyPrice'].sum().reset_index()
            category_pivot_warranty = category_warranty.pivot_table(index=['Store', category_column], columns='Month', values='WarrantyPrice', aggfunc='first').fillna(0)
            category_pivot_warranty['Change (₹)'] = category_pivot_warranty['June'] - category_pivot_warranty['May']
            
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
                change_amt = row['Change (₹)']
                change_pct = row['Change (%)']
                
                st.write(f"**{store} (Decrease: ₹{abs(change_amt):,.0f}, {change_pct:.2f}%):**")
                reasons = []
                
                store_cat = category_pivot_warranty.loc[store] if store in category_pivot_warranty.index.get_level_values(0) else pd.DataFrame()
                if not store_cat.empty:
                    decreased_cats = store_cat[store_cat['Change (₹)'] < 0]
                    for cat, cat_row in decreased_cats.iterrows():
                        reasons.append(f"{cat} warranty sales decreased by ₹{abs(cat_row['Change (₹)']):,.0f}.")
                
                if store in metrics_pivot.index:
                    may_total_sales = metrics_pivot.loc[store, 'May TotalSoldPrice']
                    june_total_sales = metrics_pivot.loc[store, 'June TotalSoldPrice']
                    if june_total_sales > may_total_sales:
                        reasons.append(f"Total sales increased (₹{june_total_sales:,.0f} vs. ₹{may_total_sales:,.0f}), potentially diluting warranty sales.")
                    
                    may_warranty_count = metrics_pivot.loc[store, 'May WarrantyCount']
                    june_warranty_count = metrics_pivot.loc[store, 'June WarrantyCount']
                    if june_warranty_count < may_warranty_count:
                        reasons.append(f"Fewer warranty units sold ({june_warranty_count:.0f} vs. {may_warranty_count:.0f}).")
                    
                    may_ahsp = metrics_pivot.loc[store, 'May AHSP']
                    june_ahsp = metrics_pivot.loc[store, 'June AHSP']
                    if june_ahsp < may_ahsp:
                        reasons.append(f"Lower average warranty selling price (₹{june_ahsp:,.2f} vs. ₹{may_ahsp:,.2f}).")
                
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
    significant_stores = store_conv_pivot[abs(store_conv_pivot['Count Conversion Change (%)']) > 2]
    if not significant_stores.empty:
        st.write("**Stores with Significant Count Conversion Changes:**")
        for store in significant_stores.index:
            change = float(store_conv_pivot.loc[store, 'Count Conversion Change (%)'])
            st.write(f"- {store}: {change:.2f}% {'increase' if change > 0 else 'decrease'}")

    significant_rbms = rbm_pivot[abs(rbm_pivot['Count Conversion Change (%)']) > 2]
    if not significant_rbms.empty:
        st.write("**RBMs with Significant Count Conversion Changes:**")
        for rbm in significant_rbms.index:
            change = float(rbm_pivot.loc[rbm, 'Count Conversion Change (%)'])
            st.write(f"- {rbm}: {change:.2f}% {'increase' if change > 0 else 'decrease'}")

    if not category_pivot.empty:
        significant_categories = category_pivot[abs(category_pivot['Count Conversion Change (%)']) > 2]
        if not significant_categories.empty:
            st.write(f"**{category_column} with Significant Count Conversion Changes:**")
            for category in significant_categories.index:
                change = float(category_pivot.loc[category, 'Count Conversion Change (%)'])
                st.write(f"- {category}: {change:.2f}% {'increase' if change > 0 else 'decrease'}")
    else:
        st.info("No significant item category changes with current filters.")

# Improvement Opportunities
with st.expander("Improvement Opportunities (June)", expanded=True):
    low_count_performers = store_summary[(store_summary['Month'] == 'June') & (store_summary['Conversion% (Count)'] < TARGET_COUNT_CONVERSION)]
    low_value_performers = store_summary[(store_summary['Month'] == 'June') & (store_summary['Conversion% (Price)'] < TARGET_VALUE_CONVERSION)]
    low_ahsp_performers = store_summary[(store_summary['Month'] == 'June') & (store_summary['AHSP'] < TARGET_AHSP)]
    
    if not low_count_performers.empty:
        st.write(f"These stores in June have below-target count conversion (target: {TARGET_COUNT_CONVERSION}%):")
        for _, row in low_count_performers.iterrows():
            st.write(f"- {row['Store']}: {row['Conversion% (Count)']:.2f}%")
    else:
        st.success(f"All stores meet or exceed the count conversion target ({TARGET_COUNT_CONVERSION}%) in June.")

    if not low_value_performers.empty:
        st.write(f"These stores in June have below-target value conversion (target: {TARGET_VALUE_CONVERSION}%):")
        for _, row in low_value_performers.iterrows():
            st.write(f"- {row['Store']}: {row['Conversion% (Price)']:.2f}%")
    else:
        st.success(f"All stores meet or exceed the value conversion target ({TARGET_VALUE_CONVERSION}%) in June.")

    if not low_ahsp_performers.empty:
        st.write(f"These stores in June have below-target AHSP (target: ₹{TARGET_AHSP:,.2f}):")
        for _, row in low_ahsp_performers.iterrows():
            st.write(f"- {row['Store']}: ₹{row['AHSP']:,.2f}")
    else:
        st.success(f"All stores meet or exceed the AHSP target (₹{TARGET_AHSP:,.2f}) in June.")

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
            
            may_future = future_summary[future_summary['Month'] == 'May']
            june_future = future_summary[future_summary['Month'] == 'June']
            
            may_future_count = may_future['Conversion% (Count)'].iloc[0] if not may_future.empty else 0
            june_future_count = june_future['Conversion% (Count)'].iloc[0] if not june_future.empty else 0
            may_future_value = may_future['Conversion% (Price)'].iloc[0] if not may_future.empty else 0
            june_future_value = june_future['Conversion% (Price)'].iloc[0] if not june_future.empty else 0
            may_future_ahsp = may_future['AHSP'].iloc[0] if not may_future.empty else 0
            june_future_ahsp = june_future['AHSP'].iloc[0] if not june_future.empty else 0
            
            st.write(f"Average count conversion in FUTURE stores (May): {may_future_count:.2f}%")
            st.write(f"Average count conversion in FUTURE stores (June): {june_future_count:.2f}%")
            st.write(f"Average value conversion in FUTURE stores (May): {may_future_value:.2f}%")
            st.write(f"Average value conversion in FUTURE stores (June): {june_future_value:.2f}%")
            st.write(f"Average AHSP in FUTURE stores (May): ₹{may_future_ahsp:,.2f}")
            st.write(f"Average AHSP in FUTURE stores (June): ₹{june_future_ahsp:,.2f}")
            
            if june_future_count < TARGET_COUNT_CONVERSION or june_future_value < TARGET_VALUE_CONVERSION or june_future_ahsp < TARGET_AHSP:
                st.warning("FUTURE stores are below target in June.")
                st.write("**Recommendations:**")
                st.write("- Conduct store-specific training to boost conversion rates.")
                st.write("- Analyze customer demographics for targeted promotions.")
                st.write("- Review warranty pricing strategy to improve AHSP.")
            else:
                st.success("FUTURE stores meet or exceed targets in June!")
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
        replacement_pivot['Warranty Sales Change (₹)'] = replacement_pivot['June WarrantyPrice'] - replacement_pivot['May WarrantyPrice']
        replacement_pivot['Count Conversion Change (%)'] = replacement_pivot['June Conversion% (Count)'] - replacement_pivot['May Conversion% (Count)']
        replacement_pivot['Value Conversion Change (%)'] = replacement_pivot['June Conversion% (Price)'] - replacement_pivot['May Conversion% (Price)']
        replacement_pivot['AHSP Change (₹)'] = replacement_pivot['June AHSP'] - replacement_pivot['May AHSP']
        
        # Prepare display
        display_cols = ['May WarrantyPrice', 'June WarrantyPrice', 'Warranty Sales Change (₹)',
                       'May Conversion% (Count)', 'June Conversion% (Count)', 'Count Conversion Change (%)',
                       'May Conversion% (Price)', 'June Conversion% (Price)', 'Value Conversion Change (%)',
                       'May AHSP', 'June AHSP', 'AHSP Change (₹)']
        
        display_df = replacement_pivot[display_cols].reset_index()
        display_df.columns = ['Replacement Category', 'May Warranty (₹)', 'June Warranty (₹)', 'Warranty Change (₹)',
                             'May Count Conv (%)', 'June Count Conv (%)', 'Count Conv Change (%)',
                             'May Value Conv (%)', 'June Value Conv (%)', 'Value Conv Change (%)',
                             'May AHSP (₹)', 'June AHSP (₹)', 'AHSP Change (₹)']
        
        st.dataframe(display_df.style.format({
            'May Warranty (₹)': '₹{:.0f}',
            'June Warranty (₹)': '₹{:.0f}',
            'Warranty Change (₹)': '₹{:.0f}',
            'May Count Conv (%)': '{:.2f}%',
            'June Count Conv (%)': '{:.2f}%',
            'Count Conv Change (%)': '{:.2f}%',
            'May Value Conv (%)': '{:.2f}%',
            'June Value Conv (%)': '{:.2f}%',
            'Value Conv Change (%)': '{:.2f}%',
            'May AHSP (₹)': '₹{:.2f}',
            'June AHSP (₹)': '₹{:.2f}',
            'AHSP Change (₹)': '₹{:.2f}'
        }).applymap(highlight_change, subset=['Warranty Change (₹)', 'Count Conv Change (%)', 'Value Conv Change (%)', 'AHSP Change (₹)']), 
        use_container_width=True)
        
        # Visualization
        fig_replacement = px.bar(replacement_sales, 
                                x='Replacement Category', 
                                y='Conversion% (Count)', 
                                color='Month', 
                                barmode='group', 
                                title='Replacement Category Count Conversion: May vs June',
                                template='plotly_white',
                                color_discrete_sequence=['#3730a3', '#06b6d4'])
        fig_replacement.update_layout(
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)',
            font=dict(family="Poppins, Inter, sans-serif", size=12, color="#1f2937"),
            showlegend=True,
            xaxis=dict(showgrid=False, tickangle=45),
            yaxis=dict(showgrid=True, gridcolor='#e5e7eb')
        )
        st.plotly_chart(fig_replacement, use_container_width=True)

# Correlation Analysis
st.markdown('<h3 class="subheader">🔗 Correlation Analysis (June)</h3>', unsafe_allow_html=True)
corr_matrix = filtered_df[filtered_df['Month'] == 'June'][['TotalSoldPrice', 'WarrantyPrice', 'TotalCount', 'WarrantyCount', 'AHSP']].corr()
fig_corr = px.imshow(
    corr_matrix,
    text_auto=True,
    title="Correlation Matrix of Key Metrics (June)",
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
