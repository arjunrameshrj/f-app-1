import streamlit as st
import pandas as pd
import plotly.express as px
import os
import hashlib
from io import BytesIO

# Streamlit page configuration
st.set_page_config(page_title="Warranty Conversion Dashboard", layout="wide", initial_sidebar_state="expanded")

# Custom CSS for professional styling
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
        .low-value-conversion {
            background-color: #fee2e2 !important;
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
        .download-btn {
            margin-top: 10px;
            margin-bottom: 20px;
        }
        .download-btn button {
            background: linear-gradient(90deg, #10b981, #059669);
            color: white;
            border: none;
            padding: 10px 20px;
            border-radius: 8px;
            font-weight: 500;
            cursor: pointer;
            transition: all 0.3s ease;
        }
        .download-btn button:hover {
            background: linear-gradient(90deg, #059669, #047857);
            transform: translateY(-2px);
            box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1);
        }
    </style>
""", unsafe_allow_html=True)

# Function to convert DataFrame to Excel for download
def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Store Performance')
        workbook = writer.book
        worksheet = writer.sheets['Store Performance']
        
        # Add formatting
        header_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'top',
            'fg_color': '#3730a3',
            'font_color': 'white',
            'border': 1,
            'align': 'center'
        })
        
        # Format the header
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)
        
        # Add conditional formatting for low conversion
        low_conversion_format = workbook.add_format({'bg_color': '#fee2e2'})
        value_conv_col = df.columns.get_loc('Value Conv (%)') if 'Value Conv (%)' in df.columns else None
        
        if value_conv_col is not None:
            worksheet.conditional_format(1, value_conv_col, len(df), value_conv_col, {
                'type': 'cell',
                'criteria': '<',
                'value': 2.0,
                'format': low_conversion_format
            })
        
        # Format numbers
        num_format = workbook.add_format({'num_format': '#,##0.00'})
        percent_format = workbook.add_format({'num_format': '0.00%'})
        currency_format = workbook.add_format({'num_format': '₹#,##0.00'})
        
        # Apply formats to appropriate columns
        for col_num, col_name in enumerate(df.columns):
            if 'Conv (%)' in col_name:
                worksheet.set_column(col_num, col_num, None, percent_format)
            elif 'AHSP' in col_name or 'Warranty Sales' in col_name:
                worksheet.set_column(col_num, col_num, None, currency_format)
            elif df[col_name].dtype in ['float64', 'int64']:
                worksheet.set_column(col_num, col_num, None, num_format)
        
        # Bold the total row if it exists
        if 'Total' in df['Store'].values:
            total_row_format = workbook.add_format({'bold': True})
            total_row_index = df.index[df['Store'] == 'Total'][0] + 1  # +1 for header
            worksheet.set_row(total_row_index, None, total_row_format)
    
    processed_data = output.getvalue()
    return processed_data

# Create data directory if it doesn't exist
DATA_DIR = "data"
if not os.path.exists(DATA_DIR):
    os.makedirs(DATA_DIR)

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
    fan_categories = ['CEILING FAN', 'PEDESTAL FAN', 'RECHARGABLE FAN', 'TABLE FAN', 'TOWER FAN', 'WALL FAN']
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
        return item_category

# Session state initialization
if 'authenticated' not in st.session_state:
    st.session_state.authenticated = False
if 'show_upload_form' not in st.session_state:
    st.session_state.show_upload_form = False
if 'user_role' not in st.session_state:
    st.session_state.user_role = "non-admin"
if 'data_loaded' not in st.session_state:
    st.session_state.data_loaded = False
if 'current_df' not in st.session_state:
    st.session_state.current_df = None

# Main dashboard
st.markdown('<h1 class="main-header">📊 Warranty Conversion Analysis Dashboard - sep</h1>', unsafe_allow_html=True)

# Load data function
required_columns = ['Item Category', 'BDM', 'RBM', 'Store', 'Staff Name', 'TotalSoldPrice', 'WarrantyPrice', 'TotalCount', 'WarrantyCount']

@st.cache_data
def load_data(file, file_path=None):
    try:
        if isinstance(file, str):  # File path
            df = pd.read_csv(file)
        elif file.name.endswith('.csv'):  # Uploaded CSV
            df = pd.read_csv(file)
        else:  # Uploaded Excel
            df = pd.read_excel(file, engine='openpyxl')
        
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            st.error(f"Missing columns in June file: {', '.join(missing_columns)}")
            return None
        
        numeric_cols = ['TotalSoldPrice', 'WarrantyPrice', 'TotalCount', 'WarrantyCount']
        for col in numeric_cols:
            df[col] = pd.to_numeric(df[col], errors='coerce')
        
        if df[numeric_cols].isna().any().any():
            st.warning("Missing or invalid values in numeric columns for June. Filling with 0.")
            df[numeric_cols] = df[numeric_cols].fillna(0)
        
        # Add replacement category column
        df['Replacement Category'] = df['Item Category'].apply(map_to_replacement_category)
        
        df['Conversion% (Count)'] = (df['WarrantyCount'] / df['TotalCount'] * 100).round(2)
        df['Conversion% (Price)'] = (df['WarrantyPrice'] / df['TotalSoldPrice'] * 100).where(df['TotalSoldPrice'] > 0, 0).round(2)
        df['AHSP'] = (df['WarrantyPrice'] / df['WarrantyCount']).where(df['WarrantyCount'] > 0, 0).round(2)
        df['Month'] = 'June'
        
        if file_path and not isinstance(file, str):
            if file.name.endswith('.csv'):
                with open(file_path, 'wb') as f:
                    f.write(file.getvalue())
            else:
                df.to_csv(file_path, index=False)
            st.success("June data saved successfully and available to all users.")
        
        return df
    except Exception as e:
        st.error(f"Error loading June file: {str(e)}")
        return None

# Load saved data or handle new upload
df = None

# Sidebar content
with st.sidebar:
    st.markdown('<h2 style="color: #3730a3; font-weight: 600;">🔍 Dashboard Controls</h2>', unsafe_allow_html=True)
    st.markdown('<hr style="border: 1px solid #e5e7eb; margin: 10px 0;">', unsafe_allow_html=True)
    
    # File upload form
    st.markdown('<h3 style="color: #3730a3; font-weight: 500;">📁 Upload June Data</h3>', unsafe_allow_html=True)
    with st.form("upload_form"):
        june_file = st.file_uploader("June Data", type=["csv", "xlsx", "xls"], key="june")
        upload_password = st.text_input("Upload Password", type="password", placeholder="Enter password to upload")
        submit_upload = st.form_submit_button("Upload", type="primary")
        if submit_upload:
            if authenticate("admin", upload_password):
                if june_file:
                    df = load_data(june_file, JUNE_FILE_PATH)
                    if df is not None:
                        st.session_state.current_df = df
                        st.session_state.data_loaded = True
                        st.rerun()
                else:
                    st.warning("Please select a file to upload.")
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

# Load saved data if no new upload
if st.session_state.current_df is None:
    if os.path.exists(JUNE_FILE_PATH):
        df = load_data(JUNE_FILE_PATH)
        if df is not None:
            st.session_state.current_df = df
            st.session_state.data_loaded = True

# If no uploaded or saved files, use sample data with speaker categories
if st.session_state.current_df is None:
    data = {
        'Item Category': ['CEILING FAN', 'PEDESTAL FAN', 'MIXER GRINDER', 'IRON BOX', 'ELECTRIC KETTLE', 'OTG', 'GARMENTS STEAMER', 'INDUCTION COOKER', 'SOUND BAR', 'PARTY SPEAKER', 'BLUETOOTH SPEAKER', 'HOME THEATRE'],
        'BDM': ['BDM1'] * 12,
        'RBM': ['RENJITH'] * 12,
        'Store': ['Kannur FUTURE', 'Store C'] * 6,
        'Staff Name': ['Staff1', 'Staff2'] * 6,
        'TotalSoldPrice': [48239177/12] * 12,
        'WarrantyPrice': [300619/12] * 12,
        'TotalCount': [5286/12] * 12,
        'WarrantyCount': [483/12] * 12,
        'Month': ['June'] * 12
    }
    df = pd.DataFrame(data)
    df['Replacement Category'] = df['Item Category'].apply(map_to_replacement_category)
    df['Conversion% (Count)'] = (df['WarrantyCount'] / df['TotalCount'] * 100).round(2)
    df['Conversion% (Price)'] = (df['WarrantyPrice'] / df['TotalSoldPrice'] * 100).round(2)
    df['AHSP'] = (df['WarrantyPrice'] / df['WarrantyCount']).where(df['WarrantyPrice'] > 0, 0).round(2)
    st.session_state.current_df = df
    st.session_state.data_loaded = True

# Now that we have data, set up the filters in sidebar
if st.session_state.data_loaded and st.session_state.current_df is not None:
    df = st.session_state.current_df
    
    with st.sidebar:
        # Sidebar filters
        st.markdown('<hr style="border: 1px solid #e5e7eb; margin: 20px 0;">', unsafe_allow_html=True)
        st.markdown('<h3 style="color: #3730a3; font-weight: 500;">⚙️ Filters</h3>', unsafe_allow_html=True)
        
        replacement_filter = st.checkbox("Show Replacement Warranty Categories Only")
        speaker_filter = st.checkbox("Show Speaker Categories Only")
        future_filter = st.checkbox("Show only FUTURE stores")
        
        # NEW: Value Conversion range filter
        st.markdown('<hr style="border: 1px solid #e5e7eb; margin: 20px 0;">', unsafe_allow_html=True)
        st.markdown('<h4 style="color: #3730a3; font-weight: 500;">📊 Value Conversion Filter</h4>', unsafe_allow_html=True)
        
        # Calculate actual value conversion range from current data
        store_summary_temp = df.groupby('Store').agg({
            'TotalSoldPrice': 'sum',
            'WarrantyPrice': 'sum'
        }).reset_index()
        store_summary_temp['Value Conv (%)'] = (store_summary_temp['WarrantyPrice'] / store_summary_temp['TotalSoldPrice'] * 100).where(store_summary_temp['TotalSoldPrice'] > 0, 0).round(2)
        
        min_conv = float(store_summary_temp['Value Conv (%)'].min())
        max_conv = float(store_summary_temp['Value Conv (%)'].max())
        
        # FIX: Handle case where min and max are the same
        if min_conv == max_conv:
            # If all values are the same, create a small range around the value
            min_conv = max(0, min_conv - 0.1)
            max_conv = max_conv + 0.1
        
        value_conv_range = st.slider(
            "Filter by Value Conversion (%)",
            min_value=min_conv,
            max_value=max_conv,
            value=(min_conv, max_conv),
            step=0.1,
            help="Filter stores based on Value Conversion percentage range"
        )
        
        # Sorting option for Store Performance - UPDATED with Value Conversion option
        st.markdown('<hr style="border: 1px solid #e5e7eb; margin: 20px 0;">', unsafe_allow_html=True)
        sort_by = st.selectbox("Sort Store Performance By", ["Count Conv (%)", "Value Conv (%)", "AHSP (₹)", "Warranty Sales (₹)", "Warranty Units"], index=0)
        sort_order = st.selectbox("Sort Order", ["Descending", "Ascending"], index=0)

    # Apply replacement or speaker filter
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

    # Apply filters
    filtered_df = df.copy()
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
        st.warning("No data matches your filters.")
        st.stop()

    # Main dashboard
    st.markdown('<h2 class="subheader">📈 OCT Performance</h2>', unsafe_allow_html=True)

    # KPI metrics
    st.markdown('<h3 class="subheader">Warranty Total Count and KPIs</h3>', unsafe_allow_html=True)
    col1, col2, col3, col4 = st.columns(4)

    june_total_warranty = filtered_df['WarrantyPrice'].sum()
    june_total_units = filtered_df['TotalCount'].sum()
    june_total_warranty_units = filtered_df['WarrantyCount'].sum()
    june_total_sales = filtered_df['TotalSoldPrice'].sum()
    june_count_conversion = (june_total_warranty_units / june_total_units * 100) if june_total_units > 0 else 0
    june_value_conversion = (june_total_warranty / june_total_sales * 100) if june_total_sales > 0 else 0
    june_ahsp = (june_total_warranty / june_total_warranty_units) if june_total_warranty_units > 0 else 0

    with col1:
        st.metric("Warranty Sales", f"₹{june_total_warranty:,.0f}")
    with col2:
        st.metric("Count Conversion", f"{june_count_conversion:.2f}%")
    with col3:
        st.metric("Value Conversion", f"{june_value_conversion:.2f}%")
    with col4:
        st.metric("AHSP", f"₹{june_ahsp:,.2f}")

    # Store Performance
    st.markdown('<h3 class="subheader">🏬 Store Performance</h3>', unsafe_allow_html=True)

    store_summary = filtered_df.groupby('Store').agg({
        'TotalSoldPrice': 'sum',
        'WarrantyPrice': 'sum',
        'TotalCount': 'sum',
        'WarrantyCount': 'sum'
    }).reset_index()

    store_summary['Conversion% (Count)'] = (store_summary['WarrantyCount'] / store_summary['TotalCount'] * 100).round(2)
    store_summary['Conversion% (Price)'] = (store_summary['WarrantyPrice'] / store_summary['TotalSoldPrice'] * 100).round(2)
    store_summary['AHSP'] = (store_summary['WarrantyPrice'] / store_summary['WarrantyCount']).where(store_summary['WarrantyCount'] > 0, 0).round(2)

    store_display = store_summary[['Store', 'Conversion% (Count)', 'Conversion% (Price)', 'AHSP', 'WarrantyPrice', 'WarrantyCount']]
    store_display.columns = ['Store', 'Count Conv (%)', 'Value Conv (%)', 'AHSP (₹)', 'Warranty Sales (₹)', 'Warranty Units']

    # Create total row
    total_row = pd.DataFrame({
        'Store': ['Total'],
        'Count Conv (%)': [(store_summary['WarrantyCount'].sum() / store_summary['TotalCount'].sum() * 100).round(2) if store_summary['TotalCount'].sum() > 0 else 0],
        'Value Conv (%)': [(store_summary['WarrantyPrice'].sum() / store_summary['TotalSoldPrice'].sum() * 100).round(2) if store_summary['TotalSoldPrice'].sum() > 0 else 0],
        'AHSP (₹)': [(store_summary['WarrantyPrice'].sum() / store_summary['WarrantyCount'].sum()).round(2) if store_summary['WarrantyCount'].sum() > 0 else 0],
        'Warranty Sales (₹)': [store_summary['WarrantyPrice'].sum()],
        'Warranty Units': [store_summary['WarrantyCount'].sum()]
    })

    # Split data into non-total and total rows
    non_total_stores = store_display[store_display['Store'] != 'Total']

    # NEW: Apply Value Conversion range filter
    if value_conv_range != (min_conv, max_conv):
        non_total_stores = non_total_stores[
            (non_total_stores['Value Conv (%)'] >= value_conv_range[0]) & 
            (non_total_stores['Value Conv (%)'] <= value_conv_range[1])
        ]

    # NEW: Sort non-total rows based on user selection
    sort_ascending = True if sort_order == "Ascending" else False

    # Map display column names to actual column names
    sort_column_mapping = {
        "Count Conv (%)": "Count Conv (%)",
        "Value Conv (%)": "Value Conv (%)", 
        "AHSP (₹)": "AHSP (₹)",
        "Warranty Sales (₹)": "Warranty Sales (₹)",
        "Warranty Units": "Warranty Units"
    }

    sort_column = sort_column_mapping.get(sort_by, "Count Conv (%)")
    non_total_stores = non_total_stores.sort_values(sort_column, ascending=sort_ascending)

    # Concatenate sorted non-total rows with total row
    store_display = pd.concat([non_total_stores, total_row], ignore_index=True)

    # Apply conditional formatting for low value conversion
    def highlight_low_value_conversion(row):
        if row['Value Conv (%)'] < 2.0 and row['Store'] != 'Total':
            return ['background-color: #fee2e2'] * len(row)
        return [''] * len(row)

    styled_store_display = store_display.style.format({
        'Count Conv (%)': '{:.2f}%',
        'Value Conv (%)': '{:.2f}%',
        'AHSP (₹)': '₹{:.2f}',
        'Warranty Sales (₹)': '₹{:.0f}',
        'Warranty Units': '{:.0f}'
    }).set_properties(**{'font-weight': 'bold'}, subset=pd.IndexSlice[store_display.index[-1], :]).apply(highlight_low_value_conversion, axis=1)

    st.dataframe(styled_store_display, use_container_width=True)

    # Add Excel download button for Store Performance
    st.markdown('<div class="download-btn">', unsafe_allow_html=True)
    excel_data = to_excel(store_display)
    st.download_button(
        label="📥 Download Store Performance as Excel",
        data=excel_data,
        file_name="store_performance_june.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    st.markdown('</div>', unsafe_allow_html=True)

    # Create filtered store summary for visualization (excluding Total row for charts)
    store_summary_for_chart = store_summary[store_summary['Store'].isin(non_total_stores['Store'])]

    if not store_summary_for_chart.empty:
        # Sort the chart data based on the same criteria
        chart_sort_column_mapping = {
            "Count Conv (%)": "Conversion% (Count)",
            "Value Conv (%)": "Conversion% (Price)",
            "AHSP (₹)": "AHSP",
            "Warranty Sales (₹)": "WarrantyPrice", 
            "Warranty Units": "WarrantyCount"
        }
        
        chart_sort_column = chart_sort_column_mapping.get(sort_by, "Conversion% (Count)")
        store_summary_for_chart = store_summary_for_chart.sort_values(chart_sort_column, ascending=sort_ascending)
        
        # Determine which column to show in the chart based on sort selection
        if sort_by == "Value Conv (%)":
            chart_y_column = 'Conversion% (Price)'
            chart_title = 'Store Value Conversion - June'
        elif sort_by == "AHSP (₹)":
            chart_y_column = 'AHSP'
            chart_title = 'Store AHSP - June'
        elif sort_by == "Warranty Sales (₹)":
            chart_y_column = 'WarrantyPrice'
            chart_title = 'Store Warranty Sales - June'
        elif sort_by == "Warranty Units":
            chart_y_column = 'WarrantyCount'
            chart_title = 'Store Warranty Units - June'
        else:
            chart_y_column = 'Conversion% (Count)'
            chart_title = 'Store Count Conversion - June'

        fig_store = px.bar(store_summary_for_chart, 
                           x='Store', 
                           y=chart_y_column, 
                           title=chart_title,
                           template='plotly_white',
                           color_discrete_sequence=['#3730a3'])
        fig_store.update_layout(
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)',
            font=dict(family="Poppins, Inter, sans-serif", size=12, color="#1f2937"),
            xaxis=dict(showgrid=False, tickangle=45),
            yaxis=dict(showgrid=True, gridcolor='#e5e7eb')
        )
        st.plotly_chart(fig_store, use_container_width=True)

    # RBM Performance
    st.markdown('<h3 class="subheader">👥 RBM Performance</h3>', unsafe_allow_html=True)

    rbm_summary = filtered_df.groupby('RBM').agg({
        'TotalSoldPrice': 'sum',
        'WarrantyPrice': 'sum',
        'TotalCount': 'sum',
        'WarrantyCount': 'sum'
    }).reset_index()

    rbm_summary['Conversion% (Count)'] = (rbm_summary['WarrantyCount'] / rbm_summary['TotalCount'] * 100).round(2)
    rbm_summary['Conversion% (Price)'] = (rbm_summary['WarrantyPrice'] / rbm_summary['TotalSoldPrice'] * 100).round(2)
    rbm_summary['AHSP'] = (rbm_summary['WarrantyPrice'] / rbm_summary['WarrantyCount']).where(rbm_summary['WarrantyCount'] > 0, 0).round(2)

    rbm_display = rbm_summary[['RBM', 'Conversion% (Count)', 'Conversion% (Price)', 'AHSP', 'WarrantyPrice', 'WarrantyCount']]
    rbm_display.columns = ['RBM', 'Count Conv (%)', 'Value Conv (%)', 'AHSP (₹)', 'Warranty Sales (₹)', 'Warranty Units']
    rbm_display = rbm_display.sort_values('Count Conv (%)', ascending=False)

    st.dataframe(rbm_display.style.format({
        'Count Conv (%)': '{:.2f}%',
        'Value Conv (%)': '{:.2f}%',
        'AHSP (₹)': '₹{:.2f}',
        'Warranty Sales (₹)': '₹{:.0f}',
        'Warranty Units': '{:.0f}'
    }), use_container_width=True)

    fig_rbm = px.bar(rbm_summary, 
                     x='RBM', 
                     y='Conversion% (Count)', 
                     title='RBM Count Conversion - June',
                     template='plotly_white',
                     color_discrete_sequence=['#3730a3'])
    fig_rbm.update_layout(
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        font=dict(family="Poppins, Inter, sans-serif", size=12, color="#1f2937"),
        xaxis=dict(showgrid=False, tickangle=45),
        yaxis=dict(showgrid=True, gridcolor='#e5e7eb')
    )
    st.plotly_chart(fig_rbm, use_container_width=True)

    # Category Performance
    st.markdown(f'<h3 class="subheader">📦 {category_column} Performance</h3>', unsafe_allow_html=True)

    category_summary = filtered_df.groupby(category_column).agg({
        'TotalSoldPrice': 'sum',
        'WarrantyPrice': 'sum',
        'TotalCount': 'sum',
        'WarrantyCount': 'sum'
    }).reset_index()

    category_summary['Conversion% (Count)'] = (category_summary['WarrantyCount'] / category_summary['TotalCount'] * 100).round(2)
    category_summary['Conversion% (Price)'] = (category_summary['WarrantyPrice'] / category_summary['TotalSoldPrice'] * 100).round(2)
    category_summary['AHSP'] = (category_summary['WarrantyPrice'] / category_summary['WarrantyCount']).where(category_summary['WarrantyCount'] > 0, 0).round(2)

    if not category_summary.empty:
        category_display = category_summary[[category_column, 'Conversion% (Count)', 'Conversion% (Price)', 'AHSP', 'WarrantyPrice', 'WarrantyCount']]
        category_display.columns = [category_column, 'Count Conv (%)', 'Value Conv (%)', 'AHSP (₹)', 'Warranty Sales (₹)', 'Warranty Units']
        category_display = category_display.sort_values('Count Conv (%)', ascending=False)

        st.dataframe(category_display.style.format({
            'Count Conv (%)': '{:.2f}%',
            'Value Conv (%)': '{:.2f}%',
            'AHSP (₹)': '₹{:.2f}',
            'Warranty Sales (₹)': '₹{:.0f}',
            'Warranty Units': '{:.0f}'
        }), use_container_width=True)

        fig_category = px.bar(category_summary, 
                              x=category_column, 
                              y='Conversion% (Count)', 
                              title=f'{category_column} Count Conversion - June',
                              template='plotly_white',
                              color_discrete_sequence=['#3730a3'])
        fig_category.update_layout(
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)',
            font=dict(family="Poppins, Inter, sans-serif", size=12, color="#1f2937"),
            xaxis=dict(showgrid=False, tickangle=45),
            yaxis=dict(showgrid=True, gridcolor='#e5e7eb')
        )
        st.plotly_chart(fig_category, use_container_width=True)
    else:
        st.warning(f"No {category_column.lower()} data available with current filters.")

    # Rest of the code remains the same (Insights section, etc.)
    # ... [Rest of the code remains unchanged]

else:
    st.warning("Please upload data to use the dashboard.")
