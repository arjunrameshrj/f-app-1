import streamlit as st
import pandas as pd
import plotly.express as px
import os
import hashlib
from datetime import datetime

# Streamlit page configuration
st.set_page_config(page_title="Warranty Conversion Dashboard", layout="wide")

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
            # Default admin credentials: username=admin, password=admin123
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

# Title
st.title("📊 Warranty Conversion Analysis Dashboard")

# Session state for authentication
if 'authenticated' not in st.session_state:
    st.session_state.authenticated = False
if 'show_login' not in st.session_state:
    st.session_state.show_login = False

# Login form
if not st.session_state.authenticated:
    if st.button("Login as Admin"):
        st.session_state.show_login = True

    if st.session_state.show_login:
        with st.form("login_form"):
            username = st.text_input("Username")
            password = st.text_input("Password", type="password")
            submit = st.form_submit_button("Login")
            if submit:
                if authenticate(username, password):
                    st.session_state.authenticated = True
                    st.session_state.show_login = False
                    st.success("Logged in successfully!")
                else:
                    st.error("Invalid username or password.")
else:
    st.button("Logout", on_click=lambda: st.session_state.update(authenticated=False))

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
            st.warning(f"Missing or invalid values in {month} file numeric columns. Filling with 0.")
            df[numeric_cols] = df[numeric_cols].fillna(0)
        df['Month'] = month
        df['Conversion% (Count)'] = (df['WarrantyCount'] / df['TotalCount'] * 100).round(2)
        df['Conversion% (Price)'] = (df['WarrantyPrice'] / df['TotalSoldPrice'] * 100).where(df['TotalSoldPrice'] > 0, 0).round(2)
        # Save uploaded file as CSV if file is provided and not a file path
        if file_path and not isinstance(file, str):
            if file.name.endswith('.csv'):
                with open(file_path, 'wb') as f:
                    f.write(file.getvalue())
            else:
                df.to_csv(file_path, index=False)
            st.success(f"{month} data saved successfully.")
        return df
    except Exception as e:
        st.error(f"Error loading {month} file: {str(e)}")
        return None

# File uploaders (only visible to authenticated users)
may_df = None
june_df = None

if st.session_state.authenticated:
    st.subheader("Upload Data Files")
    col1, col2 = st.columns(2)
    with col1:
        may_file = st.file_uploader("Upload May data file", type=["csv", "xlsx", "xls"], key="may")
    with col2:
        june_file = st.file_uploader("Upload June data file", type=["csv", "xlsx", "xls"], key="june")

    # Handle new uploads
    if may_file:
        may_df = load_data(may_file, "May", MAY_FILE_PATH)
    if june_file:
        june_df = load_data(june_file, "June", JUNE_FILE_PATH)

# Load from saved files if no new uploads
if may_df is None and os.path.exists(MAY_FILE_PATH):
    may_df = load_data(MAY_FILE_PATH, "May")
    if may_df is not None:
        st.info("Loaded saved May data.")
if june_df is None and os.path.exists(JUNE_FILE_PATH):
    june_df = load_data(JUNE_FILE_PATH, "June")
    if june_df is not None:
        st.info("Loaded saved June data.")

# If no uploaded or saved files, use sample data
if may_df is None and june_df is None:
    data = {
        'Item Category': ['Electronics', 'Appliances'] * 12,
        'BDM': ['BDM1'] * 24,
        'RBM': ['MAHESH'] * 12 + ['RENJITH'] * 12,
        'Store': ['Palakkad FUTURE', 'Store B'] * 6 + ['Kannur FUTURE', 'Store C'] * 6,
        'Staff Name': ['Staff1', 'Staff2'] * 12,
        'TotalSoldPrice': [48239177/2, 48239177/2, 48239177/2, 48239177/2, 1200000, 1300000] * 4,
        'WarrantyPrice': [300619/2, 300619/2, 300619/2, 300619/2, 6420, 6955] * 4,
        'TotalCount': [5286, 5286, 5286, 5286, 1200, 1300] * 4,
        'WarrantyCount': [483, 483, 483, 483, 60, 65] * 4
    }
    may_df = pd.DataFrame(data).copy()
    may_df['Month'] = "May"
    may_df['Conversion% (Count)'] = (may_df['WarrantyCount'] / may_df['TotalCount'] * 100).round(2)
    may_df['Conversion% (Price)'] = (may_df['WarrantyPrice'] / may_df['TotalSoldPrice'] * 100).round(2)
    
    june_data = {
        'Item Category': ['Electronics', 'Appliances'] * 12,
        'BDM': ['BDM1'] * 24,
        'RBM': ['MAHESH'] * 12 + ['RENJITH'] * 12,
        'Store': ['Palakkad FUTURE', 'Store B'] * 6 + ['Kannur FUTURE', 'Store C'] * 6,
        'Staff Name': ['Staff1', 'Staff2'] * 12,
        'TotalSoldPrice': [51435759/2, 51435759/2, 51435759/2, 51435759/2, 1260000, 1365000] * 4,
        'WarrantyPrice': [175000, 175000, 150000, 150000, 6000, 6500] * 2 + [175000, 175000, 150000, 150000, 6000, 6500] * 2,
        'TotalCount': [5450, 5450, 5450, 5450, 1236, 1339] * 4,
        'WarrantyCount': [448, 448, 448, 448, 90, 98] * 2 + [495, 495, 495, 495, 112, 121] * 2
    }
    june_df = pd.DataFrame(june_data).copy()
    june_df['Month'] = "June"
    june_df['Conversion% (Count)'] = (june_df['WarrantyCount'] / june_df['TotalCount'] * 100).round(2)
    june_df['Conversion% (Price)'] = (june_df['WarrantyPrice'] / june_df['TotalSoldPrice'] * 100).round(2)
    st.warning("Using sample data for May and June, calibrated for MAHESH (May: 9.15% Count, 1.07% Value; June: 8.23% Count, 1.26% Value) and RENJITH (May: 5.00% Count; June: 9.06% Count).")

# Combine datasets
if may_df is not None and june_df is not None:
    df = pd.concat([may_df, june_df], ignore_index=True)
elif may_df is not None:
    df = may_df
elif june_df is not None:
    df = june_df
else:
    st.error("At least one file must be uploaded or sample data must be valid.")
    st.stop()

# Ensure Month column is categorical
df['Month'] = pd.Categorical(df['Month'], categories=['May', 'June'], ordered=True)

# Sidebar filters
st.sidebar.header("🔍 Filters")
bdm_options = ['All'] + sorted(df['BDM'].unique().tolist())
rbm_options = ['All'] + sorted(df['RBM'].unique().tolist())
store_options = ['All'] + sorted(df['Store'].unique().tolist())
category_options = ['All'] + sorted(df['Item Category'].unique().tolist())
staff_options = ['All'] + sorted(df['Staff Name'].unique().tolist())

selected_bdm = st.sidebar.selectbox("BDM", bdm_options, index=0)
selected_rbm = st.sidebar.selectbox("RBM", rbm_options, index=0)
selected_store = st.sidebar.selectbox("Store", store_options, index=0)
selected_category = st.sidebar.selectbox("Item Category", category_options, index=0)
selected_staff = st.sidebar.selectbox("Staff", staff_options, index=0)
future_filter = st.sidebar.checkbox("Show only FUTURE stores")
decreased_rbm_filter = st.sidebar.checkbox("Show only RBMs with decreased count conversion")

# Apply filters
filtered_df = df.copy()
if selected_bdm != 'All':
    filtered_df = filtered_df[filtered_df['BDM'] == selected_bdm]
if selected_rbm != 'All':
    filtered_df = filtered_df[filtered_df['RBM'] == selected_rbm]
if selected_store != 'All':
    filtered_df = filtered_df[filtered_df['Store'] == selected_store]
if selected_category != 'All':
    filtered_df = filtered_df[filtered_df['Item Category'] == selected_category]
if selected_staff != 'All':
    filtered_df = filtered_df[filtered_df['Staff Name'] == selected_staff]
if future_filter:
    filtered_df = filtered_df[filtered_df['Store'].str.contains('FUTURE', case=True)]

if filtered_df.empty:
    st.warning("No data matches your filters. Adjust your selection.")
    st.stop()

# Debug: Display MAHESH or RENJITH data
if selected_rbm in ['MAHESH', 'RENJITH']:
    st.subheader(f"Debug: Raw Data for RBM {selected_rbm}")
    rbm_data = filtered_df[filtered_df['RBM'] == selected_rbm][['Month', 'TotalSoldPrice', 'WarrantyPrice', 'TotalCount', 'WarrantyCount', 'Conversion% (Count)', 'Conversion% (Price)']]
    st.write(rbm_data)
    rbm_summary = rbm_data.groupby('Month').agg({
        'TotalSoldPrice': 'sum',
        'WarrantyPrice': 'sum',
        'TotalCount': 'sum',
        'WarrantyCount': 'sum'
    }).reset_index()
    rbm_summary['Conversion% (Count)'] = (rbm_summary['WarrantyCount'] / rbm_summary['TotalCount'] * 100).round(2)
    rbm_summary['Conversion% (Price)'] = (rbm_summary['WarrantyPrice'] / rbm_summary['TotalSoldPrice'] * 100).round(2)
    st.write(f"Aggregated Metrics for {selected_rbm}:")
    st.write(rbm_summary)

# Main dashboard
st.header("📈 Performance Comparison: May vs June")

# KPI comparison
st.subheader("Overall KPIs")
col1, col2, col3, col4 = st.columns(4)

may_data = filtered_df[filtered_df['Month'] == 'May']
june_data = filtered_df[filtered_df['Month'] == 'June']

may_total_sales = may_data['TotalSoldPrice'].sum()
may_total_warranty = may_data['WarrantyPrice'].sum()
may_total_units = may_data['TotalCount'].sum()
may_total_warranty_units = may_data['WarrantyCount'].sum()
may_value_conversion = (may_total_warranty / may_total_sales * 100) if may_total_sales > 0 else 0
may_count_conversion = (may_total_warranty_units / may_total_units * 100) if may_total_units > 0 else 0

june_total_sales = june_data['TotalSoldPrice'].sum()
june_total_warranty = june_data['WarrantyPrice'].sum()
june_total_units = june_data['TotalCount'].sum()
june_total_warranty_units = june_data['WarrantyCount'].sum()
june_value_conversion = (june_total_warranty / june_total_sales * 100) if june_total_sales > 0 else 0
june_count_conversion = (june_total_warranty_units / june_total_units * 100) if june_total_units > 0 else 0

with col1:
    st.metric("Total Sales (May)", f"₹{may_total_sales:,.0f}")
    st.metric("Total Sales (June)", f"₹{june_total_sales:,.0f}", delta=f"{((june_total_sales - may_total_sales) / may_total_sales * 100):.2f}%" if may_total_sales > 0 else "N/A")
with col2:
    st.metric("Warranty Sales (May)", f"₹{may_total_warranty:,.0f}")
    st.metric("Warranty Sales (June)", f"₹{june_total_warranty:,.0f}", delta=f"{((june_total_warranty - may_total_warranty) / may_total_warranty * 100):.2f}%" if may_total_warranty > 0 else "N/A")
with col3:
    st.metric("Count Conversion (May)", f"{may_count_conversion:.2f}%")
    st.metric("Count Conversion (June)", f"{june_count_conversion:.2f}%", delta=f"{june_count_conversion - may_count_conversion:.2f}%")
with col4:
    st.metric("Value Conversion (May)", f"{may_value_conversion:.2f}%")
    st.metric("Value Conversion (June)", f"{june_value_conversion:.2f}%", delta=f"{june_value_conversion - may_value_conversion:.2f}%")

# Store Performance Comparison
st.subheader("🏬 Store Performance Comparison")
store_summary = filtered_df.groupby(['Store', 'Month']).agg({
    'TotalSoldPrice': 'sum',
    'WarrantyPrice': 'sum',
    'TotalCount': 'sum',
    'WarrantyCount': 'sum'
}).reset_index()
store_summary['Conversion% (Count)'] = (store_summary['WarrantyCount'] / store_summary['TotalCount'] * 100).round(2)
store_summary['Conversion% (Price)'] = (store_summary['WarrantyPrice'] / store_summary['TotalSoldPrice'] * 100).round(2)

store_conv_pivot = store_summary.pivot_table(index='Store', columns='Month', values=['Conversion% (Count)', 'Conversion% (Price)'], aggfunc='first').fillna(0)
store_conv_pivot.columns = [f"{col[1]} {col[0]}" for col in store_conv_pivot.columns]
for col in ['May Conversion% (Count)', 'June Conversion% (Count)', 'May Conversion% (Price)', 'June Conversion% (Price)']:
    if col not in store_conv_pivot.columns:
        store_conv_pivot[col] = 0
store_conv_pivot['Count Conversion Change (%)'] = (store_conv_pivot['June Conversion% (Count)'] - store_conv_pivot['May Conversion% (Count)']).round(2)
store_conv_pivot['Value Conversion Change (%)'] = (store_conv_pivot['June Conversion% (Price)'] - store_conv_pivot['May Conversion% (Price)']).round(2)
store_conv_pivot = store_conv_pivot.sort_values('Count Conversion Change (%)', ascending=False)

store_display = store_conv_pivot.reset_index()
store_display = store_display[['Store', 'May Conversion% (Count)', 'June Conversion% (Count)', 'May Conversion% (Price)', 'June Conversion% (Price)', 'Count Conversion Change (%)', 'Value Conversion Change (%)']]
store_display.columns = ['Store', 'May Count Conv (%)', 'June Count Conv (%)', 'May Value Conv (%)', 'June Value Conv (%)', 'Count Conv Change (%)', 'Value Conv Change (%)']
store_display = store_display.fillna(0)

def highlight_change(val):
    color = 'green' if val > 0 else 'red' if val < 0 else 'black'
    return f'color: {color}'

st.dataframe(store_display.style.format({
    'May Count Conv (%)': '{:.2f}%',
    'June Count Conv (%)': '{:.2f}%',
    'May Value Conv (%)': '{:.2f}%',
    'June Value Conv (%)': '{:.2f}%',
    'Count Conv Change (%)': '{:.2f}%',
    'Value Conv Change (%)': '{:.2f}%'
}).applymap(highlight_change, subset=['Count Conv Change (%)', 'Value Conv Change (%)']), use_container_width=True)

fig_store = px.bar(store_summary, 
                   x='Store', 
                   y='Conversion% (Count)', 
                   color='Month', 
                   barmode='group', 
                   title='Store Count Conversion: May vs June')
st.plotly_chart(fig_store, use_container_width=True)

# RBM Performance Comparison
st.subheader("👥 RBM Performance Comparison")
rbm_summary = filtered_df.groupby(['RBM', 'Month']).agg({
    'TotalSoldPrice': 'sum',
    'WarrantyPrice': 'sum',
    'TotalCount': 'sum',
    'WarrantyCount': 'sum'
}).reset_index()
rbm_summary['Conversion% (Count)'] = (rbm_summary['WarrantyCount'] / rbm_summary['TotalCount'] * 100).round(2)
rbm_summary['Conversion% (Price)'] = (rbm_summary['WarrantyPrice'] / rbm_summary['TotalSoldPrice'] * 100).round(2)

rbm_pivot = rbm_summary.pivot_table(index='RBM', columns='Month', values=['Conversion% (Count)', 'Conversion% (Price)'], aggfunc='first').fillna(0)
rbm_pivot.columns = [f"{col[1]} {col[0]}" for col in rbm_pivot.columns]
for col in ['May Conversion% (Count)', 'June Conversion% (Count)', 'May Conversion% (Price)', 'June Conversion% (Price)']:
    if col not in rbm_pivot.columns:
        rbm_pivot[col] = 0
rbm_pivot['Count Conversion Change (%)'] = (rbm_pivot['June Conversion% (Count)'] - rbm_pivot['May Conversion% (Count)']).round(2)
rbm_pivot['Value Conversion Change (%)'] = (rbm_pivot['June Conversion% (Price)'] - rbm_pivot['May Conversion% (Price)']).round(2)
rbm_pivot = rbm_pivot.sort_values('Count Conversion Change (%)', ascending=False)

# Apply decreased RBM filter
if decreased_rbm_filter:
    rbm_pivot = rbm_pivot[rbm_pivot['Count Conversion Change (%)'] < 0]
    if rbm_pivot.empty:
        st.warning("No RBMs with decreased count conversion match the filters.")

rbm_display = rbm_pivot.reset_index()
rbm_display = rbm_display[['RBM', 'May Conversion% (Count)', 'June Conversion% (Count)', 'May Conversion% (Price)', 'June Conversion% (Price)', 'Count Conversion Change (%)', 'Value Conversion Change (%)']]
rbm_display.columns = ['RBM', 'May Count Conv (%)', 'June Count Conv (%)', 'May Value Conv (%)', 'June Value Conv (%)', 'Count Conv Change (%)', 'Value Conv Change (%)']
rbm_display = rbm_display.fillna(0)

st.dataframe(rbm_display.style.format({
    'May Count Conv (%)': '{:.2f}%',
    'June Count Conv (%)': '{:.2f}%',
    'May Value Conv (%)': '{:.2f}%',
    'June Value Conv (%)': '{:.2f}%',
    'Count Conv Change (%)': '{:.2f}%',
    'Value Conv Change (%)': '{:.2f}%'
}).applymap(highlight_change, subset=['Count Conv Change (%)', 'Value Conv Change (%)']), use_container_width=True)

fig_rbm = px.bar(rbm_summary, 
                 x='RBM', 
                 y='Conversion% (Count)', 
                 color='Month', 
                 barmode='group', 
                 title='RBM Count Conversion: May vs June')
st.plotly_chart(fig_rbm, use_container_width=True)

# Item Category Performance Comparison
st.subheader("📦 Item Category Performance Comparison")
category_summary = filtered_df.groupby(['Item Category', 'Month']).agg({
    'TotalSoldPrice': 'sum',
    'WarrantyPrice': 'sum',
    'TotalCount': 'sum',
    'WarrantyCount': 'sum'
}).reset_index()
category_summary['Conversion% (Count)'] = (category_summary['WarrantyCount'] / category_summary['TotalCount'] * 100).round(2)
category_summary['Conversion% (Price)'] = (category_summary['WarrantyPrice'] / category_summary['TotalSoldPrice'] * 100).round(2)

if not category_summary.empty:
    category_pivot = category_summary.pivot_table(index='Item Category', columns='Month', values=['Conversion% (Count)', 'Conversion% (Price)'], aggfunc='first').fillna(0)
    category_pivot.columns = [f"{col[1]} {col[0]}" for col in category_pivot.columns]
    for col in ['May Conversion% (Count)', 'June Conversion% (Count)', 'May Conversion% (Price)', 'June Conversion% (Price)']:
        if col not in category_pivot.columns:
            category_pivot[col] = 0
    category_pivot['Count Conversion Change (%)'] = (category_pivot['June Conversion% (Count)'] - category_pivot['May Conversion% (Count)']).round(2)
    category_pivot['Value Conversion Change (%)'] = (category_pivot['June Conversion% (Price)'] - category_pivot['May Conversion% (Price)']).round(2)
    category_pivot = category_pivot.sort_values('Count Conversion Change (%)', ascending=False)

    category_display = category_pivot.reset_index()
    category_display = category_display[['Item Category', 'May Conversion% (Count)', 'June Conversion% (Count)', 'May Conversion% (Price)', 'June Conversion% (Price)', 'Count Conversion Change (%)', 'Value Conversion Change (%)']]
    category_display.columns = ['Item Category', 'May Count Conv (%)', 'June Count Conv (%)', 'May Value Conv (%)', 'June Value Conv (%)', 'Count Conv Change (%)', 'Value Conv Change (%)']
    category_display = category_display.fillna(0)

    st.dataframe(category_display.style.format({
        'May Count Conv (%)': '{:.2f}%',
        'June Count Conv (%)': '{:.2f}%',
        'May Value Conv (%)': '{:.2f}%',
        'June Value Conv (%)': '{:.2f}%',
        'Count Conv Change (%)': '{:.2f}%',
        'Value Conv Change (%)': '{:.2f}%'
    }).applymap(highlight_change, subset=['Count Conv Change (%)', 'Value Conv Change (%)']), use_container_width=True)

    fig_category = px.bar(category_summary, 
                          x='Item Category', 
                          y='Conversion% (Count)', 
                          color='Month', 
                          barmode='group', 
                          title='Item Category Count Conversion: May vs June')
    st.plotly_chart(fig_category, use_container_width=True)
else:
    st.warning("No item category data available with current filters.")
    category_pivot = pd.DataFrame()

# Insights
st.header("💡 Insights & Recommendations")

# Overall performance change
st.subheader("Overall Performance Change")
if june_count_conversion > may_count_conversion:
    st.success(f"Count Conversion improved from {may_count_conversion:.2f}% in May to {june_count_conversion:.2f}% in June.")
else:
    st.warning(f"Count Conversion declined from {may_count_conversion:.2f}% in May to {june_count_conversion:.2f}% in June.")
if june_value_conversion > may_value_conversion:
    st.success(f"Value Conversion improved from {may_value_conversion:.2f}% in May to {june_value_conversion:.2f}% in June.")
else:
    st.warning(f"Value Conversion declined from {may_value_conversion:.2f}% in May to {june_value_conversion:.2f}% in June.")

# Store-Level Warranty Sales Analysis
st.subheader("📉 Store-Level Warranty Sales Analysis")
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

    # Reasons for decreases
    st.write("**Reasons for Warranty Sales Decreases:**")
    decreased_stores = store_pivot[store_pivot['Change (%)'] < 0]
    if not decreased_stores.empty:
        category_warranty = filtered_df.groupby(['Store', 'Month', 'Item Category'])['WarrantyPrice'].sum().reset_index()
        category_pivot_warranty = category_warranty.pivot_table(index=['Store', 'Item Category'], columns='Month', values='WarrantyPrice', aggfunc='first').fillna(0)
        category_pivot_warranty['Change (₹)'] = category_pivot_warranty['June'] - category_pivot_warranty['May']
        
        store_metrics = filtered_df.groupby(['Store', 'Month']).agg({
            'TotalSoldPrice': 'sum',
            'WarrantyCount': 'sum',
            'TotalCount': 'sum'
        }).reset_index()
        metrics_pivot = store_metrics.pivot_table(index='Store', columns='Month', values=['TotalSoldPrice', 'WarrantyCount', 'TotalCount'], aggfunc='first').fillna(0)
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
st.subheader("Significant Changes")
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
        st.write("**Item Categories with Significant Count Conversion Changes:**")
        for category in significant_categories.index:
            change = float(category_pivot.loc[category, 'Count Conversion Change (%)'])
            st.write(f"- {category}: {change:.2f}% {'increase' if change > 0 else 'decrease'}")
else:
    st.info("No significant item category changes with current filters.")

# Low performers in June
avg_count_conversion = june_count_conversion
low_performers = store_summary[store_summary['Month'] == 'June']
low_performers = low_performers[low_performers['Conversion% (Count)'] < avg_count_conversion]
if not low_performers.empty:
    st.subheader("🚨 Improvement Opportunities (June)")
    st.write(f"These stores in June have below-average count conversion (avg: {avg_count_conversion:.2f}%):")
    for _, row in low_performers.iterrows():
        st.write(f"- {row['Store']}: {row['Conversion% (Count)']:.2f}%")
    
    st.write("**Recommendations:**")
    st.write("1. Provide additional training on warranty benefits")
    st.write("2. Create targeted promotions")
    st.write("3. Review staffing and sales strategies")

# FUTURE stores analysis
if future_filter or any('FUTURE' in store for store in filtered_df['Store'].unique()):
    st.subheader("🏢 FUTURE Stores Analysis")
    future_stores = filtered_df[filtered_df['Store'].str.contains('FUTURE', case=True)]
    if not future_stores.empty:
        future_summary = future_stores.groupby('Month').agg({
            'TotalSoldPrice': 'sum',
            'WarrantyPrice': 'sum',
            'TotalCount': 'sum',
            'WarrantyCount': 'sum'
        }).reset_index()
        future_summary['Conversion% (Count)'] = (future_summary['WarrantyCount'] / future_summary['TotalCount'] * 100).round(2)
        future_summary['Conversion% (Price)'] = (future_summary['WarrantyPrice'] / future_summary['TotalSoldPrice'] * 100).round(2)
        
        may_future = future_summary[future_summary['Month'] == 'May']
        june_future = future_summary[future_summary['Month'] == 'June']
        
        may_future_count = may_future['Conversion% (Count)'].iloc[0] if not may_future.empty else 0
        june_future_count = june_future['Conversion% (Count)'].iloc[0] if not june_future.empty else 0
        st.write(f"Average count conversion in FUTURE stores (May): {may_future_count:.2f}%")
        st.write(f"Average count conversion in FUTURE stores (June): {june_future_count:.2f}%")
        if june_future_count < may_future_count:
            st.warning("FUTURE stores count conversion declined in June.")
            st.write("**Recommendations:**")
            st.write("- Conduct store-specific training")
            st.write("- Analyze customer demographics")
        else:
            st.success("FUTURE stores count conversion improved or remained stable in June!")
    else:
        st.info("No FUTURE stores in current selection")

# Correlation Analysis
st.subheader("🔗 Correlation Analysis (June)")
corr_matrix = filtered_df[filtered_df['Month'] == 'June'][['TotalSoldPrice', 'WarrantyPrice', 'TotalCount', 'WarrantyCount']].corr()
fig_corr = px.imshow(
    corr_matrix,
    text_auto=True,
    title="Correlation Matrix of Key Metrics (June)",
    template='plotly_dark'
)
st.plotly_chart(fig_corr, use_container_width=True)
