import streamlit as st
import requests
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta, date
import pytz
import time
import json
import io
import numpy as np
from collections import Counter
import re
import xlsxwriter
from io import BytesIO

# Set page config
st.set_page_config(
    page_title="HubSpot Business Analytics",
    page_icon="📊",
    layout="wide"
)

# UI Configuration Constants
TOP_N = 10  # Limit charts to top N items
MAX_LABEL_LENGTH = 25  # Truncate long labels
CHART_HEIGHT = 420
COLOR_PALETTE = px.colors.qualitative.Set2

# CUSTOMER DEAL STAGE CONFIGURATION
CUSTOMER_DEAL_STAGES = ["admission_confirmed", "closedwon", "won"]  # Add all stages that indicate a customer

# Custom CSS for better styling
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        color: #1a1a1a;
        margin-bottom: 1rem;
    }
    .sub-header {
        font-size: 1.5rem;
        color: #666;
        margin-bottom: 1rem;
    }
    .stats-card {
        background-color: #f8f9fa;
        padding: 1rem;
        border-radius: 0.5rem;
        border-left: 4px solid #ff7a59;
        margin-bottom: 1rem;
    }
    .metric-card {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 1.5rem;
        border-radius: 0.5rem;
        color: white;
        text-align: center;
    }
    .header-container {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 2rem;
        border-radius: 0.5rem;
        color: white;
        margin-bottom: 2rem;
    }
    .warning-box {
        background-color: #fff3cd;
        color: #856404;
        padding: 1rem;
        border-radius: 0.5rem;
        margin: 1rem 0;
        border: 1px solid #ffeaa7;
    }
    .quality-high-red {
        background-color: #f8d7da !important;
        color: #721c24 !important;
        font-weight: bold;
    }
    .quality-high-green {
        background-color: #d4edda !important;
        color: #155724 !important;
        font-weight: bold;
    }
    .section-header {
        background: linear-gradient(90deg, #f8f9fa 0%, #e9ecef 100%);
        padding: 1rem;
        border-radius: 0.5rem;
        margin: 1.5rem 0 1rem 0;
        border-left: 4px solid #4dabf7;
    }
    .kpi-card {
        background: white;
        padding: 1rem;
        border-radius: 0.5rem;
        border: 1px solid #dee2e6;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
        text-align: center;
    }
    .comparison-card {
        background: linear-gradient(135deg, #fdfcfb 0%, #e2d1c3 100%);
        padding: 1.5rem;
        border-radius: 0.5rem;
        border: 1px solid #ddd;
        margin-bottom: 1rem;
    }
    .unassigned-badge {
        background-color: #ffeaa7;
        color: #856404;
        padding: 0.2rem 0.5rem;
        border-radius: 0.2rem;
        font-size: 0.8rem;
        font-weight: bold;
    }
    
    /* ✅ NEW: Fixed-size KPI Cards for TV/Executive Dashboards */
    .kpi-container {
        display: flex;
        justify-content: center;
        align-items: stretch;
        gap: 20px;
        margin: 20px 0;
        flex-wrap: wrap;
    }
    
    .kpi-box {
        min-width: 220px !important;
        max-width: 220px !important;
        height: 130px !important;
        background: linear-gradient(135deg, #667eea, #764ba2);
        border-radius: 12px;
        padding: 16px;
        text-align: center;
        color: white;
        display: flex;
        flex-direction: column;
        justify-content: center;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.1);
        transition: transform 0.2s ease;
    }
    
    .kpi-box:hover {
        transform: translateY(-3px);
        box-shadow: 0 6px 16px rgba(0, 0, 0, 0.15);
    }
    
    .kpi-title {
        font-size: 14px;
        opacity: 0.9;
        margin-bottom: 8px;
        font-weight: 500;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    
    .kpi-value {
        font-size: 34px !important;
        font-weight: 700 !important;
        line-height: 1.2 !important;
        margin: 5px 0;
    }
    
    .kpi-sub {
        font-size: 13px;
        opacity: 0.85;
        margin-top: 5px;
    }
    
    /* Alternative color schemes */
    .kpi-box-blue {
        background: linear-gradient(135deg, #4A6FA5, #166088);
    }
    
    .kpi-box-green {
        background: linear-gradient(135deg, #2E8B57, #3CB371);
    }
    
    .kpi-box-orange {
        background: linear-gradient(135deg, #FF7A59, #FFA500);
    }
    
    .kpi-box-purple {
        background: linear-gradient(135deg, #8A2BE2, #9370DB);
    }
    
    .kpi-box-teal {
        background: linear-gradient(135deg, #20B2AA, #48D1CC);
    }
    
    .kpi-box-red {
        background: linear-gradient(135deg, #DC143C, #FF6347);
    }
    
    .kpi-box-yellow {
        background: linear-gradient(135deg, #FFD700, #FFA500);
    }
    
    /* Secondary KPI Cards (smaller) */
    .secondary-kpi-container {
        display: flex;
        justify-content: center;
        gap: 15px;
        margin: 15px 0;
        flex-wrap: wrap;
    }
    
    .secondary-kpi-box {
        min-width: 150px !important;
        max-width: 150px !important;
        height: 100px !important;
        background: linear-gradient(135deg, #f8f9fa, #e9ecef);
        border-radius: 8px;
        padding: 12px;
        text-align: center;
        color: #333;
        display: flex;
        flex-direction: column;
        justify-content: center;
        border: 1px solid #dee2e6;
    }
    
    .secondary-kpi-title {
        font-size: 12px;
        opacity: 0.8;
        margin-bottom: 5px;
        font-weight: 500;
    }
    
    .secondary-kpi-value {
        font-size: 24px !important;
        font-weight: 700 !important;
        line-height: 1.2 !important;
        margin: 3px 0;
    }
    
    .secondary-kpi-sub {
        font-size: 11px;
        opacity: 0.7;
        margin-top: 3px;
    }
    
    /* Download Button Styles */
    .download-btn {
        background: linear-gradient(135deg, #28a745, #20c997);
        color: white;
        border: none;
        padding: 0.75rem 1.5rem;
        border-radius: 0.5rem;
        font-weight: 600;
        cursor: pointer;
        transition: all 0.3s ease;
        display: inline-flex;
        align-items: center;
        justify-content: center;
        gap: 0.5rem;
        text-decoration: none;
    }
    
    .download-btn:hover {
        transform: translateY(-2px);
        box-shadow: 0 4px 12px rgba(40, 167, 69, 0.3);
        color: white;
        text-decoration: none;
    }
    
    .download-btn-excel {
        background: linear-gradient(135deg, #217346, #1e7e34);
    }
    
    .download-btn-csv {
        background: linear-gradient(135deg, #007bff, #0056b3);
    }
    
    .download-btn-all {
        background: linear-gradient(135deg, #6f42c1, #6610f2);
    }
    
    .download-section {
        background: linear-gradient(135deg, #f8f9fa, #e9ecef);
        padding: 1.5rem;
        border-radius: 0.75rem;
        border: 1px solid #dee2e6;
        margin: 1.5rem 0;
    }
    
    .download-header {
        color: #2c3e50;
        font-size: 1.25rem;
        margin-bottom: 1rem;
        display: flex;
        align-items: center;
        gap: 0.5rem;
    }
    
    /* ✅ NEW: Revenue and Matrix Styles */
    .revenue-card {
        background: linear-gradient(135deg, #4CAF50, #8BC34A);
        color: white;
        padding: 1.5rem;
        border-radius: 10px;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.1);
        text-align: center;
    }
    
    .matrix-star {
        background-color: #d4edda !important;
        color: #155724 !important;
        font-weight: bold;
    }
    
    .matrix-potential {
        background-color: #cce5ff !important;
        color: #004085 !important;
        font-weight: bold;
    }
    
    .matrix-burn {
        background-color: #fff3cd !important;
        color: #856404 !important;
        font-weight: bold;
    }
    
    .matrix-weak {
        background-color: #f8d7da !important;
        color: #721c24 !important;
        font-weight: bold;
    }
    
    .revenue-kpi {
        background: linear-gradient(135deg, #FFD700, #FFA500);
    }
    
    .conversion-kpi {
        background: linear-gradient(135deg, #20B2AA, #48D1CC);
    }
    
    .deal-kpi {
        background: linear-gradient(135deg, #9370DB, #8A2BE2);
    }
    
    .warning-card {
        background: linear-gradient(135deg, #fff3cd, #ffeaa7);
        border: 1px solid #ffc107;
        border-radius: 8px;
        padding: 15px;
        margin: 10px 0;
    }
</style>
""", unsafe_allow_html=True)

# Constants
HUBSPOT_API_BASE = "https://api.hubapi.com"
IST = pytz.timezone('Asia/Kolkata')

# ✅ SECURITY: Load API key from Streamlit secrets
def get_api_key():
    """Safely load API key from Streamlit secrets"""
    try:
        # Check if secrets are available
        if "hubspot" in st.secrets and "api_key" in st.secrets["hubspot"]:
            api_key = st.secrets["hubspot"]["api_key"]
            if api_key and api_key.startswith("pat-"):
                return api_key
        
        # If no secrets or invalid, check for direct secret
        if "HUBSPOT_API_KEY" in st.secrets:
            api_key = st.secrets["HUBSPOT_API_KEY"]
            if api_key and api_key.startswith("pat-"):
                return api_key
        
        # Try environment variable as fallback
        import os
        api_key = os.getenv("HUBSPOT_API_KEY")
        if api_key and api_key.startswith("pat-"):
            return api_key
        
        # Return None if no valid key found
        return None
        
    except Exception as e:
        st.error(f"Error loading API key: {str(e)}")
        return None

# Lead Status Mapping - NO CUSTOMER HERE!
LEAD_STATUS_MAP = {
    "cold": "Cold",
    "warm": "Warm", 
    "hot": "Hot",
    "new": "New Lead",
    "open": "New Lead",
    "neutral_prospect": "Cold",
    "prospect": "Warm",  
    "hot_prospect": "Hot",
    "not_connected": "Not Connected (NC)",
    "not_interested": "Not Interested", 
    "unqualified": "Not Qualified",
    "not_qualified": "Not Qualified",
    # ❌ NO CUSTOMER HERE - Customer ONLY from DEALS
    "duplicate": "Duplicate",
    "junk": "Duplicate",
    "": "Unknown",
    None: "Unknown",
    "unknown": "Unknown",
    "other": "Unknown"
}

# ✅ NEW: KPI Rendering Functions
def render_kpi(title, value, subtitle="", color_class=""):
    """Render a professional KPI card with fixed size."""
    color_class = color_class or ""
    return f"""
    <div class="kpi-box {color_class}">
        <div class="kpi-title">{title}</div>
        <div class="kpi-value">{value}</div>
        <div class="kpi-sub">{subtitle}</div>
    </div>
    """

def render_secondary_kpi(title, value, subtitle=""):
    """Render a smaller secondary KPI card."""
    return f"""
    <div class="secondary-kpi-box">
        <div class="secondary-kpi-title">{title}</div>
        <div class="secondary-kpi-value">{value}</div>
        <div class="secondary-kpi-sub">{subtitle}</div>
    </div>
    """

def render_kpi_row(kpis, container_class="kpi-container"):
    """Render a row of KPI cards."""
    kpi_html = "".join(kpis)
    return f"""
    <div class="{container_class}">
        {kpi_html}
    </div>
    """

# ✅ NEW: Enhanced Excel Export Function
def create_excel_report(df, metrics, kpis, customers_df, date_range, date_field):
    """Create a professional Excel report with multiple sheets and formatting."""
    
    # Create an in-memory BytesIO object
    output = BytesIO()
    
    # Create Excel writer with XlsxWriter engine
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        
        # Define formats
        header_format = workbook.add_format({
            'bold': True,
            'font_size': 12,
            'bg_color': '#4F81BD',
            'font_color': 'white',
            'align': 'center',
            'valign': 'vcenter',
            'border': 1
        })
        
        subheader_format = workbook.add_format({
            'bold': True,
            'font_size': 11,
            'bg_color': '#DCE6F1',
            'align': 'center',
            'border': 1
        })
        
        number_format = workbook.add_format({
            'num_format': '#,##0',
            'align': 'right',
            'border': 1
        })
        
        percent_format = workbook.add_format({
            'num_format': '0.0%',
            'align': 'right',
            'border': 1
        })
        
        kpi_format = workbook.add_format({
            'bold': True,
            'font_size': 14,
            'align': 'center',
            'valign': 'vcenter'
        })
        
        date_format = workbook.add_format({
            'num_format': 'yyyy-mm-dd',
            'align': 'center',
            'border': 1
        })
        
        # Sheet 1: Executive Summary
        summary_data = {
            'Metric': [
                'Total Leads',
                'Deal Leads (Hot+Warm+Cold)',
                'Customers (from Deals)',
                'Hot Leads',
                'Warm Leads', 
                'Cold Leads',
                'New Leads',
                'Not Connected',
                'Not Interested',
                'Not Qualified',
                'Duplicate',
                'Lead → Deal %',
                'Lead → Customer %',
                'Deal → Customer %',
                'Total Revenue',
                'Avg Revenue per Customer'
            ],
            'Count': [
                kpis['total_leads'],
                kpis['deal_leads'],
                kpis['customer'],
                kpis['hot'],
                kpis['warm'],
                kpis['cold'],
                kpis['new_lead'],
                kpis['not_connected'],
                kpis['not_interested'],
                kpis['not_qualified'],
                kpis['duplicate'],
                kpis['lead_to_deal_pct']/100,
                kpis['lead_to_customer_pct']/100,
                kpis['deal_to_customer_pct']/100,
                kpis['total_revenue'],
                kpis['avg_revenue_per_customer']
            ]
        }
        
        summary_df = pd.DataFrame(summary_data)
        summary_df.to_excel(writer, sheet_name='Executive Summary', index=False)
        
        worksheet = writer.sheets['Executive Summary']
        
        # Apply formatting
        for col_num, value in enumerate(summary_df.columns.values):
            worksheet.write(0, col_num, value, header_format)
        
        # Set column widths
        worksheet.set_column('A:A', 35)
        worksheet.set_column('B:B', 25)
        
        # Add report header
        worksheet.merge_range('D1:F1', 'HubSpot Analytics Report', workbook.add_format({
            'bold': True, 'font_size': 16, 'align': 'center', 'valign': 'vcenter'
        }))
        
        worksheet.merge_range('D2:F2', f'Date Field: {date_field}', workbook.add_format({
            'align': 'center'
        }))
        
        worksheet.merge_range('D3:F3', f'Date Range: {date_range[0]} to {date_range[1]}', workbook.add_format({
            'align': 'center'
        }))
        
        worksheet.merge_range('D4:F4', f'Generated: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}', workbook.add_format({
            'align': 'center'
        }))
        
        # Sheet 2: Raw Lead Data
        if not df.empty:
            df.to_excel(writer, sheet_name='Raw Lead Data', index=False)
            worksheet = writer.sheets['Raw Lead Data']
            
            # Apply formatting
            for col_num, value in enumerate(df.columns.values):
                worksheet.write(0, col_num, value, header_format)
            
            # Auto-adjust column widths
            for i, col in enumerate(df.columns):
                column_len = max(df[col].astype(str).map(len).max(), len(col)) + 2
                worksheet.set_column(i, i, min(column_len, 50))
        
        # Sheet 3: Customer Data (from Deals)
        if customers_df is not None and not customers_df.empty:
            customers_df.to_excel(writer, sheet_name='Customer Data', index=False)
            worksheet = writer.sheets['Customer Data']
            
            for col_num, value in enumerate(customers_df.columns.values):
                worksheet.write(0, col_num, value, header_format)
            
            # Format revenue columns
            if 'Revenue' in customers_df.columns:
                num_rows = len(customers_df)
                for row in range(1, num_rows + 1):
                    worksheet.write(row, customers_df.columns.get_loc('Revenue'), 
                                   customers_df.iloc[row-1, customers_df.columns.get_loc('Revenue')], 
                                   workbook.add_format({'num_format': '₹#,##0', 'align': 'right'}))
        
        # Sheet 4: Course Distribution
        if 'metric_1' in metrics and not metrics['metric_1'].empty:
            metric_1 = metrics['metric_1'].copy()
            metric_1.to_excel(writer, sheet_name='Course Distribution', index=False)
            
            worksheet = writer.sheets['Course Distribution']
            for col_num, value in enumerate(metric_1.columns.values):
                worksheet.write(0, col_num, value, header_format)
        
        # Sheet 5: Owner Performance
        if 'metric_2' in metrics and not metrics['metric_2'].empty:
            metric_2 = metrics['metric_2'].copy()
            metric_2.to_excel(writer, sheet_name='Owner Performance', index=False)
            
            worksheet = writer.sheets['Owner Performance']
            for col_num, value in enumerate(metric_2.columns.values):
                worksheet.write(0, col_num, value, header_format)
        
        # Sheet 6: KPI Dashboard
        if 'metric_4' in metrics and not metrics['metric_4'].empty:
            metric_4 = metrics['metric_4'].copy()
            metric_4.to_excel(writer, sheet_name='KPI Dashboard', index=False)
            
            worksheet = writer.sheets['KPI Dashboard']
            for col_num, value in enumerate(metric_4.columns.values):
                worksheet.write(0, col_num, value, header_format)
        
        # Sheet 7: Lead Status Summary
        if not df.empty:
            status_summary = df['Lead Status'].value_counts().reset_index()
            status_summary.columns = ['Lead Status', 'Count']
            status_summary['Percentage'] = (status_summary['Count'] / len(df) * 100).round(1)
            
            status_summary.to_excel(writer, sheet_name='Lead Status Summary', index=False)
            
            worksheet = writer.sheets['Lead Status Summary']
            for col_num, value in enumerate(status_summary.columns.values):
                worksheet.write(0, col_num, value, header_format)
    
    # Get the Excel data
    excel_data = output.getvalue()
    return excel_data

# Function to validate API key
def validate_api_key(api_key):
    """Validate the HubSpot API key format."""
    if not api_key:
        return False, "❌ API key is empty"
    
    if not api_key.startswith("pat-"):
        return False, "❌ Invalid API key format. Should start with 'pat-'"
    
    if len(api_key) < 20:
        return False, "❌ API key appears too short"
    
    return True, "✅ API key format looks valid"

def test_hubspot_connection(api_key):
    """Test if the HubSpot API key is valid."""
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json"
    }
    url = f"{HUBSPOT_API_BASE}/crm/v3/objects/contacts?limit=1"
    
    try:
        response = requests.get(url, headers=headers, timeout=10)
        
        if response.status_code == 200:
            return True, "✅ Connection successful! API key is valid."
        elif response.status_code == 401:
            error_data = response.json()
            error_message = error_data.get('message', 'Unknown error')
            
            if "Invalid token" in error_message or "expired" in error_message:
                return False, "❌ API key is invalid or expired. Please check your API key."
            elif "scope" in error_message.lower():
                return False, f"❌ Missing required scopes. Error: {error_message}"
            else:
                return False, f"❌ Authentication failed. Status: {response.status_code}"
        else:
            return False, f"❌ Connection failed. Status: {response.status_code}"
    except requests.exceptions.RequestException as e:
        return False, f"❌ Connection error: {str(e)}"

@st.cache_data(ttl=3600)
def fetch_owner_mapping(api_key):
    """Fetch ALL owner ID to name mapping with pagination."""
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json"
    }

    url = f"{HUBSPOT_API_BASE}/crm/v3/owners"
    params = {"limit": 100}
    mapping = {}
    page_count = 0
    total_owners = 0

    try:
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        while True:
            response = requests.get(url, headers=headers, params=params, timeout=30)
            response.raise_for_status()

            data = response.json()
            owners = data.get("results", [])
            
            for owner in owners:
                owner_id = str(owner.get("id", ""))

                first_name = owner.get("firstName", "")
                last_name = owner.get("lastName", "")
                email = owner.get("email", "")

                # Create display name
                if first_name or last_name:
                    name = f"{first_name} {last_name}".strip()
                elif email:
                    name = email.split("@")[0]
                else:
                    name = f"Owner {owner_id}"

                mapping[owner_id] = name

            page_count += 1
            total_owners += len(owners)
            
            # Update progress
            if page_count <= 5:  # Assume max 5 pages
                progress_bar.progress(min(page_count / 5, 0.9))
            status_text.text(f"📋 Fetched {total_owners} owners (page {page_count})...")

            # Check for next page
            paging = data.get("paging", {})
            next_link = paging.get("next", {}).get("link")

            if not next_link:
                progress_bar.progress(1.0)
                status_text.text(f"✅ Owner mapping complete! Total: {total_owners} owners")
                break

            # Update for next page
            url = next_link
            params = None  # Link already has params
            
            # Small delay to avoid rate limiting
            time.sleep(0.1)
            
    except requests.exceptions.RequestException as e:
        st.warning(f"⚠️ Partial owner mapping loaded. Error: {str(e)[:100]}")
    except Exception as e:
        st.warning(f"⚠️ Could not fetch owner mapping: {str(e)[:100]}")

    return mapping

@st.cache_data(ttl=3600)
def fetch_single_owner_name(api_key, owner_id):
    """Fetch a single owner name by ID (for internal user IDs)."""
    if not owner_id:
        return None
    
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json"
    }

    url = f"{HUBSPOT_API_BASE}/crm/v3/owners/{owner_id}"

    try:
        response = requests.get(url, headers=headers, timeout=10)
        if response.status_code == 200:
            owner = response.json()
            first = owner.get("firstName", "")
            last = owner.get("lastName", "")
            email = owner.get("email", "")

            if first or last:
                return f"{first} {last}".strip()
            elif email:
                return email.split("@")[0]
            else:
                return f"Owner {owner_id}"
    except requests.exceptions.RequestException:
        # Silently fail - owner might not exist or API error
        pass
    except Exception:
        # Any other error
        pass

    return None

def date_to_hubspot_timestamp(date_obj, is_end_date=False):
    """Convert date to HubSpot timestamp (milliseconds)."""
    if isinstance(date_obj, str):
        date_obj = datetime.strptime(date_obj, "%Y-%m-%d").date()
    
    if is_end_date:
        dt = datetime.combine(date_obj, datetime.max.time())
    else:
        dt = datetime.combine(date_obj, datetime.min.time())
    
    dt_ist = IST.localize(dt)
    dt_utc = dt_ist.astimezone(pytz.UTC)
    return int(dt_utc.timestamp() * 1000)

def fetch_hubspot_contacts_with_date_filter(api_key, date_field, start_date, end_date):
    """Fetch ALL contacts from HubSpot with server-side date filtering."""
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json"
    }
    
    start_timestamp = date_to_hubspot_timestamp(start_date, is_end_date=False)
    safe_end_date = end_date + timedelta(days=1)
    end_timestamp = date_to_hubspot_timestamp(safe_end_date, is_end_date=False)
    
    all_contacts = []
    after = None
    page_count = 0
    
    # Build filter
    if date_field == "Created Date":
        filter_groups = [{
            "filters": [
                {"propertyName": "createdate", "operator": "GTE", "value": start_timestamp},
                {"propertyName": "createdate", "operator": "LTE", "value": end_timestamp}
            ]
        }]
    elif date_field == "Last Modified Date":
        filter_groups = [{
            "filters": [
                {"propertyName": "lastmodifieddate", "operator": "GTE", "value": start_timestamp},
                {"propertyName": "lastmodifieddate", "operator": "LTE", "value": end_timestamp}
            ]
        }]
    else:  # Both
        filter_groups = [
            {
                "filters": [
                    {"propertyName": "createdate", "operator": "GTE", "value": start_timestamp},
                    {"propertyName": "createdate", "operator": "LTE", "value": end_timestamp}
                ]
            },
            {
                "filters": [
                    {"propertyName": "lastmodifieddate", "operator": "GTE", "value": start_timestamp},
                    {"propertyName": "lastmodifieddate", "operator": "LTE", "value": end_timestamp}
                ]
            }
        ]
    
    url = f"{HUBSPOT_API_BASE}/crm/v3/objects/contacts/search"
    
    progress_bar = st.progress(0)
    status_text = st.empty()
    status_text.text(f"📡 Fetching contacts with {date_field} filter...")
    
    # Define properties needed - REMOVED lifecycle stage dependency
    all_properties = [
        "hs_lead_status", "lead_status", 
        "hubspot_owner_id", "hs_assigned_owner_id",
        "course", "program", "product", "service", "offering",
        "course_name", "program_name", "product_name",
        "enquired_course", "interested_course", "course_interested",
        "program_of_interest", "course_of_interest", "product_of_interest",
        "firstname", "lastname", "email", "phone", 
        "createdate", "lastmodifieddate", "hs_object_id",
        "company", "jobtitle", "country"
    ]
    
    try:
        while True:
            body = {
                "filterGroups": filter_groups,
                "properties": all_properties,
                "associations": ["owners"],
                "limit": 100,
                "sorts": [{
                    "propertyName": "createdate" if date_field == "Created Date" else "lastmodifieddate",
                    "direction": "ASCENDING"
                }]
            }
            
            if after:
                body["after"] = after
            
            response = requests.post(url, headers=headers, json=body, timeout=30)
            
            if response.status_code == 429:
                retry_after = int(response.headers.get('Retry-After', 10))
                status_text.warning(f"⚠️ Rate limited. Waiting {retry_after} seconds...")
                time.sleep(retry_after)
                continue
            
            response.raise_for_status()
            data = response.json()
            
            batch_contacts = data.get("results", [])
            
            if batch_contacts:
                all_contacts.extend(batch_contacts)
                page_count += 1
                
                progress = min(page_count / 10, 0.99) if page_count <= 10 else 0.99
                progress_bar.progress(progress)
                status_text.text(f"📥 Fetched {len(all_contacts)} contacts...")
                
                paging_info = data.get("paging", {})
                after = paging_info.get("next", {}).get("after")
                
                if not after:
                    break
                
                time.sleep(0.1)
            else:
                break
        
        progress_bar.progress(1.0)
        status_text.text(f"✅ Fetch complete! Total: {len(all_contacts)} contacts")
        
        return all_contacts, len(all_contacts)
        
    except requests.exceptions.RequestException as e:
        progress_bar.empty()
        status_text.empty()
        st.error(f"❌ Error fetching data: {e}")
        return [], 0
    except Exception as e:
        progress_bar.empty()
        status_text.empty()
        st.error(f"❌ Unexpected error: {e}")
        return [], 0

# ✅ NEW: Fetch DEALS from HubSpot
def fetch_hubspot_deals(api_key, start_date, end_date):
    """Fetch DEALS from HubSpot (customers from deal stage)."""
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json"
    }
    
    # Convert dates to timestamps
    start_timestamp = date_to_hubspot_timestamp(start_date, is_end_date=False)
    end_timestamp = date_to_hubspot_timestamp(end_date, is_end_date=True)
    
    all_deals = []
    after = None
    page_count = 0
    
    url = f"{HUBSPOT_API_BASE}/crm/v3/objects/deals/search"
    
    progress_bar = st.progress(0)
    status_text = st.empty()
    status_text.text("📡 Fetching deals (customers)...")
    
    # Build filter for customer deals
    filter_groups = [{
        "filters": [
            {
                "propertyName": "dealstage",
                "operator": "IN",
                "values": CUSTOMER_DEAL_STAGES
            },
            {
                "propertyName": "closedate",
                "operator": "GTE",
                "value": start_timestamp
            },
            {
                "propertyName": "closedate",
                "operator": "LTE",
                "value": end_timestamp
            }
        ]
    }]
    
    # Properties to fetch
    deal_properties = [
        "dealname",
        "dealstage",
        "amount",
        "hubspot_owner_id",
        "closedate",
        "createdate",
        "course",
        "program",
        "product",
        "service",
        "offering",
        "course_name",
        "program_name"
    ]
    
    try:
        while True:
            body = {
                "filterGroups": filter_groups,
                "properties": deal_properties,
                "associations": ["owners"],
                "limit": 100,
                "sorts": [{"propertyName": "closedate", "direction": "ASCENDING"}]
            }
            
            if after:
                body["after"] = after
            
            response = requests.post(url, headers=headers, json=body, timeout=30)
            
            if response.status_code == 429:
                retry_after = int(response.headers.get('Retry-After', 10))
                status_text.warning(f"⚠️ Rate limited. Waiting {retry_after} seconds...")
                time.sleep(retry_after)
                continue
            
            response.raise_for_status()
            data = response.json()
            
            batch_deals = data.get("results", [])
            
            if batch_deals:
                all_deals.extend(batch_deals)
                page_count += 1
                
                progress = min(page_count / 10, 0.99) if page_count <= 10 else 0.99
                progress_bar.progress(progress)
                status_text.text(f"💰 Fetched {len(all_deals)} customer deals...")
                
                paging_info = data.get("paging", {})
                after = paging_info.get("next", {}).get("after")
                
                if not after:
                    break
                
                time.sleep(0.1)
            else:
                break
        
        progress_bar.progress(1.0)
        status_text.text(f"✅ Deal fetch complete! Total: {len(all_deals)} customer deals")
        
        return all_deals, len(all_deals)
        
    except requests.exceptions.RequestException as e:
        progress_bar.empty()
        status_text.empty()
        st.error(f"❌ Error fetching deals: {e}")
        return [], 0
    except Exception as e:
        progress_bar.empty()
        status_text.empty()
        st.error(f"❌ Unexpected error fetching deals: {e}")
        return [], 0

def normalize_lead_status(raw_status):
    """Normalize lead status - NO CUSTOMER HERE."""
    if not raw_status:
        return "Unknown"
    
    status = str(raw_status).strip().lower()
    
    # ❌ NO CUSTOMER HERE - Customer ONLY from DEALS
    # DO NOT ADD "customer" check here
    
    if "prospect" in status:
        if "hot" in status:
            return "Hot"
        elif "warm" in status:
            return "Warm"
        elif "neutral" in status or "cold" in status:
            return "Cold"
        else:
            return "Warm"
    
    if "not_connect" in status or "nc" in status.lower():
        return "Not Connected (NC)"
    
    if "not_interest" in status:
        return "Not Interested"
    
    if "not_qualif" in status or "unqualif" in status:
        return "Not Qualified"
    
    if "duplicate" in status or "junk" in status:
        return "Duplicate"
    
    # ❌ ABSOLUTELY NO CUSTOMER HERE
    # if "customer" in status:
    #     return "Customer"  # ❌ WRONG
    
    if "new" in status or "open" in status:
        return "New Lead"
    
    if status in LEAD_STATUS_MAP:
        return LEAD_STATUS_MAP[status]
    
    return status.replace("_", " ").title()

def process_contacts_data(contacts, owner_mapping=None, api_key=None):
    """Process raw contacts data into a clean DataFrame - NO CUSTOMER HERE."""
    if not contacts:
        return pd.DataFrame()
    
    processed_data = []
    unmapped_owners = set()
    
    for contact in contacts:
        properties = contact.get("properties", {})
        
        # Extract course information
        course_info = ""
        course_fields = [
            "course", "program", "product", "service", "offering",
            "course_name", "program_name", "product_name",
            "enquired_course", "interested_course", "course_interested",
            "program_of_interest", "course_of_interest", "product_of_interest"
        ]
        
        for field in course_fields:
            if field in properties and properties[field] and str(properties[field]).strip():
                course_info = properties[field]
                break
        
        # Owner ID extraction
        owner_id = (
            properties.get("hubspot_owner_id")
            or properties.get("hs_assigned_owner_id")
            or ""
        )
        
        # Try from associations
        if not owner_id:
            associations = contact.get("associations", {})
            owners = associations.get("owners", {}).get("results", [])
            if owners:
                owner_id = str(owners[0].get("id", ""))
        
        owner_id = str(owner_id)
        
        # Map owner ID to name
        owner_name = ""
        if owner_mapping:
            if owner_id in owner_mapping:
                owner_name = owner_mapping[owner_id]
            else:
                if api_key and owner_id:
                    fetched_name = fetch_single_owner_name(api_key, owner_id)
                    if fetched_name:
                        owner_name = fetched_name
                        owner_mapping[owner_id] = fetched_name
                    else:
                        owner_name = f"❌ Unassigned ({owner_id})" if owner_id else "❌ Unassigned"
                        if owner_id:
                            unmapped_owners.add(owner_id)
                else:
                    owner_name = f"❌ Unassigned ({owner_id})" if owner_id else "❌ Unassigned"
                    if owner_id:
                        unmapped_owners.add(owner_id)
        else:
            owner_name = owner_id
        
        # ✅ CRITICAL: NO CUSTOMER from contacts!
        raw_lead_status = properties.get("hs_lead_status", "") or properties.get("lead_status", "")
        
        # Normalize lead status (Cold/Warm/Hot/etc.) - NO CUSTOMER HERE
        lead_status = normalize_lead_status(raw_lead_status)
        
        # Create full name
        full_name = f"{properties.get('firstname', '')} {properties.get('lastname', '')}".strip()
        
        processed_data.append({
            "ID": contact.get("id", ""),
            "Full Name": full_name,
            "Email": properties.get("email", ""),
            "Phone": properties.get("phone", ""),
            "Company": properties.get("company", ""),
            "Job Title": properties.get("jobtitle", ""),
            "Country": properties.get("country", ""),
            "Course/Program": course_info,
            "Course Owner": owner_name,
            "Lead Status": lead_status,  # ❌ NO CUSTOMER
            "Created Date": properties.get("createdate", ""),
            "Lead Status Raw": raw_lead_status,
            "Owner ID": owner_id
        })
    
    # Show warning if there are unmapped owners
    if unmapped_owners:
        st.warning(f"⚠️ {len(unmapped_owners)} owner IDs could not be mapped to names. Showing as '❌ Unassigned'.")
    
    df = pd.DataFrame(processed_data)
    
    if len(processed_data) > 0:
        with st.expander("🧪 Lead Validation", expanded=False):
            st.write(f"✅ Total Leads: {len(df)}")
            st.write("📊 Lead Status Distribution:")
            status_counts = df['Lead Status'].value_counts()
            st.write(status_counts)
            
            # Verify no "Customer" in leads
            customer_in_leads = (df['Lead Status'] == 'Customer').sum()
            if customer_in_leads > 0:
                st.error(f"❌ ERROR: Found {customer_in_leads} 'Customer' in leads - should be ZERO!")
            else:
                st.success("✅ PERFECT: No 'Customer' in lead data (as expected)")
    
    return df

# ✅ NEW: Process DEALS as Customers
def process_deals_as_customers(deals, owner_mapping=None, api_key=None):
    """Process raw deals data into customer DataFrame."""
    if not deals:
        return pd.DataFrame()
    
    processed_data = []
    unmapped_owners = set()
    
    for deal in deals:
        properties = deal.get("properties", {})
        
        # Extract course information from deal
        course_info = ""
        course_fields = [
            "course", "program", "product", "service", "offering",
            "course_name", "program_name", "product_name"
        ]
        
        for field in course_fields:
            if field in properties and properties[field] and str(properties[field]).strip():
                course_info = properties[field]
                break
        
        # Owner ID extraction from deal
        owner_id = properties.get("hubspot_owner_id", "")
        
        # Try from associations
        if not owner_id:
            associations = deal.get("associations", {})
            owners = associations.get("owners", {}).get("results", [])
            if owners:
                owner_id = str(owners[0].get("id", ""))
        
        owner_id = str(owner_id)
        
        # Map owner ID to name
        owner_name = ""
        if owner_mapping:
            if owner_id in owner_mapping:
                owner_name = owner_mapping[owner_id]
            else:
                if api_key and owner_id:
                    fetched_name = fetch_single_owner_name(api_key, owner_id)
                    if fetched_name:
                        owner_name = fetched_name
                        owner_mapping[owner_id] = fetched_name
                    else:
                        owner_name = f"❌ Unassigned ({owner_id})" if owner_id else "❌ Unassigned"
                        if owner_id:
                            unmapped_owners.add(owner_id)
                else:
                    owner_name = f"❌ Unassigned ({owner_id})" if owner_id else "❌ Unassigned"
                    if owner_id:
                        unmapped_owners.add(owner_id)
        else:
            owner_name = owner_id
        
        # Parse amount
        amount = 0
        amount_str = properties.get("amount", "0")
        if amount_str:
            try:
                amount = float(str(amount_str).replace(",", ""))
            except:
                amount = 0
        
        # Get close date
        close_date = properties.get("closedate", "")
        
        processed_data.append({
            "Customer ID": deal.get("id", ""),
            "Deal Name": properties.get("dealname", ""),
            "Course/Program": course_info,
            "Course Owner": owner_name,
            "Amount": amount,
            "Close Date": close_date,
            "Deal Stage": properties.get("dealstage", ""),
            "Is Customer": 1  # ✅ ALL these deals are customers
        })
    
    df = pd.DataFrame(processed_data)
    
    if len(processed_data) > 0:
        with st.expander("🧪 Customer Validation (from Deals)", expanded=False):
            st.write(f"✅ Total Customers from Deals: {len(df)}")
            st.write(f"💰 Total Revenue: ₹{df['Amount'].sum():,.2f}")
            st.write(f"🏆 Top Revenue Course: {df.groupby('Course/Program')['Amount'].sum().idxmax() if not df.empty else 'N/A'}")
            
            # Show sample data
            if not df.empty:
                st.write("Sample Customer Records:", df[['Deal Name', 'Course/Program', 'Amount', 'Close Date']].head(5))
    
    return df

def create_metric_1(df):
    """METRIC 1: Course × Lead Status (COUNT) - Distribution View - NO CUSTOMER"""
    if df.empty or 'Course/Program' not in df.columns:
        return pd.DataFrame()
    
    # Filter out empty courses
    df_course = df[df['Course/Program'].notna() & (df['Course/Program'] != '')].copy()
    
    if df_course.empty:
        return pd.DataFrame()
    
    # Clean course names
    df_course['Course_Clean'] = df_course['Course/Program'].str.strip()
    
    # Create pivot table
    pivot = pd.pivot_table(
        df_course,
        index='Course_Clean',
        columns='Lead Status',
        values='ID',
        aggfunc='count',
        fill_value=0
    )
    
    # Reset index
    pivot = pivot.reset_index().rename(columns={'Course_Clean': 'Course'})
    
    # Add total column
    if len(pivot.columns) > 1:
        status_cols = [col for col in pivot.columns if col != 'Course']
        pivot['Total'] = pivot[status_cols].sum(axis=1)
    
    return pivot

def create_metric_2(df):
    """METRIC 2: Course Owner × Lead Status (COUNT) - Performance View - NO CUSTOMER"""
    if df.empty or 'Course Owner' not in df.columns:
        return pd.DataFrame()
    
    # Filter out empty owners
    df_owner = df[df['Course Owner'].notna() & (df['Course Owner'] != '')].copy()
    
    if df_owner.empty:
        return pd.DataFrame()
    
    # Create pivot table
    pivot = pd.pivot_table(
        df_owner,
        index='Course Owner',
        columns='Lead Status',
        values='ID',
        aggfunc='count',
        fill_value=0
    )
    
    # Reset index
    pivot = pivot.reset_index()
    
    # Add total column
    if len(pivot.columns) > 1:
        status_cols = [col for col in pivot.columns if col != 'Course Owner']
        pivot['Total'] = pivot[status_cols].sum(axis=1)
    
    return pivot

def create_metric_3(df):
    """METRIC 3: Course × Lead Status (COUNT) - For charts & filters"""
    return create_metric_1(df)

# ✅ UPDATED: Merge Customer data from Deals into metrics
def create_metric_4(df_contacts, df_customers):
    """METRIC 4: Course Owner Performance SUMMARY with KPI calculations - Merges Contacts + Customers"""
    if df_contacts.empty or 'Course Owner' not in df_contacts.columns:
        return pd.DataFrame()
    
    # First get lead data from contacts
    owner_lead_pivot = create_metric_2(df_contacts)
    
    if owner_lead_pivot.empty:
        return pd.DataFrame()
    
    # Get customer data from deals
    if not df_customers.empty and 'Course Owner' in df_customers.columns:
        # Group customers by owner
        customer_by_owner = df_customers.groupby('Course Owner').agg(
            Customer_Count=('Is Customer', 'sum'),
            Customer_Revenue=('Amount', 'sum')
        ).reset_index()
    else:
        # Create empty customer data
        customer_by_owner = pd.DataFrame(columns=['Course Owner', 'Customer_Count', 'Customer_Revenue'])
    
    # Merge lead data with customer data
    result_df = owner_lead_pivot.copy()
    
    # Add customer count column (initially 0 for all owners)
    result_df['Customer'] = 0
    
    # Merge with customer data
    if not customer_by_owner.empty:
        result_df = pd.merge(result_df, customer_by_owner, on='Course Owner', how='left')
        # Fill NaN values
        result_df['Customer_Count'] = result_df['Customer_Count'].fillna(0)
        result_df['Customer_Revenue'] = result_df['Customer_Revenue'].fillna(0)
        result_df['Customer'] = result_df['Customer_Count']
    else:
        result_df['Customer_Count'] = 0
        result_df['Customer_Revenue'] = 0
    
    # Deal Leads = Hot + Warm + Cold (from contacts) + Customer (from deals)
    deal_statuses = ['Cold', 'Warm', 'Hot']
    result_df['Deal Leads'] = 0
    
    for status in deal_statuses:
        if status in result_df.columns:
            result_df['Deal Leads'] += result_df[status].fillna(0)
    
    # Add customers from deals
    result_df['Deal Leads'] += result_df['Customer']
    
    # Calculate Deal % = Deal Leads / Grand Total * 100
    if 'Total' in result_df.columns:
        result_df = result_df.rename(columns={'Total': 'Grand Total'})
        result_df['Deal %'] = (result_df['Deal Leads'] / result_df['Grand Total'] * 100).round(2)
    else:
        result_df['Deal %'] = 0
    
    # Calculate Customer % = Customer / Deal Leads * 100
    result_df['Customer %'] = np.where(
        result_df['Deal Leads'] > 0,
        (result_df['Customer'] / result_df['Deal Leads'] * 100).round(2),
        0
    )
    
    # 🔥 NEW: Calculate the three new conversion metrics
    # 1. Lead to Customer Conversion = Customer / Grand Total * 100
    if 'Grand Total' in result_df.columns:
        result_df['Lead→Customer %'] = np.where(
            result_df['Grand Total'] > 0,
            (result_df['Customer'] / result_df['Grand Total'] * 100).round(2),
            0
        )
    else:
        result_df['Lead→Customer %'] = 0
    
    # 2. Lead to Deal Conversion = Deal Leads / Grand Total * 100
    if 'Grand Total' in result_df.columns:
        result_df['Lead→Deal %'] = np.where(
            result_df['Grand Total'] > 0,
            (result_df['Deal Leads'] / result_df['Grand Total'] * 100).round(2),
            0
        )
    else:
        result_df['Lead→Deal %'] = 0
    
    # Select and order columns
    base_cols = ['Course Owner']
    status_cols = ['Cold', 'Hot', 'Warm']
    
    # Ensure status columns exist
    existing_status_cols = [col for col in status_cols if col in result_df.columns]
    
    # 🔥 UPDATED: Include all new metrics
    final_cols = base_cols + existing_status_cols + [
        'Customer', 
        'Customer_Revenue',
        'Deal Leads', 
        'Deal %', 
        'Customer %',
        'Lead→Customer %',
        'Lead→Deal %',
        'Grand Total'
    ]
    
    # Create final dataframe
    final_df = result_df[final_cols].copy()
    
    # Sort by Grand Total if available
    if 'Grand Total' in final_df.columns:
        final_df = final_df.sort_values('Grand Total', ascending=False)
    
    return final_df

# ✅ NEW: Course Revenue Analysis from DEALS
def create_course_revenue(df_customers):
    """Calculate revenue by course from customer DEAL data."""
    if df_customers.empty or 'Course/Program' not in df_customers.columns or 'Amount' not in df_customers.columns:
        return pd.DataFrame()
    
    # Filter courses with revenue
    revenue_df = df_customers[df_customers['Course/Program'].notna() & (df_customers['Course/Program'] != '')].copy()
    
    if revenue_df.empty:
        return pd.DataFrame()
    
    # Clean course names
    revenue_df['Course_Clean'] = revenue_df['Course/Program'].str.strip()
    
    # Group by course
    result_df = revenue_df.groupby('Course_Clean').agg(
        Customers=('Is Customer', 'count'),
        Revenue=('Amount', 'sum')
    ).reset_index().rename(columns={'Course_Clean': 'Course'})
    
    # Calculate revenue per customer
    result_df['Revenue per Customer'] = np.where(
        result_df['Customers'] > 0,
        (result_df['Revenue'] / result_df['Customers']).round(0),
        0
    )
    
    # Sort by revenue
    result_df = result_df.sort_values('Revenue', ascending=False)
    
    return result_df

# ✅ NEW: Volume vs Conversion Matrix with DEAL-based customers
def create_volume_conversion_matrix(metric_1, df_contacts, df_customers):
    """Create a 2x2 matrix to classify courses based on volume and conversion."""
    if metric_1.empty or 'Total' not in metric_1.columns:
        return pd.DataFrame()
    
    # Get customer counts by course from DEALS
    if not df_customers.empty and 'Course/Program' in df_customers.columns:
        customer_by_course = df_customers.groupby('Course/Program').agg(
            Customer_Count=('Is Customer', 'sum')
        ).reset_index()
    else:
        customer_by_course = pd.DataFrame(columns=['Course/Program', 'Customer_Count'])
    
    matrix_data = []
    
    for _, row in metric_1.iterrows():
        course = row['Course']
        total = row.get('Total', 0)
        
        # Get customer count for this course from DEALS
        customer_data = customer_by_course[customer_by_course['Course/Program'] == course]
        customer_count = customer_data['Customer_Count'].iloc[0] if not customer_data.empty else 0
        
        # Calculate conversion %
        conversion_pct = (customer_count / total * 100) if total > 0 else 0
        
        matrix_data.append({
            'Course': course,
            'Volume': total,
            'Conversion %': round(conversion_pct, 1),
            'Customer Count': customer_count
        })
    
    matrix_df = pd.DataFrame(matrix_data)
    
    if len(matrix_df) < 2:
        return matrix_df
    
    # Calculate thresholds (median)
    volume_threshold = matrix_df['Volume'].median()
    conversion_threshold = matrix_df['Conversion %'].median()
    
    # Classify each course
    def classify_course(row):
        if row['Volume'] >= volume_threshold and row['Conversion %'] >= conversion_threshold:
            return "⭐ Star"
        elif row['Volume'] < volume_threshold and row['Conversion %'] >= conversion_threshold:
            return "📈 Potential"
        elif row['Volume'] >= volume_threshold and row['Conversion %'] < conversion_threshold:
            return "⚠️ Burn (High Volume, Low Conversion)"
        else:
            return "❌ Weak"
    
    matrix_df['Segment'] = matrix_df.apply(classify_course, axis=1)
    
    return matrix_df

def create_comparison_data(df_contacts, df_customers, comparison_type, item1, item2):
    """Create comparison data for different comparison types."""
    if df_contacts.empty:
        return None
    
    results = {}
    
    if comparison_type == "Course vs Course":
        # Get course data from contacts
        metric_1 = create_metric_1(df_contacts)
        if not metric_1.empty:
            # Filter for selected courses
            course1_data = metric_1[metric_1['Course'] == item1] if item1 in metric_1['Course'].values else pd.DataFrame()
            course2_data = metric_1[metric_1['Course'] == item2] if item2 in metric_1['Course'].values else pd.DataFrame()
            
            results['type'] = 'course_vs_course'
            results['item1'] = item1
            results['item2'] = item2
            results['data1'] = course1_data
            results['data2'] = course2_data
            
            # Get customer data for these courses from DEALS
            if not df_customers.empty:
                course1_customers = df_customers[df_customers['Course/Program'] == item1]
                course2_customers = df_customers[df_customers['Course/Program'] == item2]
                
                customer_count1 = len(course1_customers)
                customer_count2 = len(course2_customers)
                
                # Calculate comparison metrics
                if not course1_data.empty and not course2_data.empty:
                    total1 = course1_data['Total'].values[0] if 'Total' in course1_data.columns else 1
                    total2 = course2_data['Total'].values[0] if 'Total' in course2_data.columns else 1
                    
                    # Lead to Customer %
                    results['lead_to_customer_pct1'] = round((customer_count1 / total1 * 100), 1) if total1 > 0 else 0
                    results['lead_to_customer_pct2'] = round((customer_count2 / total2 * 100), 1) if total2 > 0 else 0
    
    elif comparison_type == "Owner vs Owner":
        # Get owner data from contacts + customers
        metric_4 = create_metric_4(df_contacts, df_customers)
        if not metric_4.empty:
            # Filter for selected owners
            owner1_data = metric_4[metric_4['Course Owner'] == item1] if item1 in metric_4['Course Owner'].values else pd.DataFrame()
            owner2_data = metric_4[metric_4['Course Owner'] == item2] if item2 in metric_4['Course Owner'].values else pd.DataFrame()
            
            results['type'] = 'owner_vs_owner'
            results['item1'] = item1
            results['item2'] = item2
            results['data1'] = owner1_data
            results['data2'] = owner2_data
    
    elif comparison_type == "Course vs Owner":
        results['type'] = 'course_vs_owner'
        results['item1'] = item1  # Course
        results['item2'] = item2  # Owner
        
        # Get courses for this owner from contacts
        owner_courses = df_contacts[(df_contacts['Course Owner'] == item2) & (df_contacts['Course/Program'].notna()) & (df_contacts['Course/Program'] != '')].copy()
        
        if not owner_courses.empty:
            # Create pivot for owner's courses
            pivot = pd.pivot_table(
                owner_courses,
                index='Course/Program',
                columns='Lead Status',
                values='ID',
                aggfunc='count',
                fill_value=0
            )
            
            results['owner_courses'] = pivot.reset_index()
    
    return results

def calculate_kpis(df_contacts, df_customers):
    """Calculate key performance indicators for the dashboard."""
    if df_contacts.empty:
        return {}
    
    # Total metrics from CONTACTS
    total_leads = len(df_contacts)
    
    # Lead status breakdown from CONTACTS
    status_counts = df_contacts['Lead Status'].value_counts()
    
    cold = status_counts.get('Cold', 0)
    warm = status_counts.get('Warm', 0)
    hot = status_counts.get('Hot', 0)
    new_lead = status_counts.get('New Lead', 0)
    not_connected = status_counts.get('Not Connected (NC)', 0)
    not_interested = status_counts.get('Not Interested', 0)
    not_qualified = status_counts.get('Not Qualified', 0)
    duplicate = status_counts.get('Duplicate', 0)
    
    # ✅ CUSTOMER metrics from DEALS
    if not df_customers.empty:
        customer = len(df_customers)
        total_revenue = df_customers['Amount'].sum()
        avg_revenue_per_customer = round((total_revenue / customer), 0) if customer > 0 else 0
    else:
        customer = 0
        total_revenue = 0
        avg_revenue_per_customer = 0
    
    # Deal Leads = Hot + Warm + Cold (from contacts) + Customer (from deals)
    deal_leads = hot + warm + cold + customer
    
    # 🔥 NEW: Calculate the three new conversion metrics
    # 1. Lead to Customer Conversion
    lead_to_customer_pct = round((customer / total_leads * 100), 1) if total_leads > 0 else 0
    
    # 2. Lead to Deal Conversion
    lead_to_deal_pct = round((deal_leads / total_leads * 100), 1) if total_leads > 0 else 0
    
    # 3. Deal to Customer Conversion
    deal_to_customer_pct = round((customer / deal_leads * 100), 1) if deal_leads > 0 else 0
    
    # Deal % (same as Lead to Deal Conversion)
    deal_pct = round((deal_leads / total_leads * 100), 1) if total_leads > 0 else 0
    
    # Customer % (same as Deal to Customer Conversion)
    customer_pct = round((customer / deal_leads * 100), 1) if deal_leads > 0 else 0
    
    # Top performing metrics
    top_course = ""
    top_owner = ""
    top_revenue_course = ""
    top_revenue_amount = 0
    
    if 'Course/Program' in df_contacts.columns:
        course_counts = df_contacts['Course/Program'].value_counts()
        if not course_counts.empty:
            top_course = str(course_counts.index[0])
    
    if 'Course Owner' in df_contacts.columns:
        owner_counts = df_contacts['Course Owner'].value_counts()
        if not owner_counts.empty:
            top_owner = str(owner_counts.index[0])
    
    # ✅ NEW: Best Revenue Course from DEALS
    if not df_customers.empty and 'Course/Program' in df_customers.columns:
        revenue_by_course = df_customers.groupby('Course/Program')['Amount'].sum()
        if not revenue_by_course.empty:
            top_revenue_course = str(revenue_by_course.index[0])
            top_revenue_amount = revenue_by_course.iloc[0]
    
    # Conversion metrics
    conversion_ratio = round((customer / total_leads * 100), 1) if total_leads > 0 else 0
    dropoff_ratio = round((not_interested + not_qualified + not_connected) / total_leads * 100, 1) if total_leads > 0 else 0
    
    return {
        'total_leads': total_leads,
        'deal_leads': deal_leads,
        'cold': cold,
        'warm': warm,
        'hot': hot,
        'customer': customer,  # ✅ FROM DEALS
        'new_lead': new_lead,
        'not_connected': not_connected,
        'not_interested': not_interested,
        'not_qualified': not_qualified,
        'duplicate': duplicate,
        'deal_pct': deal_pct,
        'customer_pct': customer_pct,
        'lead_to_customer_pct': lead_to_customer_pct,
        'lead_to_deal_pct': lead_to_deal_pct,
        'deal_to_customer_pct': deal_to_customer_pct,
        'total_revenue': total_revenue,
        'avg_revenue_per_customer': avg_revenue_per_customer,
        'top_course': top_course,
        'top_owner': top_owner,
        'top_revenue_course': top_revenue_course,
        'top_revenue_amount': top_revenue_amount,
        'conversion_ratio': conversion_ratio,
        'dropoff_ratio': dropoff_ratio
    }

def main():
    # ✅ SECURITY: Get API key from secrets
    api_key = get_api_key()
    
    if not api_key:
        st.error("""
        ## 🔐 API Key Required
        
        **Please configure your HubSpot API key:**
        
        1. Create a `.streamlit/secrets.toml` file in your project directory
        2. Add your API key:
        ```toml
        [hubspot]
        api_key = "your-pat-key-here"
        ```
        
        3. **OR** set it as environment variable: `HUBSPOT_API_KEY`
        
        4. Restart the app
        
        **Security Note:** Your API key will be encrypted and never exposed in the code.
        """)
        return
    
    # Show API key info (first 10 chars for security)
    api_key_preview = api_key[:10] + "..." if len(api_key) > 10 else api_key
    st.sidebar.success(f"✅ API Key loaded: `{api_key_preview}`")
    
    # Header with gradient
    st.markdown(
        """
        <div class="header-container">
            <h1 style="margin: 0; font-size: 2.5rem;">📊 HubSpot Business Performance Dashboard</h1>
            <p style="margin: 0.5rem 0 0 0; font-size: 1.2rem; opacity: 0.9;">✅ Deal-Based Customer Analytics: Leads + Customers Separated</p>
            <p style="margin: 0.5rem 0 0 0; font-size: 1rem; opacity: 0.8;">⭐ Customers from Deals (Admission Confirmed) | Leads from Contacts</p>
        </div>
        """,
        unsafe_allow_html=True
    )
    
    # ✅ NEW: Data Source Warning
    st.markdown("""
    <div class="warning-card">
        <strong>⚠️ IMPORTANT DATA SOURCE CHANGE:</strong><br>
        • <strong>Customers</strong> now come from <strong>DEALS</strong> (Admission Confirmed stage)<br>
        • <strong>Leads</strong> come from <strong>CONTACTS</strong> only<br>
        • Customer count is 100% accurate (Deal Stage = Admission Confirmed)
    </div>
    """, unsafe_allow_html=True)
    
    # Initialize session state
    if 'contacts_df' not in st.session_state:
        st.session_state.contacts_df = None
    if 'customers_df' not in st.session_state:
        st.session_state.customers_df = None
    if 'owner_mapping' not in st.session_state:
        st.session_state.owner_mapping = None
    if 'metrics' not in st.session_state:
        st.session_state.metrics = {}
    if 'date_filter' not in st.session_state:
        st.session_state.date_filter = None
    if 'date_range' not in st.session_state:
        st.session_state.date_range = None
    if 'revenue_data' not in st.session_state:
        st.session_state.revenue_data = None
    
    # Create sidebar for configuration
    with st.sidebar:
        st.markdown("## 🔧 Configuration")
        
        # Customer Deal Stage Configuration
        st.markdown("### 🎯 Customer Deal Stages")
        default_stages = ", ".join(CUSTOMER_DEAL_STAGES)
        custom_stages = st.text_input(
            "Customer Deal Stages (comma-separated):",
            value=default_stages,
            help="Deal stages that indicate a customer (e.g., admission_confirmed, closedwon)"
        )
        
        # Update global stages
        global CUSTOMER_DEAL_STAGES
        if custom_stages:
            CUSTOMER_DEAL_STAGES = [s.strip().lower() for s in custom_stages.split(",") if s.strip()]
        
        st.info(f"✅ Customer = Deal Stage in: {', '.join(CUSTOMER_DEAL_STAGES)}")
        
        # Test connection button
        if st.button("🔗 Test API Connection", use_container_width=True):
            # First validate format
            is_valid_format, format_msg = validate_api_key(api_key)
            if not is_valid_format:
                st.error(format_msg)
            else:
                # Test connection
                is_valid, message = test_hubspot_connection(api_key)
                if is_valid:
                    st.success(message)
                else:
                    st.error(message)
        
        st.divider()
        
        st.markdown("## 📅 Date Range Filter")
        
        date_field = st.selectbox(
            "Select date field for LEADS:",
            ["Created Date", "Last Modified Date", "Both"]
        )
        
        st.markdown("### Customer Date Range")
        st.info("Customers use DEAL CLOSE DATE (not contact date)")
        
        default_end = datetime.now(IST).date()
        default_start = default_end - timedelta(days=30)
        
        start_date = st.date_input("Start date", value=default_start)
        end_date = st.date_input("End date", value=default_end)
        
        if start_date > end_date:
            st.error("Start date must be before end date!")
            return
        
        days_diff = (end_date - start_date).days + 1
        st.info(f"📅 Will fetch data from {days_diff} day(s)")
        
        st.divider()
        
        st.markdown("## ⚡ Quick Actions")
        
        if st.button("🚀 Fetch ALL Data", type="primary", use_container_width=True):
            if start_date > end_date:
                st.error("Start date must be before end date.")
            else:
                with st.spinner("Fetching data from HubSpot..."):
                    # Test connection first
                    is_valid, message = test_hubspot_connection(api_key)
                    
                    if is_valid:
                        # Store date filter info
                        st.session_state.date_filter = date_field
                        st.session_state.date_range = (start_date.strftime("%Y-%m-%d"), end_date.strftime("%Y-%m-%d"))
                        
                        # Fetch owners
                        st.info("📋 Fetching all owners...")
                        owner_mapping = fetch_owner_mapping(api_key)
                        st.session_state.owner_mapping = owner_mapping
                        
                        # Fetch CONTACTS (Leads)
                        st.info("📊 Fetching contacts (leads)...")
                        contacts, total_contacts = fetch_hubspot_contacts_with_date_filter(
                            api_key, date_field, start_date, end_date
                        )
                        
                        # ✅ NEW: Fetch DEALS (Customers)
                        st.info("💰 Fetching deals (customers)...")
                        deals, total_deals = fetch_hubspot_deals(api_key, start_date, end_date)
                        
                        if contacts:
                            # Process contacts (leads)
                            df_contacts = process_contacts_data(contacts, owner_mapping, api_key)
                            st.session_state.contacts_df = df_contacts
                            
                            # Process deals (customers)
                            df_customers = process_deals_as_customers(deals, owner_mapping, api_key)
                            st.session_state.customers_df = df_customers
                            
                            # Calculate all metrics
                            st.session_state.metrics = {
                                'metric_1': create_metric_1(df_contacts),
                                'metric_2': create_metric_2(df_contacts),
                                'metric_3': create_metric_3(df_contacts),
                                'metric_4': create_metric_4(df_contacts, df_customers)
                            }
                            
                            # Calculate revenue data from CUSTOMERS (deals)
                            st.session_state.revenue_data = create_course_revenue(df_customers)
                            
                            st.success(f"""
                            ✅ Successfully loaded:
                            • 📊 {len(contacts)} contacts (leads)
                            • 💰 {len(deals)} customers (from deals)
                            • 🎯 Customer Deal Stages: {', '.join(CUSTOMER_DEAL_STAGES)}
                            """)
                            st.rerun()
                        else:
                            st.warning("No contacts found for the selected date range.")
                    else:
                        st.error(f"Connection failed: {message}")
        
        if st.button("🔄 Refresh Analysis", use_container_width=True):
            if st.session_state.contacts_df is not None:
                df_contacts = st.session_state.contacts_df
                df_customers = st.session_state.customers_df
                
                st.session_state.metrics = {
                    'metric_1': create_metric_1(df_contacts),
                    'metric_2': create_metric_2(df_contacts),
                    'metric_3': create_metric_3(df_contacts),
                    'metric_4': create_metric_4(df_contacts, df_customers)
                }
                
                # Refresh revenue data
                st.session_state.revenue_data = create_course_revenue(df_customers)
                
                st.success("Analysis refreshed!")
                st.rerun()
        
        if st.button("🗑️ Clear All Data", use_container_width=True):
            st.session_state.clear()
            st.rerun()
        
        st.divider()
        
        # Download Section
        st.markdown("## 📥 Download Options")
        
        if st.session_state.contacts_df is not None and not st.session_state.contacts_df.empty:
            df_contacts = st.session_state.contacts_df
            df_customers = st.session_state.customers_df
            metrics = st.session_state.metrics
            
            if df_customers is not None and not df_customers.empty:
                kpis = calculate_kpis(df_contacts, df_customers)
                
                col1, col2 = st.columns(2)
                
                with col1:
                    # Download Leads Data
                    csv_leads = df_contacts.to_csv(index=False).encode('utf-8')
                    st.download_button(
                        label="📄 Leads CSV",
                        data=csv_leads,
                        file_name=f"hubspot_leads_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                        mime="text/csv",
                        use_container_width=True
                    )
                
                with col2:
                    # Download Customers Data
                    csv_customers = df_customers.to_csv(index=False).encode('utf-8')
                    st.download_button(
                        label="💰 Customers CSV",
                        data=csv_customers,
                        file_name=f"hubspot_customers_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                        mime="text/csv",
                        use_container_width=True
                    )
                
                # Excel Download Button
                if st.button("📊 Download Full Report", use_container_width=True, type="primary"):
                    with st.spinner("🔄 Generating Excel report..."):
                        try:
                            excel_data = create_excel_report(
                                df_contacts, 
                                metrics, 
                                kpis, 
                                df_customers,
                                st.session_state.date_range,
                                st.session_state.date_filter
                            )
                            
                            st.download_button(
                                label="⬇️ Download Excel Report",
                                data=excel_data,
                                file_name=f"HubSpot_Analytics_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                use_container_width=True
                            )
                            st.success("✅ Excel report generated!")
                        except Exception as e:
                            st.error(f"❌ Error: {str(e)}")
        
        st.divider()
        st.markdown("### 📊 Dashboard Logic")
        st.info("""
        **🎯 SEPARATED DATA SOURCES:**
        
        1️⃣ **LEADS** → Contacts API
        • Cold/Warm/Hot/New/NC
        • NO Customers here
        
        2️⃣ **CUSTOMERS** → Deals API  
        • Deal Stage = "Admission Confirmed"
        • Revenue from Amount field
        
        **✅ 100% ACCURATE CUSTOMER LOGIC:**
        • Customer = Deal (NOT Contact Lifecycle)
        • Deal Close Date = Customer Date
        • Revenue = Deal Amount
        
        **📈 CONVERSION METRICS:**
        • Lead→Customer % = Customers/Leads
        • Deal→Customer % = 100% (by definition)
        
        **🔧 CUSTOMER STAGES CONFIGURABLE**
        """)
    
    # Main content area
    if st.session_state.contacts_df is not None and not st.session_state.contacts_df.empty:
        df_contacts = st.session_state.contacts_df
        df_customers = st.session_state.customers_df
        metrics = st.session_state.metrics
        revenue_data = st.session_state.revenue_data
        
        # ✅ NEW: Data Source Summary
        st.markdown("### 📊 Data Source Summary")
        col_sum1, col_sum2, col_sum3 = st.columns(3)
        
        with col_sum1:
            st.metric("Total Leads", f"{len(df_contacts):,}", "From Contacts")
        
        with col_sum2:
            customer_count = len(df_customers) if df_customers is not None else 0
            st.metric("Total Customers", f"{customer_count:,}", "From Deals")
        
        with col_sum3:
            revenue = df_customers['Amount'].sum() if df_customers is not None and not df_customers.empty else 0
            st.metric("Total Revenue", f"₹{revenue:,.0f}", "From Customer Deals")
        
        # Download Center
        st.markdown('<div class="section-header"><h2>📥 Download Center</h2></div>', unsafe_allow_html=True)
        
        col_dl1, col_dl2, col_dl3 = st.columns(3)
        
        with col_dl1:
            csv_leads = df_contacts.to_csv(index=False).encode('utf-8')
            st.download_button(
                label="📄 Download Leads Data",
                data=csv_leads,
                file_name="hubspot_leads.csv",
                mime="text/csv",
                use_container_width=True
            )
        
        with col_dl2:
            if df_customers is not None and not df_customers.empty:
                csv_customers = df_customers.to_csv(index=False).encode('utf-8')
                st.download_button(
                    label="💰 Download Customers Data",
                    data=csv_customers,
                    file_name="hubspot_customers.csv",
                    mime="text/csv",
                    use_container_width=True
                )
        
        with col_dl3:
            if st.button("💎 Generate Full Report", use_container_width=True, type="primary"):
                with st.spinner("Creating report..."):
                    try:
                        kpis = calculate_kpis(df_contacts, df_customers)
                        excel_data = create_excel_report(
                            df_contacts, 
                            metrics, 
                            kpis, 
                            df_customers,
                            st.session_state.date_range,
                            st.session_state.date_filter
                        )
                        
                        st.download_button(
                            label="⬇️ Download Full Excel Report",
                            data=excel_data,
                            file_name="HubSpot_Full_Report.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
                        st.success("✅ Report ready!")
                    except Exception as e:
                        st.error(f"❌ Error: {str(e)}")
        
        st.divider()
        
        # ✅ Executive KPI Dashboard
        st.markdown('<div class="section-header"><h2>🏆 Executive Business Dashboard</h2></div>', unsafe_allow_html=True)
        
        # Calculate KPIs with SEPARATED data
        kpis = calculate_kpis(df_contacts, df_customers)
        
        # Primary KPI Row
        st.markdown(
            render_kpi_row([
                render_kpi("Total Leads", f"{kpis['total_leads']:,}", "From Contacts", "kpi-box-blue"),
                render_kpi("Deal Leads", f"{kpis['deal_leads']:,}", f"{kpis['deal_pct']}% deal conversion", "kpi-box-green"),
                render_kpi("Customers", f"{kpis['customer']:,}", "From Deals ONLY", "deal-kpi"),
                render_kpi("Total Revenue", f"₹{kpis['total_revenue']:,.0f}", f"From {kpis['customer']:,} customers", "revenue-kpi"),
            ]),
            unsafe_allow_html=True
        )
        
        # Conversion KPI Row
        st.markdown(
            render_kpi_row([
                render_secondary_kpi("Lead→Customer", f"{kpis['lead_to_customer_pct']}%", "Leads become customers"),
                render_secondary_kpi("Lead→Deal", f"{kpis['lead_to_deal_pct']}%", "Leads in pipeline"),
                render_secondary_kpi("Deal→Customer", f"{kpis['deal_to_customer_pct']}%", "Pipeline conversion"),
                render_secondary_kpi("Avg Revenue", f"₹{kpis['avg_revenue_per_customer']:,}", "Per customer"),
                render_secondary_kpi("Top Course", kpis['top_course'][:15] if kpis['top_course'] else "N/A", "Most leads"),
            ], container_class="secondary-kpi-container"),
            unsafe_allow_html=True
        )
        
        st.divider()
        
        # Global Filters
        st.markdown("### 🎛️ Global Filters")
        filter_col1, filter_col2, filter_col3 = st.columns(3)
        
        with filter_col1:
            courses = df_contacts['Course/Program'].dropna().unique()
            courses = [str(c).strip() for c in courses if str(c).strip() != '']
            selected_courses = st.multiselect(
                "Filter by Course:",
                options=courses[:50],
                default=[],
                help="Filter leads by course"
            )
        
        with filter_col2:
            owners = df_contacts['Course Owner'].dropna().unique()
            owners = [str(o).strip() for o in owners if str(o).strip() != '']
            selected_owners = st.multiselect(
                "Filter by Owner:",
                options=owners[:50],
                default=[],
                help="Filter leads by owner"
            )
        
        with filter_col3:
            lead_statuses = df_contacts['Lead Status'].dropna().unique()
            lead_statuses = [str(s).strip() for s in lead_statuses if str(s).strip() != '']
            selected_statuses = st.multiselect(
                "Filter by Lead Status:",
                options=lead_statuses,
                default=[],
                help="Filter leads by status"
            )
        
        # Apply filters to leads
        filtered_contacts = df_contacts.copy()
        
        if selected_courses:
            filtered_contacts = filtered_contacts[filtered_contacts['Course/Program'].isin(selected_courses)]
        
        if selected_owners:
            filtered_contacts = filtered_contacts[filtered_contacts['Course Owner'].isin(selected_owners)]
        
        if selected_statuses:
            filtered_contacts = filtered_contacts[filtered_contacts['Lead Status'].isin(selected_statuses)]
        
        # Apply filters to customers (if any)
        filtered_customers = df_customers.copy() if df_customers is not None else pd.DataFrame()
        
        if selected_courses and filtered_customers is not None and not filtered_customers.empty:
            filtered_customers = filtered_customers[filtered_customers['Course/Program'].isin(selected_courses)]
        
        if selected_owners and filtered_customers is not None and not filtered_customers.empty:
            filtered_customers = filtered_customers[filtered_customers['Course Owner'].isin(selected_owners)]
        
        # Show filter info
        filter_info = []
        if selected_courses:
            filter_info.append(f"{len(selected_courses)} courses")
        if selected_owners:
            filter_info.append(f"{len(selected_owners)} owners")
        if selected_statuses:
            filter_info.append(f"{len(selected_statuses)} statuses")
        
        if filter_info:
            st.info(f"📊 Showing {len(filtered_contacts):,} leads (filtered by: {', '.join(filter_info)})")
            
            # Show filtered KPIs
            filtered_kpis = calculate_kpis(filtered_contacts, filtered_customers)
            
            st.markdown(
                render_kpi_row([
                    render_kpi("Filtered Leads", f"{filtered_kpis['total_leads']:,}", f"{filtered_kpis['total_leads']/kpis['total_leads']*100:.1f}% of total", "kpi-box-orange"),
                    render_kpi("Filtered Customers", f"{filtered_kpis['customer']:,}", f"{filtered_kpis['lead_to_customer_pct']}% conversion", "deal-kpi"),
                    render_kpi("Filtered Revenue", f"₹{filtered_kpis['total_revenue']:,.0f}", f"From {filtered_kpis['customer']:,} customers", "revenue-kpi"),
                ]),
                unsafe_allow_html=True
            )
        
        # Create tabs
        tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
            "📊 Lead Distribution", 
            "👤 Owner Performance", 
            "🎯 Conversion Analysis", 
            "📈 KPI Dashboard",
            "💰 Revenue Analysis",
            "🆚 Comparison View"
        ])
        
        # SECTION 1: Lead Distribution
        with tab1:
            st.markdown('<div class="section-header"><h3>📊 Lead Distribution (Contacts)</h3></div>', unsafe_allow_html=True)
            
            if not filtered_contacts.empty:
                # Lead Status Distribution
                st.markdown("#### Lead Status Distribution")
                
                status_counts = filtered_contacts['Lead Status'].value_counts().reset_index()
                status_counts.columns = ['Lead Status', 'Count']
                status_counts['Percentage'] = (status_counts['Count'] / len(filtered_contacts) * 100).round(1)
                
                fig1 = px.pie(
                    status_counts,
                    values='Count',
                    names='Lead Status',
                    title='Lead Status Distribution',
                    hole=0.3,
                    color_discrete_sequence=COLOR_PALETTE
                )
                fig1.update_traces(textposition='inside', textinfo='percent+label')
                st.plotly_chart(fig1, use_container_width=True)
                
                # Course Distribution
                st.markdown("#### Top Courses by Lead Volume")
                
                course_counts = filtered_contacts['Course/Program'].value_counts().head(10).reset_index()
                course_counts.columns = ['Course', 'Count']
                
                fig2 = px.bar(
                    course_counts,
                    x='Course',
                    y='Count',
                    title='Top 10 Courses by Lead Volume',
                    color='Count',
                    color_continuous_scale='Viridis'
                )
                fig2.update_layout(xaxis_tickangle=-45)
                st.plotly_chart(fig2, use_container_width=True)
                
                # Raw Data
                st.markdown("#### Lead Data")
                st.dataframe(filtered_contacts, use_container_width=True, height=300)
        
        # SECTION 2: Owner Performance
        with tab2:
            st.markdown('<div class="section-header"><h3>👤 Owner Performance (Leads)</h3></div>', unsafe_allow_html=True)
            
            if not filtered_contacts.empty:
                metric_2 = create_metric_2(filtered_contacts)
                
                if not metric_2.empty:
                    # Owner Performance Chart
                    st.markdown("#### Owner Performance by Lead Volume")
                    
                    top_owners = metric_2.head(10).copy()
                    top_owners['Course Owner'] = top_owners['Course Owner'].str.slice(0, 20)
                    
                    fig = px.bar(
                        top_owners,
                        x='Course Owner',
                        y='Total',
                        title='Top 10 Owners by Lead Volume',
                        color='Total',
                        color_continuous_scale='Blues'
                    )
                    fig.update_layout(xaxis_tickangle=-45)
                    st.plotly_chart(fig, use_container_width=True)
                    
                    # Owner Details
                    st.markdown("#### Owner Performance Details")
                    st.dataframe(metric_2, use_container_width=True, height=300)
        
        # SECTION 3: Conversion Analysis
        with tab3:
            st.markdown('<div class="section-header"><h3>🎯 Conversion Analysis (Leads → Customers)</h3></div>', unsafe_allow_html=True)
            
            # Volume vs Conversion Matrix
            st.markdown("#### 📉 Volume vs Conversion Matrix")
            
            metric_1 = create_metric_1(filtered_contacts)
            matrix_df = create_volume_conversion_matrix(metric_1, filtered_contacts, filtered_customers)
            
            if not matrix_df.empty:
                # Apply conditional formatting
                def color_matrix(val):
                    if val == "⭐ Star":
                        return 'background-color: #d4edda; color: #155724; font-weight: bold'
                    elif val == "📈 Potential":
                        return 'background-color: #cce5ff; color: #004085; font-weight: bold'
                    elif "⚠️ Burn" in val:
                        return 'background-color: #fff3cd; color: #856404; font-weight: bold'
                    elif val == "❌ Weak":
                        return 'background-color: #f8d7da; color: #721c24; font-weight: bold'
                    return ''
                
                styled_matrix = matrix_df.style.applymap(color_matrix, subset=['Segment'])
                st.dataframe(styled_matrix, use_container_width=True, height=350)
                
                # Conversion Funnel
                st.markdown("#### 📊 Conversion Funnel")
                
                funnel_data = {
                    'Stage': ['Leads', 'Deal Pipeline', 'Customers'],
                    'Count': [
                        kpis['total_leads'],
                        kpis['deal_leads'],
                        kpis['customer']
                    ],
                    'Conversion': [
                        '100%',
                        f"{kpis['lead_to_deal_pct']}%",
                        f"{kpis['lead_to_customer_pct']}%"
                    ]
                }
                
                funnel_df = pd.DataFrame(funnel_data)
                
                fig = px.funnel(
                    funnel_df,
                    x='Count',
                    y='Stage',
                    title='Lead to Customer Conversion Funnel',
                    color='Stage',
                    text='Conversion'
                )
                st.plotly_chart(fig, use_container_width=True)
        
        # SECTION 4: KPI Dashboard
        with tab4:
            st.markdown('<div class="section-header"><h3>📈 KPI Dashboard (Leads + Customers)</h3></div>', unsafe_allow_html=True)
            
            metric_4 = create_metric_4(filtered_contacts, filtered_customers)
            
            if not metric_4.empty:
                # Top 3 Owners KPI
                top_owners_kpi = metric_4.head(3)
                
                owner_perf_kpis = []
                for _, row in top_owners_kpi.iterrows():
                    owner_name = row['Course Owner'][:12] + "..." if len(row['Course Owner']) > 12 else row['Course Owner']
                    lead_to_customer_pct = row.get('Lead→Customer %', 0)
                    revenue = row.get('Customer_Revenue', 0)
                    
                    owner_perf_kpis.append(render_secondary_kpi(
                        owner_name,
                        f"{lead_to_customer_pct}%",
                        f"₹{revenue:,.0f} revenue"
                    ))
                
                st.markdown("#### 🥇 Top 3 Owners by Performance")
                st.markdown(
                    render_kpi_row(owner_perf_kpis, container_class="secondary-kpi-container"),
                    unsafe_allow_html=True
                )
                
                # KPI Table with conditional formatting
                st.markdown("#### Owner KPI Table")
                
                def highlight_lead_to_customer(val):
                    if isinstance(val, (int, float)):
                        if val < 3:
                            return 'background-color: #f8d7da; color: #721c24; font-weight: bold'
                        elif val < 8:
                            return 'background-color: #fff3cd; color: #856404; font-weight: bold'
                        else:
                            return 'background-color: #d4edda; color: #155724; font-weight: bold'
                    return ''
                
                display_df = metric_4.style.applymap(highlight_lead_to_customer, subset=['Lead→Customer %'])
                st.dataframe(display_df, use_container_width=True, height=400)
        
        # SECTION 5: Revenue Analysis
        with tab5:
            st.markdown('<div class="section-header"><h3>💰 Revenue Analysis (From Customer Deals)</h3></div>', unsafe_allow_html=True)
            
            if revenue_data is not None and not revenue_data.empty:
                # Revenue KPIs
                top_revenue = revenue_data.iloc[0] if not revenue_data.empty else None
                total_revenue = revenue_data['Revenue'].sum()
                total_customers = revenue_data['Customers'].sum()
                
                if top_revenue is not None:
                    st.markdown(
                        render_kpi_row([
                            render_kpi("Best Revenue Course", top_revenue['Course'][:20], f"₹{top_revenue['Revenue']:,.0f} revenue", "revenue-kpi"),
                            render_kpi("Total Revenue", f"₹{total_revenue:,.0f}", f"{total_customers} customers", "kpi-box-green"),
                            render_kpi("Avg Revenue/Customer", f"₹{revenue_data['Revenue per Customer'].mean():,.0f}", "Average", "kpi-box-purple"),
                        ]),
                        unsafe_allow_html=True
                    )
                
                # Revenue Chart
                st.markdown("#### Revenue by Course")
                
                top_revenue_chart = revenue_data.head(10).copy()
                top_revenue_chart['Course'] = top_revenue_chart['Course'].str.slice(0, 25)
                
                fig = px.bar(
                    top_revenue_chart,
                    x='Course',
                    y='Revenue',
                    title='Top 10 Courses by Revenue',
                    color='Revenue',
                    color_continuous_scale='Viridis',
                    text='Revenue'
                )
                fig.update_traces(texttemplate='₹%{text:,.0f}', textposition='outside')
                fig.update_layout(xaxis_tickangle=-45)
                st.plotly_chart(fig, use_container_width=True)
                
                # Revenue Table
                st.markdown("#### Revenue Data")
                display_revenue = revenue_data.copy()
                display_revenue['Revenue'] = display_revenue['Revenue'].apply(lambda x: f"₹{x:,.0f}")
                display_revenue['Revenue per Customer'] = display_revenue['Revenue per Customer'].apply(lambda x: f"₹{x:,.0f}")
                st.dataframe(display_revenue, use_container_width=True, height=300)
        
        # SECTION 6: Comparison View
        with tab6:
            st.markdown('<div class="section-header"><h3>🆚 Comparison View</h3></div>', unsafe_allow_html=True)
            
            # Comparison controls
            col_a, col_b = st.columns(2)
            
            with col_a:
                comparison_type = st.selectbox(
                    "Comparison Type:",
                    ["Course vs Course", "Owner vs Owner"]
                )
            
            with col_b:
                if comparison_type == "Course vs Course":
                    courses = filtered_contacts['Course/Program'].dropna().unique()
                    courses = [str(c).strip() for c in courses if str(c).strip() != '']
                    
                    if courses:
                        item1 = st.selectbox("Select Course 1:", courses[:20])
                        remaining = [c for c in courses if c != item1]
                        item2 = st.selectbox("Select Course 2:", ["Select..."] + remaining[:19])
                
                elif comparison_type == "Owner vs Owner":
                    owners = filtered_contacts['Course Owner'].dropna().unique()
                    owners = [str(o).strip() for o in owners if str(o).strip() != '']
                    
                    if owners:
                        item1 = st.selectbox("Select Owner 1:", owners[:20])
                        remaining = [o for o in owners if o != item1]
                        item2 = st.selectbox("Select Owner 2:", ["Select..."] + remaining[:19])
            
            # Perform comparison
            if item1 and item2 and item1 != "Select..." and item2 != "Select...":
                comparison_results = create_comparison_data(
                    filtered_contacts, filtered_customers, comparison_type, item1, item2
                )
                
                if comparison_results:
                    st.markdown(f"### Comparing: **{item1}** vs **{item2}**")
                    
                    if comparison_results['type'] == 'course_vs_course':
                        # Course Comparison Chart
                        if 'lead_to_customer_pct1' in comparison_results:
                            comp_data = pd.DataFrame({
                                'Metric': ['Lead→Customer %'],
                                item1[:20]: [comparison_results['lead_to_customer_pct1']],
                                item2[:20]: [comparison_results['lead_to_customer_pct2']]
                            })
                            
                            fig = px.bar(
                                comp_data.melt(id_vars=['Metric'], var_name='Course', value_name='Percentage'),
                                x='Metric',
                                y='Percentage',
                                color='Course',
                                barmode='group',
                                title='Conversion Rate Comparison',
                                text='Percentage'
                            )
                            fig.update_traces(texttemplate='%{text:.1f}%', textposition='outside')
                            st.plotly_chart(fig, use_container_width=True)
    
    else:
        # Welcome screen
        st.markdown(
            """
            <div style='text-align: center; padding: 3rem;'>
                <h2>👋 Welcome to HubSpot Business Performance Dashboard</h2>
                <p style='font-size: 1.1rem; color: #666; margin: 1rem 0;'>
                    <strong>✅ SEPARATED DATA SOURCES:</strong> Leads from Contacts | Customers from Deals
                </p>
                
                <div style='margin-top: 2rem; background-color: #f8f9fa; padding: 2rem; border-radius: 0.5rem;'>
                    <h4>🎯 NEW DATA ARCHITECTURE:</h4>
                    
                    <div style='display: flex; justify-content: center; gap: 2rem; margin-top: 1rem;'>
                        <div style='text-align: left; background-color: #e8f4fd; padding: 1rem; border-radius: 0.5rem; width: 45%;'>
                            <h5>📊 LEADS (Contacts)</h5>
                            <ul>
                                <li>Cold / Warm / Hot leads</li>
                                <li>New leads, Not Connected</li>
                                <li>Not Interested, Not Qualified</li>
                                <li>❌ NO Customers here</li>
                            </ul>
                        </div>
                        
                        <div style='text-align: left; background-color: #e8f4fd; padding: 1rem; border-radius: 0.5rem; width: 45%;'>
                            <h5>💰 CUSTOMERS (Deals)</h5>
                            <ul>
                                <li>Deal Stage = "Admission Confirmed"</li>
                                <li>Revenue from Deal Amount</li>
                                <li>Close Date = Customer Date</li>
                                <li>✅ ONLY Customers here</li>
                            </ul>
                        </div>
                    </div>
                    
                    <div style='margin-top: 2rem; padding: 1rem; background-color: #d4edda; border-radius: 0.5rem;'>
                        <h5>🚀 Getting Started:</h5>
                        <ol style='text-align: left; margin-left: 25%;'>
                            <li>Configure Customer Deal Stages in sidebar</li>
                            <li>Set date range for leads & customers</li>
                            <li>Click "Fetch ALL Data"</li>
                            <li>Check Executive KPI Dashboard</li>
                            <li>Explore Volume vs Conversion Matrix</li>
                            <li>Analyze Revenue from Customer Deals</li>
                        </ol>
                    </div>
                    
                    <div style='margin-top: 2rem; padding: 1rem; background-color: #fff3cd; border-radius: 0.5rem;'>
                        <h5>⚙️ Customer Deal Stage Configuration:</h5>
                        <p>Default: <code>admission_confirmed, closedwon, won</code></p>
                        <p>Change in sidebar to match your HubSpot deal stages</p>
                    </div>
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )

if __name__ == "__main__":
    main()
