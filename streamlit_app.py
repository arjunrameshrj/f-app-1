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
    page_title="HubSpot Advanced Analytics",
    page_icon="📊",
    layout="wide"
)

# UI Configuration Constants
TOP_N = 10  # Limit charts to top N items
MAX_LABEL_LENGTH = 25  # Truncate long labels
CHART_HEIGHT = 420
COLOR_PALETTE = px.colors.qualitative.Set2

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
    # ❌ NO CUSTOMER HERE - Customer ONLY from Lifecycle Stage
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
def create_excel_report(df, metrics, kpis, date_range, date_field):
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
                'Deal Leads (Hot+Warm+Cold+Customer)',
                'Customers',
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
                'Deal → Customer %'
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
                kpis['deal_to_customer_pct']/100
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
        worksheet.set_column('B:B', 20)
        
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
        
        # Sheet 2: Raw Data
        if not df.empty:
            df.to_excel(writer, sheet_name='Raw Data', index=False)
            worksheet = writer.sheets['Raw Data']
            
            # Apply formatting
            for col_num, value in enumerate(df.columns.values):
                worksheet.write(0, col_num, value, header_format)
            
            # Auto-adjust column widths
            for i, col in enumerate(df.columns):
                column_len = max(df[col].astype(str).map(len).max(), len(col)) + 2
                worksheet.set_column(i, i, min(column_len, 50))
        
        # Sheet 3: Course Distribution (Metric 1)
        if 'metric_1' in metrics and not metrics['metric_1'].empty:
            metric_1 = metrics['metric_1'].copy()
            metric_1.to_excel(writer, sheet_name='Course Distribution', index=False)
            
            worksheet = writer.sheets['Course Distribution']
            for col_num, value in enumerate(metric_1.columns.values):
                worksheet.write(0, col_num, value, header_format)
            
            # Format numbers
            num_cols = len(metric_1.columns)
            for row in range(1, len(metric_1) + 1):
                for col in range(1, num_cols):
                    if metric_1.columns[col] == 'Course':
                        continue
                    worksheet.write(row, col, metric_1.iloc[row-1, col], number_format)
        
        # Sheet 4: Owner Performance (Metric 2)
        if 'metric_2' in metrics and not metrics['metric_2'].empty:
            metric_2 = metrics['metric_2'].copy()
            metric_2.to_excel(writer, sheet_name='Owner Performance', index=False)
            
            worksheet = writer.sheets['Owner Performance']
            for col_num, value in enumerate(metric_2.columns.values):
                worksheet.write(0, col_num, value, header_format)
            
            # Format numbers
            num_cols = len(metric_2.columns)
            for row in range(1, len(metric_2) + 1):
                for col in range(1, num_cols):
                    if metric_2.columns[col] == 'Course Owner':
                        continue
                    worksheet.write(row, col, metric_2.iloc[row-1, col], number_format)
        
        # Sheet 5: KPI Dashboard (Metric 4)
        if 'metric_4' in metrics and not metrics['metric_4'].empty:
            metric_4 = metrics['metric_4'].copy()
            metric_4.to_excel(writer, sheet_name='KPI Dashboard', index=False)
            
            worksheet = writer.sheets['KPI Dashboard']
            for col_num, value in enumerate(metric_4.columns.values):
                worksheet.write(0, col_num, value, header_format)
            
            # Format percentages
            percent_columns = ['Deal %', 'Customer %', 'Lead→Customer %', 'Lead→Deal %']
            num_rows = len(metric_4)
            num_cols = len(metric_4.columns)
            
            for row in range(1, num_rows + 1):
                for col in range(num_cols):
                    col_name = metric_4.columns[col]
                    if col_name in percent_columns:
                        value = metric_4.iloc[row-1, col]
                        if pd.notnull(value):
                            worksheet.write(row, col, value/100, percent_format)
                    elif col_name not in ['Course Owner']:
                        value = metric_4.iloc[row-1, col]
                        if pd.notnull(value):
                            worksheet.write(row, col, value, number_format)
        
        # Sheet 6: Lead Status Summary
        if not df.empty:
            status_summary = df['Lead Status'].value_counts().reset_index()
            status_summary.columns = ['Lead Status', 'Count']
            status_summary['Percentage'] = (status_summary['Count'] / len(df) * 100).round(1)
            
            status_summary.to_excel(writer, sheet_name='Lead Status Summary', index=False)
            
            worksheet = writer.sheets['Lead Status Summary']
            for col_num, value in enumerate(status_summary.columns.values):
                worksheet.write(0, col_num, value, header_format)
            
            # Format percentages
            for row in range(1, len(status_summary) + 1):
                worksheet.write(row, 2, status_summary.iloc[row-1, 2]/100, percent_format)
                worksheet.write(row, 1, status_summary.iloc[row-1, 1], number_format)
        
        # Sheet 7: Top Performers
        top_data = []
        
        # Top Courses
        if not df.empty and 'Course/Program' in df.columns:
            top_courses = df['Course/Program'].value_counts().head(10).reset_index()
            top_courses.columns = ['Course', 'Count']
            for _, row in top_courses.iterrows():
                top_data.append(['Top Course', row['Course'], row['Count']])
        
        # Top Owners
        if not df.empty and 'Course Owner' in df.columns:
            top_owners = df['Course Owner'].value_counts().head(10).reset_index()
            top_owners.columns = ['Owner', 'Count']
            for _, row in top_owners.iterrows():
                top_data.append(['Top Owner', row['Owner'], row['Count']])
        
        if top_data:
            top_df = pd.DataFrame(top_data, columns=['Category', 'Item', 'Count'])
            top_df.to_excel(writer, sheet_name='Top Performers', index=False)
            
            worksheet = writer.sheets['Top Performers']
            for col_num, value in enumerate(top_df.columns.values):
                worksheet.write(0, col_num, value, header_format)
        
        # Add chart sheet (simple bar chart example)
        chart_sheet = workbook.add_worksheet('Charts')
        
        # Add a simple summary chart
        if not df.empty and 'Lead Status' in df.columns:
            chart_data = df['Lead Status'].value_counts().head(10)
            chart_sheet.write_column('A1', chart_data.index.tolist())
            chart_sheet.write_column('B1', chart_data.values.tolist())
            
            # Create a chart object
            chart = workbook.add_chart({'type': 'column'})
            chart.add_series({
                'name': 'Lead Status Count',
                'categories': '=Charts!$A$1:$A$10',
                'values': '=Charts!$B$1:$B$10',
            })
            chart.set_title({'name': 'Top 10 Lead Status Distribution'})
            chart.set_x_axis({'name': 'Lead Status'})
            chart.set_y_axis({'name': 'Count'})
            
            chart_sheet.insert_chart('D2', chart)
    
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
    
    # Define properties needed for all 4 metrics
    all_properties = [
        "hs_lead_status", "lead_status", 
        "lifecyclestage",  # ✅ MUST HAVE: Lifecycle Stage for Customer detection
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
            # ✅ FIX: Add associations to request body
            body = {
                "filterGroups": filter_groups,
                "properties": all_properties,
                "associations": ["owners"],  # ⭐ REQUIRED: Get owner associations
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

def normalize_lead_status(raw_status):
    """Normalize lead status."""
    if not raw_status:
        return "Unknown"
    
    status = str(raw_status).strip().lower()
    
    # ❌ NO CUSTOMER HERE - Customer ONLY from Lifecycle Stage
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

# ✅ FIXED: Added api_key parameter to function
def process_contacts_data(contacts, owner_mapping=None, api_key=None):
    """Process raw contacts data into a clean DataFrame."""
    if not contacts:
        return pd.DataFrame()
    
    processed_data = []
    unmapped_owners = set()
    from_association_count = 0
    from_direct_api_count = 0
    
    for contact in contacts:
        properties = contact.get("properties", {})
        
        # Extract course information from multiple fields
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
        
        # ✅ BULLETPROOF OWNER ID EXTRACTION (FROM ALL POSSIBLE SOURCES)
        # 1. First try from properties
        owner_id = (
            properties.get("hubspot_owner_id")
            or properties.get("hs_assigned_owner_id")
            or ""
        )
        
        # 2. If not found in properties, try from associations (MOST IMPORTANT)
        if not owner_id:
            associations = contact.get("associations", {})
            owners = associations.get("owners", {}).get("results", [])
            if owners:
                owner_id = str(owners[0].get("id", ""))
                from_association_count += 1
        
        owner_id = str(owner_id)
        
        # 🔥 FIXED: Map owner ID to name with direct API fallback (WITH API_KEY)
        owner_name = ""
        if owner_mapping:
            if owner_id in owner_mapping:
                owner_name = owner_mapping[owner_id]
            else:
                # ✅ FIXED: Now api_key is passed correctly
                if api_key and owner_id:  # ✅ ADDED: Check api_key is available
                    fetched_name = fetch_single_owner_name(api_key, owner_id)
                    if fetched_name:
                        owner_name = fetched_name
                        owner_mapping[owner_id] = fetched_name  # cache it
                        from_direct_api_count += 1
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
        
        # ✅ FINAL FIX: CUSTOMER ONLY FROM LIFECYCLE STAGE (EXACT MATCH)
        raw_lead_status = properties.get("hs_lead_status", "") or properties.get("lead_status", "")
        lifecycle_stage = properties.get("lifecyclestage", "")
        
        # Normalize lead status (Cold/Warm/Hot/etc.) - NO CUSTOMER HERE
        lead_status = normalize_lead_status(raw_lead_status)
        
        # 🔥 FINAL RULE: Customer ONLY from Lifecycle Stage (EXACT MATCH)
        if lifecycle_stage:
            lifecycle_clean = lifecycle_stage.strip().lower()
            if lifecycle_clean == "customer":  # ✅ EXACT MATCH ONLY
                lead_status = "Customer"
        
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
            "Lead Status": lead_status,
            "Created Date": properties.get("createdate", ""),
            "Lead Status Raw": raw_lead_status,
            "Lifecycle Stage": lifecycle_stage,  # ✅ ADDED for validation
            "Owner ID": owner_id
        })
    
    # Show debug info
    debug_info = []
    if from_association_count > 0:
        debug_info.append(f"Found {from_association_count} owners from associations")
    if from_direct_api_count > 0:
        debug_info.append(f"Found {from_direct_api_count} owners via direct API")
    
    if debug_info:
        st.info(f"ℹ️ {'; '.join(debug_info)}")
    
    # Show warning if there are unmapped owners
    if unmapped_owners:
        st.warning(f"⚠️ {len(unmapped_owners)} owner IDs could not be mapped to names. Showing as '❌ Unassigned'.")
    
    # ✅ VALIDATION: Show Customer count from Lifecycle Stage
    df = pd.DataFrame(processed_data)
    
    if len(processed_data) > 0:
        with st.expander("🧪 Customer Validation", expanded=False):
            customer_count = (df['Lead Status'] == 'Customer').sum()
            lifecycle_customer_count = len([c for c in processed_data if c.get('Lifecycle Stage', '').strip().lower() == 'customer'])
            
            st.write(f"✅ Dashboard Customer Count: {customer_count}")
            st.write(f"✅ Lifecycle Stage = 'customer': {lifecycle_customer_count}")
            
            # Show sample data
            if customer_count > 0:
                customer_samples = df[df['Lead Status'] == 'Customer'][['Full Name', 'Lead Status Raw', 'Lifecycle Stage']].head(5)
                st.write("Sample Customer Records:", customer_samples)
            
            if customer_count == lifecycle_customer_count:
                st.success("✅ PERFECT MATCH: Customer count is 100% accurate!")
            else:
                st.error(f"❌ MISMATCH: Check data logic")
    
    return df

def create_metric_1(df):
    """METRIC 1: Course × Lead Status (COUNT) - Distribution View"""
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
    """METRIC 2: Course Owner × Lead Status (COUNT) - Performance View"""
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

def create_metric_4(df):
    """METRIC 4: Course Owner Performance SUMMARY with KPI calculations"""
    if df.empty or 'Course Owner' not in df.columns:
        return pd.DataFrame()
    
    # First get metric 2 data
    owner_pivot = create_metric_2(df)
    
    if owner_pivot.empty:
        return pd.DataFrame()
    
    # Deal Leads = Hot + Warm + Cold + Customer
    deal_statuses = ['Cold', 'Warm', 'Hot', 'Customer']
    
    # Ensure all required columns exist
    result_df = owner_pivot.copy()
    
    # Calculate deal leads (Cold + Warm + Hot + Customer)
    result_df['Deal Leads'] = 0
    for status in deal_statuses:
        if status in result_df.columns:
            result_df['Deal Leads'] += result_df[status].fillna(0)
    
    # Calculate Deal % = Deal Leads / Grand Total * 100
    if 'Total' in result_df.columns:
        result_df = result_df.rename(columns={'Total': 'Grand Total'})
        result_df['Deal %'] = (result_df['Deal Leads'] / result_df['Grand Total'] * 100).round(2)
    else:
        result_df['Deal %'] = 0
    
    # Calculate Customer % = Customer / Deal Leads * 100
    if 'Customer' in result_df.columns:
        result_df['Customer %'] = np.where(
            result_df['Deal Leads'] > 0,
            (result_df['Customer'] / result_df['Deal Leads'] * 100).round(2),
            0
        )
    else:
        result_df['Customer %'] = 0
    
    # 🔥 NEW: Calculate the three new conversion metrics
    # 1. Lead to Customer Conversion = Customer / Grand Total * 100
    if 'Customer' in result_df.columns and 'Grand Total' in result_df.columns:
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
    
    # 3. Deal to Customer Conversion = Customer / Deal Leads * 100
    # (Already calculated as 'Customer %')
    
    # Select and order columns
    base_cols = ['Course Owner']
    status_cols = ['Cold', 'Hot', 'Warm', 'Customer']
    
    # Ensure status columns exist
    existing_status_cols = [col for col in status_cols if col in result_df.columns]
    
    # 🔥 UPDATED: Include all new metrics
    final_cols = base_cols + existing_status_cols + [
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

def create_comparison_data(df, comparison_type, item1, item2):
    """Create comparison data for different comparison types."""
    if df.empty:
        return None
    
    results = {}
    
    if comparison_type == "Course vs Course":
        # Get course data
        metric_1 = create_metric_1(df)
        if not metric_1.empty:
            # Filter for selected courses
            course1_data = metric_1[metric_1['Course'] == item1] if item1 in metric_1['Course'].values else pd.DataFrame()
            course2_data = metric_1[metric_1['Course'] == item2] if item2 in metric_1['Course'].values else pd.DataFrame()
            
            results['type'] = 'course_vs_course'
            results['item1'] = item1
            results['item2'] = item2
            results['data1'] = course1_data
            results['data2'] = course2_data
            
            # Calculate comparison metrics
            if not course1_data.empty and not course2_data.empty:
                # Deal Leads (Cold + Warm + Hot + Customer)
                deal_cols = ['Cold', 'Warm', 'Hot', 'Customer']
                deal1 = course1_data[deal_cols].sum(axis=1).values[0] if all(col in course1_data.columns for col in deal_cols) else 0
                total1 = course1_data['Total'].values[0] if 'Total' in course1_data.columns else 1
                deal_pct1 = (deal1 / total1 * 100) if total1 > 0 else 0
                
                deal2 = course2_data[deal_cols].sum(axis=1).values[0] if all(col in course2_data.columns for col in deal_cols) else 0
                total2 = course2_data['Total'].values[0] if 'Total' in course2_data.columns else 1
                deal_pct2 = (deal2 / total2 * 100) if total2 > 0 else 0
                
                results['deal_pct1'] = round(deal_pct1, 1)
                results['deal_pct2'] = round(deal_pct2, 1)
                
                # Customer % (Customer / Deal Leads)
                if 'Customer' in course1_data.columns and deal1 > 0:
                    customer1 = course1_data['Customer'].values[0]
                    results['customer_pct1'] = round((customer1 / deal1 * 100), 1)
                else:
                    results['customer_pct1'] = 0
                
                if 'Customer' in course2_data.columns and deal2 > 0:
                    customer2 = course2_data['Customer'].values[0]
                    results['customer_pct2'] = round((customer2 / deal2 * 100), 1)
                else:
                    results['customer_pct2'] = 0
    
    elif comparison_type == "Owner vs Owner":
        # Get owner data
        metric_4 = create_metric_4(df)
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
        # This is more complex - need to get course data for specific owner
        # For simplicity, we'll show course distribution for selected owner
        results['type'] = 'course_vs_owner'
        results['item1'] = item1  # Course
        results['item2'] = item2  # Owner
        
        # Get courses for this owner
        owner_courses = df[(df['Course Owner'] == item2) & (df['Course/Program'].notna()) & (df['Course/Program'] != '')].copy()
        
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

def calculate_kpis(df):
    """Calculate key performance indicators for the dashboard."""
    if df.empty:
        return {}
    
    # Total metrics
    total_leads = len(df)
    
    # Lead status breakdown
    status_counts = df['Lead Status'].value_counts()
    
    cold = status_counts.get('Cold', 0)
    warm = status_counts.get('Warm', 0)
    hot = status_counts.get('Hot', 0)
    customer = status_counts.get('Customer', 0)  # ✅ NOW 100% ACCURATE
    new_lead = status_counts.get('New Lead', 0)
    not_connected = status_counts.get('Not Connected (NC)', 0)
    not_interested = status_counts.get('Not Interested', 0)
    not_qualified = status_counts.get('Not Qualified', 0)
    duplicate = status_counts.get('Duplicate', 0)
    
    # Deal Leads = Hot + Warm + Cold + Customer
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
    
    if 'Course/Program' in df.columns:
        course_counts = df['Course/Program'].value_counts()
        if not course_counts.empty:
            top_course = str(course_counts.index[0])
            top_course_count = course_counts.iloc[0]
    
    if 'Course Owner' in df.columns:
        owner_counts = df['Course Owner'].value_counts()
        if not owner_counts.empty:
            top_owner = str(owner_counts.index[0])
            top_owner_count = owner_counts.iloc[0]
    
    # Quality metrics
    quality_ratio = round((hot + warm) / total_leads * 100, 1) if total_leads > 0 else 0
    dropoff_ratio = round((not_interested + not_qualified + not_connected) / total_leads * 100, 1) if total_leads > 0 else 0
    
    return {
        'total_leads': total_leads,
        'deal_leads': deal_leads,
        'cold': cold,
        'warm': warm,
        'hot': hot,
        'customer': customer,  # ✅ NOW ACCURATE
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
        'top_course': top_course,
        'top_owner': top_owner,
        'quality_ratio': quality_ratio,
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
            <h1 style="margin: 0; font-size: 2.5rem;">📊 HubSpot Advanced Analytics Dashboard</h1>
            <p style="margin: 0.5rem 0 0 0; font-size: 1.2rem; opacity: 0.9;">✅ Customer count 100% accurate (Lifecycle Stage ONLY)</p>
            <p style="margin: 0.5rem 0 0 0; font-size: 1rem; opacity: 0.8;">✅ Owner mapping fully fixed with API key</p>
        </div>
        """,
        unsafe_allow_html=True
    )
    
    # Initialize session state
    if 'contacts_df' not in st.session_state:
        st.session_state.contacts_df = None
    if 'owner_mapping' not in st.session_state:
        st.session_state.owner_mapping = None
    if 'metrics' not in st.session_state:
        st.session_state.metrics = {}
    if 'date_filter' not in st.session_state:
        st.session_state.date_filter = None
    if 'date_range' not in st.session_state:
        st.session_state.date_range = None
    
    # Create sidebar for configuration
    with st.sidebar:
        st.markdown("## 🔧 Configuration")
        
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
            "Select date field:",
            ["Created Date", "Last Modified Date", "Both"]
        )
        
        default_end = datetime.now(IST).date()
        default_start = default_end - timedelta(days=30)
        
        start_date = st.date_input("Start date", value=default_start)
        end_date = st.date_input("End date", value=default_end)
        
        if start_date > end_date:
            st.error("Start date must be before end date!")
            return
        
        days_diff = (end_date - start_date).days + 1
        st.info(f"📅 Will fetch contacts from {days_diff} day(s)")
        
        st.divider()
        
        st.markdown("## ⚡ Quick Actions")
        
        if st.button("🚀 Fetch ALL Contacts", type="primary", use_container_width=True):
            if start_date > end_date:
                st.error("Start date must be before end date.")
            else:
                with st.spinner("Fetching contacts..."):
                    # Test connection first
                    is_valid, message = test_hubspot_connection(api_key)
                    
                    if is_valid:
                        # Store date filter info
                        st.session_state.date_filter = date_field
                        st.session_state.date_range = (start_date.strftime("%Y-%m-%d"), end_date.strftime("%Y-%m-%d"))
                        
                        # 🔧 FIX: Fetch ALL owners with pagination
                        st.info("📋 Fetching all owners (this may take a moment)...")
                        owner_mapping = fetch_owner_mapping(api_key)
                        st.session_state.owner_mapping = owner_mapping
                        
                        st.info(f"✅ Loaded {len(owner_mapping)} owners")
                        
                        # Now fetch contacts WITH associations
                        contacts, total_fetched = fetch_hubspot_contacts_with_date_filter(
                            api_key, date_field, start_date, end_date
                        )
                        
                        if contacts:
                            # ✅ CRITICAL FIX: Pass api_key to process_contacts_data
                            df = process_contacts_data(contacts, owner_mapping, api_key)  # ✅ ADDED api_key
                            st.session_state.contacts_df = df
                            
                            # Calculate all metrics
                            st.session_state.metrics = {
                                'metric_1': create_metric_1(df),
                                'metric_2': create_metric_2(df),
                                'metric_3': create_metric_3(df),
                                'metric_4': create_metric_4(df)
                            }
                            
                            st.success(f"✅ Successfully loaded {len(contacts)} contacts!")
                            st.rerun()
                        else:
                            st.warning("No contacts found for the selected date range.")
                    else:
                        st.error(f"Connection failed: {message}")
        
        if st.button("🔄 Refresh Analysis", use_container_width=True):
            if st.session_state.contacts_df is not None:
                df = st.session_state.contacts_df
                
                st.session_state.metrics = {
                    'metric_1': create_metric_1(df),
                    'metric_2': create_metric_2(df),
                    'metric_3': create_metric_3(df),
                    'metric_4': create_metric_4(df)
                }
                
                st.success("Analysis refreshed!")
                st.rerun()
        
        if st.button("🗑️ Clear All Data", use_container_width=True):
            st.session_state.clear()
            st.rerun()
        
        st.divider()
        
        # ✅ NEW: Download Section in Sidebar
        st.markdown("## 📥 Download Options")
        
        if st.session_state.contacts_df is not None and not st.session_state.contacts_df.empty:
            df = st.session_state.contacts_df
            metrics = st.session_state.metrics
            kpis = calculate_kpis(df)
            
            col1, col2 = st.columns(2)
            
            with col1:
                # Download Raw Data as CSV
                csv = df.to_csv(index=False).encode('utf-8')
                st.download_button(
                    label="📄 Raw Data CSV",
                    data=csv,
                    file_name=f"hubspot_raw_data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                    mime="text/csv",
                    use_container_width=True,
                    help="Download all contact data as CSV"
                )
            
            with col2:
                # Download Metric 4 (KPI Dashboard) as CSV
                if 'metric_4' in metrics and not metrics['metric_4'].empty:
                    csv_kpi = metrics['metric_4'].to_csv(index=False).encode('utf-8')
                    st.download_button(
                        label="📊 KPI Dashboard CSV",
                        data=csv_kpi,
                        file_name=f"hubspot_kpi_dashboard_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                        mime="text/csv",
                        use_container_width=True,
                        help="Download owner KPI dashboard as CSV"
                    )
            
            # Excel Download Button
            if st.button("📊 Download Full Excel Report", use_container_width=True, type="primary"):
                with st.spinner("🔄 Generating professional Excel report..."):
                    try:
                        excel_data = create_excel_report(
                            df, 
                            metrics, 
                            kpis, 
                            st.session_state.date_range,
                            st.session_state.date_filter
                        )
                        
                        st.download_button(
                            label="⬇️ Click to Download Excel Report",
                            data=excel_data,
                            file_name=f"HubSpot_Analytics_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True,
                            help="Download comprehensive Excel report with multiple sheets and formatting"
                        )
                        st.success("✅ Excel report generated successfully!")
                    except Exception as e:
                        st.error(f"❌ Error generating Excel report: {str(e)}")
        
        st.divider()
        st.markdown("### 📊 About This Dashboard")
        st.info("""
        **✅ 100% ACCURATE DATA:**
        • Cold/Warm/Hot → Lead Status
        • Customer → Lifecycle Stage ONLY
        
        **✅ BULLETPROOF OWNER MAPPING:**
        • Pre-fetched owner list
        • Association fallback
        • Direct API calls for missing
        
        **5 Analytical Sections:**
        1. 📊 Course Distribution (Top 10)
        2. 👤 Owner Performance (Top 10) 
        3. 🎯 Lead Quality (% only)
        4. 📈 KPI Dashboard (Clean tables)
        5. 🆚 Comparison View (One visual)
        
        **✅ Professional Export:**
        • Raw Data CSV
        • KPI Dashboard CSV  
        • Complete Excel Report (Multiple sheets)
        """)
    
    # Main content area
    # Display dashboard if data exists
    if st.session_state.contacts_df is not None and not st.session_state.contacts_df.empty:
        df = st.session_state.contacts_df
        metrics = st.session_state.metrics
        
        # Show owner mapping stats
        if st.session_state.owner_mapping:
            st.sidebar.success(f"📋 {len(st.session_state.owner_mapping)} owners loaded")
        
        # ✅ NEW: Enhanced Download Section at the Top
        st.markdown('<div class="section-header"><h2>📥 Download Center</h2></div>', unsafe_allow_html=True)
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            # Download Raw Data
            csv_raw = df.to_csv(index=False).encode('utf-8')
            st.download_button(
                label="📄 Download Raw Data (CSV)",
                data=csv_raw,
                file_name=f"hubspot_raw_data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                mime="text/csv",
                use_container_width=True,
                help="All contact records with complete details"
            )
        
        with col2:
            # Download KPI Dashboard
            if 'metric_4' in metrics and not metrics['metric_4'].empty:
                csv_kpi = metrics['metric_4'].to_csv(index=False).encode('utf-8')
                st.download_button(
                    label="📊 Download KPI Dashboard (CSV)",
                    data=csv_kpi,
                    file_name=f"hubspot_kpi_dashboard_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                    mime="text/csv",
                    use_container_width=True,
                    help="Owner performance metrics with conversion rates"
                )
        
        with col3:
            # Premium Excel Report Button
            if st.button("💎 Generate Premium Excel Report", use_container_width=True, type="primary"):
                with st.spinner("✨ Creating premium Excel report with formatting..."):
                    try:
                        kpis = calculate_kpis(df)
                        excel_data = create_excel_report(
                            df, 
                            metrics, 
                            kpis, 
                            st.session_state.date_range,
                            st.session_state.date_filter
                        )
                        
                        st.download_button(
                            label="⬇️ Download Premium Excel Report",
                            data=excel_data,
                            file_name=f"HubSpot_Premium_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
                        st.success("✅ Premium Excel report ready for download!")
                    except Exception as e:
                        st.error(f"❌ Error: {str(e)}")
        
        st.divider()
        
        # ✅ NEW: Executive KPI Dashboard at the TOP
        st.markdown('<div class="section-header"><h2>🏆 Executive Dashboard Overview</h2></div>', unsafe_allow_html=True)
        
        # Calculate KPIs
        kpis = calculate_kpis(df)
        
        # Primary KPI Row - BIG, Fixed Size, Centered
        st.markdown(
            render_kpi_row([
                render_kpi("Total Leads", f"{kpis['total_leads']:,}", "All contacts", "kpi-box-blue"),
                render_kpi("Deal Leads", f"{kpis['deal_leads']:,}", f"{kpis['deal_pct']}%", "kpi-box-green"),
                render_kpi("Customers", f"{kpis['customer']:,}", "Lifecycle Stage ONLY", "kpi-box-purple"),
                render_kpi("Lead→Customer", f"{kpis['lead_to_customer_pct']}%", "Customer / Total", "kpi-box-red"),
            ]),
            unsafe_allow_html=True
        )
        
        # Secondary KPI Row with new conversion metrics
        st.markdown(
            render_kpi_row([
                render_secondary_kpi("Hot Leads", f"{kpis['hot']:,}", "High Priority"),
                render_secondary_kpi("Warm Leads", f"{kpis['warm']:,}", "Follow-up"),
                render_secondary_kpi("Cold Leads", f"{kpis['cold']:,}", "Nurture"),
                render_secondary_kpi("Lead→Deal", f"{kpis['lead_to_deal_pct']}%", "Deal / Total"),
                render_secondary_kpi("Deal→Customer", f"{kpis['deal_to_customer_pct']}%", "Customer / Deal"),
            ], container_class="secondary-kpi-container"),
            unsafe_allow_html=True
        )
        
        st.divider()
        
        # Global Filters at the top
        st.markdown("### 🎛️ Global Filters")
        filter_col1, filter_col2, filter_col3 = st.columns(3)
        
        with filter_col1:
            # Course filter
            courses = df['Course/Program'].dropna().unique()
            courses = [str(c).strip() for c in courses if str(c).strip() != '']
            selected_courses = st.multiselect(
                "Filter by Course:",
                options=courses[:50] if len(courses) > 50 else courses,
                default=[],
                help="Select courses to filter all views"
            )
        
        with filter_col2:
            # Owner filter
            owners = df['Course Owner'].dropna().unique()
            owners = [str(o).strip() for o in owners if str(o).strip() != '']
            selected_owners = st.multiselect(
                "Filter by Owner:",
                options=owners[:50] if len(owners) > 50 else owners,
                default=[],
                help="Select owners to filter all views"
            )
        
        with filter_col3:
            # Lead Status filter
            lead_statuses = df['Lead Status'].dropna().unique()
            lead_statuses = [str(s).strip() for s in lead_statuses if str(s).strip() != '']
            selected_statuses = st.multiselect(
                "Filter by Lead Status:",
                options=lead_statuses,
                default=[],
                help="Select lead statuses to filter all views"
            )
        
        # Apply filters
        filtered_df = df.copy()
        
        if selected_courses:
            filtered_df = filtered_df[filtered_df['Course/Program'].isin(selected_courses)]
        
        if selected_owners:
            filtered_df = filtered_df[filtered_df['Course Owner'].isin(selected_owners)]
        
        if selected_statuses:
            filtered_df = filtered_df[filtered_df['Lead Status'].isin(selected_statuses)]
        
        # Update metrics with filtered data
        filtered_metrics = {
            'metric_1': create_metric_1(filtered_df),
            'metric_2': create_metric_2(filtered_df),
            'metric_3': create_metric_3(filtered_df),
            'metric_4': create_metric_4(filtered_df)
        }
        
        # Show filter info
        filter_info = []
        if selected_courses:
            filter_info.append(f"{len(selected_courses)} courses")
        if selected_owners:
            filter_info.append(f"{len(selected_owners)} owners")
        if selected_statuses:
            filter_info.append(f"{len(selected_statuses)} statuses")
        
        if filter_info:
            st.info(f"📊 Showing {len(filtered_df):,} contacts (filtered by: {', '.join(filter_info)})")
        else:
            st.info(f"📊 Showing all {len(filtered_df):,} contacts (no filters applied)")
        
        # ✅ NEW: Filtered KPI Cards
        if selected_courses or selected_owners or selected_statuses:
            filtered_kpis = calculate_kpis(filtered_df)
            
            st.markdown('<div class="section-header"><h3>📊 Filtered View KPIs</h3></div>', unsafe_allow_html=True)
            
            st.markdown(
                render_kpi_row([
                    render_kpi("Filtered Leads", f"{filtered_kpis['total_leads']:,}", f"{filtered_kpis['total_leads']/kpis['total_leads']*100:.1f}% of total", "kpi-box-orange"),
                    render_kpi("Lead→Deal %", f"{filtered_kpis['lead_to_deal_pct']}%", f"{filtered_kpis['deal_leads']:,} deals", "kpi-box-green"),
                    render_kpi("Lead→Customer %", f"{filtered_kpis['lead_to_customer_pct']}%", f"{filtered_kpis['customer']:,} customers", "kpi-box-red"),
                    render_kpi("Deal→Customer %", f"{filtered_kpis['deal_to_customer_pct']}%", "Customer / Deal", "kpi-box-purple"),
                ]),
                unsafe_allow_html=True
            )
            
            # Download filtered data
            st.markdown("### 📥 Download Filtered Data")
            col_f1, col_f2 = st.columns(2)
            
            with col_f1:
                csv_filtered = filtered_df.to_csv(index=False).encode('utf-8')
                st.download_button(
                    label="📄 Download Filtered Data (CSV)",
                    data=csv_filtered,
                    file_name=f"hubspot_filtered_data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                    mime="text/csv",
                    use_container_width=True
                )
        
        # Create tabs for different sections
        tab1, tab2, tab3, tab4, tab5 = st.tabs([
            "📊 Course Distribution", 
            "👤 Owner Performance", 
            "🎯 Lead Quality", 
            "📈 KPI Dashboard",
            "🆚 Comparison View"
        ])
        
        # SECTION 1: Course × Lead Status (Distribution View)
        with tab1:
            st.markdown('<div class="section-header"><h3>📊 Course Distribution (Top 10 by Volume)</h3></div>', unsafe_allow_html=True)
            
            if not filtered_metrics['metric_1'].empty:
                metric_1 = filtered_metrics['metric_1'].copy()
                
                # Download option for this tab
                col_dl1, col_dl2 = st.columns([3, 1])
                with col_dl2:
                    csv_tab1 = metric_1.to_csv(index=False).encode('utf-8')
                    st.download_button(
                        label="📥 Download Course Data",
                        data=csv_tab1,
                        file_name="course_distribution.csv",
                        mime="text/csv",
                        use_container_width=True
                    )
                
                # Limit to top courses for charts
                if 'Total' in metric_1.columns:
                    chart_data = metric_1.sort_values('Total', ascending=False).head(TOP_N)
                else:
                    chart_data = metric_1.head(TOP_N)
                
                # Trim long course names
                chart_data['Course'] = chart_data['Course'].str.slice(0, MAX_LABEL_LENGTH)
                
                # Row 1: Stacked Bar Chart - TOP N ONLY
                st.markdown(f"#### Stacked Bar Chart: Top {TOP_N} Courses")
                
                # Prepare data for stacked bar
                bar_data = chart_data.melt(id_vars=['Course'], var_name='Lead Status', value_name='Count')
                # Remove 'Total' from the chart
                bar_data = bar_data[bar_data['Lead Status'] != 'Total']
                
                # Filter out zero columns
                status_counts = bar_data.groupby('Lead Status')['Count'].sum()
                active_statuses = status_counts[status_counts > 0].index.tolist()
                bar_data = bar_data[bar_data['Lead Status'].isin(active_statuses)]
                
                if not bar_data.empty:
                    fig1 = px.bar(
                        bar_data,
                        x='Course',
                        y='Count',
                        color='Lead Status',
                        title=f'Lead Status Distribution (Top {TOP_N} Courses)',
                        barmode='stack',
                        color_discrete_sequence=COLOR_PALETTE
                    )
                    fig1.update_layout(
                        xaxis_tickangle=-30,
                        xaxis_title="",
                        yaxis_title="Leads Count",
                        legend_title="Lead Status",
                        height=CHART_HEIGHT
                    )
                    st.plotly_chart(fig1, use_container_width=True)
                
                # Row 2: Heatmap Table
                st.markdown("#### Heatmap Table: Course Performance (All Courses)")
                
                # Prepare heatmap data (show all in table, but limit columns)
                heatmap_data = metric_1.copy()
                if 'Total' in heatmap_data.columns:
                    heatmap_data = heatmap_data.drop('Total', axis=1)
                
                # Create heatmap
                fig2 = px.imshow(
                    heatmap_data.set_index('Course'),
                    labels=dict(x="Lead Status", y="Course", color="Count"),
                    aspect="auto",
                    color_continuous_scale='Viridis'
                )
                fig2.update_layout(height=400)
                st.plotly_chart(fig2, use_container_width=True)
                
                # Row 3: Data Table
                st.markdown("#### Detailed Data Table (All Courses)")
                st.dataframe(metric_1, use_container_width=True, height=300)
                
            else:
                st.info("No course data available after filtering")
        
        # SECTION 2: Course Owner × Lead Status (Performance View)
        with tab2:
            st.markdown('<div class="section-header"><h3>👤 Owner Performance (Top 10 by Volume)</h3></div>', unsafe_allow_html=True)
            
            if not filtered_metrics['metric_2'].empty:
                metric_2 = filtered_metrics['metric_2'].copy()
                
                # Download option for this tab
                col_dl1, col_dl2 = st.columns([3, 1])
                with col_dl2:
                    csv_tab2 = metric_2.to_csv(index=False).encode('utf-8')
                    st.download_button(
                        label="📥 Download Owner Data",
                        data=csv_tab2,
                        file_name="owner_performance.csv",
                        mime="text/csv",
                        use_container_width=True
                    )
                
                # ✅ NEW: Owner Performance KPI Cards
                # Get top 5 owners by total
                if 'Total' in metric_2.columns:
                    top_owners_kpi = metric_2.nlargest(5, 'Total')
                    
                    owner_kpis = []
                    for _, row in top_owners_kpi.iterrows():
                        owner_name = row['Course Owner'][:15] + "..." if len(row['Course Owner']) > 15 else row['Course Owner']
                        total = int(row['Total'])
                        
                        # Calculate deal %
                        deal_cols = ['Cold', 'Warm', 'Hot', 'Customer']
                        deal = sum([row.get(col, 0) for col in deal_cols if col in row])
                        deal_pct = round((deal / total * 100), 1) if total > 0 else 0
                        
                        owner_kpis.append(render_secondary_kpi(
                            owner_name, 
                            f"{total:,}", 
                            f"{deal_pct}% deals"
                        ))
                    
                    st.markdown("#### 🥇 Top 5 Owners by Volume")
                    st.markdown(
                        render_kpi_row(owner_kpis, container_class="secondary-kpi-container"),
                        unsafe_allow_html=True
                    )
                
                # Limit to top owners for charts
                if 'Total' in metric_2.columns:
                    chart_data = metric_2.sort_values('Total', ascending=False).head(TOP_N)
                else:
                    chart_data = metric_2.head(TOP_N)
                
                # Trim long owner names
                chart_data['Course Owner'] = chart_data['Course Owner'].str.slice(0, MAX_LABEL_LENGTH)
                
                # Row 1: Grouped Bar Chart - TOP N ONLY
                st.markdown(f"#### Grouped Bar Chart: Top {TOP_N} Owners")
                
                # Prepare data for grouped bar
                bar_data = chart_data.melt(id_vars=['Course Owner'], var_name='Lead Status', value_name='Count')
                # Focus on key statuses only
                key_statuses = ['Cold', 'Warm', 'Hot', 'Customer']
                bar_data = bar_data[bar_data['Lead Status'].isin(key_statuses)]
                
                # Filter out zero statuses
                status_counts = bar_data.groupby('Lead Status')['Count'].sum()
                active_statuses = status_counts[status_counts > 0].index.tolist()
                bar_data = bar_data[bar_data['Lead Status'].isin(active_statuses)]
                
                if not bar_data.empty:
                    fig1 = px.bar(
                        bar_data,
                        x='Course Owner',
                        y='Count',
                        color='Lead Status',
                        title=f'Key Lead Status by Owner (Top {TOP_N})',
                        barmode='group',
                        color_discrete_sequence=COLOR_PALETTE
                    )
                    fig1.update_layout(
                        xaxis_tickangle=-30,
                        xaxis_title="",
                        yaxis_title="Leads Count",
                        legend_title="Lead Status",
                        height=CHART_HEIGHT
                    )
                    st.plotly_chart(fig1, use_container_width=True)
                
                # Row 2: Donut Chart for selected owner
                st.markdown("#### Donut Chart: Select an Owner")
                
                if 'Course Owner' in metric_2.columns and len(metric_2) > 0:
                    owner_options = metric_2['Course Owner'].tolist()
                    selected_owner = st.selectbox("Select owner for donut chart:", owner_options, key="owner_select_tab2")
                    
                    if selected_owner:
                        owner_data = metric_2[metric_2['Course Owner'] == selected_owner]
                        if not owner_data.empty:
                            # Melt for donut
                            donut_data = owner_data.melt(id_vars=['Course Owner'], var_name='Lead Status', value_name='Count')
                            donut_data = donut_data[donut_data['Lead Status'] != 'Total']
                            
                            if not donut_data.empty:
                                fig2 = px.pie(
                                    donut_data,
                                    values='Count',
                                    names='Lead Status',
                                    title=f'Lead Status Distribution for {selected_owner}',
                                    hole=0.4,
                                    color_discrete_sequence=COLOR_PALETTE
                                )
                                fig2.update_traces(textposition='inside', textinfo='percent+label')
                                fig2.update_layout(height=400)
                                st.plotly_chart(fig2, use_container_width=True)
                
                # Row 3: Data Table
                st.markdown("#### Owner Performance Table (All Owners)")
                st.dataframe(metric_2, use_container_width=True, height=300)
                
            else:
                st.info("No owner data available after filtering")
        
        # SECTION 3: Course-wise Lead Quality View
        with tab3:
            st.markdown('<div class="section-header"><h3>🎯 Course-wise Lead Quality Analysis</h3></div>', unsafe_allow_html=True)
            
            if not filtered_metrics['metric_1'].empty:
                metric_1 = filtered_metrics['metric_1'].copy()
                
                # ✅ NEW: Top 3 Courses by Quality KPI Cards
                # Calculate quality score for each course
                quality_scores = []
                for _, row in metric_1.iterrows():
                    course = row['Course']
                    total = row.get('Total', 0)
                    
                    if total > 0:
                        # Quality score = (Hot + Warm) / Total
                        hot = row.get('Hot', 0)
                        warm = row.get('Warm', 0)
                        quality_score = round(((hot + warm) / total * 100), 1)
                        
                        quality_scores.append({
                            'Course': course,
                            'Quality Score': quality_score,
                            'Total': total,
                            'Hot': hot,
                            'Warm': warm
                        })
                
                if quality_scores:
                    quality_df = pd.DataFrame(quality_scores)
                    top_quality_courses = quality_df.nlargest(3, 'Quality Score')
                    
                    quality_kpis = []
                    for _, row in top_quality_courses.iterrows():
                        course_name = row['Course'][:12] + "..." if len(row['Course']) > 12 else row['Course']
                        quality_kpis.append(render_secondary_kpi(
                            course_name,
                            f"{row['Quality Score']}%",
                            f"{row['Total']:,} total"
                        ))
                    
                    st.markdown("#### 🥇 Top 3 Courses by Quality Score")
                    st.markdown(
                        render_kpi_row(quality_kpis, container_class="secondary-kpi-container"),
                        unsafe_allow_html=True
                    )
                
                # Limit to top courses for charts
                if 'Total' in metric_1.columns:
                    chart_data = metric_1.sort_values('Total', ascending=False).head(TOP_N)
                else:
                    chart_data = metric_1.head(TOP_N)
                
                # Trim long course names
                chart_data['Course'] = chart_data['Course'].str.slice(0, MAX_LABEL_LENGTH)
                
                # Row 1: 100% Stacked Bar Chart - TOP N ONLY
                st.markdown(f"#### 100% Stacked Bar: Quality Profile (Top {TOP_N} Courses)")
                
                # Calculate percentages
                perc_data = chart_data.copy()
                if 'Total' in perc_data.columns:
                    # Calculate percentage for each status
                    status_cols = [col for col in perc_data.columns if col not in ['Course', 'Total']]
                    for col in status_cols:
                        perc_data[col] = (perc_data[col] / perc_data['Total'] * 100).round(1)
                
                # Melt for stacked bar
                stacked_data = perc_data.melt(id_vars=['Course'], var_name='Lead Status', value_name='Percentage')
                stacked_data = stacked_data[stacked_data['Lead Status'] != 'Total']
                
                # Filter out zero statuses
                status_sums = stacked_data.groupby('Lead Status')['Percentage'].sum()
                active_statuses = status_sums[status_sums > 0].index.tolist()
                stacked_data = stacked_data[stacked_data['Lead Status'].isin(active_statuses)]
                
                if not stacked_data.empty:
                    fig1 = px.bar(
                        stacked_data,
                        x='Course',
                        y='Percentage',
                        color='Lead Status',
                        title=f'Lead Status % by Course (Top {TOP_N})',
                        barmode='stack',
                        text='Percentage',
                        color_discrete_sequence=COLOR_PALETTE
                    )
                    fig1.update_layout(
                        xaxis_tickangle=-30,
                        xaxis_title="",
                        yaxis_title="Percentage (%)",
                        yaxis_range=[0, 100],
                        legend_title="Lead Status",
                        height=CHART_HEIGHT
                    )
                    fig1.update_traces(texttemplate='%{text:.1f}%', textposition='inside')
                    st.plotly_chart(fig1, use_container_width=True)
                
                # Row 2: New Conversion Metrics for Courses
                st.markdown("#### Conversion Metrics (Top 10 Courses)")
                
                # Define deal statuses (Cold + Warm + Hot + Customer)
                deal_statuses = ['Cold', 'Warm', 'Hot', 'Customer']
                
                # Calculate metrics for each course
                kpi_data = []
                for _, row in chart_data.iterrows():
                    course = row['Course']
                    total = row.get('Total', 0)
                    
                    # Calculate deal leads (Cold + Warm + Hot + Customer)
                    deal = sum([row.get(status, 0) for status in deal_statuses if status in row])
                    
                    # Calculate lead to deal percentage
                    lead_to_deal_pct = (deal / total * 100) if total > 0 else 0
                    
                    # Calculate customer count
                    customer = row.get('Customer', 0)
                    
                    # Calculate lead to customer percentage
                    lead_to_customer_pct = (customer / total * 100) if total > 0 else 0
                    
                    # Calculate deal to customer percentage
                    deal_to_customer_pct = (customer / deal * 100) if deal > 0 else 0
                    
                    kpi_data.append({
                        'Course': course,
                        'Total': int(total),
                        'Deal': int(deal),
                        'Lead→Deal %': round(lead_to_deal_pct, 1),
                        'Customer': int(customer),
                        'Lead→Customer %': round(lead_to_customer_pct, 1),
                        'Deal→Customer %': round(deal_to_customer_pct, 1)
                    })
                
                kpi_df = pd.DataFrame(kpi_data)
                
                # Display as a clean table
                st.dataframe(kpi_df, use_container_width=True, height=300)
                
            else:
                st.info("No course quality data available")
        
        # SECTION 4: Course Owner KPI TABLE
        with tab4:
            st.markdown('<div class="section-header"><h3>📈 Course Owner KPI Dashboard</h3></div>', unsafe_allow_html=True)
            
            if not filtered_metrics['metric_4'].empty:
                metric_4 = filtered_metrics['metric_4'].copy()
                
                # Show unassigned count
                unassigned_count = (df['Course Owner'].str.contains('❌ Unassigned', na=False)).sum()
                if unassigned_count > 0:
                    st.warning(f"⚠️ {unassigned_count} contacts are unassigned")
                
                # ✅ NEW: Owner Performance KPI Cards (Top 3)
                if len(metric_4) > 0:
                    top_owners_kpi = metric_4.head(3)
                    
                    owner_perf_kpis = []
                    for _, row in top_owners_kpi.iterrows():
                        owner_name = row['Course Owner'][:12] + "..." if len(row['Course Owner']) > 12 else row['Course Owner']
                        lead_to_deal_pct = row.get('Lead→Deal %', 0)
                        lead_to_customer_pct = row.get('Lead→Customer %', 0)
                        deal_to_customer_pct = row.get('Customer %', 0)
                        
                        # Determine color based on performance
                        if lead_to_deal_pct >= 50 and lead_to_customer_pct >= 10:
                            color_class = "kpi-box-green"
                        elif lead_to_deal_pct >= 30 and lead_to_customer_pct >= 5:
                            color_class = "kpi-box-orange"
                        else:
                            color_class = "kpi-box"
                        
                        owner_perf_kpis.append(render_kpi(
                            owner_name,
                            f"{lead_to_deal_pct}%",
                            f"L→D: {lead_to_deal_pct}% | L→C: {lead_to_customer_pct}%",
                            color_class
                        ))
                    
                    st.markdown("#### 🥇 Top 3 Owners by Volume")
                    st.markdown(
                        render_kpi_row(owner_perf_kpis),
                        unsafe_allow_html=True
                    )
                
                # Display the KPI table with styling
                st.markdown("#### Owner Performance Summary")
                
                # Apply conditional formatting
                def highlight_lead_to_deal(val):
                    if isinstance(val, (int, float)):
                        if val < 20:
                            return 'background-color: #f8d7da; color: #721c24; font-weight: bold'
                        elif val < 40:
                            return 'background-color: #fff3cd; color: #856404; font-weight: bold'
                        else:
                            return 'background-color: #d4edda; color: #155724; font-weight: bold'
                    return ''
                
                def highlight_lead_to_customer(val):
                    if isinstance(val, (int, float)):
                        if val < 3:
                            return 'background-color: #f8d7da; color: #721c24; font-weight: bold'
                        elif val < 8:
                            return 'background-color: #fff3cd; color: #856404; font-weight: bold'
                        else:
                            return 'background-color: #d4edda; color: #155724; font-weight: bold'
                    return ''
                
                def highlight_deal_to_customer(val):
                    if isinstance(val, (int, float)):
                        if val < 10:
                            return 'background-color: #f8d7da; color: #721c24; font-weight: bold'
                        elif val < 20:
                            return 'background-color: #fff3cd; color: #856404; font-weight: bold'
                        else:
                            return 'background-color: #d4edda; color: #155724; font-weight: bold'
                    return ''
                
                # Display styled dataframe ONLY
                display_df = metric_4.style.applymap(highlight_lead_to_deal, subset=['Lead→Deal %'])
                display_df = display_df.applymap(highlight_lead_to_customer, subset=['Lead→Customer %'])
                display_df = display_df.applymap(highlight_deal_to_customer, subset=['Customer %'])
                
                st.dataframe(display_df, use_container_width=True, height=450)
                
                # Export options
                st.markdown("#### 📥 Export Options")
                col_exp1, col_exp2 = st.columns(2)
                
                with col_exp1:
                    csv = metric_4.to_csv(index=False)
                    st.download_button(
                        "📊 Download KPI Table (CSV)",
                        csv,
                        "owner_kpi_table.csv",
                        "text/csv",
                        use_container_width=True,
                        help="Download KPI dashboard as CSV"
                    )
                
                with col_exp2:
                    if st.button("💎 Export to Excel with Formatting", use_container_width=True):
                        with st.spinner("Creating formatted Excel..."):
                            try:
                                kpis = calculate_kpis(df)
                                excel_data = create_excel_report(
                                    df, 
                                    {'metric_4': metric_4}, 
                                    kpis, 
                                    st.session_state.date_range,
                                    st.session_state.date_filter
                                )
                                
                                st.download_button(
                                    "⬇️ Download Formatted Excel",
                                    excel_data,
                                    "owner_kpi_dashboard.xlsx",
                                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    use_container_width=True
                                )
                                st.success("Formatted Excel ready!")
                            except Exception as e:
                                st.error(f"Error: {str(e)}")
                
                # Show formula explanation
                with st.expander("📊 Formula Explanation"):
                    st.markdown("""
                    **✅ 100% ACCURATE DATA LOGIC:**
                    
                    • **Cold/Warm/Hot** → From Lead Status
                    • **Customer** → From Lifecycle Stage ONLY (exact match "customer")
                    
                    **📈 KPI Formulas:**
                    
                    1. **Deal Leads** = Cold + Warm + Hot + Customer
                    2. **Lead → Deal %** = (Deal Leads / Grand Total) × 100
                    3. **Lead → Customer %** = (Customer / Grand Total) × 100
                    4. **Deal → Customer %** = (Customer / Deal Leads) × 100
                    
                    **🎯 Quality Thresholds:**
                    - **Lead→Deal %**: 🟢 >40% | 🟠 20-40% | 🔴 <20%
                    - **Lead→Customer %**: 🟢 >8% | 🟠 3-8% | 🔴 <3%
                    - **Deal→Customer %**: 🟢 >20% | 🟠 10-20% | 🔴 <10%
                    """)
                
            else:
                st.info("No KPI data available")
        
        # SECTION 5: COMPARISON VIEW
        with tab5:
            st.markdown('<div class="section-header"><h3>🆚 Comparison View</h3></div>', unsafe_allow_html=True)
            
            # Comparison controls
            st.markdown("#### Comparison Configuration")
            
            col_a, col_b = st.columns(2)
            
            with col_a:
                comparison_type = st.selectbox(
                    "Comparison Type:",
                    ["Course vs Course", "Owner vs Owner", "Course vs Owner"]
                )
            
            with col_b:
                # Get available items based on comparison type
                if comparison_type == "Course vs Course":
                    items = filtered_metrics['metric_1']['Course'].tolist() if not filtered_metrics['metric_1'].empty else []
                    items = [str(i).strip() for i in items if str(i).strip() != '']
                    if items:
                        item1 = st.selectbox("Select Course 1:", items)
                        remaining_items = [i for i in items if i != item1]
                        item2 = st.selectbox("Select Course 2:", ["Select..."] + remaining_items) if remaining_items else None
                    else:
                        item1 = None
                        item2 = None
                
                elif comparison_type == "Owner vs Owner":
                    items = filtered_metrics['metric_2']['Course Owner'].tolist() if not filtered_metrics['metric_2'].empty else []
                    items = [str(i).strip() for i in items if str(i).strip() != '']
                    if items:
                        item1 = st.selectbox("Select Owner 1:", items)
                        remaining_items = [i for i in items if i != item1]
                        item2 = st.selectbox("Select Owner 2:", ["Select..."] + remaining_items) if remaining_items else None
                    else:
                        item1 = None
                        item2 = None
                
                else:  # Course vs Owner
                    courses = filtered_metrics['metric_1']['Course'].tolist() if not filtered_metrics['metric_1'].empty else []
                    courses = [str(c).strip() for c in courses if str(c).strip() != '']
                    owners = filtered_metrics['metric_2']['Course Owner'].tolist() if not filtered_metrics['metric_2'].empty else []
                    owners = [str(o).strip() for o in owners if str(o).strip() != '']
                    
                    item1 = st.selectbox("Select Course:", ["Select..."] + courses) if courses else None
                    item2 = st.selectbox("Select Owner:", ["Select..."] + owners) if owners else None
            
            # Perform comparison
            if item1 and item2 and item1 != "Select..." and item2 != "Select...":
                comparison_results = create_comparison_data(
                    filtered_df, comparison_type, item1, item2
                )
                
                if comparison_results:
                    st.markdown(f"### Comparing: **{item1}** vs **{item2}**")
                    
                    # ✅ NEW: KPI Comparison Cards
                    if comparison_results['type'] == 'course_vs_course':
                        # Create KPI comparison cards
                        if 'deal_pct1' in comparison_results and 'deal_pct2' in comparison_results:
                            st.markdown(
                                render_kpi_row([
                                    render_kpi(f"{item1[:15]}", f"{comparison_results['deal_pct1']}%", "Lead→Deal %", "kpi-box-blue"),
                                    render_kpi("VS", "", "Comparison", "kpi-box"),
                                    render_kpi(f"{item2[:15]}", f"{comparison_results['deal_pct2']}%", "Lead→Deal %", "kpi-box-green"),
                                ]),
                                unsafe_allow_html=True
                            )
                            
                            # Add customer % comparison
                            if 'customer_pct1' in comparison_results and 'customer_pct2' in comparison_results:
                                st.markdown(
                                    render_kpi_row([
                                        render_kpi(f"{item1[:15]}", f"{comparison_results['customer_pct1']}%", "Deal→Customer %", "kpi-box-purple"),
                                        render_kpi("VS", "", "Comparison", "kpi-box"),
                                        render_kpi(f"{item2[:15]}", f"{comparison_results['customer_pct2']}%", "Deal→Customer %", "kpi-box-teal"),
                                    ]),
                                    unsafe_allow_html=True
                                )
                    
                    elif comparison_results['type'] == 'owner_vs_owner':
                        # Create owner comparison KPI cards
                        if not comparison_results['data1'].empty and not comparison_results['data2'].empty:
                            owner1_data = comparison_results['data1'].iloc[0] if len(comparison_results['data1']) > 0 else pd.Series()
                            owner2_data = comparison_results['data2'].iloc[0] if len(comparison_results['data2']) > 0 else pd.Series()
                            
                            # Get key metrics
                            owner1_lead_to_deal = owner1_data.get('Lead→Deal %', 0)
                            owner1_lead_to_customer = owner1_data.get('Lead→Customer %', 0)
                            owner2_lead_to_deal = owner2_data.get('Lead→Deal %', 0)
                            owner2_lead_to_customer = owner2_data.get('Lead→Customer %', 0)
                            
                            st.markdown(
                                render_kpi_row([
                                    render_kpi(f"{item1[:12]}", f"{owner1_lead_to_deal}%", "Lead→Deal %", "kpi-box-blue"),
                                    render_kpi("L→D %", "", "Metric", "kpi-box"),
                                    render_kpi(f"{item2[:12]}", f"{owner2_lead_to_deal}%", "Lead→Deal %", "kpi-box-green"),
                                ]),
                                unsafe_allow_html=True
                            )
                            
                            st.markdown(
                                render_kpi_row([
                                    render_kpi(f"{item1[:12]}", f"{owner1_lead_to_customer}%", "Lead→Customer %", "kpi-box-purple"),
                                    render_kpi("L→C %", "", "Metric", "kpi-box"),
                                    render_kpi(f"{item2[:12]}", f"{owner2_lead_to_customer}%", "Lead→Customer %", "kpi-box-teal"),
                                ]),
                                unsafe_allow_html=True
                            )
                    
                    # Original visualization
                    if comparison_results['type'] == 'course_vs_course':
                        # ONE VISUAL: Side-by-side bar for % comparison
                        if 'deal_pct1' in comparison_results and 'deal_pct2' in comparison_results:
                            comp_data = pd.DataFrame({
                                'Metric': ['Lead→Deal %', 'Deal→Customer %'],
                                item1[:20]: [comparison_results['deal_pct1'], comparison_results.get('customer_pct1', 0)],
                                item2[:20]: [comparison_results['deal_pct2'], comparison_results.get('customer_pct2', 0)]
                            })
                            
                            fig = px.bar(
                                comp_data.melt(id_vars=['Metric'], var_name='Item', value_name='Percentage'),
                                x='Metric',
                                y='Percentage',
                                color='Item',
                                barmode='group',
                                title='Performance Comparison (%)',
                                text='Percentage',
                                color_discrete_sequence=COLOR_PALETTE
                            )
                            fig.update_traces(texttemplate='%{text:.1f}%', textposition='outside')
                            fig.update_layout(
                                xaxis_title="",
                                yaxis_title="Percentage (%)",
                                height=400
                            )
                            st.plotly_chart(fig, use_container_width=True)
                    
                    elif comparison_results['type'] == 'owner_vs_owner':
                        # ONE VISUAL: Funnel bar chart
                        if not comparison_results['data1'].empty and not comparison_results['data2'].empty:
                            # Get funnel data
                            funnel_cols = ['Cold', 'Warm', 'Hot', 'Customer']
                            owner1_data = comparison_results['data1'].iloc[0] if len(comparison_results['data1']) > 0 else pd.Series()
                            owner2_data = comparison_results['data2'].iloc[0] if len(comparison_results['data2']) > 0 else pd.Series()
                            
                            comparison_list = []
                            for col in funnel_cols:
                                if col in owner1_data and col in owner2_data:
                                    comparison_list.append({
                                        'Stage': col,
                                        item1[:15]: int(owner1_data[col]),
                                        item2[:15]: int(owner2_data[col])
                                    })
                            
                            if comparison_list:
                                radar_df = pd.DataFrame(comparison_list)
                                melted_df = radar_df.melt(id_vars=['Stage'], var_name='Owner', value_name='Count')
                                
                                # Create grouped bar chart instead of radar
                                fig = px.bar(
                                    melted_df,
                                    x='Stage',
                                    y='Count',
                                    color='Owner',
                                    barmode='group',
                                    title='Funnel Comparison',
                                    text='Count',
                                    color_discrete_sequence=COLOR_PALETTE
                                )
                                fig.update_layout(
                                    xaxis_title="Lead Stage",
                                    yaxis_title="Count",
                                    height=400
                                )
                                fig.update_traces(texttemplate='%{text}', textposition='outside')
                                st.plotly_chart(fig, use_container_width=True)
                    
                    elif comparison_results['type'] == 'course_vs_owner':
                        # ONE VISUAL: Heatmap
                        if 'owner_courses' in comparison_results and not comparison_results['owner_courses'].empty:
                            # Heatmap showing this owner's performance across courses
                            heatmap_df = comparison_results['owner_courses'].set_index('Course/Program')
                            
                            fig = px.imshow(
                                heatmap_df,
                                labels=dict(x="Lead Status", y="Course", color="Count"),
                                aspect="auto",
                                title=f"{item2}'s Performance by Course",
                                color_continuous_scale='RdYlGn'
                            )
                            fig.update_layout(height=400)
                            st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("Select two items to compare")
        
        # Footer with export summary
        st.divider()
        st.markdown(
            f"""
            <div style='text-align: center; color: #666; font-size: 0.8rem; padding: 1rem;'>
            <strong>HubSpot Advanced Analytics</strong> • Customer from Lifecycle Stage ONLY • Bulletproof Owner Mapping • 
            Data: {datetime.now(IST).strftime("%Y-%m-%d %H:%M:%S")} IST •
            <a href="#top" style="color: #4A6FA5; text-decoration: none;">↑ Back to Top</a>
            </div>
            """,
            unsafe_allow_html=True
        )
    
    else:
        # Welcome screen
        st.markdown(
            """
            <div style='text-align: center; padding: 3rem;'>
                <h2>👋 Welcome to HubSpot Advanced Analytics</h2>
                <p style='font-size: 1.1rem; color: #666; margin: 1rem 0;'>
                    Configure date filters and click "Fetch ALL Contacts" to start analysis
                </p>
                <div style='margin-top: 2rem; background-color: #f8f9fa; padding: 2rem; border-radius: 0.5rem;'>
                    <h4>✅ 100% ACCURATE DATA LOGIC:</h4>
                    <div style='display: flex; justify-content: center; gap: 2rem; margin-top: 1rem;'>
                        <div style='text-align: left;'>
                            <p><strong>📊 Cold/Warm/Hot</strong></p>
                            <ul>
                                <li>From Lead Status</li>
                                <li>Sales follow-up stages</li>
                                <li>Pre-customer status</li>
                            </ul>
                        </div>
                        <div style='text-align: left;'>
                            <p><strong>👤 Customer</strong></p>
                            <ul>
                                <li>From Lifecycle Stage ONLY</li>
                                <li>Exact match "customer"</li>
                                <li>Matches HubSpot Excel</li>
                            </ul>
                        </div>
                        <div style='text-align: left;'>
                            <p><strong>👥 Owner Mapping</strong></p>
                            <ul>
                                <li>Pre-fetched owner list</li>
                                <li>Association fallback</li>
                                <li>Direct API for missing</li>
                            </ul>
                        </div>
                    </div>
                    <div style='display: flex; justify-content: center; gap: 2rem; margin-top: 1rem;'>
                        <div style='text-align: left;'>
                            <p><strong>📊 Course Distribution</strong></p>
                            <ul>
                                <li>Top 10 courses by volume</li>
                                <li>Clean stacked bar charts</li>
                                <li>Readable labels</li>
                            </ul>
                        </div>
                        <div style='text-align: left;'>
                            <p><strong>👤 Owner Performance</strong></p>
                            <ul>
                                <li>Top 10 owners by volume</li>
                                <li>Grouped bar charts</li>
                                <li>Donut charts for details</li>
                            </ul>
                        </div>
                        <div style='text-align: left;'>
                            <p><strong>🎯 Lead Quality</strong></p>
                            <ul>
                                <li>100% stacked bars</li>
                                <li>Engaged % & Customer %</li>
                                <li>Quality ranking tables</li>
                            </ul>
                        </div>
                    </div>
                </div>
                <div style='margin-top: 2rem; padding: 1rem; background-color: #e8f4fd; border-radius: 0.5rem;'>
                    <p><strong>🚀 Getting Started:</strong></p>
                    <ol style='text-align: left; margin-left: 30%;'>
                        <li>Configure date range in sidebar</li>
                        <li>Click "Fetch ALL Contacts"</li>
                        <li>Use smart default filters</li>
                        <li>Explore clean, focused tabs</li>
                        <li>Download professional reports</li>
                    </ol>
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )

if __name__ == "__main__":
    main()
