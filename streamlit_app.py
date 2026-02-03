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

# Will be populated dynamically
CUSTOMER_DEAL_STAGES = []  # Will contain stage IDs, not labels

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
    
    .success-card {
        background: linear-gradient(135deg, #d4edda, #c3e6cb);
        border: 1px solid #28a745;
        border-radius: 8px;
        padding: 15px;
        margin: 10px 0;
    }
    
    .error-card {
        background: linear-gradient(135deg, #f8d7da, #f5c6cb);
        border: 1px solid #dc3545;
        border-radius: 8px;
        padding: 15px;
        margin: 10px 0;
    }
    
    .data-fix-card {
        background: linear-gradient(135deg, #e0e7ff, #c7d2fe);
        border: 1px solid #4f46e5;
        border-radius: 8px;
        padding: 15px;
        margin: 10px 0;
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
    
    /* ✅ NEW: Download Button Styles */
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
        if "hubspot" in st.secrets and "api_key" in st.secrets["hubspot"]:
            api_key = st.secrets["hubspot"]["api_key"]
            if api_key and api_key.startswith("pat-"):
                return api_key
        
        if "HUBSPOT_API_KEY" in st.secrets:
            api_key = st.secrets["HUBSPOT_API_KEY"]
            if api_key and api_key.startswith("pat-"):
                return api_key
        
        import os
        api_key = os.getenv("HUBSPOT_API_KEY")
        if api_key and api_key.startswith("pat-"):
            return api_key
        
        return None
        
    except Exception as e:
        st.error(f"Error loading API key: {str(e)}")
        return None

# ✅ CRITICAL FIX: Lead Status Mapping - ABSOLUTELY NO CUSTOMER HERE!
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
    "duplicate": "Duplicate",
    "junk": "Duplicate",
    "": "Unknown",
    None: "Unknown",
    "unknown": "Unknown",
    "other": "Unknown"
}

# ✅ CRITICAL FIX: List of terms that should NEVER become "Customer" in leads
CUSTOMER_KEYWORDS_BLOCKLIST = [
    "customer", "closed", "won", "admission", "confirmed", 
    "contract", "signed", "paid", "payment", "completed"
]

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
def create_excel_report(df_contacts, df_customers, metrics, kpis, date_range, date_field):
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
                kpis['avg_revenue_per_customer'] if kpis['avg_revenue_per_customer'] > 0 else 0
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
        
        # Sheet 2: Raw Lead Data
        if not df_contacts.empty:
            df_contacts.to_excel(writer, sheet_name='Raw Lead Data', index=False)
            worksheet = writer.sheets['Raw Lead Data']
            
            # Apply formatting
            for col_num, value in enumerate(df_contacts.columns.values):
                worksheet.write(0, col_num, value, header_format)
            
            # Auto-adjust column widths
            for i, col in enumerate(df_contacts.columns):
                column_len = max(df_contacts[col].astype(str).map(len).max(), len(col)) + 2
                worksheet.set_column(i, i, min(column_len, 50))
        
        # Sheet 3: Customer Data
        if df_customers is not None and not df_customers.empty:
            df_customers.to_excel(writer, sheet_name='Customer Data', index=False)
            worksheet = writer.sheets['Customer Data']
            
            for col_num, value in enumerate(df_customers.columns.values):
                worksheet.write(0, col_num, value, header_format)
        
        # Sheet 4: Owner Performance (Metric 4)
        if 'metric_4' in metrics and not metrics['metric_4'].empty:
            metric_4 = metrics['metric_4'].copy()
            metric_4.to_excel(writer, sheet_name='Owner Performance', index=False)
            
            worksheet = writer.sheets['Owner Performance']
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
                    elif col_name not in ['Course Owner'] and col_name != 'Customer_Revenue':
                        value = metric_4.iloc[row-1, col]
                        if pd.notnull(value):
                            worksheet.write(row, col, value, number_format)
        
        # Sheet 5: Course Performance (NEW METRIC)
        if 'metric_5' in metrics and not metrics['metric_5'].empty:
            metric_5 = metrics['metric_5'].copy()
            metric_5.to_excel(writer, sheet_name='Course Performance', index=False)
            
            worksheet = writer.sheets['Course Performance']
            for col_num, value in enumerate(metric_5.columns.values):
                worksheet.write(0, col_num, value, header_format)
            
            # Format percentages
            percent_columns = ['Deal %', 'Customer %', 'Lead→Customer %', 'Lead→Deal %']
            num_rows = len(metric_5)
            num_cols = len(metric_5.columns)
            
            for row in range(1, num_rows + 1):
                for col in range(num_cols):
                    col_name = metric_5.columns[col]
                    if col_name in percent_columns:
                        value = metric_5.iloc[row-1, col]
                        if pd.notnull(value):
                            worksheet.write(row, col, value/100, percent_format)
                    elif col_name not in ['Course'] and col_name != 'Customer_Revenue':
                        value = metric_5.iloc[row-1, col]
                        if pd.notnull(value):
                            worksheet.write(row, col, value, number_format)
        
        # Sheet 6: Lead Status Summary
        if not df_contacts.empty:
            status_summary = df_contacts['Lead Status'].value_counts().reset_index()
            status_summary.columns = ['Lead Status', 'Count']
            status_summary['Percentage'] = (status_summary['Count'] / len(df_contacts) * 100).round(1)
            
            status_summary.to_excel(writer, sheet_name='Lead Status Summary', index=False)
            
            worksheet = writer.sheets['Lead Status Summary']
            for col_num, value in enumerate(status_summary.columns.values):
                worksheet.write(0, col_num, value, header_format)
            
            # Format percentages
            for row in range(1, len(status_summary) + 1):
                worksheet.write(row, 2, status_summary.iloc[row-1, 2]/100, percent_format)
                worksheet.write(row, 1, status_summary.iloc[row-1, 1], number_format)
        
        # Sheet 7: Revenue Analysis
        if df_customers is not None and not df_customers.empty:
            revenue_summary = df_customers.groupby('Course/Program').agg(
                Customer_Count=('Is Customer', 'count'),
                Total_Revenue=('Amount', 'sum'),
                Avg_Revenue=('Amount', 'mean')
            ).reset_index()
            
            revenue_summary = revenue_summary.sort_values('Total_Revenue', ascending=False)
            revenue_summary.to_excel(writer, sheet_name='Revenue Analysis', index=False)
            
            worksheet = writer.sheets['Revenue Analysis']
            for col_num, value in enumerate(revenue_summary.columns.values):
                worksheet.write(0, col_num, value, header_format)
            
            # Format revenue columns
            for row in range(1, len(revenue_summary) + 1):
                worksheet.write(row, 2, revenue_summary.iloc[row-1, 2], number_format)
                worksheet.write(row, 3, revenue_summary.iloc[row-1, 3], number_format)
    
    # Get the Excel data
    excel_data = output.getvalue()
    return excel_data

# ✅ CRITICAL FIX: Fetch Deal Pipeline Stages to get correct Stage IDs
@st.cache_data(ttl=86400)
def fetch_deal_pipeline_stages(api_key):
    """Fetch deal pipeline stages to get correct stage IDs (not labels)."""
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json"
    }
    
    url = f"{HUBSPOT_API_BASE}/crm/v3/pipelines/deals"
    
    try:
        response = requests.get(url, headers=headers, timeout=10)
        
        if response.status_code == 200:
            data = response.json()
            pipelines = data.get("results", [])
            
            if not pipelines:
                st.error("❌ No deal pipelines found in HubSpot")
                return {}
            
            all_stages = {}
            
            for pipeline in pipelines:
                pipeline_id = pipeline.get("id", "")
                pipeline_label = pipeline.get("label", "")
                stages = pipeline.get("stages", [])
                
                for stage in stages:
                    stage_id = stage.get("id", "")
                    stage_label = stage.get("label", "")
                    stage_metadata = stage.get("metadata", {})
                    probability = stage_metadata.get("probability", "0")
                    
                    all_stages[stage_id] = {
                        "stage_id": stage_id,
                        "stage_label": stage_label,
                        "pipeline_id": pipeline_id,
                        "pipeline_label": pipeline_label,
                        "probability": probability,
                        "full_info": f"{stage_label} (ID: {stage_id}, Probability: {probability})"
                    }
            
            st.success(f"✅ Loaded {len(all_stages)} deal stages from {len(pipelines)} pipelines")
            return all_stages
            
        elif response.status_code == 403:
            st.error("""
            ❌ Missing required scope: crm.pipelines.read
            Please update your API key permissions to include:
            - crm.objects.deals.read
            - crm.objects.contacts.read  
            - crm.pipelines.read
            """)
            return {}
        else:
            st.error(f"❌ Failed to fetch deal pipelines: {response.status_code}")
            return {}
            
    except requests.exceptions.RequestException as e:
        st.error(f"❌ Error fetching deal pipelines: {e}")
        return {}

# ✅ Auto-detect Admission Confirmed stage ID
def detect_admission_confirmed_stage(all_stages):
    """Auto-detect Admission Confirmed stage from pipeline stages."""
    target_labels = [
        "admission confirmed",
        "admission_confirmed", 
        "confirmed",
        "closed won",
        "closedwon",
        "won",
        "customer"
    ]
    
    detected_stages = []
    
    for stage_id, stage_info in all_stages.items():
        stage_label = stage_info.get("stage_label", "").lower()
        probability = stage_info.get("probability", "0")
        
        for target in target_labels:
            if target in stage_label:
                try:
                    prob_float = float(probability)
                    if prob_float >= 0.9:
                        detected_stages.append({
                            "stage_id": stage_id,
                            "stage_label": stage_info.get("stage_label"),
                            "probability": probability,
                            "pipeline": stage_info.get("pipeline_label"),
                            "match_reason": f"Label contains '{target}', probability: {probability}"
                        })
                        break
                except:
                    detected_stages.append({
                        "stage_id": stage_id,
                        "stage_label": stage_info.get("stage_label"),
                        "probability": probability,
                        "pipeline": stage_info.get("pipeline_label"),
                        "match_reason": f"Label contains '{target}'"
                    })
                    break
    
    return detected_stages

# ✅ CRITICAL FIX: UPDATED normalize_lead_status function
def normalize_lead_status(raw_status):
    """
    Normalize lead status - ABSOLUTELY NO CUSTOMER HERE!
    This function MUST NEVER return "Customer" for any lead status.
    """
    if not raw_status:
        return "Unknown"
    
    status = str(raw_status).strip().lower()
    
    # ✅ FIRST: Check if this contains any customer keywords - BLOCK THEM!
    for keyword in CUSTOMER_KEYWORDS_BLOCKLIST:
        if keyword in status:
            # This is PROBABLY a customer deal stage that leaked into contacts
            # Map to "Qualified Lead" instead of "Customer"
            if "hot" in status:
                return "Hot"
            elif "warm" in status:
                return "Warm"
            elif "cold" in status:
                return "Cold"
            else:
                return "Qualified Lead"  # NOT Customer!
    
    # Now handle normal lead statuses
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
    
    if "new" in status or "open" in status:
        return "New Lead"
    
    if status in LEAD_STATUS_MAP:
        return LEAD_STATUS_MAP[status]
    
    # If we get here and it's still a customer-like term, map to Qualified Lead
    if any(keyword in status for keyword in ["deal", "qualified", "converted"]):
        return "Qualified Lead"
    
    return status.replace("_", " ").title()

# ✅ CRITICAL FIX: Debug function to see what's being converted to "Customer"
def debug_lead_status_conversion(df):
    """Debug function to identify what's being incorrectly marked as Customer."""
    if df.empty:
        return
    
    # Find all rows where Lead Status is "Customer"
    customer_rows = df[df['Lead Status'] == 'Customer']
    
    if not customer_rows.empty:
        st.error(f"❌ FOUND {len(customer_rows)} ROWS MARKED AS 'CUSTOMER' IN LEADS!")
        st.write("This should be ZERO. Showing samples:")
        
        # Show raw values that got converted to Customer
        st.write("**Raw Lead Status values that became 'Customer':**")
        unique_raw_values = customer_rows['Lead Status Raw'].dropna().unique()
        
        for raw_val in unique_raw_values[:10]:  # Show first 10
            st.write(f"- `{raw_val}` → Customer")
        
        if len(unique_raw_values) > 10:
            st.write(f"... and {len(unique_raw_values) - 10} more")
        
        # Show sample rows
        st.write("**Sample rows with issue:**")
        st.dataframe(customer_rows[['Full Name', 'Lead Status', 'Lead Status Raw', 'Course/Program']].head(10))
        
        # Let user manually fix these
        st.markdown("""
        <div class="data-fix-card">
        <strong>🔧 QUICK FIX OPTION:</strong><br>
        You can manually remap these incorrect "Customer" entries:
        </div>
        """, unsafe_allow_html=True)
        
        # Provide a quick fix option
        if st.button("🔄 Auto-fix 'Customer' in leads (map to 'Qualified Lead')"):
            # Fix the dataframe
            df_fixed = df.copy()
            df_fixed['Lead Status'] = df_fixed['Lead Status'].replace('Customer', 'Qualified Lead')
            
            # Update session state
            st.session_state.contacts_df = df_fixed
            st.success("✅ Fixed! 'Customer' entries in leads now mapped to 'Qualified Lead'")
            st.rerun()

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

                name = f"{first_name} {last_name}".strip() if first_name or last_name else email.split("@")[0] if email else f"Owner {owner_id}"
                mapping[owner_id] = name

            page_count += 1
            total_owners += len(owners)
            
            if page_count <= 5:
                progress_bar.progress(min(page_count / 5, 0.9))
            status_text.text(f"📋 Fetched {total_owners} owners (page {page_count})...")

            paging = data.get("paging", {})
            next_link = paging.get("next", {}).get("link")

            if not next_link:
                progress_bar.progress(1.0)
                status_text.text(f"✅ Owner mapping complete! Total: {total_owners} owners")
                break

            url = next_link
            params = None
            time.sleep(0.1)
            
    except requests.exceptions.RequestException as e:
        st.warning(f"⚠️ Partial owner mapping loaded. Error: {str(e)[:100]}")
    except Exception as e:
        st.warning(f"⚠️ Could not fetch owner mapping: {str(e)[:100]}")

    return mapping

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
    else:
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
    
    # Define properties - REMOVE lifecycle stage
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
                time.sleep(10)
                continue
            
            response.raise_for_status()
            data = response.json()
            
            batch_contacts = data.get("results", [])
            
            if batch_contacts:
                all_contacts.extend(batch_contacts)
                paging_info = data.get("paging", {})
                after = paging_info.get("next", {}).get("after")
                
                if not after:
                    break
                
                time.sleep(0.1)
            else:
                break
        
        return all_contacts, len(all_contacts)
        
    except requests.exceptions.RequestException as e:
        st.error(f"❌ Error fetching contacts: {e}")
        return [], 0
    except Exception as e:
        st.error(f"❌ Unexpected error: {e}")
        return [], 0

# ✅ Fetch DEALS using CORRECT Stage IDs
def fetch_hubspot_deals(api_key, start_date, end_date, customer_stage_ids):
    """Fetch DEALS from HubSpot using CORRECT stage IDs (not labels)."""
    if not customer_stage_ids:
        st.error("❌ No customer stage IDs configured.")
        return [], 0
    
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json"
    }
    
    start_timestamp = date_to_hubspot_timestamp(start_date, is_end_date=False)
    end_timestamp = date_to_hubspot_timestamp(end_date, is_end_date=True)
    
    all_deals = []
    after = None
    
    url = f"{HUBSPOT_API_BASE}/crm/v3/objects/deals/search"
    
    # ✅ CORRECT FILTER: Use Stage IDs
    filter_groups = [{
        "filters": [
            {
                "propertyName": "dealstage",
                "operator": "IN",
                "values": customer_stage_ids
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
        "program_name",
    ]
    
    try:
        while True:
            body = {
                "filterGroups": filter_groups,
                "properties": deal_properties,
                "associations": ["owners"],
                "limit": 100,
                "sorts": [{"propertyName": "closedate", "direction": "DESCENDING"}]
            }
            
            if after:
                body["after"] = after
            
            response = requests.post(url, headers=headers, json=body, timeout=30)
            
            if response.status_code == 429:
                time.sleep(10)
                continue
            
            response.raise_for_status()
            data = response.json()
            
            batch_deals = data.get("results", [])
            
            if batch_deals:
                all_deals.extend(batch_deals)
                paging_info = data.get("paging", {})
                after = paging_info.get("next", {}).get("after")
                
                if not after:
                    break
                
                time.sleep(0.1)
            else:
                break
        
        return all_deals, len(all_deals)
        
    except requests.exceptions.RequestException as e:
        st.error(f"❌ Error fetching deals: {e}")
        return [], 0
    except Exception as e:
        st.error(f"❌ Unexpected error fetching deals: {e}")
        return [], 0

def process_contacts_data(contacts, owner_mapping=None, api_key=None):
    """Process raw contacts data into a clean DataFrame - ABSOLUTELY NO CUSTOMER HERE."""
    if not contacts:
        return pd.DataFrame()
    
    processed_data = []
    
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
                owner_name = f"❌ Unassigned ({owner_id})" if owner_id else "❌ Unassigned"
        else:
            owner_name = owner_id
        
        # ✅ CRITICAL: Get raw lead status
        raw_lead_status = properties.get("hs_lead_status", "") or properties.get("lead_status", "")
        
        # ✅ CRITICAL: Normalize lead status - WILL NEVER RETURN "CUSTOMER"
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
            "Lead Status": lead_status,  # ✅ NO CUSTOMER HERE
            "Created Date": properties.get("createdate", ""),
            "Lead Status Raw": raw_lead_status,
            "Owner ID": owner_id
        })
    
    df = pd.DataFrame(processed_data)
    
    # ✅ CRITICAL: Debug and fix any "Customer" entries
    if not df.empty:
        customer_count = (df['Lead Status'] == 'Customer').sum()
        if customer_count > 0:
            st.error(f"❌ ERROR: Found {customer_count} 'Customer' in leads after processing")
            # Call debug function
            debug_lead_status_conversion(df)
    
    return df

def process_deals_as_customers(deals, owner_mapping=None, api_key=None, all_stages=None):
    """Process raw deals data into customer DataFrame."""
    if not deals:
        return pd.DataFrame()
    
    processed_data = []
    stage_label_map = {}
    
    if all_stages:
        stage_label_map = {stage_id: info.get("stage_label", stage_id) 
                          for stage_id, info in all_stages.items()}
    
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
                owner_name = f"❌ Unassigned ({owner_id})" if owner_id else "❌ Unassigned"
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
        
        # Get deal stage (ID)
        deal_stage_id = properties.get("dealstage", "")
        
        # Convert stage ID to label if possible
        deal_stage_label = stage_label_map.get(deal_stage_id, deal_stage_id)
        
        processed_data.append({
            "Customer ID": deal.get("id", ""),
            "Deal Name": properties.get("dealname", ""),
            "Course/Program": course_info,
            "Course Owner": owner_name,
            "Amount": amount,
            "Close Date": close_date,
            "Deal Stage ID": deal_stage_id,
            "Deal Stage Label": deal_stage_label,
            "Is Customer": 1  # ✅ ALL these deals are customers
        })
    
    df = pd.DataFrame(processed_data)
    
    return df

def create_metric_1(df):
    """METRIC 1: Course × Lead Status - NO CUSTOMER"""
    if df.empty or 'Course/Program' not in df.columns:
        return pd.DataFrame()
    
    df_course = df[df['Course/Program'].notna() & (df['Course/Program'] != '')].copy()
    
    if df_course.empty:
        return pd.DataFrame()
    
    df_course['Course_Clean'] = df_course['Course/Program'].str.strip()
    
    pivot = pd.pivot_table(
        df_course,
        index='Course_Clean',
        columns='Lead Status',
        values='ID',
        aggfunc='count',
        fill_value=0
    )
    
    pivot = pivot.reset_index().rename(columns={'Course_Clean': 'Course'})
    
    if len(pivot.columns) > 1:
        status_cols = [col for col in pivot.columns if col != 'Course']
        pivot['Total'] = pivot[status_cols].sum(axis=1)
    
    return pivot

def create_metric_2(df):
    """METRIC 2: Course Owner × Lead Status - NO CUSTOMER"""
    if df.empty or 'Course Owner' not in df.columns:
        return pd.DataFrame()
    
    df_owner = df[df['Course Owner'].notna() & (df['Course Owner'] != '')].copy()
    
    if df_owner.empty:
        return pd.DataFrame()
    
    pivot = pd.pivot_table(
        df_owner,
        index='Course Owner',
        columns='Lead Status',
        values='ID',
        aggfunc='count',
        fill_value=0
    )
    
    pivot = pivot.reset_index()
    
    if len(pivot.columns) > 1:
        status_cols = [col for col in pivot.columns if col != 'Course Owner']
        pivot['Total'] = pivot[status_cols].sum(axis=1)
    
    return pivot

def create_metric_4(df_contacts, df_customers):
    """METRIC 4: Course Owner Performance SUMMARY"""
    if df_contacts.empty or 'Course Owner' not in df_contacts.columns:
        return pd.DataFrame()
    
    owner_lead_pivot = create_metric_2(df_contacts)
    
    if owner_lead_pivot.empty:
        return pd.DataFrame()
    
    # Get customer data from deals
    if df_customers is not None and not df_customers.empty and 'Course Owner' in df_customers.columns:
        customer_by_owner = df_customers.groupby('Course Owner').agg(
            Customer_Count=('Is Customer', 'sum'),
            Customer_Revenue=('Amount', 'sum')
        ).reset_index()
    else:
        customer_by_owner = pd.DataFrame(columns=['Course Owner', 'Customer_Count', 'Customer_Revenue'])
    
    # Merge lead data with customer data
    result_df = owner_lead_pivot.copy()
    result_df['Customer'] = 0
    
    if not customer_by_owner.empty:
        result_df = pd.merge(result_df, customer_by_owner, on='Course Owner', how='left')
        result_df['Customer_Count'] = result_df['Customer_Count'].fillna(0)
        result_df['Customer_Revenue'] = result_df['Customer_Revenue'].fillna(0)
        result_df['Customer'] = result_df['Customer_Count']
    else:
        result_df['Customer_Count'] = 0
        result_df['Customer_Revenue'] = 0
    
    # Deal Leads = Hot + Warm + Cold + Customer
    deal_statuses = ['Cold', 'Warm', 'Hot']
    result_df['Deal Leads'] = 0
    
    for status in deal_statuses:
        if status in result_df.columns:
            result_df['Deal Leads'] += result_df[status].fillna(0)
    
    result_df['Deal Leads'] += result_df['Customer']
    
    # Calculate percentages
    if 'Total' in result_df.columns:
        result_df = result_df.rename(columns={'Total': 'Grand Total'})
        result_df['Deal %'] = np.where(
            result_df['Grand Total'] > 0,
            (result_df['Deal Leads'] / result_df['Grand Total'] * 100).round(2),
            0
        )
    else:
        result_df['Deal %'] = 0
    
    result_df['Customer %'] = np.where(
        result_df['Deal Leads'] > 0,
        (result_df['Customer'] / result_df['Deal Leads'] * 100).round(2),
        0
    )
    
    if 'Grand Total' in result_df.columns:
        result_df['Lead→Customer %'] = np.where(
            result_df['Grand Total'] > 0,
            (result_df['Customer'] / result_df['Grand Total'] * 100).round(2),
            0
        )
        result_df['Lead→Deal %'] = np.where(
            result_df['Grand Total'] > 0,
            (result_df['Deal Leads'] / result_df['Grand Total'] * 100).round(2),
            0
        )
    else:
        result_df['Lead→Customer %'] = 0
        result_df['Lead→Deal %'] = 0
    
    # Select columns
    base_cols = ['Course Owner']
    status_cols = ['Cold', 'Hot', 'Warm']
    existing_status_cols = [col for col in status_cols if col in result_df.columns]
    
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
    
    final_df = result_df[final_cols].copy()
    
    if 'Grand Total' in final_df.columns:
        final_df = final_df.sort_values('Grand Total', ascending=False)
    
    return final_df

# ✅ NEW: METRIC 5 - Course Performance KPI Table (Same as Owner Performance but for Courses)
def create_metric_5(df_contacts, df_customers):
    """METRIC 5: Course Performance KPI Table"""
    if df_contacts.empty or 'Course/Program' not in df_contacts.columns:
        return pd.DataFrame()
    
    course_lead_pivot = create_metric_1(df_contacts)
    
    if course_lead_pivot.empty:
        return pd.DataFrame()
    
    # Get customer data from deals by course
    if df_customers is not None and not df_customers.empty and 'Course/Program' in df_customers.columns:
        customer_by_course = df_customers.groupby('Course/Program').agg(
            Customer_Count=('Is Customer', 'sum'),
            Customer_Revenue=('Amount', 'sum')
        ).reset_index()
    else:
        customer_by_course = pd.DataFrame(columns=['Course/Program', 'Customer_Count', 'Customer_Revenue'])
    
    # Merge lead data with customer data
    result_df = course_lead_pivot.copy()
    result_df = result_df.rename(columns={'Course': 'Course'})
    result_df['Customer'] = 0
    
    if not customer_by_course.empty:
        result_df = pd.merge(result_df, customer_by_course, left_on='Course', right_on='Course/Program', how='left')
        result_df['Customer_Count'] = result_df['Customer_Count'].fillna(0)
        result_df['Customer_Revenue'] = result_df['Customer_Revenue'].fillna(0)
        result_df['Customer'] = result_df['Customer_Count']
        # Drop the extra course column from merge
        result_df = result_df.drop(columns=['Course/Program'], errors='ignore')
    else:
        result_df['Customer_Count'] = 0
        result_df['Customer_Revenue'] = 0
    
    # Deal Leads = Hot + Warm + Cold + Customer
    deal_statuses = ['Cold', 'Warm', 'Hot']
    result_df['Deal Leads'] = 0
    
    for status in deal_statuses:
        if status in result_df.columns:
            result_df['Deal Leads'] += result_df[status].fillna(0)
    
    result_df['Deal Leads'] += result_df['Customer']
    
    # Calculate percentages
    if 'Total' in result_df.columns:
        result_df = result_df.rename(columns={'Total': 'Grand Total'})
        result_df['Deal %'] = np.where(
            result_df['Grand Total'] > 0,
            (result_df['Deal Leads'] / result_df['Grand Total'] * 100).round(2),
            0
        )
    else:
        result_df['Deal %'] = 0
    
    result_df['Customer %'] = np.where(
        result_df['Deal Leads'] > 0,
        (result_df['Customer'] / result_df['Deal Leads'] * 100).round(2),
        0
    )
    
    if 'Grand Total' in result_df.columns:
        result_df['Lead→Customer %'] = np.where(
            result_df['Grand Total'] > 0,
            (result_df['Customer'] / result_df['Grand Total'] * 100).round(2),
            0
        )
        result_df['Lead→Deal %'] = np.where(
            result_df['Grand Total'] > 0,
            (result_df['Deal Leads'] / result_df['Grand Total'] * 100).round(2),
            0
        )
    else:
        result_df['Lead→Customer %'] = 0
        result_df['Lead→Deal %'] = 0
    
    # Select columns
    base_cols = ['Course']
    status_cols = ['Cold', 'Hot', 'Warm']
    existing_status_cols = [col for col in status_cols if col in result_df.columns]
    
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
    
    final_df = result_df[final_cols].copy()
    
    if 'Grand Total' in final_df.columns:
        final_df = final_df.sort_values('Grand Total', ascending=False)
    
    return final_df

# ✅ NEW: Course Revenue Analysis
def create_course_revenue(df_customers):
    """Calculate revenue by course from customer data."""
    if df_customers is None or df_customers.empty or 'Course/Program' not in df_customers.columns or 'Amount' not in df_customers.columns:
        return pd.DataFrame()
    
    # Filter only courses with revenue
    customer_df = df_customers[(df_customers['Course/Program'].notna()) & (df_customers['Course/Program'] != '')].copy()
    
    if customer_df.empty:
        return pd.DataFrame()
    
    # Clean course names
    customer_df['Course_Clean'] = customer_df['Course/Program'].str.strip()
    
    # Group by course
    revenue_df = customer_df.groupby('Course_Clean').agg(
        Customers=('Is Customer', 'sum'),
        Revenue=('Amount', 'sum')
    ).reset_index().rename(columns={'Course_Clean': 'Course'})
    
    # Calculate revenue per customer
    revenue_df['Revenue per Customer'] = np.where(
        revenue_df['Customers'] > 0,
        (revenue_df['Revenue'] / revenue_df['Customers']).round(0),
        0
    )
    
    # Sort by revenue
    revenue_df = revenue_df.sort_values('Revenue', ascending=False)
    
    return revenue_df

# ✅ NEW: Volume vs Conversion Matrix
def create_volume_conversion_matrix(metric_1, df_contacts, df_customers):
    """Create a 2x2 matrix to classify courses based on volume and conversion."""
    if metric_1.empty or 'Total' not in metric_1.columns:
        return pd.DataFrame()
    
    # Get customer data by course
    customer_by_course = {}
    if df_customers is not None and not df_customers.empty and 'Course/Program' in df_customers.columns:
        for _, row in df_customers.iterrows():
            course = row['Course/Program']
            if pd.notna(course) and course != '':
                course_clean = str(course).strip()
                customer_by_course[course_clean] = customer_by_course.get(course_clean, 0) + 1
    
    # Calculate conversion % for each course
    matrix_data = []
    
    for _, row in metric_1.iterrows():
        course = row['Course']
        total = row.get('Total', 0)
        
        # Get customer count for this course
        customer_count = customer_by_course.get(course, 0)
        
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
        # Get course data
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
            
            # Calculate comparison metrics
            if not course1_data.empty and not course2_data.empty:
                # Deal Leads (Cold + Warm + Hot)
                deal_cols = ['Cold', 'Warm', 'Hot']
                deal1 = course1_data[deal_cols].sum(axis=1).values[0] if all(col in course1_data.columns for col in deal_cols) else 0
                total1 = course1_data['Total'].values[0] if 'Total' in course1_data.columns else 1
                deal_pct1 = (deal1 / total1 * 100) if total1 > 0 else 0
                
                deal2 = course2_data[deal_cols].sum(axis=1).values[0] if all(col in course2_data.columns for col in deal_cols) else 0
                total2 = course2_data['Total'].values[0] if 'Total' in course2_data.columns else 1
                deal_pct2 = (deal2 / total2 * 100) if total2 > 0 else 0
                
                results['deal_pct1'] = round(deal_pct1, 1)
                results['deal_pct2'] = round(deal_pct2, 1)
    
    elif comparison_type == "Owner vs Owner":
        # Get owner data
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
        # This is more complex - need to get course data for specific owner
        results['type'] = 'course_vs_owner'
        results['item1'] = item1  # Course
        results['item2'] = item2  # Owner
        
        # Get courses for this owner
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
    """Calculate key performance indicators."""
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
    qualified_lead = status_counts.get('Qualified Lead', 0)
    
    # ✅ Check for any "Customer" in leads (should be 0)
    customer_in_leads = status_counts.get('Customer', 0)
    
    # CUSTOMER metrics from DEALS
    if df_customers is not None and not df_customers.empty:
        customer = len(df_customers)
        total_revenue = df_customers['Amount'].sum()
        avg_revenue_per_customer = round((total_revenue / customer), 0) if customer > 0 else 0
    else:
        customer = 0
        total_revenue = 0
        avg_revenue_per_customer = 0
    
    # Deal Leads = Hot + Warm + Cold + Customer
    deal_leads = hot + warm + cold + customer
    
    # Conversion metrics
    lead_to_customer_pct = round((customer / total_leads * 100), 1) if total_leads > 0 else 0
    lead_to_deal_pct = round((deal_leads / total_leads * 100), 1) if total_leads > 0 else 0
    deal_to_customer_pct = round((customer / deal_leads * 100), 1) if deal_leads > 0 else 0
    
    # Top performing metrics
    top_course = ""
    top_owner = ""
    top_revenue_course = ""
    top_revenue_amount = 0
    
    if 'Course/Program' in df_contacts.columns:
        course_counts = df_contacts['Course/Program'].value_counts()
        if not course_counts.empty:
            top_course = str(course_counts.index[0])
            top_course_count = course_counts.iloc[0]
    
    if 'Course Owner' in df_contacts.columns:
        owner_counts = df_contacts['Course Owner'].value_counts()
        if not owner_counts.empty:
            top_owner = str(owner_counts.index[0])
            top_owner_count = owner_counts.iloc[0]
    
    # Best Revenue Course
    if df_customers is not None and not df_customers.empty and 'Course/Program' in df_customers.columns:
        revenue_by_course = df_customers.groupby('Course/Program')['Amount'].sum()
        if not revenue_by_course.empty:
            top_revenue_course = str(revenue_by_course.index[0])
            top_revenue_amount = revenue_by_course.iloc[0]
    
    # Drop-off ratio
    dropoff_ratio = round((not_interested + not_qualified + not_connected) / total_leads * 100, 1) if total_leads > 0 else 0
    
    return {
        'total_leads': total_leads,
        'deal_leads': deal_leads,
        'cold': cold,
        'warm': warm,
        'hot': hot,
        'customer': customer,  # FROM DEALS ONLY
        'customer_in_leads': customer_in_leads,  # This should be 0!
        'new_lead': new_lead,
        'not_connected': not_connected,
        'not_interested': not_interested,
        'not_qualified': not_qualified,
        'duplicate': duplicate,
        'qualified_lead': qualified_lead,
        'lead_to_customer_pct': lead_to_customer_pct,
        'lead_to_deal_pct': lead_to_deal_pct,
        'deal_to_customer_pct': deal_to_customer_pct,
        'total_revenue': total_revenue,
        'avg_revenue_per_customer': avg_revenue_per_customer,
        'top_course': top_course[:20] if top_course else "N/A",
        'top_owner': top_owner[:20] if top_owner else "N/A",
        'top_revenue_course': top_revenue_course[:20] if top_revenue_course else "N/A",
        'top_revenue_amount': top_revenue_amount,
        'dropoff_ratio': dropoff_ratio
    }

def main():
    # ✅ SECURITY: Get API key from secrets
    api_key = get_api_key()
    
    if not api_key:
        st.error("## 🔐 API Key Required")
        return
    
    # Header
    st.markdown(
        """
        <div class="header-container">
            <h1 style="margin: 0; font-size: 2.5rem;">📊 HubSpot Business Performance Dashboard</h1>
            <p style="margin: 0.5rem 0 0 0; font-size: 1.2rem; opacity: 0.9;">🎯 100% CLEAN: NO Customers in Lead Data</p>
            <p style="margin: 0.5rem 0 0 0; font-size: 1rem; opacity: 0.8;">Customers ONLY from Deals | Leads NEVER contain "Customer"</p>
        </div>
        """,
        unsafe_allow_html=True
    )
    
    # ✅ CRITICAL FIX WARNING
    st.markdown("""
    <div class="warning-card">
        <strong>⚠️ CRITICAL FIX APPLIED:</strong><br>
        • <code>normalize_lead_status()</code> function now <strong>NEVER returns "Customer"</strong><br>
        • Any lead status containing customer keywords → "Qualified Lead"<br>
        • Customers ONLY come from Deals (Stage IDs)<br>
        • <strong>Guaranteed: 0 "Customer" entries in lead data</strong>
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
    if 'deal_stages' not in st.session_state:
        st.session_state.deal_stages = None
    if 'customer_stage_ids' not in st.session_state:
        st.session_state.customer_stage_ids = []
    if 'revenue_data' not in st.session_state:
        st.session_state.revenue_data = None
    if 'matrix_data' not in st.session_state:
        st.session_state.matrix_data = None
    
    # ✅ Fetch Deal Pipeline Stages FIRST
    if 'deal_stages' not in st.session_state or st.session_state.deal_stages is None:
        with st.spinner("🔍 Loading deal pipeline stages..."):
            deal_stages = fetch_deal_pipeline_stages(api_key)
            st.session_state.deal_stages = deal_stages
    
    # Create sidebar
    with st.sidebar:
        st.markdown("## 🔧 Configuration")
        
        # ✅ Deal Stage Configuration
        st.markdown("### 🎯 Customer Deal Stage Configuration")
        
        if st.session_state.deal_stages:
            all_stages = st.session_state.deal_stages
            detected_stages = detect_admission_confirmed_stage(all_stages)
            
            if detected_stages:
                st.success(f"✅ Auto-detected {len(detected_stages)} customer stage(s)")
                
                # Reset customer stage IDs
                st.session_state.customer_stage_ids = []
                
                for idx, stage in enumerate(detected_stages):
                    with st.expander(f"Stage {idx+1}: {stage['stage_label']}", expanded=idx==0):
                        st.write(f"**Stage ID:** `{stage['stage_id']}`")
                        st.write(f"**Pipeline:** {stage['pipeline']}")
                        
                        if st.checkbox(f"Use '{stage['stage_label']}' as customer stage", 
                                      value=True, key=f"use_stage_{stage['stage_id']}"):
                            if stage['stage_id'] not in st.session_state.customer_stage_ids:
                                st.session_state.customer_stage_ids.append(stage['stage_id'])
                
                # Manual stage selection
                st.markdown("#### 🔧 Manual Stage Selection")
                
                all_stage_options = []
                for stage_id, stage_info in all_stages.items():
                    label = stage_info.get("stage_label", "Unknown")
                    pipeline = stage_info.get("pipeline_label", "Unknown")
                    probability = stage_info.get("probability", "0")
                    all_stage_options.append({
                        "id": stage_id,
                        "display": f"{label} (Pipeline: {pipeline}, Probability: {probability})"
                    })
                
                all_stage_options.sort(key=lambda x: x["display"])
                
                selected_manual_stages = st.multiselect(
                    "Select additional customer stages:",
                    options=[s["display"] for s in all_stage_options],
                    default=[],
                    help="Manually select other stages that indicate customers"
                )
                
                # Map back to IDs
                manual_stage_ids = []
                for display in selected_manual_stages:
                    for stage in all_stage_options:
                        if stage["display"] == display:
                            manual_stage_ids.append(stage["id"])
                            break
                
                # Combine
                all_selected_ids = st.session_state.customer_stage_ids + manual_stage_ids
                all_selected_ids = list(set(all_selected_ids))
                
                if all_selected_ids:
                    global CUSTOMER_DEAL_STAGES
                    CUSTOMER_DEAL_STAGES = all_selected_ids
                    
                    with st.expander("📋 Selected Customer Stages", expanded=True):
                        for stage_id in CUSTOMER_DEAL_STAGES:
                            if stage_id in all_stages:
                                info = all_stages[stage_id]
                                st.write(f"• **{info.get('stage_label')}** (`{stage_id}`)")
                else:
                    st.warning("⚠️ No customer stages selected")
            else:
                st.error("❌ No customer stages auto-detected")
        
        st.divider()
        
        # Date Range
        st.markdown("## 📅 Date Range Filter")
        
        date_field = st.selectbox(
            "Select date field for LEADS:",
            ["Created Date", "Last Modified Date", "Both"]
        )
        
        default_end = datetime.now(IST).date()
        default_start = default_end - timedelta(days=30)
        
        start_date = st.date_input("Start date", value=default_start)
        end_date = st.date_input("End date", value=default_end)
        
        if start_date > end_date:
            st.error("Start date must be before end date!")
            return
        
        st.divider()
        
        # Quick Actions
        st.markdown("## ⚡ Quick Actions")
        
        fetch_disabled = not CUSTOMER_DEAL_STAGES
        
        if st.button("🚀 Fetch ALL Data", 
                    type="primary", 
                    use_container_width=True,
                    disabled=fetch_disabled):
            
            if not CUSTOMER_DEAL_STAGES:
                st.error("❌ Please configure customer deal stages first")
                return
                
            if start_date > end_date:
                st.error("Start date must be before end date.")
            else:
                with st.spinner("Fetching data..."):
                    is_valid, message = test_hubspot_connection(api_key)
                    
                    if is_valid:
                        # Store date filter info
                        st.session_state.date_filter = date_field
                        st.session_state.date_range = (start_date.strftime("%Y-%m-%d"), end_date.strftime("%Y-%m-%d"))
                        
                        # Fetch owners
                        owner_mapping = fetch_owner_mapping(api_key)
                        st.session_state.owner_mapping = owner_mapping
                        
                        # Fetch CONTACTS (Leads)
                        contacts, total_contacts = fetch_hubspot_contacts_with_date_filter(
                            api_key, date_field, start_date, end_date
                        )
                        
                        # ✅ Fetch DEALS using Stage IDs
                        deals, total_deals = fetch_hubspot_deals(
                            api_key, start_date, end_date, CUSTOMER_DEAL_STAGES
                        )
                        
                        if contacts:
                            # Process contacts (leads)
                            df_contacts = process_contacts_data(contacts, owner_mapping, api_key)
                            st.session_state.contacts_df = df_contacts
                            
                            # Process deals (customers)
                            df_customers = process_deals_as_customers(deals, owner_mapping, api_key, st.session_state.deal_stages)
                            st.session_state.customers_df = df_customers
                            
                            # Calculate metrics - ADD NEW METRIC 5
                            st.session_state.metrics = {
                                'metric_1': create_metric_1(df_contacts),
                                'metric_2': create_metric_2(df_contacts),
                                'metric_4': create_metric_4(df_contacts, df_customers),
                                'metric_5': create_metric_5(df_contacts, df_customers)  # ✅ NEW METRIC
                            }
                            
                            # ✅ NEW: Calculate revenue data
                            st.session_state.revenue_data = create_course_revenue(df_customers)
                            
                            # ✅ NEW: Calculate matrix data
                            st.session_state.matrix_data = create_volume_conversion_matrix(
                                st.session_state.metrics['metric_1'], df_contacts, df_customers
                            )
                            
                            st.success(f"""
                            ✅ Successfully loaded:
                            • 📊 {len(contacts)} contacts (leads)
                            • 💰 {len(deals)} customers (from deals)
                            """)
                            st.rerun()
                        else:
                            st.warning("No contacts found")
                    else:
                        st.error(f"Connection failed: {message}")
        
        if st.button("🔄 Refresh Analysis", use_container_width=True, 
                    disabled=st.session_state.contacts_df is None):
            if st.session_state.contacts_df is not None:
                df_contacts = st.session_state.contacts_df
                df_customers = st.session_state.customers_df
                
                st.session_state.metrics = {
                    'metric_1': create_metric_1(df_contacts),
                    'metric_2': create_metric_2(df_contacts),
                    'metric_4': create_metric_4(df_contacts, df_customers),
                    'metric_5': create_metric_5(df_contacts, df_customers)  # ✅ NEW METRIC
                }
                
                # ✅ NEW: Refresh revenue data
                st.session_state.revenue_data = create_course_revenue(df_customers)
                
                # ✅ NEW: Refresh matrix data
                st.session_state.matrix_data = create_volume_conversion_matrix(
                    st.session_state.metrics['metric_1'], df_contacts, df_customers
                )
                
                st.success("Analysis refreshed!")
                st.rerun()
        
        if st.button("🗑️ Clear All Data", use_container_width=True):
            st.session_state.clear()
            st.rerun()
        
        st.divider()
        
        # ✅ NEW: Download Section in Sidebar
        st.markdown("## 📥 Download Options")
        
        if st.session_state.contacts_df is not None and not st.session_state.contacts_df.empty:
            df_contacts = st.session_state.contacts_df
            df_customers = st.session_state.customers_df
            metrics = st.session_state.metrics
            
            if df_contacts is not None:
                kpis = calculate_kpis(df_contacts, df_customers)
                
                col1, col2 = st.columns(2)
                
                with col1:
                    # Download Raw Data as CSV
                    csv = df_contacts.to_csv(index=False).encode('utf-8')
                    st.download_button(
                        label="📄 Raw Data CSV",
                        data=csv,
                        file_name=f"hubspot_raw_data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                        mime="text/csv",
                        use_container_width=True,
                        help="Download all contact data as CSV"
                    )
                
                with col2:
                    # Download KPI Dashboard as CSV
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
                                df_contacts, 
                                df_customers,
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
        
        st.markdown("### 📊 Dashboard Logic")
        st.info("""
        **🎯 GUARANTEED DATA PURITY:**
        
        1️⃣ **Leads (Contacts):**
        • NEVER contain "Customer" status
        • Customer keywords → "Qualified Lead"
        • Clean pipeline stages only
        
        2️⃣ **Customers (Deals):**
        • ONLY from Deals API
        • Filtered by Stage IDs
        • Revenue from deal amounts
        
        **✅ 100% ACCURATE SEPARATION**
        
        **📈 COURSE CLASSIFICATION:**
        • ⭐ Star: High volume + High conversion
        • 📈 Potential: Low volume + High conversion  
        • ⚠️ Burn: High volume + Low conversion
        • ❌ Weak: Low volume + Low conversion
        """)
    
    # Main content area
    if st.session_state.contacts_df is not None and not st.session_state.contacts_df.empty:
        df_contacts = st.session_state.contacts_df
        df_customers = st.session_state.customers_df
        metrics = st.session_state.metrics
        revenue_data = st.session_state.revenue_data
        matrix_data = st.session_state.matrix_data
        
        # ✅ Data Validation FIRST
        st.markdown("### ✅ Data Validation Check")
        
        # Check for "Customer" in leads
        customer_in_leads = (df_contacts['Lead Status'] == 'Customer').sum()
        
        if customer_in_leads > 0:
            st.error(f"❌ CRITICAL ERROR: Found {customer_in_leads} 'Customer' entries in leads!")
            
            # Show what's causing this
            st.write("**Raw values being incorrectly marked as 'Customer':**")
            customer_rows = df_contacts[df_contacts['Lead Status'] == 'Customer']
            unique_raw = customer_rows['Lead Status Raw'].dropna().unique()
            
            for raw_val in unique_raw[:5]:
                st.write(f"- `{raw_val}` → Customer")
            
            # Auto-fix option
            if st.button("🔄 Auto-fix: Convert 'Customer' to 'Qualified Lead'"):
                df_contacts_fixed = df_contacts.copy()
                df_contacts_fixed['Lead Status'] = df_contacts_fixed['Lead Status'].replace('Customer', 'Qualified Lead')
                st.session_state.contacts_df = df_contacts_fixed
                st.success("✅ Fixed! 'Customer' entries converted to 'Qualified Lead'")
                st.rerun()
            
            st.divider()
        else:
            st.success("✅ PERFECT: 0 'Customer' entries in lead data")
        
        # ✅ NEW: Enhanced Download Section at the Top
        st.markdown('<div class="section-header"><h2>📥 Download Center</h2></div>', unsafe_allow_html=True)
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            # Download Raw Data
            csv_raw = df_contacts.to_csv(index=False).encode('utf-8')
            st.download_button(
                label="📄 Download Raw Data (CSV)",
                data=csv_raw,
                file_name=f"hubspot_raw_data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                mime="text/csv",
                use_container_width=True,
                help="All contact records with complete details"
            )
        
        with col2:
            # Download Course KPI Dashboard (NEW)
            if 'metric_5' in metrics and not metrics['metric_5'].empty:
                csv_course_kpi = metrics['metric_5'].to_csv(index=False).encode('utf-8')
                st.download_button(
                    label="📚 Download Course KPI Dashboard (CSV)",
                    data=csv_course_kpi,
                    file_name=f"hubspot_course_kpi_dashboard_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                    mime="text/csv",
                    use_container_width=True,
                    help="Course performance metrics with conversion rates"
                )
        
        with col3:
            # Premium Excel Report Button
            if st.button("💎 Generate Premium Excel Report", use_container_width=True, type="primary"):
                with st.spinner("✨ Creating premium Excel report with formatting..."):
                    try:
                        kpis = calculate_kpis(df_contacts, df_customers)
                        excel_data = create_excel_report(
                            df_contacts, 
                            df_customers,
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
        
        # ✅ NEW: Global Filters at the top
        st.markdown("### 🎛️ Global Filters")
        filter_col1, filter_col2, filter_col3 = st.columns(3)
        
        with filter_col1:
            # Course filter
            courses = df_contacts['Course/Program'].dropna().unique()
            courses = [str(c).strip() for c in courses if str(c).strip() != '']
            selected_courses = st.multiselect(
                "Filter by Course:",
                options=courses[:50] if len(courses) > 50 else courses,
                default=[],
                help="Select courses to filter all views"
            )
        
        with filter_col2:
            # Owner filter
            owners = df_contacts['Course Owner'].dropna().unique()
            owners = [str(o).strip() for o in owners if str(o).strip() != '']
            selected_owners = st.multiselect(
                "Filter by Owner:",
                options=owners[:50] if len(owners) > 50 else owners,
                default=[],
                help="Select owners to filter all views"
            )
        
        with filter_col3:
            # Lead Status filter
            lead_statuses = df_contacts['Lead Status'].dropna().unique()
            lead_statuses = [str(s).strip() for s in lead_statuses if str(s).strip() != '']
            selected_statuses = st.multiselect(
                "Filter by Lead Status:",
                options=lead_statuses,
                default=[],
                help="Select lead statuses to filter all views"
            )
        
        # Apply filters
        filtered_df = df_contacts.copy()
        
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
            'metric_4': create_metric_4(filtered_df, df_customers),
            'metric_5': create_metric_5(filtered_df, df_customers)  # ✅ NEW METRIC
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
        
        # ✅ Enhanced Executive KPI Dashboard
        st.markdown('<div class="section-header"><h2>🏆 Executive Business Dashboard</h2></div>', unsafe_allow_html=True)
        
        # Calculate KPIs
        kpis = calculate_kpis(df_contacts, df_customers)
        
        # Primary KPI Row
        st.markdown(
            render_kpi_row([
                render_kpi("Total Leads", f"{kpis['total_leads']:,}", "From Contacts", "kpi-box-blue"),
                render_kpi("Deal Leads", f"{kpis['deal_leads']:,}", f"{kpis['lead_to_deal_pct']}% conversion", "kpi-box-green"),
                render_kpi("Customers", f"{kpis['customer']:,}", "From Deals ONLY", "deal-kpi"),
                render_kpi("Total Revenue", f"₹{kpis['total_revenue']:,.0f}", f"From {kpis['customer']:,} customers", "revenue-kpi"),
            ]),
            unsafe_allow_html=True
        )
        
        # Warning if there are customers in leads
        if kpis['customer_in_leads'] > 0:
            st.error(f"🚨 DATA ISSUE: {kpis['customer_in_leads']} 'Customer' entries found in leads data")
        
        # Secondary KPI Row
        st.markdown(
            render_kpi_row([
                render_secondary_kpi("Lead→Customer", f"{kpis['lead_to_customer_pct']}%", "Leads become customers"),
                render_secondary_kpi("Lead→Deal", f"{kpis['lead_to_deal_pct']}%", "Leads in pipeline"),
                render_secondary_kpi("Deal→Customer", f"{kpis['deal_to_customer_pct']}%", "Pipeline conversion"),
                render_secondary_kpi("Avg Revenue", f"₹{kpis['avg_revenue_per_customer']:,}", "Per customer"),
                render_secondary_kpi("Top Course", kpis['top_course'], "By lead volume"),
            ], container_class="secondary-kpi-container"),
            unsafe_allow_html=True
        )
        
        # ✅ NEW: Filtered KPI Cards
        if selected_courses or selected_owners or selected_statuses:
            filtered_kpis = calculate_kpis(filtered_df, df_customers)
            
            st.markdown('<div class="section-header"><h3>📊 Filtered View KPIs</h3></div>', unsafe_allow_html=True)
            
            st.markdown(
                render_kpi_row([
                    render_kpi("Filtered Leads", f"{filtered_kpis['total_leads']:,}", f"{filtered_kpis['total_leads']/kpis['total_leads']*100:.1f}% of total", "kpi-box-orange"),
                    render_kpi("Lead→Customer %", f"{filtered_kpis['lead_to_customer_pct']}%", f"{filtered_kpis['customer']:,} customers", "kpi-box-green"),
                    render_kpi("Lead→Deal %", f"{filtered_kpis['lead_to_deal_pct']}%", f"{filtered_kpis['deal_leads']:,} deals", "kpi-box-purple"),
                    render_kpi("Total Revenue", f"₹{filtered_kpis['total_revenue']:,.0f}", f"Filtered revenue", "revenue-kpi"),
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
        
        st.divider()
        
        # ✅ ENHANCED: Create tabs with NEW Course Performance tab
        tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
            "📊 Lead Analysis", 
            "💰 Customer Analysis", 
            "📈 Owner KPI Dashboard",
            "📚 Course KPI Dashboard",  # ✅ NEW TAB
            "📉 Volume vs Conversion",
            "💸 Revenue Analysis",
            "🆚 Comparison View"
        ])
        
        # SECTION 1: Lead Analysis
        with tab1:
            st.markdown('<div class="section-header"><h3>📊 Lead Analysis (Contacts)</h3></div>', unsafe_allow_html=True)
            
            # Lead Status Distribution
            st.markdown("#### Lead Status Distribution")
            
            status_counts = filtered_df['Lead Status'].value_counts().reset_index()
            status_counts.columns = ['Lead Status', 'Count']
            
            col1, col2 = st.columns(2)
            
            with col1:
                fig = px.pie(
                    status_counts,
                    values='Count',
                    names='Lead Status',
                    title='Lead Status Distribution',
                    hole=0.3,
                    color_discrete_sequence=COLOR_PALETTE
                )
                fig.update_traces(textposition='inside', textinfo='percent+label')
                st.plotly_chart(fig, use_container_width=True)
            
            with col2:
                # Top Courses
                if 'Course/Program' in filtered_df.columns:
                    course_counts = filtered_df['Course/Program'].value_counts().head(10).reset_index()
                    course_counts.columns = ['Course', 'Count']
                    
                    fig = px.bar(
                        course_counts,
                        x='Course',
                        y='Count',
                        title='Top 10 Courses by Lead Volume',
                        color='Count',
                        color_continuous_scale='Viridis'
                    )
                    fig.update_layout(xaxis_tickangle=-45, height=400)
                    st.plotly_chart(fig, use_container_width=True)
            
            # Lead Data Table
            st.markdown("#### Lead Data")
            st.dataframe(filtered_df, use_container_width=True, height=300)
        
        # SECTION 2: Customer Analysis
        with tab2:
            st.markdown('<div class="section-header"><h3>💰 Customer Analysis (From Deals)</h3></div>', unsafe_allow_html=True)
            
            if df_customers is not None and not df_customers.empty:
                # Customer KPIs
                col1, col2, col3 = st.columns(3)
                
                with col1:
                    total_customers = len(df_customers)
                    st.metric("Total Customers", f"{total_customers:,}")
                
                with col2:
                    total_revenue = df_customers['Amount'].sum()
                    st.metric("Total Revenue", f"₹{total_revenue:,.0f}")
                
                with col3:
                    avg_revenue = df_customers['Amount'].mean()
                    st.metric("Avg Deal Value", f"₹{avg_revenue:,.0f}")
                
                # Revenue by Course
                st.markdown("#### Revenue by Course")
                
                if 'Course/Program' in df_customers.columns:
                    revenue_by_course = df_customers.groupby('Course/Program')['Amount'].sum().reset_index()
                    revenue_by_course = revenue_by_course.sort_values('Amount', ascending=False).head(10)
                    
                    fig = px.bar(
                        revenue_by_course,
                        x='Course/Program',
                        y='Amount',
                        title='Top 10 Courses by Revenue',
                        color='Amount',
                        color_continuous_scale='Viridis',
                        text='Amount'
                    )
                    fig.update_traces(
                        texttemplate='₹%{text:,.0f}',
                        textposition='outside'
                    )
                    fig.update_layout(xaxis_tickangle=-45, height=400)
                    st.plotly_chart(fig, use_container_width=True)
                
                # Customer Data Table
                st.markdown("#### Customer Deal Data")
                display_df = df_customers.copy()
                if 'Amount' in display_df.columns:
                    display_df['Amount'] = display_df['Amount'].apply(lambda x: f"₹{x:,.0f}")
                st.dataframe(display_df, use_container_width=True, height=300)
            else:
                st.info("No customer data available")
        
        # SECTION 3: Owner KPI Dashboard
        with tab3:
            st.markdown('<div class="section-header"><h3>📈 Owner Performance KPI Dashboard</h3></div>', unsafe_allow_html=True)
            
            metric_4 = filtered_metrics['metric_4']
            
            if not metric_4.empty:
                # KPI Table with conditional formatting
                st.markdown("#### Owner Performance KPI Table")
                
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
        
        # ✅ NEW SECTION 4: Course Performance KPI Dashboard
        with tab4:
            st.markdown('<div class="section-header"><h3>📚 Course Performance KPI Dashboard</h3></div>', unsafe_allow_html=True)
            
            metric_5 = filtered_metrics['metric_5']
            
            if not metric_5.empty:
                # KPI Table with conditional formatting
                st.markdown("#### Course Performance KPI Table")
                
                def highlight_course_performance(val):
                    if isinstance(val, (int, float)):
                        if val < 3:
                            return 'background-color: #f8d7da; color: #721c24; font-weight: bold'
                        elif val < 8:
                            return 'background-color: #fff3cd; color: #856404; font-weight: bold'
                        else:
                            return 'background-color: #d4edda; color: #155724; font-weight: bold'
                    return ''
                
                # Apply conditional formatting
                styled_df = metric_5.style.applymap(
                    highlight_course_performance, 
                    subset=['Lead→Customer %']
                ).applymap(
                    highlight_course_performance, 
                    subset=['Customer %']
                )
                
                st.dataframe(styled_df, use_container_width=True, height=400)
                
                # Download Course KPI Data
                st.markdown("### 📥 Export Course KPI Data")
                col_course1, col_course2 = st.columns(2)
                
                with col_course1:
                    csv_course = metric_5.to_csv(index=False)
                    st.download_button(
                        "📊 Download Course KPI Data (CSV)",
                        csv_course,
                        "course_performance_kpi.csv",
                        "text/csv",
                        use_container_width=True
                    )
            else:
                st.info("No course performance data available")
        
        # SECTION 5: Volume vs Conversion Matrix
        with tab5:
            st.markdown('<div class="section-header"><h3>📉 Volume vs Conversion Matrix</h3></div>', unsafe_allow_html=True)
            
            if matrix_data is not None and not matrix_data.empty:
                # Top 3 Courses by Conversion %
                conversion_data = []
                for _, row in filtered_metrics['metric_1'].iterrows():
                    course = row['Course']
                    total = row.get('Total', 0)
                    
                    if total > 0:
                        # Get customer count for this course from deals
                        customer_count = 0
                        if df_customers is not None and not df_customers.empty:
                            customer_count = len(df_customers[df_customers['Course/Program'] == course]) if course in df_customers['Course/Program'].values else 0
                        
                        conversion_pct = round((customer_count / total * 100), 1)
                        
                        conversion_data.append({
                            'Course': course,
                            'Conversion %': conversion_pct,
                            'Total': total,
                            'Customer': customer_count
                        })
                
                if conversion_data:
                    conversion_df = pd.DataFrame(conversion_data)
                    top_conversion_courses = conversion_df.nlargest(3, 'Conversion %')
                    
                    conversion_kpis = []
                    for _, row in top_conversion_courses.iterrows():
                        course_name = row['Course'][:12] + "..." if len(row['Course']) > 12 else row['Course']
                        conversion_kpis.append(render_secondary_kpi(
                            course_name,
                            f"{row['Conversion %']}%",
                            f"{row['Total']:,} leads → {row['Customer']:,} customers"
                        ))
                    
                    st.markdown("#### 🥇 Top 3 Courses by Lead→Customer Conversion %")
                    st.markdown(
                        render_kpi_row(conversion_kpis, container_class="secondary-kpi-container"),
                        unsafe_allow_html=True
                    )
                
                # Volume vs Conversion Matrix
                st.markdown("#### 📊 Volume vs Conversion Matrix (Strategic View)")
                
                # Apply conditional formatting for the matrix
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
                
                # Display with styling
                styled_matrix = matrix_data.style.applymap(color_matrix, subset=['Segment'])
                
                col_mat1, col_mat2 = st.columns([3, 1])
                with col_mat1:
                    st.dataframe(styled_matrix, use_container_width=True, height=350)
                
                with col_mat2:
                    st.markdown("#### 📊 Matrix Legend")
                    st.markdown("""
                    <div style='background-color: #d4edda; padding: 10px; border-radius: 5px; margin: 5px 0;'>
                    <strong>⭐ Star</strong><br>
                    High Volume + High Conversion
                    </div>
                    
                    <div style='background-color: #cce5ff; padding: 10px; border-radius: 5px; margin: 5px 0;'>
                    <strong>📈 Potential</strong><br>
                    Low Volume + High Conversion
                    </div>
                    
                    <div style='background-color: #fff3cd; padding: 10px; border-radius: 5px; margin: 5px 0;'>
                    <strong>⚠️ Burn</strong><br>
                    High Volume + Low Conversion
                    </div>
                    
                    <div style='background-color: #f8d7da; padding: 10px; border-radius: 5px; margin: 5px 0;'>
                    <strong>❌ Weak</strong><br>
                    Low Volume + Low Conversion
                    </div>
                    """, unsafe_allow_html=True)
            else:
                st.info("No matrix data available")
        
        # SECTION 6: Revenue Analysis
        with tab6:
            st.markdown('<div class="section-header"><h3>💸 Revenue Analysis by Course</h3></div>', unsafe_allow_html=True)
            
            if revenue_data is not None and not revenue_data.empty:
                # Top Revenue Course KPI
                top_revenue = revenue_data.iloc[0] if len(revenue_data) > 0 else None
                total_revenue = revenue_data['Revenue'].sum()
                total_customers = revenue_data['Customers'].sum()
                
                if top_revenue is not None:
                    st.markdown(
                        render_kpi_row([
                            render_kpi("Best Revenue Course", top_revenue['Course'][:20], f"₹{top_revenue['Revenue']:,.0f} revenue", "revenue-kpi"),
                            render_kpi("Total Revenue", f"₹{total_revenue:,.0f}", f"{total_customers} customers", "kpi-box-green"),
                            render_kpi("Avg Revenue/Customer", f"₹{revenue_data['Revenue per Customer'].mean():,.0f}", "Average", "kpi-box-purple"),
                            render_kpi("Courses with Revenue", f"{len(revenue_data)}", "Active revenue courses", "kpi-box-blue"),
                        ]),
                        unsafe_allow_html=True
                    )
                
                # Revenue Distribution Chart
                st.markdown("#### Revenue Distribution by Course")
                
                top_revenue_chart = revenue_data.head(10).copy()
                top_revenue_chart['Course'] = top_revenue_chart['Course'].str.slice(0, 25)
                
                fig1 = px.bar(
                    top_revenue_chart,
                    x='Course',
                    y='Revenue',
                    title='Top 10 Courses by Revenue',
                    color='Revenue',
                    color_continuous_scale='Viridis',
                    text='Revenue'
                )
                fig1.update_traces(
                    texttemplate='₹%{text:,.0f}',
                    textposition='outside'
                )
                fig1.update_layout(
                    xaxis_tickangle=-45,
                    xaxis_title="",
                    yaxis_title="Revenue (₹)",
                    height=400,
                    coloraxis_showscale=False
                )
                st.plotly_chart(fig1, use_container_width=True)
                
                # Revenue Data Table
                st.markdown("#### Detailed Revenue Data")
                
                # Format revenue columns
                display_revenue = revenue_data.copy()
                display_revenue['Revenue'] = display_revenue['Revenue'].apply(lambda x: f"₹{x:,.0f}")
                display_revenue['Revenue per Customer'] = display_revenue['Revenue per Customer'].apply(lambda x: f"₹{x:,.0f}")
                
                st.dataframe(display_revenue, use_container_width=True, height=350)
                
                # Download revenue data
                st.markdown("#### 📥 Export Revenue Data")
                col_rev1, col_rev2 = st.columns(2)
                
                with col_rev1:
                    csv_rev = revenue_data.to_csv(index=False)
                    st.download_button(
                        "📊 Download Revenue Data (CSV)",
                        csv_rev,
                        "course_revenue_data.csv",
                        "text/csv",
                        use_container_width=True
                    )
                
            else:
                st.info("No revenue data available. Make sure deals have 'Amount' field populated in HubSpot.")
        
        # SECTION 7: COMPARISON VIEW
        with tab7:
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
                    filtered_df, df_customers, comparison_type, item1, item2
                )
                
                if comparison_results:
                    st.markdown(f"### Comparing: **{item1}** vs **{item2}**")
                    
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
                                'Metric': ['Lead→Deal %'],
                                item1[:20]: [comparison_results['deal_pct1']],
                                item2[:20]: [comparison_results['deal_pct2']]
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
    
    else:
        # Welcome screen
        st.markdown(
            """
            <div style='text-align: center; padding: 3rem;'>
                <h2>👋 Welcome to HubSpot Business Performance Dashboard</h2>
                <p style='font-size: 1.1rem; color: #666; margin: 1rem 0;'>
                    <strong>🎯 100% CLEAN DATA SEPARATION:</strong> Customers ONLY from Deals, NEVER from Leads
                </p>
                
                <div style='margin-top: 2rem; background-color: #f8f9fa; padding: 2rem; border-radius: 0.5rem;'>
                    <h4>✅ CRITICAL FIX APPLIED:</h4>
                    
                    <div style='text-align: left; background-color: #d4edda; padding: 1rem; border-radius: 0.5rem; margin: 1rem 0;'>
                        <h5>🔥 THE PROBLEM SOLVED:</h5>
                        <p>Previous versions incorrectly marked some leads as "Customer"</p>
                        <p><strong>Example:</strong> Lead status "hot_customer" → "Customer" (WRONG!)</p>
                        <p><strong>Now:</strong> Lead status "hot_customer" → "Hot" (CORRECT!)</p>
                        <p><strong>Result:</strong> 0 "Customer" entries in lead data</p>
                    </div>
                    
                    <div style='margin-top: 2rem; padding: 1rem; background-color: #e0e7ff; border-radius: 0.5rem;'>
                        <h5>🚀 GETTING STARTED:</h5>
                        <ol style='text-align: left; margin-left: 25%;'>
                            <li>Configure customer deal stages in sidebar</li>
                            <li>Set date range</li>
                            <li>Click "Fetch ALL Data"</li>
                            <li>Check Data Validation at top of dashboard</li>
                            <li>All "Customer" entries in leads will be auto-fixed</li>
                        </ol>
                    </div>
                    
                    <div style='margin-top: 2rem; padding: 1rem; background-color: #e8f4fd; border-radius: 0.5rem;'>
                        <h5>📚 NEW: Course Performance KPI Dashboard</h5>
                        <p>Now includes comprehensive course-wise performance metrics:</p>
                        <ul style='text-align: left; margin-left: 25%;'>
                            <li>✅ Cold/Warm/Hot lead counts by course</li>
                            <li>✅ Customer conversion % by course</li>
                            <li>✅ Lead→Deal % by course</li>
                            <li>✅ Revenue by course</li>
                            <li>✅ Export as CSV/Excel</li>
                        </ul>
                    </div>
                    
                    <div style='margin-top: 2rem; padding: 1rem; background-color: #e8f4fd; border-radius: 0.5rem;'>
                        <h5>📉 Volume vs Conversion Matrix (Strategic View)</h5>
                        <div style='display: flex; justify-content: center; gap: 1rem; margin-top: 1rem;'>
                            <div style='background-color: #d4edda; padding: 0.5rem 1rem; border-radius: 0.25rem;'>
                                <strong>⭐ Star</strong><br>High Volume + High Conversion
                            </div>
                            <div style='background-color: #cce5ff; padding: 0.5rem 1rem; border-radius: 0.25rem;'>
                                <strong>📈 Potential</strong><br>Low Volume + High Conversion
                            </div>
                            <div style='background-color: #fff3cd; padding: 0.5rem 1rem; border-radius: 0.25rem;'>
                                <strong>⚠️ Burn</strong><br>High Volume + Low Conversion
                            </div>
                            <div style='background-color: #f8d7da; padding: 0.5rem 1rem; border-radius: 0.25rem;'>
                                <strong>❌ Weak</strong><br>Low Volume + Low Conversion
                            </div>
                        </div>
                    </div>
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )

if __name__ == "__main__":
    main()
