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
        'lead_to_customer_pct': lead_to_customer_pct,
        'lead_to_deal_pct': lead_to_deal_pct,
        'deal_to_customer_pct': deal_to_customer_pct,
        'total_revenue': total_revenue,
        'avg_revenue_per_customer': avg_revenue_per_customer
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
                            
                            # Calculate metrics
                            st.session_state.metrics = {
                                'metric_4': create_metric_4(df_contacts, df_customers)
                            }
                            
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
                    'metric_4': create_metric_4(df_contacts, df_customers)
                }
                
                st.success("Analysis refreshed!")
                st.rerun()
        
        if st.button("🗑️ Clear All Data", use_container_width=True):
            st.session_state.clear()
            st.rerun()
        
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
        """)
    
    # Main content area
    if st.session_state.contacts_df is not None and not st.session_state.contacts_df.empty:
        df_contacts = st.session_state.contacts_df
        df_customers = st.session_state.customers_df
        
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
        
        # Data Source Summary
        col_sum1, col_sum2, col_sum3 = st.columns(3)
        
        with col_sum1:
            st.metric("Total Leads", f"{len(df_contacts):,}", "From Contacts")
        
        with col_sum2:
            customer_count = len(df_customers) if df_customers is not None else 0
            st.metric("Total Customers", f"{customer_count:,}", "From Deals")
        
        with col_sum3:
            revenue = df_customers['Amount'].sum() if df_customers is not None and not df_customers.empty else 0
            st.metric("Total Revenue", f"₹{revenue:,.0f}", "From Customer Deals")
        
        # ✅ Executive KPI Dashboard
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
        
        # Conversion KPI Row
        st.markdown(
            render_kpi_row([
                render_secondary_kpi("Lead→Customer", f"{kpis['lead_to_customer_pct']}%", "Leads become customers"),
                render_secondary_kpi("Lead→Deal", f"{kpis['lead_to_deal_pct']}%", "Leads in pipeline"),
                render_secondary_kpi("Deal→Customer", f"{kpis['deal_to_customer_pct']}%", "Pipeline conversion"),
                render_secondary_kpi("Avg Revenue", f"₹{kpis['avg_revenue_per_customer']:,}", "Per customer"),
            ], container_class="secondary-kpi-container"),
            unsafe_allow_html=True
        )
        
        st.divider()
        
        # Create tabs
        tab1, tab2, tab3 = st.tabs([
            "📊 Lead Analysis", 
            "💰 Customer Analysis", 
            "📈 KPI Dashboard"
        ])
        
        # SECTION 1: Lead Analysis
        with tab1:
            st.markdown('<div class="section-header"><h3>📊 Lead Analysis (Contacts)</h3></div>', unsafe_allow_html=True)
            
            # Lead Status Distribution
            st.markdown("#### Lead Status Distribution")
            
            status_counts = df_contacts['Lead Status'].value_counts().reset_index()
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
                if 'Course/Program' in df_contacts.columns:
                    course_counts = df_contacts['Course/Program'].value_counts().head(10).reset_index()
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
            st.dataframe(df_contacts, use_container_width=True, height=300)
        
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
        
        # SECTION 3: KPI Dashboard
        with tab3:
            st.markdown('<div class="section-header"><h3>📈 KPI Dashboard (Leads + Customers)</h3></div>', unsafe_allow_html=True)
            
            metric_4 = create_metric_4(df_contacts, df_customers)
            
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
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )

if __name__ == "__main__":
    main()
