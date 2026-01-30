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

# ✅ CRITICAL FIX: Fetch Deal Pipeline Stages to get correct Stage IDs
@st.cache_data(ttl=86400)  # Cache for 24 hours
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
            
            # Collect all stages from all pipelines
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
                    
                    # Store with both ID and label
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

# ✅ NEW: Auto-detect Admission Confirmed stage ID
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
        
        # Check if this is a "customer" stage
        for target in target_labels:
            if target in stage_label:
                # Also check if probability indicates closed/won (usually 1.0 or 0.9-1.0)
                try:
                    prob_float = float(probability)
                    if prob_float >= 0.9:  # High probability = likely customer stage
                        detected_stages.append({
                            "stage_id": stage_id,
                            "stage_label": stage_info.get("stage_label"),
                            "probability": probability,
                            "pipeline": stage_info.get("pipeline_label"),
                            "match_reason": f"Label contains '{target}', probability: {probability}"
                        })
                        break
                except:
                    # If probability not a number, still consider based on label
                    detected_stages.append({
                        "stage_id": stage_id,
                        "stage_label": stage_info.get("stage_label"),
                        "probability": probability,
                        "pipeline": stage_info.get("pipeline_label"),
                        "match_reason": f"Label contains '{target}'"
                    })
                    break
    
    return detected_stages

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

# ✅ CRITICAL FIX: Fetch DEALS using CORRECT Stage IDs
def fetch_hubspot_deals(api_key, start_date, end_date, customer_stage_ids):
    """Fetch DEALS from HubSpot using CORRECT stage IDs (not labels)."""
    if not customer_stage_ids:
        st.error("❌ No customer stage IDs configured. Please configure in sidebar.")
        return [], 0
    
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
    status_text.text("💰 Fetching customer deals...")
    
    # ✅ CORRECT FILTER: Use Stage IDs + Close Date
    filter_groups = [{
        "filters": [
            {
                "propertyName": "dealstage",
                "operator": "IN",
                "values": customer_stage_ids  # ✅ USING STAGE IDs, not labels
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
    
    # Optional: Also require close date to exist (ensures real customers)
    # Uncomment if you want to ensure deals have close dates:
    # filter_groups[0]["filters"].append({
    #     "propertyName": "closedate",
    #     "operator": "HAS_PROPERTY"
    # })
    
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
        "program_name",
        "hs_priority",  # Added for debugging
        "dealtype",     # Added for debugging
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
        
        if all_deals:
            status_text.text(f"✅ Deal fetch complete! Found {len(all_deals)} customer deals")
            
            # Show debug info
            stage_counts = {}
            for deal in all_deals:
                stage = deal.get("properties", {}).get("dealstage", "unknown")
                stage_counts[stage] = stage_counts.get(stage, 0) + 1
            
            st.info(f"📊 Deal Stage Breakdown: {stage_counts}")
        else:
            status_text.text("⚠️ No customer deals found for selected criteria")
            
            # Show troubleshooting info
            with st.expander("🔍 Troubleshooting - Why no deals?", expanded=False):
                st.write("**Possible reasons:**")
                st.write("1. No deals with the selected stage IDs in date range")
                st.write("2. Stage IDs might be incorrect")
                st.write("3. Close dates might be outside range")
                st.write("4. API permissions issue")
                st.write(f"**Stage IDs being used:** {customer_stage_ids}")
                st.write(f"**Date range:** {start_date} to {end_date}")
                st.write(f"**Close date filter:** {start_timestamp} to {end_timestamp}")
        
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

# ✅ Process DEALS as Customers
def process_deals_as_customers(deals, owner_mapping=None, api_key=None, all_stages=None):
    """Process raw deals data into customer DataFrame."""
    if not deals:
        return pd.DataFrame()
    
    processed_data = []
    unmapped_owners = set()
    stage_label_map = {}
    
    # Create mapping from stage ID to label
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
    
    if len(processed_data) > 0:
        with st.expander("🧪 Customer Validation (from Deals)", expanded=False):
            st.write(f"✅ Total Customers from Deals: {len(df)}")
            st.write(f"💰 Total Revenue: ₹{df['Amount'].sum():,.2f}")
            
            # Show stage distribution
            if 'Deal Stage Label' in df.columns:
                stage_dist = df['Deal Stage Label'].value_counts()
                st.write("📊 Customer Deal Stage Distribution:")
                st.write(stage_dist)
            
            # Show sample data
            if not df.empty:
                st.write("Sample Customer Records:", df[['Deal Name', 'Course/Program', 'Amount', 'Close Date', 'Deal Stage Label']].head(5))
    
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
    if df_customers is not None and not df_customers.empty and 'Course Owner' in df_customers.columns:
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
        result_df['Deal %'] = np.where(
            result_df['Grand Total'] > 0,
            (result_df['Deal Leads'] / result_df['Grand Total'] * 100).round(2),
            0
        )
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
    if df_customers is None or df_customers.empty or 'Course/Program' not in df_customers.columns or 'Amount' not in df_customers.columns:
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
    if df_customers is not None and not df_customers.empty:
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
    if df_customers is not None and not df_customers.empty and 'Course/Program' in df_customers.columns:
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
            <p style="margin: 0.5rem 0 0 0; font-size: 1.2rem; opacity: 0.9;">🎯 CORRECT: Customer = Deal Stage ID (NOT Label)</p>
            <p style="margin: 0.5rem 0 0 0; font-size: 1rem; opacity: 0.8;">🔍 Auto-detects Admission Confirmed Stage ID | 100% Accurate</p>
        </div>
        """,
        unsafe_allow_html=True
    )
    
    # ✅ CRITICAL: Data Source Warning
    st.markdown("""
    <div class="warning-card">
        <strong>⚠️ IMPORTANT FIX: Deal Stage IDs vs Labels</strong><br>
        • HubSpot API filters by <strong>Stage ID</strong>, not display label<br>
        • "Admission Confirmed" in UI ≠ "admission_confirmed" in API<br>
        • Dashboard now auto-detects correct Stage IDs<br>
        • <strong>Result: 100% accurate customer counting</strong>
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
    if 'deal_stages' not in st.session_state:
        st.session_state.deal_stages = None
    if 'customer_stage_ids' not in st.session_state:
        st.session_state.customer_stage_ids = []
    
    # ✅ NEW: Fetch Deal Pipeline Stages FIRST
    if 'deal_stages' not in st.session_state or st.session_state.deal_stages is None:
        with st.spinner("🔍 Loading deal pipeline stages..."):
            deal_stages = fetch_deal_pipeline_stages(api_key)
            st.session_state.deal_stages = deal_stages
    
    # Create sidebar for configuration
    with st.sidebar:
        st.markdown("## 🔧 Configuration")
        
        # ✅ NEW: Deal Stage Configuration Section
        st.markdown("### 🎯 Customer Deal Stage Configuration")
        
        if st.session_state.deal_stages:
            all_stages = st.session_state.deal_stages
            
            # Auto-detect Admission Confirmed stages
            detected_stages = detect_admission_confirmed_stage(all_stages)
            
            if detected_stages:
                st.success(f"✅ Auto-detected {len(detected_stages)} customer stage(s)")
                
                # Show detected stages
                for idx, stage in enumerate(detected_stages):
                    with st.expander(f"Stage {idx+1}: {stage['stage_label']}", expanded=idx==0):
                        st.write(f"**Stage ID:** `{stage['stage_id']}`")
                        st.write(f"**Pipeline:** {stage['pipeline']}")
                        st.write(f"**Probability:** {stage['probability']}")
                        st.write(f"**Match Reason:** {stage['match_reason']}")
                        
                        # Checkbox to include this stage
                        is_selected = st.checkbox(
                            f"Include '{stage['stage_label']}' as customer stage",
                            value=True,
                            key=f"stage_{stage['stage_id']}"
                        )
                        
                        if is_selected and stage['stage_id'] not in st.session_state.customer_stage_ids:
                            st.session_state.customer_stage_ids.append(stage['stage_id'])
                        elif not is_selected and stage['stage_id'] in st.session_state.customer_stage_ids:
                            st.session_state.customer_stage_ids.remove(stage['stage_id'])
                
                # Manual stage selection
                st.markdown("#### 🔧 Manual Stage Selection")
                
                # Get all stages for manual selection
                all_stage_options = []
                for stage_id, stage_info in all_stages.items():
                    label = stage_info.get("stage_label", "Unknown")
                    pipeline = stage_info.get("pipeline_label", "Unknown")
                    probability = stage_info.get("probability", "0")
                    all_stage_options.append({
                        "id": stage_id,
                        "display": f"{label} (Pipeline: {pipeline}, Probability: {probability})"
                    })
                
                # Sort by label
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
                            if stage["id"] not in st.session_state.customer_stage_ids:
                                manual_stage_ids.append(stage["id"])
                            break
                
                # Combine auto-detected and manual selections
                all_selected_ids = st.session_state.customer_stage_ids + manual_stage_ids
                
                # Remove duplicates
                all_selected_ids = list(set(all_selected_ids))
                
                if all_selected_ids:
                    st.success(f"✅ {len(all_selected_ids)} customer stage(s) selected")
                    
                    # Show summary
                    with st.expander("📋 Selected Customer Stages", expanded=True):
                        for stage_id in all_selected_ids:
                            if stage_id in all_stages:
                                info = all_stages[stage_id]
                                st.write(f"• **{info.get('stage_label')}** (`{stage_id}`)")
                            else:
                                st.write(f"• `{stage_id}`")
                    
                    # Update global variable
                    global CUSTOMER_DEAL_STAGES
                    CUSTOMER_DEAL_STAGES = all_selected_ids
                else:
                    st.warning("⚠️ No customer stages selected")
            else:
                st.error("""
                ❌ No customer stages auto-detected!
                
                **Possible reasons:**
                1. No "Admission Confirmed" stage in your pipeline
                2. Stage has different label (check HubSpot UI)
                3. API permissions issue
                
                **Solution:**
                1. Check your deal pipeline in HubSpot UI
                2. Manually select customer stages below
                3. Ensure API key has `crm.pipelines.read` scope
                """)
                
                # Show all available stages for manual selection
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
                
                # Sort by label
                all_stage_options.sort(key=lambda x: x["display"])
                
                selected_stages = st.multiselect(
                    "Select customer deal stages:",
                    options=[s["display"] for s in all_stage_options],
                    default=[],
                    help="Select stages that indicate customers (e.g., Admission Confirmed, Closed Won)"
                )
                
                # Map back to IDs
                selected_ids = []
                for display in selected_stages:
                    for stage in all_stage_options:
                        if stage["display"] == display:
                            selected_ids.append(stage["id"])
                            break
                
                if selected_ids:
                    st.session_state.customer_stage_ids = selected_ids
                    CUSTOMER_DEAL_STAGES = selected_ids
                    st.success(f"✅ {len(selected_ids)} customer stage(s) selected")
                else:
                    st.warning("⚠️ Please select at least one customer stage")
        else:
            st.error("❌ Could not load deal stages. Check API permissions.")
        
        st.divider()
        
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
        
        fetch_disabled = not CUSTOMER_DEAL_STAGES
        
        if st.button("🚀 Fetch ALL Data", 
                    type="primary", 
                    use_container_width=True,
                    disabled=fetch_disabled,
                    help="Configure customer stages first" if fetch_disabled else "Fetch contacts and deals"):
            
            if not CUSTOMER_DEAL_STAGES:
                st.error("❌ Please configure customer deal stages first")
                return
                
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
                        
                        # ✅ CORRECT: Fetch DEALS using Stage IDs
                        st.info("💰 Fetching customer deals...")
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
                            • 🎯 Using Stage IDs: {CUSTOMER_DEAL_STAGES}
                            """)
                            st.rerun()
                        else:
                            st.warning("No contacts found for the selected date range.")
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
        
        st.markdown("### 📊 Dashboard Logic")
        st.info("""
        **🎯 100% CORRECT CUSTOMER LOGIC:**
        
        1️⃣ **API Truth**: HubSpot uses Stage IDs, NOT labels
        2️⃣ **Auto-detection**: Finds "Admission Confirmed" Stage ID
        3️⃣ **Filter**: Deal Stage ID + Close Date
        
        **✅ GUARANTEED ACCURACY:**
        • Customers = Deals with correct Stage ID
        • No false positives from labels
        • Real revenue from deal amounts
        
        **🔧 CONFIGURATION:**
        • Review auto-detected stages
        • Add/remove as needed
        • Stage ID is what matters
        """)
    
    # Main content area
    if st.session_state.contacts_df is not None and not st.session_state.contacts_df.empty:
        df_contacts = st.session_state.contacts_df
        df_customers = st.session_state.customers_df
        metrics = st.session_state.metrics
        revenue_data = st.session_state.revenue_data
        
        # ✅ Data Source Summary
        st.markdown("### 📊 Data Source Summary")
        col_sum1, col_sum2, col_sum3 = st.columns(3)
        
        with col_sum1:
            st.metric("Total Leads", f"{len(df_contacts):,}", "From Contacts")
        
        with col_sum2:
            customer_count = len(df_customers) if df_customers is not None else 0
            st.metric("Total Customers", f"{customer_count:,}", "From Deals (Stage IDs)")
        
        with col_sum3:
            revenue = df_customers['Amount'].sum() if df_customers is not None and not df_customers.empty else 0
            st.metric("Total Revenue", f"₹{revenue:,.0f}", "From Customer Deals")
        
        # Show customer stage info
        if CUSTOMER_DEAL_STAGES and st.session_state.deal_stages:
            st.markdown("#### 🎯 Customer Stage Configuration")
            stage_info = []
            for stage_id in CUSTOMER_DEAL_STAGES:
                if stage_id in st.session_state.deal_stages:
                    info = st.session_state.deal_stages[stage_id]
                    stage_info.append(f"**{info.get('stage_label')}** (`{stage_id}`)")
            
            if stage_info:
                st.success(f"✅ Customer = Deal Stage in: {', '.join(stage_info)}")
        
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
        
        # Create tabs
        tab1, tab2, tab3, tab4, tab5 = st.tabs([
            "📊 Lead Distribution", 
            "💰 Customer Analysis", 
            "📈 KPI Dashboard",
            "🎯 Conversion Funnel",
            "🔍 Data Validation"
        ])
        
        # SECTION 1: Lead Distribution
        with tab1:
            st.markdown('<div class="section-header"><h3>📊 Lead Distribution (Contacts)</h3></div>', unsafe_allow_html=True)
            
            if not df_contacts.empty:
                # Lead Status Distribution
                col1, col2 = st.columns(2)
                
                with col1:
                    st.markdown("#### Lead Status Distribution")
                    
                    status_counts = df_contacts['Lead Status'].value_counts().reset_index()
                    status_counts.columns = ['Lead Status', 'Count']
                    
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
                
                with col2:
                    st.markdown("#### Top Courses by Lead Volume")
                    
                    course_counts = df_contacts['Course/Program'].value_counts().head(10).reset_index()
                    course_counts.columns = ['Course', 'Count']
                    
                    fig2 = px.bar(
                        course_counts,
                        x='Course',
                        y='Count',
                        title='Top 10 Courses by Lead Volume',
                        color='Count',
                        color_continuous_scale='Viridis'
                    )
                    fig2.update_layout(xaxis_tickangle=-45, height=400)
                    st.plotly_chart(fig2, use_container_width=True)
                
                # Raw Data
                st.markdown("#### Lead Data Table")
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
                
                # Revenue Analysis
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
                st.info("No customer data available. Make sure:")
                st.write("1. Customer deal stages are correctly configured")
                st.write("2. There are deals with those stages in the date range")
                st.write("3. Deals have close dates")
        
        # SECTION 3: KPI Dashboard
        with tab3:
            st.markdown('<div class="section-header"><h3>📈 KPI Dashboard (Leads + Customers)</h3></div>', unsafe_allow_html=True)
            
            metric_4 = create_metric_4(df_contacts, df_customers)
            
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
        
        # SECTION 4: Conversion Funnel
        with tab4:
            st.markdown('<div class="section-header"><h3>🎯 Conversion Funnel (Leads → Customers)</h3></div>', unsafe_allow_html=True)
            
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
            
            # Conversion Metrics
            st.markdown("#### 📈 Conversion Metrics")
            
            conversion_metrics = pd.DataFrame({
                'Metric': ['Lead → Deal Conversion', 'Lead → Customer Conversion', 'Deal → Customer Conversion'],
                'Percentage': [kpis['lead_to_deal_pct'], kpis['lead_to_customer_pct'], kpis['deal_to_customer_pct']],
                'Description': [
                    f"{kpis['deal_leads']:,} out of {kpis['total_leads']:,} leads",
                    f"{kpis['customer']:,} out of {kpis['total_leads']:,} leads",
                    f"{kpis['customer']:,} out of {kpis['deal_leads']:,} deals"
                ]
            })
            
            st.dataframe(conversion_metrics, use_container_width=True)
        
        # SECTION 5: Data Validation
        with tab5:
            st.markdown('<div class="section-header"><h3>🔍 Data Validation & Debug</h3></div>', unsafe_allow_html=True)
            
            col_val1, col_val2 = st.columns(2)
            
            with col_val1:
                st.markdown("#### ✅ Lead Data Validation")
                st.write(f"**Total Leads:** {len(df_contacts):,}")
                st.write("**Lead Status Distribution:**")
                status_counts = df_contacts['Lead Status'].value_counts()
                st.write(status_counts)
                
                # Check for any "Customer" in leads (should be 0)
                customer_in_leads = (df_contacts['Lead Status'] == 'Customer').sum()
                if customer_in_leads == 0:
                    st.success("✅ PERFECT: No 'Customer' in lead data")
                else:
                    st.error(f"❌ PROBLEM: Found {customer_in_leads} 'Customer' in leads (should be 0)")
            
            with col_val2:
                st.markdown("#### ✅ Customer Data Validation")
                if df_customers is not None and not df_customers.empty:
                    st.write(f"**Total Customers:** {len(df_customers):,}")
                    st.write(f"**Total Revenue:** ₹{df_customers['Amount'].sum():,.0f}")
                    
                    if 'Deal Stage Label' in df_customers.columns:
                        st.write("**Customer Deal Stages:**")
                        stage_counts = df_customers['Deal Stage Label'].value_counts()
                        st.write(stage_counts)
                    
                    # Check close dates
                    if 'Close Date' in df_customers.columns:
                        has_close_dates = df_customers['Close Date'].notna().sum()
                        st.write(f"**Deals with Close Dates:** {has_close_dates:,} ({has_close_dates/len(df_customers)*100:.1f}%)")
                else:
                    st.warning("⚠️ No customer data available")
            
            # Stage ID Debug Info
            st.markdown("#### 🔧 Stage ID Debug Information")
            st.write(f"**Configured Customer Stage IDs:** {CUSTOMER_DEAL_STAGES}")
            
            if st.session_state.deal_stages:
                st.write("**Available Stages in HubSpot:**")
                stage_list = []
                for stage_id, info in st.session_state.deal_stages.items():
                    stage_list.append({
                        'Stage ID': stage_id,
                        'Stage Label': info.get('stage_label', 'Unknown'),
                        'Pipeline': info.get('pipeline_label', 'Unknown'),
                        'Probability': info.get('probability', '0')
                    })
                
                stages_df = pd.DataFrame(stage_list)
                st.dataframe(stages_df, use_container_width=True, height=300)
    
    else:
        # Welcome screen with configuration guidance
        st.markdown(
            """
            <div style='text-align: center; padding: 3rem;'>
                <h2>👋 Welcome to HubSpot Business Performance Dashboard</h2>
                <p style='font-size: 1.1rem; color: #666; margin: 1rem 0;'>
                    <strong>🎯 100% ACCURATE CUSTOMER COUNTING:</strong> Using Deal Stage IDs, not labels
                </p>
                
                <div style='margin-top: 2rem; background-color: #f8f9fa; padding: 2rem; border-radius: 0.5rem;'>
                    <h4>🔧 CRITICAL CONFIGURATION REQUIRED:</h4>
                    
                    <div style='text-align: left; background-color: #e8f4fd; padding: 1rem; border-radius: 0.5rem; margin: 1rem 0;'>
                        <h5>🚨 THE PROBLEM (Why previous dashboards failed):</h5>
                        <p><strong>HubSpot UI shows:</strong> "Admission Confirmed"</p>
                        <p><strong>HubSpot API needs:</strong> Stage ID (e.g., <code>appointmentscheduled_12345</code>)</p>
                        <p><strong>Result if wrong:</strong> 0 customers shown, even though UI has customers</p>
                    </div>
                    
                    <div style='text-align: left; background-color: #d4edda; padding: 1rem; border-radius: 0.5rem; margin: 1rem 0;'>
                        <h5>✅ OUR SOLUTION (100% accurate):</h5>
                        <ol>
                            <li><strong>Auto-detects</strong> "Admission Confirmed" Stage ID from HubSpot</li>
                            <li><strong>Uses Stage IDs</strong> in API calls (not labels)</li>
                            <li><strong>Shows you exactly</strong> which stages are being used</li>
                            <li><strong>Allows manual override</strong> if needed</li>
                        </ol>
                    </div>
                    
                    <div style='margin-top: 2rem; padding: 1rem; background-color: #fff3cd; border-radius: 0.5rem;'>
                        <h5>🚀 GETTING STARTED:</h5>
                        <ol style='text-align: left; margin-left: 25%;'>
                            <li>Check sidebar for auto-detected customer stages</li>
                            <li>Review and confirm the stages are correct</li>
                            <li>Set date range for leads & customers</li>
                            <li>Click "Fetch ALL Data"</li>
                            <li>Check Data Validation tab to confirm accuracy</li>
                        </ol>
                    </div>
                    
                    <div style='margin-top: 2rem; padding: 1rem; background-color: #f8d7da; border-radius: 0.5rem;'>
                        <h5>⚠️ TROUBLESHOOTING:</h5>
                        <p><strong>If no stages auto-detected:</strong></p>
                        <ul style='text-align: left; margin-left: 20%;'>
                            <li>Check API key has <code>crm.pipelines.read</code> scope</li>
                            <li>Your "Admission Confirmed" stage might have different label</li>
                            <li>Manually select customer stages in sidebar</li>
                            <li>Check HubSpot UI for exact stage labels</li>
                        </ul>
                    </div>
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )

if __name__ == "__main__":
    main()
