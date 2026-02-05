import streamlit as st
import requests
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta, date
import pytz
import time
import numpy as np
from io import BytesIO

# Set page config
st.set_page_config(
    page_title="HubSpot Business Analytics",
    page_icon="📊",
    layout="wide"
)

# Custom CSS
st.markdown("""
<style>
    .header-container {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 2rem;
        border-radius: 0.5rem;
        color: white;
        margin-bottom: 2rem;
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
    .kpi-container {
        display: flex;
        justify-content: center;
        align-items: stretch;
        gap: 20px;
        margin: 20px 0;
        flex-wrap: wrap;
    }
    .kpi-box {
        min-width: 220px;
        max-width: 220px;
        height: 130px;
        background: linear-gradient(135deg, #667eea, #764ba2);
        border-radius: 12px;
        padding: 16px;
        text-align: center;
        color: white;
        display: flex;
        flex-direction: column;
        justify-content: center;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.1);
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
        font-size: 34px;
        font-weight: 700;
        line-height: 1.2;
        margin: 5px 0;
    }
    .kpi-sub {
        font-size: 13px;
        opacity: 0.85;
        margin-top: 5px;
    }
    .kpi-box-green {
        background: linear-gradient(135deg, #2E8B57, #3CB371);
    }
    .kpi-box-blue {
        background: linear-gradient(135deg, #4A6FA5, #166088);
    }
</style>
""", unsafe_allow_html=True)

# Constants
HUBSPOT_API_BASE = "https://api.hubapi.com"
IST = pytz.timezone('Asia/Kolkata')
CUSTOMER_DEAL_STAGES = []

# CRITICAL FIX: Lead Status Mapping - NO CUSTOMER HERE!
LEAD_STATUS_MAP = {
    "cold": "Cold",
    "warm": "Warm", 
    "hot": "Hot",
    "new": "New Lead",
    "open": "New Lead",
    "prospect": "Warm",  
    "hot_prospect": "Hot",
    "not_connected": "Not Connected",
    "not_interested": "Not Interested", 
    "unqualified": "Not Qualified",
    "duplicate": "Duplicate",
    "": "Unknown",
    None: "Unknown"
}

# CRITICAL FIX: Customer keywords blocklist
CUSTOMER_KEYWORDS_BLOCKLIST = [
    "customer", "closed", "won", "admission", "confirmed"
]

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
        
        return None
    except:
        return None

def normalize_lead_status(raw_status):
    """Normalize lead status - NEVER returns "Customer" for leads"""
    if not raw_status:
        return "Unknown"
    
    status = str(raw_status).strip().lower()
    
    # BLOCK customer keywords
    for keyword in CUSTOMER_KEYWORDS_BLOCKLIST:
        if keyword in status:
            if "hot" in status:
                return "Hot"
            elif "warm" in status:
                return "Warm"
            elif "cold" in status:
                return "Cold"
            else:
                return "Qualified Lead"
    
    # Handle normal statuses
    if "prospect" in status:
        if "hot" in status:
            return "Hot"
        elif "warm" in status:
            return "Warm"
        else:
            return "Warm"
    
    if "not_connect" in status or "nc" in status.lower():
        return "Not Connected"
    
    if "not_interest" in status:
        return "Not Interested"
    
    if "not_qualif" in status:
        return "Not Qualified"
    
    if "duplicate" in status:
        return "Duplicate"
    
    if "new" in status or "open" in status:
        return "New Lead"
    
    if status in LEAD_STATUS_MAP:
        return LEAD_STATUS_MAP[status]
    
    return status.replace("_", " ").title()

def fetch_deal_pipeline_stages(api_key):
    """Fetch deal pipeline stages"""
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
            
            all_stages = {}
            for pipeline in pipelines:
                pipeline_id = pipeline.get("id", "")
                pipeline_label = pipeline.get("label", "")
                stages = pipeline.get("stages", [])
                
                for stage in stages:
                    stage_id = stage.get("id", "")
                    stage_label = stage.get("label", "")
                    
                    all_stages[stage_id] = {
                        "stage_id": stage_id,
                        "stage_label": stage_label,
                        "pipeline_label": pipeline_label
                    }
            
            return all_stages
        return {}
    except:
        return {}

def detect_admission_confirmed_stage(all_stages):
    """Auto-detect Admission Confirmed stage"""
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
        
        for target in target_labels:
            if target in stage_label:
                detected_stages.append({
                    "stage_id": stage_id,
                    "stage_label": stage_info.get("stage_label"),
                    "pipeline": stage_info.get("pipeline_label")
                })
                break
    
    return detected_stages

def date_to_hubspot_timestamp(date_obj, is_end_date=False):
    """Convert date to HubSpot timestamp"""
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
    """Fetch contacts from HubSpot"""
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json"
    }
    
    start_timestamp = date_to_hubspot_timestamp(start_date, is_end_date=False)
    safe_end_date = end_date + timedelta(days=1)
    end_timestamp = date_to_hubspot_timestamp(safe_end_date, is_end_date=False)
    
    all_contacts = []
    after = None
    
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
    all_properties = [
        "hs_lead_status", "lead_status", 
        "hubspot_owner_id", "hs_assigned_owner_id",
        "course", "program", "product", "service",
        "firstname", "lastname", "email", "phone", 
        "createdate", "lastmodifieddate", "hs_object_id"
    ]
    
    try:
        while True:
            body = {
                "filterGroups": filter_groups,
                "properties": all_properties,
                "limit": 100,
                "sorts": [{
                    "propertyName": "createdate",
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
    except:
        return [], 0

def fetch_hubspot_deals(api_key, start_date, end_date, customer_stage_ids):
    """Fetch deals from HubSpot"""
    if not customer_stage_ids:
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
        "course",
        "program",
        "product"
    ]
    
    try:
        while True:
            body = {
                "filterGroups": filter_groups,
                "properties": deal_properties,
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
    except:
        return [], 0

def process_contacts_data(contacts):
    """Process raw contacts data"""
    if not contacts:
        return pd.DataFrame()
    
    processed_data = []
    
    for contact in contacts:
        properties = contact.get("properties", {})
        
        # Extract course information
        course_info = ""
        course_fields = [
            "course", "program", "product", "service"
        ]
        
        for field in course_fields:
            if field in properties and properties[field] and str(properties[field]).strip():
                course_info = properties[field]
                break
        
        # Get raw lead status
        raw_lead_status = properties.get("hs_lead_status", "") or properties.get("lead_status", "")
        
        # Normalize lead status - NO CUSTOMER HERE
        lead_status = normalize_lead_status(raw_lead_status)
        
        # Create full name
        full_name = f"{properties.get('firstname', '')} {properties.get('lastname', '')}".strip()
        
        processed_data.append({
            "ID": contact.get("id", ""),
            "Full Name": full_name,
            "Email": properties.get("email", ""),
            "Phone": properties.get("phone", ""),
            "Course/Program": course_info,
            "Lead Status": lead_status,
            "Lead Status Raw": raw_lead_status,
            "Created Date": properties.get("createdate", "")
        })
    
    df = pd.DataFrame(processed_data)
    return df

def process_deals_as_customers(deals, all_stages=None):
    """Process raw deals data into customer DataFrame"""
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
        course_fields = ["course", "program", "product"]
        
        for field in course_fields:
            if field in properties and properties[field] and str(properties[field]).strip():
                course_info = properties[field]
                break
        
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
            "Deal Name": properties.get("dealname", ""),
            "Course/Program": course_info,
            "Amount": amount,
            "Close Date": close_date,
            "Is Customer": 1
        })
    
    df = pd.DataFrame(processed_data)
    return df

def create_owner_performance(df_contacts, df_customers):
    """Create owner performance summary"""
    if df_contacts.empty:
        return pd.DataFrame()
    
    # Simple aggregation
    if df_customers is not None and not df_customers.empty:
        # Merge customer data if available
        result_df = df_contacts.copy()
        # Add customer count placeholder
        result_df['Customer'] = 0
    else:
        result_df = df_contacts.copy()
    
    return result_df

def calculate_kpis(df_contacts, df_customers):
    """Calculate key performance indicators"""
    if df_contacts.empty:
        return {}
    
    # Total metrics
    total_leads = len(df_contacts)
    
    # Lead status breakdown
    status_counts = df_contacts['Lead Status'].value_counts()
    
    cold = status_counts.get('Cold', 0)
    warm = status_counts.get('Warm', 0)
    hot = status_counts.get('Hot', 0)
    new_lead = status_counts.get('New Lead', 0)
    not_connected = status_counts.get('Not Connected', 0)
    not_interested = status_counts.get('Not Interested', 0)
    not_qualified = status_counts.get('Not Qualified', 0)
    duplicate = status_counts.get('Duplicate', 0)
    
    # Customer metrics from DEALS
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
        'customer': customer,
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

def render_kpi(title, value, subtitle="", color_class=""):
    """Render a KPI card"""
    color_class = color_class or ""
    return f"""
    <div class="kpi-box {color_class}">
        <div class="kpi-title">{title}</div>
        <div class="kpi-value">{value}</div>
        <div class="kpi-sub">{subtitle}</div>
    </div>
    """

def main():
    # Get API key
    api_key = get_api_key()
    
    if not api_key:
        st.error("## 🔐 API Key Required")
        st.info("Please configure your HubSpot API key in Streamlit secrets")
        return
    
    # Header
    st.markdown(
        """
        <div class="header-container">
            <h1 style="margin: 0; font-size: 2.5rem;">📊 HubSpot Business Dashboard</h1>
            <p style="margin: 0.5rem 0 0 0; font-size: 1.2rem; opacity: 0.9;">🎯 100% CLEAN: NO Customers in Lead Data</p>
            <p style="margin: 0.5rem 0 0 0; font-size: 1rem; opacity: 0.8;">Customers ONLY from Deals</p>
        </div>
        """,
        unsafe_allow_html=True
    )
    
    # CRITICAL FIX BANNER
    st.markdown("""
    <div class="warning-card">
        <strong>✅ CRITICAL FIX APPLIED:</strong><br>
        • Lead status function now <strong>NEVER returns "Customer"</strong><br>
        • Any customer keywords → "Qualified Lead"<br>
        • Customers ONLY come from Deals<br>
        • <strong>Guaranteed: 0 "Customer" entries in lead data</strong>
    </div>
    """, unsafe_allow_html=True)
    
    # Initialize session state
    if 'contacts_df' not in st.session_state:
        st.session_state.contacts_df = None
    if 'customers_df' not in st.session_state:
        st.session_state.customers_df = None
    if 'deal_stages' not in st.session_state:
        st.session_state.deal_stages = None
    if 'customer_stage_ids' not in st.session_state:
        st.session_state.customer_stage_ids = []
    
    # Fetch Deal Pipeline Stages
    if 'deal_stages' not in st.session_state or st.session_state.deal_stages is None:
        with st.spinner("🔍 Loading deal pipeline stages..."):
            deal_stages = fetch_deal_pipeline_stages(api_key)
            st.session_state.deal_stages = deal_stages
    
    # Create sidebar
    with st.sidebar:
        st.markdown("## 🔧 Configuration")
        
        # Deal Stage Configuration
        st.markdown("### 🎯 Customer Deal Stages")
        
        if st.session_state.deal_stages:
            all_stages = st.session_state.deal_stages
            detected_stages = detect_admission_confirmed_stage(all_stages)
            
            if detected_stages:
                st.success(f"✅ Auto-detected {len(detected_stages)} customer stage(s)")
                
                # Reset customer stage IDs
                st.session_state.customer_stage_ids = []
                
                for idx, stage in enumerate(detected_stages):
                    if st.checkbox(f"Use '{stage['stage_label']}'", 
                                  value=True, key=f"use_stage_{stage['stage_id']}"):
                        if stage['stage_id'] not in st.session_state.customer_stage_ids:
                            st.session_state.customer_stage_ids.append(stage['stage_id'])
                
                if st.session_state.customer_stage_ids:
                    global CUSTOMER_DEAL_STAGES
                    CUSTOMER_DEAL_STAGES = st.session_state.customer_stage_ids
                else:
                    st.warning("⚠️ No customer stages selected")
            else:
                st.error("❌ No customer stages auto-detected")
        
        st.divider()
        
        # Date Range
        st.markdown("## 📅 Date Range")
        
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
        
        # Fetch Data Button
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
                    # Fetch CONTACTS (Leads)
                    contacts, total_contacts = fetch_hubspot_contacts_with_date_filter(
                        api_key, date_field, start_date, end_date
                    )
                    
                    # Fetch DEALS using Stage IDs
                    deals, total_deals = fetch_hubspot_deals(
                        api_key, start_date, end_date, CUSTOMER_DEAL_STAGES
                    )
                    
                    if contacts:
                        # Process contacts (leads)
                        df_contacts = process_contacts_data(contacts)
                        st.session_state.contacts_df = df_contacts
                        
                        # Process deals (customers)
                        df_customers = process_deals_as_customers(deals, st.session_state.deal_stages)
                        st.session_state.customers_df = df_customers
                        
                        st.success(f"""
                        ✅ Successfully loaded:
                        • 📊 {len(contacts)} contacts (leads)
                        • 💰 {len(deals)} customers (from deals)
                        """)
                        st.rerun()
                    else:
                        st.warning("No contacts found")
        
        if st.button("🗑️ Clear Data", use_container_width=True):
            st.session_state.clear()
            st.rerun()
    
    # Main content area
    if st.session_state.contacts_df is not None and not st.session_state.contacts_df.empty:
        df_contacts = st.session_state.contacts_df
        df_customers = st.session_state.customers_df
        
        # Data Validation
        customer_in_leads = (df_contacts['Lead Status'] == 'Customer').sum()
        
        if customer_in_leads > 0:
            st.error(f"❌ FOUND {customer_in_leads} 'Customer' entries in leads!")
            
            # Show raw values
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
        
        # Calculate KPIs
        kpis = calculate_kpis(df_contacts, df_customers)
        
        # Display KPIs
        st.markdown("## 🏆 Executive Dashboard")
        
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.metric("Total Leads", f"{kpis['total_leads']:,}")
        with col2:
            st.metric("Deal Leads", f"{kpis['deal_leads']:,}")
        with col3:
            st.metric("Customers", f"{kpis['customer']:,}")
        with col4:
            st.metric("Total Revenue", f"₹{kpis['total_revenue']:,.0f}")
        
        # Tabs for different views
        tab1, tab2, tab3 = st.tabs(["📊 Lead Analysis", "💰 Customer Analysis", "📈 Performance"])
        
        with tab1:
            st.markdown("### Lead Status Distribution")
            
            status_counts = df_contacts['Lead Status'].value_counts().reset_index()
            status_counts.columns = ['Lead Status', 'Count']
            
            fig = px.pie(
                status_counts,
                values='Count',
                names='Lead Status',
                title='Lead Status Distribution',
                hole=0.3
            )
            st.plotly_chart(fig, use_container_width=True)
            
            st.markdown("### Lead Data")
            st.dataframe(df_contacts, use_container_width=True, height=300)
        
        with tab2:
            if df_customers is not None and not df_customers.empty:
                col1, col2 = st.columns(2)
                
                with col1:
                    st.metric("Total Customers", f"{len(df_customers):,}")
                
                with col2:
                    st.metric("Avg Deal Value", f"₹{df_customers['Amount'].mean():,.0f}")
                
                st.markdown("### Customer Deal Data")
                st.dataframe(df_customers, use_container_width=True, height=300)
            else:
                st.info("No customer data available")
        
        with tab3:
            st.markdown("### Conversion Metrics")
            
            col1, col2, col3 = st.columns(3)
            
            with col1:
                st.metric("Lead→Customer %", f"{kpis['lead_to_customer_pct']}%")
            
            with col2:
                st.metric("Lead→Deal %", f"{kpis['lead_to_deal_pct']}%")
            
            with col3:
                st.metric("Deal→Customer %", f"{kpis['deal_to_customer_pct']}%")
            
            if df_customers is not None and not df_customers.empty:
                st.markdown("### Revenue by Course")
                
                if 'Course/Program' in df_customers.columns:
                    revenue_by_course = df_customers.groupby('Course/Program')['Amount'].sum().reset_index()
                    revenue_by_course = revenue_by_course.sort_values('Amount', ascending=False).head(10)
                    
                    fig = px.bar(
                        revenue_by_course,
                        x='Course/Program',
                        y='Amount',
                        title='Top 10 Courses by Revenue'
                    )
                    fig.update_layout(xaxis_tickangle=-45)
                    st.plotly_chart(fig, use_container_width=True)
    
    else:
        # Welcome screen
        st.markdown("""
        <div style='text-align: center; padding: 3rem;'>
            <h2>👋 Welcome to HubSpot Business Dashboard</h2>
            <p style='font-size: 1.1rem; color: #666; margin: 1rem 0;'>
                <strong>🎯 100% CLEAN DATA SEPARATION:</strong> Customers ONLY from Deals
            </p>
        </div>
        """, unsafe_allow_html=True)
        
        st.markdown("""
        <div style='background-color: #e0e7ff; padding: 1rem; border-radius: 0.5rem; margin: 1rem 0;'>
            <h5 style='margin: 0 0 0.5rem 0;'>🚀 GETTING STARTED:</h5>
            <ol style='margin: 0.5rem 0 0.5rem 1rem;'>
                <li>Configure customer deal stages in sidebar</li>
                <li>Set date range</li>
                <li>Click "Fetch ALL Data"</li>
            </ol>
        </div>
        """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
