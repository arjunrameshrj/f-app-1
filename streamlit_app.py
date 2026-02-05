import streamlit as st
import pandas as pd
import requests
from datetime import datetime, timedelta
import pytz
from typing import Dict, List, Optional, Tuple

# Page configuration
st.set_page_config(
    page_title="KPI Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for cleaner UI
st.markdown("""
    <style>
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        color: #1A73E8;
        margin-bottom: 1rem;
    }
    .kpi-card {
        background-color: #f8f9fa;
        border-radius: 10px;
        padding: 20px;
        border-left: 5px solid #1A73E8;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
        margin-bottom: 20px;
    }
    .kpi-title {
        font-size: 1rem;
        color: #5f6368;
        margin-bottom: 5px;
    }
    .kpi-value {
        font-size: 2rem;
        font-weight: bold;
        color: #202124;
    }
    .kpi-change {
        font-size: 0.9rem;
        margin-top: 5px;
    }
    .positive {
        color: #0d9d58;
    }
    .negative {
        color: #ea4335;
    }
    .stDateInput > div > div > input {
        font-size: 14px;
    }
    </style>
""", unsafe_allow_html=True)

# Constants
HUBSPOT_ACCESS_TOKEN = "your_hubspot_access_token_here"
HUBSPOT_BASE_URL = "https://api.hubapi.com"

# Helper function to check if payment date is outside dashboard range
def is_outside_range(payment_date: Optional[str], start_date: datetime.date, end_date: datetime.date) -> bool:
    """Check if payment date is outside the selected date range."""
    if not payment_date:
        return False
    try:
        # Convert HubSpot timestamp (milliseconds) to date
        payment_dt = datetime.fromtimestamp(int(payment_date) / 1000).date()
        return payment_dt < start_date or payment_dt > end_date
    except:
        return False

def safe_float(value, default=0.0) -> float:
    """Safely convert value to float, handling None, strings with commas, etc."""
    if value is None:
        return default
    try:
        # Remove commas and convert
        if isinstance(value, str):
            value = value.replace(",", "").strip()
        return float(value) if value else default
    except:
        return default

def fetch_hubspot_deals(start_date: datetime, end_date: datetime, token: str) -> List[Dict]:
    """Fetch deals from HubSpot with all necessary properties."""
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }
    
    # Convert dates to milliseconds for HubSpot filter
    start_timestamp = int(start_date.timestamp() * 1000)
    end_timestamp = int(end_date.timestamp() * 1000)
    
    all_deals = []
    after = None
    has_more = True
    
    # Define all properties needed
    deal_properties = [
        "dealname",
        "dealstage",
        "amount",
        "hubspot_owner_id",
        "closedate",
        "createdate",
        # Partial payment properties
        "partial_amount",
        "partial_payment_amount",
        "offline_payment_amount",
        "online_payment_amount",
        "partial_payment_date",
        "offline_payment_date",
        "online_payment_date",
        "course",
        "program"
    ]
    
    while has_more:
        params = {
            "limit": 100,
            "properties": deal_properties,
            "archived": "false"
        }
        
        if after:
            params["after"] = after
            
        # Try to get deals created or closed in date range
        filter_groups = [{
            "filters": [
                {
                    "propertyName": "createdate",
                    "operator": "BETWEEN",
                    "highValue": str(end_timestamp),
                    "value": str(start_timestamp)
                },
                {
                    "propertyName": "closedate",
                    "operator": "BETWEEN",
                    "highValue": str(end_timestamp),
                    "value": str(start_timestamp)
                }
            ]
        }]
        
        params["filterGroups"] = filter_groups
        
        try:
            response = requests.post(
                f"{HUBSPOT_BASE_URL}/crm/v3/objects/deals/search",
                headers=headers,
                json=params
            )
            
            if response.status_code == 200:
                data = response.json()
                all_deals.extend(data.get("results", []))
                
                paging = data.get("paging", {})
                if "next" in paging:
                    after = paging["next"]["after"]
                else:
                    has_more = False
            else:
                st.error(f"Error fetching deals: {response.status_code}")
                has_more = False
                
        except Exception as e:
            st.error(f"Exception fetching deals: {str(e)}")
            has_more = False
    
    return all_deals

def process_deals_as_customers(deals_data: List[Dict], start_date: datetime.date, end_date: datetime.date) -> pd.DataFrame:
    """Process deals data into customer format with corrected revenue calculation."""
    customers = []
    
    for deal in deals_data:
        properties = deal.get("properties", {})
        
        # Get deal amount and clean it
        deal_amount = safe_float(properties.get("amount", 0))
        
        # Get partial payment amounts
        partial_amount = safe_float(properties.get("partial_payment_amount", 0))
        offline_amount = safe_float(properties.get("offline_payment_amount", 0))
        online_amount = safe_float(properties.get("online_payment_amount", 0))
        
        # Get partial payment dates
        partial_date = properties.get("partial_payment_date")
        offline_date = properties.get("offline_payment_date")
        online_date = properties.get("online_payment_date")
        
        # CORRECTED REVENUE LOGIC:
        # Subtract only if payment date is OUTSIDE dashboard range
        if is_outside_range(partial_date, start_date, end_date):
            deal_amount -= partial_amount
        
        if is_outside_range(offline_date, start_date, end_date):
            deal_amount -= offline_amount
        
        if is_outside_range(online_date, start_date, end_date):
            deal_amount -= online_amount
        
        # Ensure deal amount is not negative
        deal_amount = max(deal_amount, 0)
        
        # Get other properties
        deal_stage = properties.get("dealstage", "Unknown")
        course = properties.get("course", "Not specified")
        program = properties.get("program", "Not specified")
        
        # Determine customer status based on deal stage
        if deal_stage in ["closedwon", "closedwon不用谢", "contractsent", "63020224"]:
            status = "Active"
        elif deal_stage in ["closedlost", "closedlost不用谢"]:
            status = "Inactive"
        else:
            status = "Pending"
        
        # Get close date
        close_date = properties.get("closedate")
        if close_date:
            try:
                close_date = datetime.fromtimestamp(int(close_date) / 1000).strftime("%Y-%m-%d")
            except:
                close_date = "N/A"
        else:
            close_date = "N/A"
        
        customers.append({
            "Customer Name": properties.get("dealname", "Unknown"),
            "Status": status,
            "Amount": deal_amount,
            "Course": course,
            "Program": program,
            "Close Date": close_date
        })
    
    return pd.DataFrame(customers)

def calculate_kpis(df: pd.DataFrame, previous_df: Optional[pd.DataFrame] = None) -> Dict:
    """Calculate KPIs from customer data."""
    kpis = {}
    
    # Total Revenue (CORRECTED - excludes partial payments outside date range)
    total_revenue = df["Amount"].sum()
    kpis["Total Revenue"] = total_revenue
    
    # Customer Counts
    active_customers = df[df["Status"] == "Active"].shape[0]
    inactive_customers = df[df["Status"] == "Inactive"].shape[0]
    pending_customers = df[df["Status"] == "Pending"].shape[0]
    total_customers = df.shape[0]
    
    kpis["Active Customers"] = active_customers
    kpis["Inactive Customers"] = inactive_customers
    kpis["Pending Customers"] = pending_customers
    kpis["Total Customers"] = total_customers
    
    # Average Revenue Per Customer
    avg_revenue = total_revenue / total_customers if total_customers > 0 else 0
    kpis["Avg Revenue/Customer"] = avg_revenue
    
    # Course-wise revenue
    course_revenue = df.groupby("Course")["Amount"].sum().sort_values(ascending=False)
    kpis["Course Revenue"] = course_revenue
    
    # Program-wise revenue
    program_revenue = df.groupby("Program")["Amount"].sum().sort_values(ascending=False)
    kpis["Program Revenue"] = program_revenue
    
    # Calculate changes if previous data is provided
    if previous_df is not None:
        prev_total_revenue = previous_df["Amount"].sum()
        prev_total_customers = previous_df.shape[0]
        
        revenue_change = ((total_revenue - prev_total_revenue) / prev_total_revenue * 100) if prev_total_revenue > 0 else 0
        customer_change = ((total_customers - prev_total_customers) / prev_total_customers * 100) if prev_total_customers > 0 else 0
        
        kpis["Revenue Change %"] = revenue_change
        kpis["Customer Change %"] = customer_change
    
    return kpis

def display_kpi_card(title: str, value, change: Optional[float] = None, prefix: str = "$", is_percent: bool = False):
    """Display a KPI card with optional change indicator."""
    col1, col2 = st.columns([3, 1])
    
    with col1:
        st.markdown(f'<div class="kpi-title">{title}</div>', unsafe_allow_html=True)
        
        if is_percent:
            formatted_value = f"{value:.1f}%"
        elif prefix == "$":
            formatted_value = f"{prefix}{value:,.0f}"
        else:
            formatted_value = f"{prefix}{value:,}"
        
        st.markdown(f'<div class="kpi-value">{formatted_value}</div>', unsafe_allow_html=True)
    
    if change is not None:
        with col2:
            change_symbol = "↗️" if change > 0 else "↘️" if change < 0 else "➡️"
            change_class = "positive" if change > 0 else "negative" if change < 0 else ""
            change_text = f"{change:+.1f}%"
            
            st.markdown(
                f'<div class="kpi-change {change_class}">{change_symbol} {change_text}</div>',
                unsafe_allow_html=True
            )

# Main Dashboard
def main():
    # Header
    st.markdown('<div class="main-header">📊 KPI Dashboard</div>', unsafe_allow_html=True)
    
    # Date Range Selector
    col1, col2, col3, col4 = st.columns([2, 2, 1, 1])
    
    with col1:
        start_date = st.date_input(
            "Start Date",
            value=datetime.now().date() - timedelta(days=30),
            key="start_date"
        )
    
    with col2:
        end_date = st.date_input(
            "End Date",
            value=datetime.now().date(),
            key="end_date"
        )
    
    with col3:
        st.write("")  # Spacer
        fetch_button = st.button("🚀 Fetch Data", type="primary", use_container_width=True)
    
    with col4:
        st.write("")  # Spacer
        if st.button("🔄 Compare with Previous", use_container_width=True):
            st.session_state.compare_mode = not st.session_state.get("compare_mode", False)
    
    # Validate date range
    if start_date > end_date:
        st.error("Start date must be before end date.")
        st.stop()
    
    # Initialize session state
    if "current_data" not in st.session_state:
        st.session_state.current_data = None
    if "previous_data" not in st.session_state:
        st.session_state.previous_data = None
    
    # Fetch and process data when button is clicked
    if fetch_button:
        with st.spinner("Fetching data from HubSpot..."):
            # Fetch current period data
            deals = fetch_hubspot_deals(
                start_date=datetime.combine(start_date, datetime.min.time()),
                end_date=datetime.combine(end_date, datetime.max.time()),
                token=HUBSPOT_ACCESS_TOKEN
            )
            
            if deals:
                current_df = process_deals_as_customers(deals, start_date, end_date)
                st.session_state.current_data = current_df
                
                # Calculate previous period data for comparison
                previous_start = start_date - (end_date - start_date) - timedelta(days=1)
                previous_end = start_date - timedelta(days=1)
                
                previous_deals = fetch_hubspot_deals(
                    start_date=datetime.combine(previous_start, datetime.min.time()),
                    end_date=datetime.combine(previous_end, datetime.max.time()),
                    token=HUBSPOT_ACCESS_TOKEN
                )
                
                if previous_deals:
                    previous_df = process_deals_as_customers(previous_deals, previous_start, previous_end)
                    st.session_state.previous_data = previous_df
                else:
                    st.session_state.previous_data = None
                
                st.success(f"Successfully fetched {len(current_df)} customers!")
            else:
                st.error("No data found for the selected date range.")
                st.stop()
    
    # Display KPIs if data exists
    if st.session_state.current_data is not None:
        current_df = st.session_state.current_data
        previous_df = st.session_state.previous_data
        
        # Calculate KPIs
        current_kpis = calculate_kpis(current_df)
        previous_kpis = calculate_kpis(previous_df) if previous_df is not None else None
        
        # Display KPI Cards
        st.markdown("### 📈 Key Performance Indicators")
        
        # Row 1: Revenue KPIs
        col1, col2, col3 = st.columns(3)
        
        with col1:
            change = None
            if previous_kpis:
                change = current_kpis.get("Revenue Change %", 0)
            display_kpi_card("Total Revenue", current_kpis["Total Revenue"], change)
        
        with col2:
            display_kpi_card("Avg Revenue/Customer", current_kpis["Avg Revenue/Customer"], None, "$")
        
        with col3:
            change = None
            if previous_kpis:
                change = current_kpis.get("Customer Change %", 0)
            display_kpi_card("Total Customers", current_kpis["Total Customers"], change, "", False)
        
        # Row 2: Customer Status KPIs
        col1, col2, col3 = st.columns(3)
        
        with col1:
            display_kpi_card("Active Customers", current_kpis["Active Customers"], None, "", False)
        
        with col2:
            display_kpi_card("Inactive Customers", current_kpis["Inactive Customers"], None, "", False)
        
        with col3:
            display_kpi_card("Pending Customers", current_kpis["Pending Customers"], None, "", False)
        
        # Revenue Breakdown
        st.markdown("---")
        st.markdown("### 📊 Revenue Breakdown")
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("##### By Course")
            if not current_kpis["Course Revenue"].empty:
                course_df = current_kpis["Course Revenue"].reset_index()
                course_df.columns = ["Course", "Revenue"]
                course_df["Revenue"] = course_df["Revenue"].apply(lambda x: f"${x:,.0f}")
                st.dataframe(course_df, use_container_width=True, hide_index=True)
            else:
                st.info("No course revenue data available.")
        
        with col2:
            st.markdown("##### By Program")
            if not current_kpis["Program Revenue"].empty:
                program_df = current_kpis["Program Revenue"].reset_index()
                program_df.columns = ["Program", "Revenue"]
                program_df["Revenue"] = program_df["Revenue"].apply(lambda x: f"${x:,.0f}")
                st.dataframe(program_df, use_container_width=True, hide_index=True)
            else:
                st.info("No program revenue data available.")
        
        # Customer Details Table
        st.markdown("---")
        st.markdown("### 👥 Customer Details")
        
        # Add filters
        col1, col2, col3 = st.columns(3)
        
        with col1:
            status_filter = st.multiselect(
                "Filter by Status",
                options=["Active", "Inactive", "Pending"],
                default=["Active", "Inactive", "Pending"]
            )
        
        with col2:
            course_filter = st.multiselect(
                "Filter by Course",
                options=sorted(current_df["Course"].unique()),
                default=[]
            )
        
        with col3:
            min_amount, max_amount = st.slider(
                "Filter by Amount",
                min_value=0,
                max_value=int(current_df["Amount"].max()) if not current_df.empty else 10000,
                value=(0, int(current_df["Amount"].max()) if not current_df.empty else 10000),
                step=1000
            )
        
        # Apply filters
        filtered_df = current_df.copy()
        
        if status_filter:
            filtered_df = filtered_df[filtered_df["Status"].isin(status_filter)]
        
        if course_filter:
            filtered_df = filtered_df[filtered_df["Course"].isin(course_filter)]
        
        filtered_df = filtered_df[
            (filtered_df["Amount"] >= min_amount) & 
            (filtered_df["Amount"] <= max_amount)
        ]
        
        # Display filtered table
        if not filtered_df.empty:
            # Format Amount column
            display_df = filtered_df.copy()
            display_df["Amount"] = display_df["Amount"].apply(lambda x: f"${x:,.0f}")
            
            st.dataframe(
                display_df,
                use_container_width=True,
                column_config={
                    "Customer Name": st.column_config.TextColumn("Customer", width="large"),
                    "Status": st.column_config.TextColumn("Status", width="small"),
                    "Amount": st.column_config.TextColumn("Amount", width="medium"),
                    "Course": st.column_config.TextColumn("Course", width="medium"),
                    "Program": st.column_config.TextColumn("Program", width="medium"),
                    "Close Date": st.column_config.TextColumn("Close Date", width="small")
                },
                hide_index=True
            )
            
            # Export option
            csv = filtered_df.to_csv(index=False).encode('utf-8')
            st.download_button(
                label="📥 Export as CSV",
                data=csv,
                file_name=f"customers_{start_date}_to_{end_date}.csv",
                mime="text/csv",
                use_container_width=True
            )
        else:
            st.info("No customers match the selected filters.")
    
    else:
        # Initial state - show instructions
        st.info("👆 Select a date range and click 'Fetch Data' to load your KPI dashboard.")
        
        # Quick stats about what the dashboard does
        with st.expander("📖 How this dashboard works"):
            st.markdown("""
            ### Business Logic Applied:
            
            **Total Revenue Calculation (Corrected):**
            - ✅ Counts full deal amount by default
            - ✅ **Subtracts partial payments** ONLY if payment date is **OUTSIDE** selected date range
            - ✅ **Keeps partial payments** if payment date is **INSIDE** selected date range
            - ✅ Never shows negative revenue (safeguard)
            
            **Example:**
            - Deal Amount: $50,000
            - Partial Payment: $15,000 on March 15
            - Dashboard Date Range: April 1-5
            - **Result:** $35,000 (subtracts March payment)
            
            **KPI Metrics:**
            1. **Total Revenue** - Corrected for partial payments
            2. **Customer Counts** - Active/Inactive/Pending
            3. **Avg Revenue/Customer**
            4. **Course-wise Revenue**
            5. **Program-wise Revenue**
            
            **Data Source:** HubSpot Deals with partial payment tracking
            """)

if __name__ == "__main__":
    main()
