import os
import toml
import logging
from datetime import datetime, timedelta
import pandas as pd
from flask import Flask, render_template, request, jsonify
from flask_caching import Cache
from google.analytics.data_v1beta import BetaAnalyticsDataClient
from google.analytics.data_v1beta.types import DateRange, Metric, Dimension, RunReportRequest, RunRealtimeReportRequest, OrderBy
from google.oauth2 import service_account
import requests
import pytz
import concurrent.futures
import time
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
from whitenoise import WhiteNoise


# Configure Logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = Flask(__name__)
app.wsgi_app = WhiteNoise(app.wsgi_app, root='static/', prefix='static/')
app.config['TEMPLATES_AUTO_RELOAD'] = True
app.jinja_env.auto_reload = True
cache = Cache(app, config={'CACHE_TYPE': 'SimpleCache'})

# ---------------- CONFIGURATION ----------------
PROPERTY_ID = "281698779"
KEY_PATH = "ga4-streamlit-connect-21d2d2cc35d6.json"
SECRETS_PATH = ".streamlit/secrets.toml"
CONTENT_CALENDAR_URL = "https://script.google.com/macros/s/AKfycbxShw7pKMN19c7Ou4qQb6BY_cKxNlOJAnpA2gbtngXhNig0VrU39cQ5j7e_mfEkbJqJbg/exec"

# [OK] Calibration Constants from nn.txt
LEAD_STATUS_MAP = {
    "not_interest": "Not Interested",
    "not_qualif": "Not Qualified",
    "unqualif": "Not Qualified",
    "duplicate": "Duplicate",
    "junk": "Duplicate",
    "new": "New Lead",
    "open": "New Lead",
    "qualified": "Not Qualified",
    "not_connect": "Not Connected (NC)",
    "nc": "Not Connected (NC)"
}

CUSTOMER_KEYWORDS_BLOCKLIST = [
    "admission", "confirmed", "customer", "closed won", "closedwon", "won", "converted"
]

# [OK] Statuses that should NOT count towards Total Leads or Deal Leads
LEAD_EXCLUSION_STATUSES = [
    "Not Connected", "Not Connected (NC)", "Not Interested", "Not Qualified", "Duplicate", "Closed Lost", "None", "Unknown"
]

# Load Secrets
try:
    if os.path.exists(SECRETS_PATH):
        secrets = toml.load(SECRETS_PATH)
    else:
        # Fallback to reconstructing secrets from Environment Variables (for AWS)
        secrets = {
            "kajabi": {
                "client_id": os.environ.get("KAJABI_CLIENT_ID"),
                "client_secret": os.environ.get("KAJABI_CLIENT_SECRET")
            },
            "hubspot": {
                "token": os.environ.get("HUBSPOT_TOKEN")
            },
            "ga4": {
                "property_id": os.environ.get("GA4_PROPERTY_ID", PROPERTY_ID)
            },
            "ga4_key": {
                "type": os.environ.get("GA4_TYPE"),
                "project_id": os.environ.get("GA4_PROJECT_ID"),
                "private_key_id": os.environ.get("GA4_PRIVATE_KEY_ID"),
                "private_key": os.environ.get("GA4_PRIVATE_KEY", "").replace("\\n", "\n"),
                "client_email": os.environ.get("GA4_CLIENT_EMAIL"),
                "client_id": os.environ.get("GA4_CLIENT_ID"),
                "auth_uri": os.environ.get("GA4_AUTH_URI"),
                "token_uri": os.environ.get("GA4_TOKEN_URI"),
                "auth_provider_x509_cert_url": os.environ.get("GA4_AUTH_CERT_URL"),
                "client_x509_cert_url": os.environ.get("GA4_CLIENT_CERT_URL")
            }
        }
except Exception as e:
    logger.error(f"Failed to load secrets: {e}")
    secrets = {}

# ---------------- AUTHENTICATION ----------------
def get_ga4_client():
    if os.path.exists(KEY_PATH):
        try:
            credentials = service_account.Credentials.from_service_account_file(KEY_PATH)
            return BetaAnalyticsDataClient(credentials=credentials)
        except Exception as e:
            logger.error(f"Local key error: {e}")

    if "ga4_key" in secrets:
        try:
            credentials = service_account.Credentials.from_service_account_info(secrets["ga4_key"])
            return BetaAnalyticsDataClient(credentials=credentials)
        except Exception as e:
            logger.error(f"Secrets key error: {e}")
    return None

client = get_ga4_client()

# ---------------- HELPERS ----------------
def calculate_cumulative(daily_data, key="New Users"):
    """Build cumulative sum list for worm graph."""
    if not daily_data:
        return []
    df_c = pd.DataFrame(daily_data)
    if df_c.empty or key not in df_c.columns:
        return []
    df_c['Cumulative'] = df_c[key].cumsum()
    return df_c[['Date', 'Cumulative']].to_dict('records')


def get_comparison_dates(offset=0, period='monthly', c_start1=None, c_end1=None, c_start2=None, c_end2=None):
    """Return (m1_start, m1_end, m2_start, m2_end, base_date) using IST explicitly.
    period can be 'daily', 'weekly', 'monthly', 'quarterly', 'custom', 'all'."""
    ist = pytz.timezone('Asia/Kolkata')
    now_ist = datetime.now(ist)
    today = now_ist.date()
    
    if period == 'daily':
        target_date = today - timedelta(days=offset)
        m1_start = target_date
        m1_end = target_date
        m2_start = target_date - timedelta(days=1)
        m2_end = m2_start
        return m1_start, m1_end, m2_start, m2_end, now_ist
        
    elif period == 'weekly':
        # Monday = 0, Sunday = 6
        day_of_week = today.weekday()
        this_monday = today - timedelta(days=day_of_week)
        # Shift weeks backwards by offset
        target_monday = this_monday - timedelta(weeks=offset)
        
        m1_start = target_monday
        # if this is the current week, end today, else end on Sunday
        if offset == 0:
            m1_end = today
            days_passed = day_of_week
        else:
            m1_end = target_monday + timedelta(days=6)
            days_passed = 6
            
        m2_start = target_monday - timedelta(weeks=1)
        m2_end = m2_start + timedelta(days=days_passed)
        return m1_start, m1_end, m2_start, m2_end, now_ist
        
    elif period == 'quarterly':
        target_date = now_ist.date()
        # Find current quarter start month (1, 4, 7, 10)
        curr_q_start_month = 3 * ((target_date.month - 1) // 3) + 1
        curr_q_start_date = target_date.replace(month=curr_q_start_month, day=1)
        
        # Apply offset by going back 3 months
        for _ in range(offset):
            # Go to the last day of the previous quarter
            curr_q_start_date = (curr_q_start_date - timedelta(days=1)).replace(day=1)
            # Find the new quarter start
            curr_q_start_month = 3 * ((curr_q_start_date.month - 1) // 3) + 1
            curr_q_start_date = curr_q_start_date.replace(month=curr_q_start_month, day=1)
            
        m1_start = curr_q_start_date
        
        # End of the targeted quarter
        end_month = m1_start.month + 2
        if end_month > 12:
            q_end_date = m1_start.replace(year=m1_start.year + 1, month=end_month - 12, day=1)
            q_end_date = (q_end_date.replace(month=q_end_date.month % 12 + 1) - timedelta(days=1))
        else:
            q_end_date = m1_start.replace(month=end_month)
            if q_end_date.month == 12:
                q_end_date = q_end_date.replace(year=q_end_date.year + 1, month=1, day=1) - timedelta(days=1)
            else:
                q_end_date = q_end_date.replace(month=q_end_date.month + 1, day=1) - timedelta(days=1)

        # If it's the current ongoing quarter, end at today
        if m1_start.year == today.year and curr_q_start_month <= today.month <= curr_q_start_month + 2 and offset == 0:
            m1_end = today
        else:
            m1_end = q_end_date
            
        days_passed = (m1_end - m1_start).days
        
        # Month 2 (Previous quarter comparison)
        prev_q_last_day = m1_start - timedelta(days=1)
        prev_q_start_month = 3 * ((prev_q_last_day.month - 1) // 3) + 1
        m2_start = prev_q_last_day.replace(month=prev_q_start_month, day=1)
        
        m2_end = m2_start + timedelta(days=days_passed)
        if m2_end > prev_q_last_day:
            m2_end = prev_q_last_day

        return m1_start, m1_end, m2_start, m2_end, now_ist
        
    elif period == 'custom':
        # Default to today if parsing fails
        try:
            m1_start = datetime.strptime(c_start1, '%Y-%m-%d').date() if c_start1 else today
            m1_end = datetime.strptime(c_end1, '%Y-%m-%d').date() if c_end1 else today
            m2_start = datetime.strptime(c_start2, '%Y-%m-%d').date() if c_start2 else m1_start - timedelta(days=1)
            m2_end = datetime.strptime(c_end2, '%Y-%m-%d').date() if c_end2 else m2_start
        except Exception:
            m1_start = m1_end = m2_start = m2_end = today
            
        return m1_start, m1_end, m2_start, m2_end, now_ist
            
    elif period == 'all':
        m1_start = datetime(2020, 1, 1).date()
        m1_end = today
        m2_start = m1_start - timedelta(days=1) # dummy
        m2_end = m2_start # dummy
        return m1_start, m1_end, m2_start, m2_end, now_ist
        
    else: # monthly
        target_date = now_ist
        if offset > 0:
            for _ in range(offset):
                target_date = target_date.replace(day=1) - timedelta(days=1)
        
        # Month 1 (Current or selected)
        m1_start = target_date.date().replace(day=1)
        if target_date.year == now_ist.year and target_date.month == now_ist.month:
            m1_end = now_ist.date()
        else:
            # Last day of that month
            if target_date.month == 12:
                next_m = target_date.replace(year=target_date.year + 1, month=1, day=1)
            else:
                next_m = target_date.replace(month=target_date.month + 1, day=1)
            m1_end = (next_m - timedelta(days=1)).date()
            
        days_passed = (m1_end - m1_start).days
        
        # Month 2 (Previous month comparison)
        prev_m_last_day = m1_start - timedelta(days=1)
        m2_start = prev_m_last_day.replace(day=1)
        
        m2_end = m2_start + timedelta(days=days_passed)
        if m2_end > prev_m_last_day:
            m2_end = prev_m_last_day
            
        return m1_start, m1_end, m2_start, m2_end, now_ist


def _chart_end_date(end_date_obj, ist):
    """Return end date for chart. Cumulative charts never drop, so always include today."""
    return end_date_obj


# ---------------- GA4 ----------------
@cache.cached(timeout=30, key_prefix='active_users')
def get_active_users():
    if not client:
        return 0
    try:
        req = RunRealtimeReportRequest(
            property=f"properties/{PROPERTY_ID}",
            metrics=[Metric(name="activeUsers")]
        )
        resp = client.run_realtime_report(req, timeout=5)
        return int(resp.rows[0].metric_values[0].value) if resp and resp.rows else 0
    except Exception:
        return 0


@cache.memoize(timeout=600)
def get_discover_metrics(start_date, end_date):
    if not client:
        return 0
    try:
        req = RunReportRequest(
            property=f"properties/{PROPERTY_ID}",
            date_ranges=[DateRange(start_date=str(start_date), end_date=str(end_date))],
            metrics=[Metric(name="newUsers")]
        )
        resp = client.run_report(req)
        return int(resp.rows[0].metric_values[0].value) if resp.rows else 0
    except Exception:
        return 0


@cache.memoize(timeout=600)
def get_daily_new_users(start_date, end_date):
    if not client:
        return []
    try:
        req = RunReportRequest(
            property=f"properties/{PROPERTY_ID}",
            date_ranges=[DateRange(start_date=str(start_date), end_date=str(end_date))],
            dimensions=[Dimension(name="date")],
            metrics=[Metric(name="newUsers")],
            order_bys=[OrderBy(dimension=OrderBy.DimensionOrderBy(dimension_name="date"), desc=False)]
        )
        resp = client.run_report(req)
        data = []
        for row in resp.rows:
            d = datetime.strptime(row.dimension_values[0].value, "%Y%m%d").strftime('%Y-%m-%d')
            data.append({"Date": d, "New Users": int(row.metric_values[0].value)})
        return data
    except Exception:
        return []


@cache.memoize(timeout=600)
def get_traffic_sources(start_date, end_date):
    if not client:
        return pd.DataFrame()
    try:
        req = RunReportRequest(
            property=f"properties/{PROPERTY_ID}",
            date_ranges=[DateRange(start_date=str(start_date), end_date=str(end_date))],
            dimensions=[Dimension(name="sessionSourceMedium")],
            metrics=[Metric(name="newUsers")],
            order_bys=[OrderBy(dimension=OrderBy.DimensionOrderBy(dimension_name="newUsers"), desc=True)],
            limit=8
        )
        resp = client.run_report(req)
        data = [{"Channel": r.dimension_values[0].value, "New Users": int(r.metric_values[0].value)} for r in resp.rows]
        return pd.DataFrame(data)
    except Exception:
        return pd.DataFrame()


# ---------------- HUBSPOT ----------------
@cache.memoize(timeout=600)
def get_hubspot_token():
    try:
        # Check both secrets and os.environ for token
        token = secrets.get("hubspot", {}).get("token") or os.environ.get("HUBSPOT_TOKEN")
        return token.strip() if token else None
    except Exception:
        return None


# ---------------- STATUS NORMALIZATION ----------------
def normalize_lead_status(raw_status, close_date=None, start_date=None, end_date=None):
    """
    Standardized status normalization from nn.txt.
    Returns "Qualified Lead" only if the contact converted within the date range.
    """
    if not raw_status:
        return "Unknown"
    
    status = str(raw_status).strip().lower()
    
    # Check for Closed Lost
    if "closed lost" in status or "closed_lost" in status:
        return "Closed Lost"
    
    # Check for Customer Keywords
    for keyword in CUSTOMER_KEYWORDS_BLOCKLIST:
        if keyword in status:
            is_in_range = True
            if close_date and start_date and end_date:
                try:
                    if isinstance(close_date, str):
                        if close_date.isdigit():
                            c_dt = datetime.fromtimestamp(int(close_date)/1000).date()
                        else:
                            c_dt = datetime.strptime(close_date[:10], "%Y-%m-%d").date()
                    else:
                        c_dt = close_date
                    
                    if not (start_date <= c_dt <= end_date):
                        is_in_range = False
                except:
                    pass
            
            if is_in_range:
                return "Qualified Lead"
            
            # Downgrade if out of range to avoid double counting converted leads as active pipeline
            if "hot" in status: return "Hot"
            elif "warm" in status: return "Warm"
            else: return "Hot"

    # Map Pipeline Statuses
    if "prospect" in status:
        if "hot" in status: return "Hot"
        elif "warm" in status: return "Warm"
        elif "cold" in status or "neutral" in status: return "Cold"
        else: return "Warm"
    
    if "qualified" in status:
        return "Not Qualified" # Per nn.txt logic
        
    return LEAD_STATUS_MAP.get(status, status.replace("_", " ").title())

EXCLUDED_OWNERS = ["Mahalekshmi M J", "Sreeja Anoop", "arya.krishnan", "Devi Krishna", "Aneesha S", "Sonia William"]

@cache.memoize(timeout=3600)
def fetch_hubspot_owners():
    """Fetch HubSpot owner mapping IDs to names."""
    token = get_hubspot_token()
    if not token:
        return {}
    
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    url = "https://api.hubapi.com/crm/v3/owners"
    mapping = {}
    
    try:
        while True:
            response = requests.get(url, headers=headers, params={"limit": 100}, timeout=15)
            if response.status_code == 200:
                data = response.json()
                for owner in data.get("results", []):
                    o_id = str(owner.get("id"))
                    fn = owner.get("firstName", "")
                    ln = owner.get("lastName", "")
                    email = owner.get("email", "")
                    name = f"{fn} {ln}".strip() if fn or ln else email.split("@")[0] if email else f"Owner {o_id}"
                    mapping[o_id] = name
                
                paging = data.get("paging", {})
                next_link = paging.get("next", {}).get("link")
                if not next_link:
                    break
                url = next_link
            else:
                break
    except Exception as e:
        print(f"Error fetching owners: {e}")
    return mapping

@cache.memoize(timeout=3600)  # Cache pipeline stages for 1 hour
def fetch_deal_pipeline_stages():
    """Fetch deal pipeline stages to get correct stage IDs."""
    token = get_hubspot_token()
    if not token:
        return {}
    
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    url = "https://api.hubapi.com/crm/v3/pipelines/deals"
    
    try:
        response = requests.get(url, headers=headers, timeout=10)
        if response.status_code == 200:
            data = response.json()
            pipelines = data.get("results", [])
            all_stages = {}
            for pipeline in pipelines:
                p_id = pipeline.get("id")
                p_label = pipeline.get("label")
                for stage in pipeline.get("stages", []):
                    s_id = stage.get("id")
                    s_label = stage.get("label")
                    prob = stage.get("metadata", {}).get("probability", "0")
                    all_stages[s_id] = {
                        "id": s_id,
                        "label": s_label,
                        "pipeline_id": p_id,
                        "pipeline_label": p_label,
                        "probability": prob
                    }
            return all_stages
    except Exception as e:
        print(f"Error fetching pipelines: {e}")
    return {}

# HubSpot Session with retries
hs_session = requests.Session()
retries = Retry(total=3, backoff_factor=1, status_forcelist=[429, 500, 502, 503, 504])
hs_session.mount('https://', HTTPAdapter(max_retries=retries))

def hubspot_api_request(method, url, **kwargs):
    """Reliable HubSpot API request with exponential backoff for 429 and 5xx."""
    max_retries = 3
    for attempt in range(max_retries):
        try:
            resp = hs_session.request(method, url, **kwargs)
            if resp.status_code == 429:
                # Manual backoff for 429 just in case HTTPAdapter missed it
                wait = (attempt + 1) * 2
                logger.warning(f"HubSpot Rate Limit (429). Retrying in {wait}s...")
                time.sleep(wait)
                continue
            if resp.status_code >= 500:
                wait = attempt + 1
                logger.warning(f"HubSpot 5xx Error ({resp.status_code}). Retrying in {wait}s...")
                time.sleep(wait)
                continue
            return resp
        except Exception as e:
            if attempt == max_retries - 1:
                raise e
            time.sleep(attempt + 1)
    return None

def get_hubspot_deals(target_month_start, target_month_end, course_filter=None):
    """
    Fetch closed/won HubSpot deals in the date range.
    Returns (count, total_revenue, daily_trend_list, course_breakdown, course_breakdown_ungrouped, owner_breakdown).
    """
    token = get_hubspot_token()
    if not token:
        logger.warning("HubSpot Token missing - returning 0")
        return 0, 0, [], {}, {}, {}

    # OLD LOGIC: Use hardcoded admission stages
    customer_stages = ["closedwon", "1884422889", "2208152296", "1955461874", "1955461879", "contractsent"]
    owner_map = fetch_hubspot_owners()
    
    url = "https://api.hubapi.com/crm/v3/objects/deals/search"
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}

    ist = pytz.timezone('Asia/Kolkata')
    dt_start_ist = ist.localize(datetime.combine(target_month_start, datetime.min.time()))
    start_ts = int(dt_start_ist.astimezone(pytz.UTC).timestamp() * 1000)

    dt_end_ist = ist.localize(datetime.combine(target_month_end, datetime.max.time()))
    end_ts = int(dt_end_ist.astimezone(pytz.UTC).timestamp() * 1000)

    all_results = []
    # If course filter is applied, we can't easily filter by "Main Course" in HubSpot Search
    # because Main Course is a result of get_main_course() (Python logic).
    # So we fetch all won deals and filter in memory.
    
    # ... (rest of fetch logic remains same)
    after = None
    page_count = 1

    while True:
        body = {
            "filterGroups": [{
                "filters": [
                    {"propertyName": "dealstage", "operator": "IN", "values": customer_stages},
                    {"propertyName": "closedate", "operator": "BETWEEN", "value": start_ts, "highValue": end_ts}
                ]
            }],
            "properties": ["amount", "closedate", "dealstage", "dealname", "hubspot_owner_id",
                           "course", "program", "product", "service", "offering",
                           "course_name", "program_name", "product_name",
                           "partial_amount", "hs_v2_date_entered_2107527928", "hs_v2_date_entered_2171957962"],
            "limit": 100
        }
        if after:
            body['after'] = after

        response = hubspot_api_request("POST", url, headers=headers, json=body, timeout=25)
        if response and response.status_code == 200:
            data = response.json()
            results = data.get('results', [])
            
            # Filter out EXCLUDED_OWNERS
            for r in results:
                o_name = owner_map.get(str(r['properties'].get('hubspot_owner_id', '')), "")
                if o_name in EXCLUDED_OWNERS:
                    continue
                all_results.append(r)

            paging = data.get('paging', {})
            after = paging.get('next', {}).get('after')
            if not after:
                break
            
            page_count += 1
            time.sleep(0.4) # Respect secondary limits
    # Filtering and Breakdown
    total_val = 0
    course_breakdown = {}
    course_breakdown_ungrouped = {}
    owner_breakdown = {}
    daily_counts = {}
    
    for r in all_results:
        props = r['properties']
        
        # Determine Course Name for this deal
        course_name = ''
        course_fields = ["course", "program", "product", "service", "offering",
                         "course_name", "program_name", "product_name"]
        for field in course_fields:
            val = props.get(field)
            if val and str(val).strip():
                course_name = str(val).strip()
                break
        
        if not course_name:
            course_name = "Other"

        main_course = get_main_course(course_name)
        
        # APPLY GLOBAL FILTER
        if course_filter and main_course.lower() != course_filter.lower():
            continue

        amt = 0
        amt_str = props.get('amount')
        if amt_str:
            try:
                amt = float(str(amt_str).replace(",", ""))
            except: pass
            
        # Partial Payment Deduction Logic
        partial_amount = 0
        partial_amount_str = props.get("partial_amount", "0")
        if partial_amount_str:
            try:
                partial_amount = float(str(partial_amount_str).replace(",", ""))
            except: pass
            
        if partial_amount > 0:
            partial_ts_online = props.get("hs_v2_date_entered_2107527928", "")
            partial_ts_offline = props.get("hs_v2_date_entered_2171957962", "")
            partial_ts = partial_ts_online or partial_ts_offline
            
            if partial_ts:
                try:
                    partial_dt = datetime.fromisoformat(partial_ts.replace('Z', '+00:00'))
                    if partial_dt < dt_start_ist.astimezone(pytz.UTC):
                        amt = amt - partial_amount
                except Exception:
                    pass
        
        total_val += amt
        
        # Course Breakdown
        if main_course not in course_breakdown:
            course_breakdown[main_course] = {"count": 0, "revenue": 0}
        course_breakdown[main_course]["count"] += 1
        course_breakdown[main_course]["revenue"] += amt

        if course_name not in course_breakdown_ungrouped:
            course_breakdown_ungrouped[course_name] = {"count": 0, "revenue": 0}
        course_breakdown_ungrouped[course_name]["count"] += 1
        course_breakdown_ungrouped[course_name]["revenue"] += amt
        
        # Owner Breakdown
        o_id = str(props.get('hubspot_owner_id', ''))
        o_name = owner_map.get(o_id, f"Unknown ({o_id})" if o_id else "Unassigned")
        if o_name not in owner_breakdown:
            owner_breakdown[o_name] = {"count": 0, "revenue": 0}
        owner_breakdown[o_name]["count"] += 1
        owner_breakdown[o_name]["revenue"] += amt

        # Daily trend should track admission counts, not revenue, for the worm graph.
        try:
            close_str = props.get('closedate')
            if close_str:
                if str(close_str).isdigit():
                    dt = datetime.fromtimestamp(int(close_str)/1000, tz=pytz.UTC)
                else:
                    dt = datetime.fromisoformat(str(close_str).replace('Z', '+00:00'))
                date_str = dt.astimezone(ist).strftime('%Y-%m-%d')
                daily_counts[date_str] = daily_counts.get(date_str, 0) + 1
        except: pass

    # Format daily_trend_list
    daily_trend_list = []
    curr = target_month_start
    while curr <= target_month_end:
        ds = curr.strftime('%Y-%m-%d')
        daily_trend_list.append({"Date": ds, "Count": daily_counts.get(ds, 0)})
        curr += timedelta(days=1)

    # Count = sum of all filtered results (already filtered in the loop above via continue)
    total_count = sum(v["count"] for v in course_breakdown.values())
    return total_count, total_val, daily_trend_list, course_breakdown, course_breakdown_ungrouped, owner_breakdown

def get_hubspot_contacts(target_month_start, target_month_end, lead_status="Trial Phase", course_filter=None):
    """
    Fetch HubSpot contacts with a specific lead status (Trial Phase).
    Returns (count, course_breakdown_dict, daily_trend_list).
    """
    token = get_hubspot_token()
    if not token:
        logger.warning("HubSpot Token missing - returning 0 for contacts")
        return 0, {}, []

    url = "https://api.hubapi.com/crm/v3/objects/contacts/search"
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}

    # Define course properties to check (matching ff.txt logic)
    course_props = [
        "course", "program", "product", "service", "offering",
        "course_name", "program_name", "product_name",
        "enquired_course", "interested_course", "course_interested",
        "program_of_interest", "course_of_interest", "product_of_interest"
    ]
    all_props = ["hs_lead_status", "hubspot_owner_id", "createdate"] + course_props
    
    ist = pytz.timezone('Asia/Kolkata')
    dt_start_ist = ist.localize(datetime.combine(target_month_start, datetime.min.time()))
    start_ts = int(dt_start_ist.astimezone(pytz.UTC).timestamp() * 1000)

    dt_end_ist = ist.localize(datetime.combine(target_month_end, datetime.max.time()))
    end_ts = int(dt_end_ist.astimezone(pytz.UTC).timestamp() * 1000)

    all_contacts = []
    after = None
    page_count = 1

    while True:
        body = {
            "filterGroups": [{
                "filters": [
                    {"propertyName": "hs_lead_status", "operator": "EQ", "value": lead_status},
                    {"propertyName": "createdate", "operator": "BETWEEN", "value": start_ts, "highValue": end_ts}
                ]
            }],
            "properties": all_props,
            "limit": 100
        }
        if after:
            body['after'] = after

        response = hubspot_api_request("POST", url, headers=headers, json=body, timeout=25)
        if response and response.status_code == 200:
            data = response.json()
            results = data.get('results', [])
            all_contacts.extend(results)

            paging = data.get('paging', {})
            after = paging.get('next', {}).get('after')
            if not after:
                break
            
            page_count += 1
            time.sleep(0.4)
        else:
            msg = f"HubSpot Contacts Fetch failed on page {page_count}"
            logger.error(msg)
            raise Exception(msg)

    # Process results
    course_counts = {}
    daily_trend = {}
    
    for c in all_contacts:
        props = c.get('properties', {})
        
        # 1. Daily trend
        create_str = props.get('createdate')
        if create_str:
            try:
                # HubSpot search returns timestamps for createdate
                if create_str.isdigit():
                    dt = datetime.fromtimestamp(int(create_str)/1000, tz=pytz.UTC)
                else:
                    dt = datetime.fromisoformat(create_str.replace('Z', '+00:00'))
                date_str = dt.astimezone(ist).strftime('%Y-%m-%d')
                daily_trend[date_str] = daily_trend.get(date_str, 0) + 1
            except:
                pass

        # 2. Course breakdown
        raw_course = "Unknown Course"
        for field in course_props:
            val = props.get(field)
            if val and str(val).strip():
                raw_course = str(val).strip()
                break
        
        main_course = get_main_course(raw_course)
        
        # APPLY GLOBAL FILTER
        if course_filter and main_course.lower() != course_filter.lower():
            continue
            
        course_counts[main_course] = course_counts.get(main_course, 0) + 1

    # Format trend for worm graph
    sorted_trend = []
    curr = target_month_start
    while curr <= target_month_end:
        ds = curr.strftime('%Y-%m-%d')
        sorted_trend.append({"Date": ds, "Admissions": daily_trend.get(ds, 0)})
        curr += timedelta(days=1)

    return len(all_contacts), course_counts, sorted_trend

def _fetch_all_results_paginated(url, headers, body, results_key='results'):
    """Helper to fetch all pages of a HubSpot search/list."""
    all_results = []
    after = None
    while True:
        if after: body['after'] = after
        try:
            resp = hubspot_api_request("POST", url, headers=headers, json=body, timeout=25).json()
            results = resp.get(results_key, [])
            all_results.extend(results)
            paging = resp.get('paging', {})
            after = paging.get('next', {}).get('after')
            if not after: break
            # Removed time.sleep(0.3) for speed - HubSpot rate limits are usually high enough for these bursts
        except Exception as e:
            logger.error(f"HubSpot Pagination Error: {e}")
            break
    return all_results

LEAD_CONTACT_COURSE_FIELDS = [
    "course", "enquired_course", "program", "program_of_interest",
    "product", "service", "offering",
    "course_name", "program_name", "product_name",
    "interested_course", "course_interested",
    "course_of_interest", "product_of_interest"
]

LEAD_DEAL_COURSE_FIELDS = [
    "course", "program", "product", "service", "offering",
    "course_name", "program_name", "product_name"
]

def _parse_hubspot_datetime(raw_value):
    """Parse HubSpot timestamps that may be epoch milliseconds or ISO strings."""
    if not raw_value:
        return None
    try:
        text = str(raw_value)
        if text.isdigit():
            return datetime.fromtimestamp(int(text) / 1000, tz=pytz.UTC)
        return datetime.fromisoformat(text.replace('Z', '+00:00'))
    except Exception:
        return None

def _resolve_contact_owner_id(contact):
    """Resolve contact owner using the same fallback order as the reference analytics."""
    properties = contact.get("properties", {})
    owner_id = (
        properties.get("hubspot_owner_id")
        or properties.get("hs_assigned_owner_id")
        or ""
    )
    return str(owner_id) if owner_id else ""

def _is_excluded_owner_id(owner_id, owner_map):
    """Check whether an owner id maps to an excluded owner name."""
    if not owner_id:
        return False
    owner_name = owner_map.get(str(owner_id), "")
    return owner_name in EXCLUDED_OWNERS

def _chunked(items, size):
    """Yield fixed-size chunks from a list."""
    for idx in range(0, len(items), size):
        yield items[idx:idx + size]

def _fetch_contact_status_history_map(contact_ids, headers):
    """Fetch lead-status history in safe HubSpot-sized batches."""
    history_by_id = {}
    if not contact_ids:
        return history_by_id

    def fetch_batch(batch_ids):
        batch_history = {}
        body = {
            "inputs": [{"id": contact_id} for contact_id in batch_ids],
            "propertiesWithHistory": ["hs_lead_status", "lead_status"]
        }
        try:
            resp = hubspot_api_request("POST", batch_url, headers=headers, json=body, timeout=25)
            if resp.status_code != 200:
                logger.warning(f"Contact history batch failed ({resp.status_code}): {resp.text[:300]}")
                return batch_history
            for result in resp.json().get("results", []):
                batch_history[str(result.get("id"))] = result.get("propertiesWithHistory", {})
        except Exception as e:
            logger.warning(f"Failed to load contact status history batch: {e}")
        return batch_history

    batch_url = "https://api.hubspot.com/crm/v3/objects/contacts/batch/read"
    batches = list(_chunked(list(contact_ids), 50))
    with concurrent.futures.ThreadPoolExecutor(max_workers=5) as executor:
        futures = [executor.submit(fetch_batch, batch_ids) for batch_ids in batches]
        for future in concurrent.futures.as_completed(futures):
            history_by_id.update(future.result())
    return history_by_id

@cache.memoize(timeout=300)
def get_hubspot_contact_status_histories(contact_ids):
    """Fetch HubSpot lead-status history only for contacts that need a historical cutoff lookup."""
    if not contact_ids:
        return {}

    token = get_hubspot_token()
    if not token:
        return {}

    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    unique_ids = [str(contact_id) for contact_id in dict.fromkeys(contact_ids) if contact_id]
    return _fetch_contact_status_history_map(unique_ids, headers)

def _get_contact_status_as_of(contact, end_date, properties_with_history=None):
    """Return the last known lead status on or before the snapshot cutoff."""
    history = properties_with_history or contact.get("propertiesWithHistory", {})
    ist = pytz.timezone('Asia/Kolkata')
    cutoff_local = ist.localize(datetime.combine(end_date, datetime.max.time()))
    cutoff_utc = cutoff_local.astimezone(pytz.UTC)

    for prop_name in ("hs_lead_status", "lead_status"):
        history_entries = history.get(prop_name, [])
        valid_entries = []
        for entry in history_entries:
            ts = _parse_hubspot_datetime(entry.get("timestamp"))
            value = entry.get("value")
            if ts and value not in (None, "") and ts <= cutoff_utc:
                valid_entries.append((ts, value))
        if valid_entries:
            valid_entries.sort(key=lambda item: item[0], reverse=True)
            return valid_entries[0][1]

    props = contact.get("properties", {})
    return props.get("hs_lead_status") or props.get("lead_status") or ""

@cache.memoize(timeout=300)
def get_hubspot_lead_contacts(start_date, end_date):
    """Fetch raw HubSpot contacts once so multiple lead views can share the same snapshot."""
    token = get_hubspot_token()
    if not token:
        return []

    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    ist = pytz.timezone('Asia/Kolkata')
    dt_start_ist = ist.localize(datetime.combine(start_date, datetime.min.time()))
    start_ts = int(dt_start_ist.astimezone(pytz.UTC).timestamp() * 1000)
    dt_end_ist = ist.localize(datetime.combine(end_date, datetime.max.time()))
    end_ts = int(dt_end_ist.astimezone(pytz.UTC).timestamp() * 1000)
    owner_map = fetch_hubspot_owners()

    contacts_url = "https://api.hubspot.com/crm/v3/objects/contacts/search"
    contacts_body = {
        "filterGroups": [{
            "filters": [
                {"propertyName": "createdate", "operator": "BETWEEN", "value": start_ts, "highValue": end_ts}
            ]
        }],
        "properties": ["hs_lead_status", "lead_status", "createdate", "closedate", "lastmodifieddate", "hubspot_owner_id", "hs_assigned_owner_id"] + LEAD_CONTACT_COURSE_FIELDS,
        "limit": 100
    }
    all_contacts = _fetch_all_results_paginated(contacts_url, headers, contacts_body)
    return [
        contact for contact in all_contacts
        if not _is_excluded_owner_id(_resolve_contact_owner_id(contact), owner_map)
    ]

def _build_hubspot_lead_snapshot(all_contacts, start_date, end_date, course_filter=None, include_status_metrics=True):
    """Convert raw HubSpot contacts into lead KPIs, trend data, and deal-status split."""
    ist = pytz.timezone('Asia/Kolkata')
    cutoff_local = ist.localize(datetime.combine(end_date, datetime.max.time()))
    cutoff_utc = cutoff_local.astimezone(pytz.UTC)
    total_leads = 0
    deal_leads = 0
    daily_trend = {}
    course_counts = {}
    deal_status_breakdown = {"Hot": 0, "Warm": 0, "Cold": 0, "Other": 0}
    relevant_contacts = []
    history_candidate_ids = []

    for contact in all_contacts:
        props = contact.get('properties', {})
        created_dt = _parse_hubspot_datetime(props.get('createdate'))
        if not created_dt:
            continue

        created_local = created_dt.astimezone(ist).date()
        if not (start_date <= created_local <= end_date):
            continue

        contact_course = "Other"
        for field in LEAD_CONTACT_COURSE_FIELDS:
            val = props.get(field)
            if val:
                contact_course = get_main_course(str(val))
                break

        if course_filter:
            match = False
            for field in LEAD_CONTACT_COURSE_FIELDS:
                val = props.get(field, "")
                if val and get_main_course(str(val)).lower() == course_filter.lower():
                    match = True
                    break
            if not match:
                continue

        relevant_contacts.append((contact, props, created_local, contact_course))
        if include_status_metrics:
            updated_dt = _parse_hubspot_datetime(contact.get("updatedAt") or props.get("lastmodifieddate"))
            if updated_dt and updated_dt > cutoff_utc and contact.get("id"):
                history_candidate_ids.append(str(contact.get("id")))

    status_history_map = {}
    if include_status_metrics and history_candidate_ids:
        status_history_map = get_hubspot_contact_status_histories(tuple(dict.fromkeys(history_candidate_ids)))

    for contact, props, created_local, contact_course in relevant_contacts:
        raw_status = props.get('hs_lead_status', '')
        status = "Unknown"
        if include_status_metrics:
            raw_status = _get_contact_status_as_of(
                contact,
                end_date,
                status_history_map.get(str(contact.get("id")))
            )
            close_date = props.get('closedate')
            status = normalize_lead_status(raw_status, close_date, start_date, end_date)

        total_leads += 1
        if contact_course not in course_counts:
            course_counts[contact_course] = {"total": 0, "deal": 0}
        course_counts[contact_course]["total"] += 1

        if include_status_metrics and status not in LEAD_EXCLUSION_STATUSES:
            deal_leads += 1
            course_counts[contact_course]["deal"] += 1
            if status in deal_status_breakdown:
                deal_status_breakdown[status] += 1
            else:
                deal_status_breakdown["Other"] += 1

        date_str = created_local.strftime('%Y-%m-%d')
        daily_trend[date_str] = daily_trend.get(date_str, 0) + 1

    course_breakdown = [
        {"course": k, "total": v["total"], "deal": v["deal"]}
        for k, v in sorted(course_counts.items(), key=lambda x: x[1]["total"], reverse=True)
    ]

    sorted_trend = []
    curr = start_date
    while curr <= end_date:
        ds = curr.strftime('%Y-%m-%d')
        sorted_trend.append({"Date": ds, "Admissions": daily_trend.get(ds, 0)})
        curr += timedelta(days=1)

    return {
        "total_leads": total_leads,
        "deal_leads": deal_leads,
        "customer_count": 0,
        "course_breakdown": course_breakdown,
        "trend": sorted_trend,
        "deal_status_breakdown": deal_status_breakdown
    }

@cache.memoize(timeout=300)
def get_hubspot_customer_count(start_date, end_date, course_filter=None):
    """Count admissions/customers from won deals for LEAD KPI alignment."""
    token = get_hubspot_token()
    if not token:
        return 0

    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    owner_map = fetch_hubspot_owners()
    ist = pytz.timezone('Asia/Kolkata')
    dt_start_ist = ist.localize(datetime.combine(start_date, datetime.min.time()))
    start_ts = int(dt_start_ist.astimezone(pytz.UTC).timestamp() * 1000)
    dt_end_ist = ist.localize(datetime.combine(end_date, datetime.max.time()))
    end_ts = int(dt_end_ist.astimezone(pytz.UTC).timestamp() * 1000)

    deals_url = "https://api.hubspot.com/crm/v3/objects/deals/search"
    customer_stages = ["closedwon", "1884422889", "2208152296", "1955461874", "1955461879", "contractsent"]
    deals_body = {
        "filterGroups": [{
            "filters": [
                {"propertyName": "dealstage", "operator": "IN", "values": customer_stages},
                {"propertyName": "closedate", "operator": "BETWEEN", "value": start_ts, "highValue": end_ts}
            ]
        }],
        "properties": ["closedate", "hubspot_owner_id"] + LEAD_DEAL_COURSE_FIELDS,
        "limit": 100
    }

    all_deals = _fetch_all_results_paginated(deals_url, headers, deals_body)
    all_deals = [
        deal for deal in all_deals
        if not _is_excluded_owner_id(deal.get("properties", {}).get("hubspot_owner_id", ""), owner_map)
    ]
    if not course_filter:
        return len(all_deals)

    filtered_count = 0
    for deal in all_deals:
        props = deal.get("properties", {})
        for field in LEAD_DEAL_COURSE_FIELDS:
            val = props.get(field, "")
            if val and get_main_course(str(val)).lower() == course_filter.lower():
                filtered_count += 1
                break
    return filtered_count

@cache.memoize(timeout=300)
def get_hubspot_leads(start_date, end_date, course_filter=None, include_customer_count=False):
    """
    Fetch and calibrate Total Leads and Deal Leads.
    Now includes concurrency for performance and course filtering.
    """
    token = get_hubspot_token()
    if not token:
        return {
            "total_leads": 0,
            "deal_leads": 0,
            "customer_count": 0,
            "trend": [],
            "deal_status_breakdown": {"Hot": 0, "Warm": 0, "Cold": 0, "Other": 0}
        }

    all_contacts = get_hubspot_lead_contacts(start_date, end_date)
    result = _build_hubspot_lead_snapshot(all_contacts, start_date, end_date, course_filter)

    if include_customer_count:
        headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
        ist = pytz.timezone('Asia/Kolkata')
        dt_start_ist = ist.localize(datetime.combine(start_date, datetime.min.time()))
        start_ts = int(dt_start_ist.astimezone(pytz.UTC).timestamp() * 1000)
        dt_end_ist = ist.localize(datetime.combine(end_date, datetime.max.time()))
        end_ts = int(dt_end_ist.astimezone(pytz.UTC).timestamp() * 1000)
        deals_url = "https://api.hubspot.com/crm/v3/objects/deals/search"
        customer_stages = ["closedwon", "1884422889", "2208152296", "1955461874", "1955461879"]
        deals_filters = [
            {"propertyName": "dealstage", "operator": "IN", "values": customer_stages},
            {"propertyName": "closedate", "operator": "BETWEEN", "value": start_ts, "highValue": end_ts}
        ]
        deals_body = {
            "filterGroups": [{"filters": deals_filters}],
            "properties": ["amount", "closedate", "dealstage"] + LEAD_DEAL_COURSE_FIELDS,
            "limit": 100
        }
        all_deals = _fetch_all_results_paginated(deals_url, headers, deals_body)
        filtered_deals_count = 0
        for deal in all_deals:
            if course_filter:
                match = False
                props = deal.get('properties', {})
                for field in LEAD_DEAL_COURSE_FIELDS:
                    val = props.get(field, "")
                    if val and get_main_course(str(val)).lower() == course_filter.lower():
                        match = True
                        break
                if not match:
                    continue
            filtered_deals_count += 1
        result["customer_count"] = filtered_deals_count

    return result


# ---------------- KAJABI ----------------
KAJABI_CLIENT_ID = secrets.get("kajabi", {}).get("client_id")
KAJABI_CLIENT_SECRET = secrets.get("kajabi", {}).get("client_secret")


@cache.memoize(timeout=3500)
def get_kajabi_token():
    url = "https://api.kajabi.com/v1/oauth/token"
    try:
        response = requests.post(
            url,
            data={
                "grant_type": "client_credentials",
                "client_id": KAJABI_CLIENT_ID,
                "client_secret": KAJABI_CLIENT_SECRET
            },
            headers={"Content-Type": "application/x-www-form-urlencoded"},
            timeout=10
        )
        if response.status_code == 200:
            return response.json().get('access_token')
        return None
    except Exception:
        return None


@cache.memoize(timeout=600)
def get_kajabi_new_customers(target_month_start, target_month_end):
    """
    Fetch Kajabi customers created in the date range.
    Returns (count, customer_list, total_global_count, daily_trend_list).
    """
    token = get_kajabi_token()
    if not token:
        return 0, [], 0, []

    url = "https://api.kajabi.com/v1/customers"
    headers = {"Authorization": f"Bearer {token}"}
    ist = pytz.timezone('Asia/Kolkata')

    dt_start_ist = ist.localize(datetime.combine(target_month_start, datetime.min.time()))
    dt_end_ist = ist.localize(datetime.combine(target_month_end, datetime.max.time()))

    all_customers = []
    next_url = f"{url}?limit=100&include=offers"
    page_count = 0
    total_global = 0

    while next_url and page_count < 50:
        try:
            resp = requests.get(next_url, headers=headers, timeout=15)
            if resp.status_code != 200:
                break
            data = resp.json()
            if page_count == 0:
                total_global = data.get('meta', {}).get('total', 0)

            batch = data.get('data', [])
            if not batch:
                break

            should_stop = False
            for c in batch:
                attrs = c.get('attributes', {})
                u_str = attrs.get('updated_at')
                if u_str:
                    try:
                        u_date = datetime.fromisoformat(u_str.replace('Z', '+00:00')).astimezone(ist)
                        if u_date < dt_start_ist:
                            should_stop = True
                    except Exception:
                        pass

                c_date_str = attrs.get('created_at')
                if c_date_str:
                    try:
                        c_date = datetime.fromisoformat(c_date_str.replace('Z', '+00:00')).astimezone(ist)
                        if dt_start_ist <= c_date <= dt_end_ist:
                            all_customers.append(c)
                    except Exception:
                        pass

            if should_stop:
                break
            next_url = data.get('links', {}).get('next')
            page_count += 1
        except Exception:
            break

    # Build daily trend
    daily_counts = {}
    for c in all_customers:
        try:
            c_date_str = c['attributes'].get('created_at')
            if c_date_str:
                c_date = datetime.fromisoformat(c_date_str.replace('Z', '+00:00')).astimezone(ist)
                d_str = c_date.strftime('%Y-%m-%d')
                daily_counts[d_str] = daily_counts.get(d_str, 0) + 1
        except Exception:
            continue

    chart_end = _chart_end_date(dt_end_ist.date(), ist)
    sorted_daily = []
    curr = dt_start_ist.date()
    while curr <= chart_end:
        d_str = curr.strftime('%Y-%m-%d')
        sorted_daily.append({"Date": d_str, "Count": daily_counts.get(d_str, 0)})
        curr += timedelta(days=1)

    return len(all_customers), all_customers, total_global, sorted_daily


@cache.memoize(timeout=600)
def get_kajabi_sales(target_month_start, target_month_end):
    """Fetch Kajabi purchases and return (total_revenue, purchase_list)."""
    token = get_kajabi_token()
    if not token:
        return 0.0, []

    url = "https://api.kajabi.com/v1/purchases"
    headers = {"Authorization": f"Bearer {token}"}
    ist = pytz.timezone('Asia/Kolkata')

    dt_start_ist = ist.localize(datetime.combine(target_month_start, datetime.min.time()))
    dt_end_ist = ist.localize(datetime.combine(target_month_end, datetime.max.time()))

    revenue = 0.0
    filtered = []
    next_url = f"{url}?limit=100"
    page_count = 0

    while next_url and page_count < 50:
        try:
            resp = requests.get(next_url, headers=headers, timeout=15)
            if resp.status_code != 200:
                break
            data = resp.json()
            batch = data.get('data', [])
            if not batch:
                break

            should_stop = False
            for p in batch:
                attrs = p.get('attributes', {})
                p_date_str = attrs.get('created_at')
                if p_date_str:
                    try:
                        p_date = datetime.fromisoformat(p_date_str.replace('Z', '+00:00')).astimezone(ist)
                        if p_date < dt_start_ist:
                            should_stop = True
                            continue
                        if dt_start_ist <= p_date <= dt_end_ist:
                            amt = float(attrs.get('amount_in_cents', 0) or 0) / 100
                            revenue += amt
                            filtered.append(p)
                    except Exception:
                        pass

            if should_stop:
                break
            next_url = data.get('links', {}).get('next')
            page_count += 1
        except Exception:
            break

    return revenue, filtered


@cache.memoize(timeout=3600)
def get_kajabi_products():
    """Fetch all Kajabi products with member counts."""
    token = get_kajabi_token()
    if not token:
        return []
    try:
        url = "https://api.kajabi.com/v1/products"
        headers = {"Authorization": f"Bearer {token}"}
        all_products = []
        next_url = f"{url}?limit=100"
        while next_url:
            resp = requests.get(next_url, headers=headers, timeout=15)
            if resp.status_code != 200:
                break
            data = resp.json()
            for prod in data.get('data', []):
                attrs = prod.get('attributes', {})
                all_products.append({
                    "Name": attrs.get('title', 'Unknown'),
                    "Members": attrs.get('members_aggregate_count', 0)
                })
            next_url = data.get('links', {}).get('next')
        return all_products
    except Exception:
        return []


@cache.memoize(timeout=600)
def get_kajabi_active_customers(start_date):
    """Return (active_count, total_customers) where active = logged in since start_date."""
    token = get_kajabi_token()
    if not token:
        return 0, 0

    url = "https://api.kajabi.com/v1/customers"
    headers = {"Authorization": f"Bearer {token}"}
    ist = pytz.timezone('Asia/Kolkata')
    dt_limit_ist = ist.localize(datetime.combine(start_date, datetime.min.time()))

    active_count = 0
    total_customers = 0
    next_url = f"{url}?limit=100"
    page = 0

    while next_url and page < 50:
        try:
            resp = requests.get(next_url, headers=headers, timeout=15)
            if resp.status_code != 200:
                break
            data = resp.json()
            if page == 0:
                total_customers = data.get('meta', {}).get('total', 0)

            batch = data.get('data', [])
            if not batch:
                break

            stop_fetching = False
            for c in batch:
                attrs = c.get('attributes', {})
                u_str = attrs.get('updated_at')
                if u_str:
                    try:
                        u_date = datetime.fromisoformat(u_str.replace('Z', '+00:00')).astimezone(ist)
                        if u_date < dt_limit_ist:
                            stop_fetching = True
                    except Exception:
                        pass

                lr_str = attrs.get('last_request_at')
                if lr_str:
                    try:
                        lr_date = datetime.fromisoformat(lr_str.replace('Z', '+00:00')).astimezone(ist)
                        if lr_date >= dt_limit_ist:
                            active_count += 1
                    except Exception:
                        pass

            if stop_fetching:
                break
            next_url = data.get('links', {}).get('next')
            page += 1
        except Exception:
            break

    return active_count, total_customers


@cache.memoize(timeout=3600)
def get_kajabi_offers():
    """Fetch all Kajabi offers and return {offer_id: offer_title} map."""
    token = get_kajabi_token()
    if not token:
        return {}
    try:
        url = "https://api.kajabi.com/v1/offers"
        headers = {"Authorization": f"Bearer {token}"}
        offers_map = {}
        next_url = f"{url}?limit=100"
        while next_url:
            resp = requests.get(next_url, headers=headers, timeout=15)
            if resp.status_code != 200:
                break
            data = resp.json()
            for item in data.get('data', []):
                offers_map[item['id']] = item['attributes']['title']
            next_url = data.get('links', {}).get('next')
        return offers_map
    except Exception:
        return {}


# ---------------- GOOGLE SHEETS (RENEW) ----------------
RENEW_SHEET_URL = "https://script.google.com/macros/s/AKfycbyid7u5OIJbemqEyawmvRsJyF6XmplsjNw-u9DqDuI7dm59hxSuykJOk2Yeeyc5riDtfg/exec"


@cache.memoize(timeout=600)
def get_renew_sheet_data(target_month_start, target_month_end, course_filter=None):
    """Fetch renewal data from Google Sheets. Returns (count, total_revenue, dataframe)."""
    if not RENEW_SHEET_URL or "script.google.com" not in RENEW_SHEET_URL:
        return 0, 0, pd.DataFrame()
    try:
        response = requests.get(RENEW_SHEET_URL, timeout=15)
        data = response.json()
        df = pd.DataFrame(data)
        
        # New Column structure handling
        date_col = 'Payment Paid Date/ Initial Amount Paid Date'
        fee_col = 'Fee Amount'
        course_col = 'Course'
        
        if fee_col in df.columns:
            df[fee_col] = pd.to_numeric(df[fee_col], errors='coerce').fillna(0)
        else:
            df[fee_col] = 0
            
        if date_col in df.columns:
            # Shift UTC to IST correctly
            dt_series = pd.to_datetime(df[date_col].astype(str).str.strip(), errors='coerce', utc=True)
            df['Parsed Date'] = dt_series.dt.tz_convert('Asia/Kolkata').dt.date
            
            mask = (df['Parsed Date'] >= target_month_start) & (df['Parsed Date'] <= target_month_end)
            
            # Apply Course Filtering
            if course_filter and course_col in df.columns:
                # Map each row's course to main course category
                df['Main Course'] = df[course_col].apply(lambda x: get_main_course(str(x)))
                mask &= (df['Main Course'].str.lower() == course_filter.lower())

            df_filtered = df.loc[mask]
            
            # Final verification log
            nimmy_list = df[df['Student Name'].astype(str).str.contains('NIMMY', case=False)]
            if not nimmy_list.empty:
                logger.info(f"NIMMY Final Check: {nimmy_list[['Student Name', 'Parsed Date']].to_dict('records')}")

            logger.info(f"Renew filtered rows: {len(df_filtered)} for range {target_month_start} to {target_month_end}")
            return len(df_filtered), df_filtered[fee_col].sum(), df_filtered
        return 0, 0, pd.DataFrame()
    except Exception as e:
        logger.error(f"Renew sheet fetch error: {e}")
        return 0, 0, pd.DataFrame()


# ---------------- GOOGLE SHEETS (ADVOCATE) ----------------
ADVOCATE_SHEET_URL = "https://script.google.com/macros/s/AKfycbyz-jowsvuW712EK6JFJu3OIR0PEAPWOv3cf-_RoMELmUtUzYTRNaY33qL62s4bBAo_cQ/exec"


@cache.memoize(timeout=600)
def get_advocate_sheet_data():
    """Fetch advocate data from Google Sheets. Returns list of dicts."""
    if not ADVOCATE_SHEET_URL or "script.google.com" not in ADVOCATE_SHEET_URL:
        return []
    try:
        response = requests.get(ADVOCATE_SHEET_URL, timeout=30)
        data = response.json()
        if isinstance(data, dict) and 'error' in data:
            logger.error(f"Advocate sheet error: {data['error']}")
            return []
        return data
    except Exception as e:
        logger.error(f"Advocate sheet fetch error: {e}")
        return []


# ---------------- COURSE GROUPING ----------------
# Map individual course variants to main course categories
MAIN_COURSE_KEYWORDS = [
    ("OET", ["oet"]),
    ("PTE", ["pte", "pearson"]),
    ("IELTS", ["ielts"]),
    ("German", ["german"]),
    ("Prometric", ["prometric"]),
    ("Nclex-RN", ["nclex"]),
    ("CBT", ["cbt"]),
    ("DM", ["dm ", "digital marketing"]),
    ("Fluency", ["fluency", "english fluency"]),
    ("Media", ["media"]),
    ("Yoga", ["yoga"]),
    ("Prometric-Oman", ["prometric-oman", "prometric oman"]),
]

def get_main_course(course_name):
    """Map individual course name to its main course category."""
    if not course_name:
        return "Other"
    lower = course_name.lower().strip()
    # Special case: Prometric-Oman before generic Prometric
    if "oman" in lower and "prometric" in lower:
        return "Prometric"
    for main_name, keywords in MAIN_COURSE_KEYWORDS:
        for kw in keywords:
            if kw in lower:
                return main_name
    return course_name  # Return original if no match

def group_course_breakdown(breakdown):
    """Aggregate a course_breakdown dict {name: {count, revenue}} by main course category."""
    grouped = {}
    for course_name, info in breakdown.items():
        main = get_main_course(course_name)
        if main not in grouped:
            grouped[main] = {"count": 0, "revenue": 0}
        grouped[main]["count"] += info["count"]
        grouped[main]["revenue"] += info["revenue"]
    return grouped

# ---------------- FLASK ROUTES ----------------
@app.route('/')
def index():
    offset = request.args.get('offset', 0, type=int)
    period = request.args.get('period', 'monthly')
    c_start1 = request.args.get('c_start1')
    c_end1 = request.args.get('c_end1')
    c_start2 = request.args.get('c_start2')
    c_end2 = request.args.get('c_end2')
    m1_start, m1_end, m2_start, m2_end, date_obj = get_comparison_dates(offset, period, c_start1, c_end1, c_start2, c_end2)
    
    if period == 'all':
        main_label = "ALL TIME"
        r1_label = f"{m1_start.strftime('%b %Y')} - {m1_end.strftime('%b %Y')}"
        r2_label = "N/A"
    elif period == 'quarterly':
        q_num = (m1_start.month - 1) // 3 + 1
        main_label = f"Q{q_num} {m1_start.year}"
        r1_label = f"{m1_start.strftime('%b %d')} - {m1_end.strftime('%b %d')}"
        r2_label = f"{m2_start.strftime('%b %d')} - {m2_end.strftime('%b %d')}"
    elif period == 'custom':
        main_label = "CUSTOM RANGE"
        r1_label = f"{m1_start.strftime('%b %d, %Y')} - {m1_end.strftime('%b %d, %Y')}"
        r2_label = f"{m2_start.strftime('%b %d, %Y')} - {m2_end.strftime('%b %d, %Y')}"
    elif period == 'weekly':
        main_label = f"Week of {m1_start.strftime('%b %d')}"
        r1_label = f"{m1_start.strftime('%b %d')} - {m1_end.strftime('%b %d')}"
        r2_label = f"{m2_start.strftime('%b %d')} - {m2_end.strftime('%b %d')}"
    elif period == 'daily':
        main_label = m1_start.strftime('%B %d, %Y')
        r1_label = "Today" if offset == 0 else m1_start.strftime('%b %d')
        r2_label = m2_start.strftime('%b %d')
    else:
        main_label = m1_start.strftime('%B %Y')
        r1_label = f"{m1_start.strftime('%b %d')} - {m1_end.strftime('%b %d')}"
        r2_label = f"{m2_start.strftime('%b %d')} - {m2_end.strftime('%b %d')}"

    return render_template(
        'index.html',
        active_users='—',
        current_period=period,
        current_month=main_label,
        mtd_range=r1_label,
        prev_range=r2_label,
        month_offset=offset
    )


@app.route('/api/active-users')
@cache.cached(timeout=30)
def api_active_users():
    return jsonify({"count": get_active_users()})


@app.route('/api/discover')
@cache.cached(timeout=300, query_string=True)
def api_discover():
    offset = request.args.get('offset', 0, type=int)
    period = request.args.get('period', 'monthly')
    c_start1 = request.args.get('c_start1')
    c_end1 = request.args.get('c_end1')
    c_start2 = request.args.get('c_start2')
    c_end2 = request.args.get('c_end2')
    m1_start, m1_end, m2_start, m2_end, _ = get_comparison_dates(offset, period, c_start1, c_end1, c_start2, c_end2)
    m2_full_end = m1_start - timedelta(days=1)  # Full previous month for worm graph

    try:
        with concurrent.futures.ThreadPoolExecutor() as executor:
            f_m1 = executor.submit(get_discover_metrics, m1_start, m1_end)
            f_m2 = executor.submit(get_discover_metrics, m2_start, m2_end)
            f_trend = executor.submit(get_daily_new_users, m1_start, m1_end)
            f_trend_prev = executor.submit(get_daily_new_users, m2_start, m2_full_end)

        m1_val = f_m1.result()
        m2_val = f_m2.result()
        trend = f_trend.result()
        trend_prev = f_trend_prev.result()
        sources = get_traffic_sources(m1_start, m1_end)

        delta = m1_val - m2_val
        pct = (delta / m2_val * 100) if m2_val > 0 else 0

        return jsonify({
            "m1_val": f"{m1_val:,}",
            "m2_val": f"{m2_val:,}",
            "delta_pct": pct,
            "trend": trend,
            "worm_m1": calculate_cumulative(trend, "New Users"),
            "worm_m2": calculate_cumulative(trend_prev, "New Users"),
            "sources": sources.to_dict('records')
        })
    except Exception as e:
        logger.error(f"API DISCOVER ERROR: {e}")
        return jsonify({"error": "Failed to fetch discover metrics", "details": str(e)}), 500


@app.route('/api/try')
@cache.cached(timeout=300, query_string=True)
def api_try():
    """Return filtered lead and deal counts for the LEAD phase."""
    offset = request.args.get('offset', 0, type=int)
    period = request.args.get('period', 'monthly')
    c_start1 = request.args.get('c_start1')
    c_end1 = request.args.get('c_end1')
    c_start2 = request.args.get('c_start2')
    c_end2 = request.args.get('c_end2')
    course = request.args.get('course')

    m1_start, m1_end, m2_start, m2_end, _ = get_comparison_dates(offset, period, c_start1, c_end1, c_start2, c_end2)
    m2_full_end = m1_start - timedelta(days=1)

    try:
        # Reuse the previous full-month contact snapshot for both previous MTD and full-month trend.
        with concurrent.futures.ThreadPoolExecutor(max_workers=4) as executor:
            f_m1_contacts = executor.submit(get_hubspot_lead_contacts, m1_start, m1_end)
            f_m2_full_contacts = executor.submit(get_hubspot_lead_contacts, m2_start, m2_full_end)
            f_m1_customers = executor.submit(get_hubspot_customer_count, m1_start, m1_end, course)
            f_m2_customers = executor.submit(get_hubspot_customer_count, m2_start, m2_end, course)

            m1_contacts = f_m1_contacts.result()
            m2_full_contacts = f_m2_full_contacts.result()
            m1_customer_count = f_m1_customers.result()
            m2_customer_count = f_m2_customers.result()

        r1 = _build_hubspot_lead_snapshot(m1_contacts, m1_start, m1_end, course_filter=course)
        r2 = _build_hubspot_lead_snapshot(m2_full_contacts, m2_start, m2_end, course_filter=course)
        r2_f = _build_hubspot_lead_snapshot(
            m2_full_contacts,
            m2_start,
            m2_full_end,
            course_filter=course,
            include_status_metrics=False
        )

        m1_leads = r1["total_leads"]
        m1_deals = r1["deal_status_breakdown"].get("Hot", 0) + r1["deal_status_breakdown"].get("Warm", 0) + r1["deal_status_breakdown"].get("Cold", 0) + m1_customer_count
        m1_trend = r1["trend"]

        m2_leads = r2["total_leads"]
        m2_deals = r2["deal_status_breakdown"].get("Hot", 0) + r2["deal_status_breakdown"].get("Warm", 0) + r2["deal_status_breakdown"].get("Cold", 0) + m2_customer_count

        m2_trend = r2_f["trend"]

        m1_course_breakdown = r1.get("course_breakdown", [])
        m2_course_breakdown = r2.get("course_breakdown", [])
        m1_deal_split = r1.get("deal_status_breakdown", {})
        m2_deal_split = r2.get("deal_status_breakdown", {})
    except Exception as e:
        logger.error(f"API TRY ERROR: {e}")
        return jsonify({"error": "Failed to fetch HubSpot leads", "details": str(e)}), 500

    lead_delta = m1_leads - m2_leads
    lead_pct = (lead_delta / m2_leads * 100) if m2_leads > 0 else 0
    
    deal_delta = m1_deals - m2_deals
    deal_pct = (deal_delta / m2_deals * 100) if m2_deals > 0 else 0

    return jsonify({
        "m1_leads": m1_leads,
        "m2_leads": m2_leads,
        "lead_delta": f"{lead_delta:+}",
        "lead_pct": round(lead_pct, 1),
        "m1_deals": m1_deals,
        "m2_deals": m2_deals,
        "deal_delta": f"{deal_delta:+}",
        "deal_pct": round(deal_pct, 1),
        "m1_course_breakdown": m1_course_breakdown,
        "m2_course_breakdown": m2_course_breakdown,
        "m1_deal_split": m1_deal_split,
        "m2_deal_split": m2_deal_split,
        "m1_customer_count": m1_customer_count,
        "m2_customer_count": m2_customer_count,
        "worm_m1": calculate_cumulative(m1_trend, "Admissions"),
        "worm_m2": calculate_cumulative(m2_trend, "Admissions"),
        "trend": m1_trend
    })


@app.route('/api/buy')
def api_buy():
    offset = request.args.get('offset', 0, type=int)
    period = request.args.get('period', 'monthly')
    c_start1 = request.args.get('c_start1')
    c_end1 = request.args.get('c_end1')
    c_start2 = request.args.get('c_start2')
    c_end2 = request.args.get('c_end2')
    course = request.args.get('course')
    m1_start, m1_end, m2_start, m2_end, _ = get_comparison_dates(offset, period, c_start1, c_end1, c_start2, c_end2)
    m2_full_end = m1_start - timedelta(days=1)  # Full previous month end for worm graph

    # Fetch HubSpot data concurrently for speed
    try:
        with concurrent.futures.ThreadPoolExecutor() as executor:
            f_m1 = executor.submit(get_hubspot_deals, m1_start, m1_end, course_filter=course)
            f_m2 = executor.submit(get_hubspot_deals, m2_start, m2_end, course_filter=course)
            f_m2_f = executor.submit(get_hubspot_deals, m2_start, m2_full_end, course_filter=course)
            
            m1_count, m1_val, m1_trend, m1_course_breakdown, m1_course_breakdown_ungrouped, m1_owner_breakdown = f_m1.result()
            m2_count, m2_val, _, m2_course_breakdown, m2_course_breakdown_ungrouped, m2_owner_breakdown = f_m2.result()
            _, _, m2_trend, _, _, _ = f_m2_f.result()
    except Exception as e:
        logger.error(f"API BUY ERROR: {e}")
        return jsonify({"error": "Failed to fetch HubSpot data", "details": str(e)}), 500

    delta_count = m1_count - m2_count
    delta_val = m1_val - m2_val
    pct_val = (delta_val / m2_val * 100) if m2_val > 0 else 0

    # Merge this month + previous month course breakdown (already grouped at source)
    all_courses = set(list(m1_course_breakdown.keys()) + list(m2_course_breakdown.keys()))
    course_list = []
    for course in all_courses:
        m1_info = m1_course_breakdown.get(course, {"count": 0, "revenue": 0})
        m2_info = m2_course_breakdown.get(course, {"count": 0, "revenue": 0})
        course_list.append({
            "course": course,
            "m1_count": m1_info["count"],
            "m1_revenue": m1_info["revenue"],
            "m2_count": m2_info["count"],
            "m2_revenue": m2_info["revenue"]
        })
    course_list.sort(key=lambda x: x["m1_count"], reverse=True)

    # Merge this month + previous month course breakdown ungrouped
    all_courses_ungrouped = set(list(m1_course_breakdown_ungrouped.keys()) + list(m2_course_breakdown_ungrouped.keys()))
    course_list_ungrouped = []
    for course in all_courses_ungrouped:
        m1_info = m1_course_breakdown_ungrouped.get(course, {"count": 0, "revenue": 0})
        m2_info = m2_course_breakdown_ungrouped.get(course, {"count": 0, "revenue": 0})
        course_list_ungrouped.append({
            "course": course,
            "m1_count": m1_info["count"],
            "m1_revenue": m1_info["revenue"],
            "m2_count": m2_info["count"],
            "m2_revenue": m2_info["revenue"]
        })
    course_list_ungrouped.sort(key=lambda x: x["m1_count"], reverse=True)

    # Merge this month + previous month owner breakdown
    all_owners = set(list(m1_owner_breakdown.keys()) + list(m2_owner_breakdown.keys()))
    owner_list = []
    for owner in all_owners:
        m1_info = m1_owner_breakdown.get(owner, {"count": 0, "revenue": 0})
        m2_info = m2_owner_breakdown.get(owner, {"count": 0, "revenue": 0})
        owner_list.append({
            "owner": owner,
            "m1_count": m1_info["count"],
            "m1_revenue": m1_info["revenue"],
            "m2_count": m2_info["count"],
            "m2_revenue": m2_info["revenue"]
        })
    owner_list.sort(key=lambda x: x["m1_count"], reverse=True)

    return jsonify({
        "m1_count": m1_count,
        "m1_val": f"₹{m1_val:,.0f}",
        "m2_count": m2_count,
        "m2_val": f"₹{m2_val:,.0f}",
        "delta_count": f"{delta_count:+}",
        "pct_val": pct_val,
        "worm_m1": calculate_cumulative(m1_trend, "Count"),
        "worm_m2": calculate_cumulative(m2_trend, "Count"),
        "course_breakdown": course_list,
        "course_breakdown_ungrouped": course_list_ungrouped,
        "owner_breakdown": owner_list
    })


@app.route('/api/use')
def api_use():
    offset = request.args.get('offset', 0, type=int)
    period = request.args.get('period', 'monthly')
    c_start1 = request.args.get('c_start1')
    c_end1 = request.args.get('c_end1')
    c_start2 = request.args.get('c_start2')
    c_end2 = request.args.get('c_end2')
    course = request.args.get('course')
    m1_start, m1_end, m2_start, m2_end, _ = get_comparison_dates(offset, period, c_start1, c_end1, c_start2, c_end2)
    m2_full_end = m1_start - timedelta(days=1)  # Full previous month end for worm graph

    with concurrent.futures.ThreadPoolExecutor() as executor:
        f_m1_cust = executor.submit(get_kajabi_new_customers, m1_start, m1_end)
        f_m2_cust = executor.submit(get_kajabi_new_customers, m2_start, m2_end)          # MTD count for KPI
        f_m2_cust_full = executor.submit(get_kajabi_new_customers, m2_start, m2_full_end) # Full month for graph
        f_m1_sales = executor.submit(get_kajabi_sales, m1_start, m1_end)
        f_m2_sales = executor.submit(get_kajabi_sales, m2_start, m2_end)
        f_prod = executor.submit(get_kajabi_products)
        f_active = executor.submit(get_kajabi_active_customers, m1_start)
        f_offers = executor.submit(get_kajabi_offers)

    m1_count, m1_list, total, m1_trend = f_m1_cust.result()
    m2_count, _, _, _ = f_m2_cust.result()          # KPI: MTD count only
    _, _, _, m2_trend = f_m2_cust_full.result()      # Graph: full previous month trend
    m1_rev, _ = f_m1_sales.result()
    m2_rev, _ = f_m2_sales.result()
    products = f_prod.result()
    active, _ = f_active.result()
    offers_map = f_offers.result()

    # Build MTD course enrollment counts with FILTERING
    offer_counts = {}
    filtered_m1_list = []
    for c in m1_list:
        match = False
        row_offers = []
        for o_item in c.get('relationships', {}).get('offers', {}).get('data', []):
            oid = o_item.get('id')
            if oid:
                name = offers_map.get(oid, f"Offer {oid}")
                main = get_main_course(name)
                if course and main.lower() != course.lower():
                    continue
                match = True
                offer_counts[name] = offer_counts.get(name, 0) + 1
        if not course or match:
            filtered_m1_list.append(c)

    mtd_courses = sorted(
        [{"Course": k, "Enrollments": v} for k, v in offer_counts.items()],
        key=lambda x: x['Enrollments'],
        reverse=True
    )[:10]

    if course:
        # Re-calculate count based on filter
        m1_count = len(filtered_m1_list)
        # Revenue filtering is harder without sale-level course info,
        # but we can estimate or leave as is if not available.
        pass

    delta_count = m1_count - m2_count
    delta_rev = m1_rev - m2_rev

    return jsonify({
        "total_customers": f"{total:,}",
        "m1_new": m1_count,
        "new_delta": f"{delta_count:+}",
        "m2_new": m2_count,
        "m1_rev": f"₹{m1_rev:,.2f}",
        "rev_delta": f"₹{delta_rev:,.2f}",
        "active_learners": f"{active:,}",
        "products": sorted(products, key=lambda x: x['Members'], reverse=True)[:10],
        "mtd_courses": mtd_courses,
        "worm_m1": calculate_cumulative(m1_trend, "Count"),
        "worm_m2": calculate_cumulative(m2_trend, "Count")
    })


@app.route('/api/renew')
def api_renew():
    offset = request.args.get('offset', 0, type=int)
    period = request.args.get('period', 'monthly')
    c_start1 = request.args.get('c_start1')
    c_end1 = request.args.get('c_end1')
    c_start2 = request.args.get('c_start2')
    c_end2 = request.args.get('c_end2')
    course = request.args.get('course')
    m1_start, m1_end, m2_start, m2_end, _ = get_comparison_dates(offset, period, c_start1, c_end1, c_start2, c_end2)

    with concurrent.futures.ThreadPoolExecutor() as executor:
        f_m1 = executor.submit(get_renew_sheet_data, m1_start, m1_end, course_filter=course)
        f_m2 = executor.submit(get_renew_sheet_data, m2_start, m2_end, course_filter=course)

    m1_count, m1_val, m1_df = f_m1.result()
    m2_count, m2_val, m2_df = f_m2.result()

    delta_count = m1_count - m2_count
    delta_val = m1_val - m2_val
    pct_val = (delta_val / m2_val * 100) if m2_val > 0 else 0

    recent = []

    def get_sanitized_breakdown(df, col):
        if df.empty or col not in df.columns:
            return []
        s = df[col].astype(str).str.strip().str.upper()
        counts = s.value_counts().to_dict()
        items = [{"name": str(k), "count": int(v)} for k, v in counts.items()]
        items.sort(key=lambda x: x["count"], reverse=True)
        return items

    m1_courses = get_sanitized_breakdown(m1_df, "Course")
    m2_courses = get_sanitized_breakdown(m2_df, "Course")
    m1_packages = get_sanitized_breakdown(m1_df, "Package")
    m2_packages = get_sanitized_breakdown(m2_df, "Package")

    if not m1_df.empty:
        # Columns: Student Name, Course, Package, Lead Owner, Fee Amount, Paid Amount, Pending Amount, Payment Paid Date/ Initial Amount Paid Date, Pending Paid Date
        cols_to_show = ["Student Name", "Course", "Package", "Lead Owner", "Fee Amount", "Paid Amount", "Pending Amount", "Payment Paid Date/ Initial Amount Paid Date", "Pending Paid Date"]
        # Filter to only existing columns to avoid errors
        available_cols = [c for c in cols_to_show if c in m1_df.columns]
        recent = m1_df[available_cols].to_dict('records')

    return jsonify({
        "m1_count": m1_count,
        "m1_val": f"₹{m1_val:,.0f}",
        "m2_count": m2_count,
        "m2_val": f"₹{m2_val:,.0f}",
        "delta_count": f"{delta_count:+}",
        "pct_val": pct_val,
        "recent_renewals": recent,
        "m1_course_breakdown": m1_courses,
        "m2_course_breakdown": m2_courses,
        "m1_package_breakdown": m1_packages,
        "m2_package_breakdown": m2_packages,
        "v": "2.3"
    })


def _parse_advocate_passout(row):
    """Parse PASS OUT MONTH + PASS OUT DATE into a (year, month_num, date_obj, month_label) tuple."""
    import re, calendar
    month_str = (row.get("PASS OUT MONTH") or "").strip().upper()
    date_str = (row.get("PASS OUT DATE") or "").strip()

    # Map month name → number
    month_map = {m.upper(): i for i, m in enumerate(calendar.month_name) if m}
    # Also handle abbreviations
    month_map.update({m.upper(): i for i, m in enumerate(calendar.month_abbr) if m})

    month_num = None
    year = None

    # Try to extract month from PASS OUT MONTH (e.g. "February", "JANUARY", "January 2026", "Feb 2026")
    if month_str:
        parts = month_str.split()
        for p in parts:
            if p in month_map:
                month_num = month_map[p]
            elif p.isdigit() and len(p) == 4:
                year = int(p)

    # Try to parse date for a more precise date and to extract year if missing
    date_obj = None
    if date_str:
        for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%d-%m-%Y", "%b %d", "%d %b", "%d %b %Y"):
            try:
                date_obj = datetime.strptime(date_str, fmt)
                if date_obj.year < 2000:
                    # Format didn't include year; we'll fill it from context
                    date_obj = None
                else:
                    if not year:
                        year = date_obj.year
                    if not month_num:
                        month_num = date_obj.month
                break
            except ValueError:
                continue
        # Also try raw date patterns like "Feb 7", "14/02/2026", "Jan 18"
        if not date_obj and date_str:
            match = re.search(r'(\d{1,2})[/\-](\d{1,2})[/\-](\d{2,4})', date_str)
            if match:
                d, m, y = int(match.group(1)), int(match.group(2)), int(match.group(3))
                if y < 100: y += 2000
                try:
                    date_obj = datetime(y, m, d)
                    if not year: year = y
                    if not month_num: month_num = m
                except ValueError:
                    try:
                        date_obj = datetime(y, d, m)  # Try swapped day/month
                        if not year: year = y
                        if not month_num: month_num = m
                    except ValueError:
                        pass

    # Default year to current year if still missing
    if not year:
        year = datetime.now().year

    # Build month label
    if month_num:
        month_label = f"{calendar.month_name[month_num].upper()} {year}"
    else:
        month_label = "UNKNOWN"

    return year, month_num, date_obj, month_label


@app.route('/api/advocate')
@cache.cached(timeout=300, query_string=True)
def api_advocate():
    import calendar
    offset = request.args.get('offset', 0, type=int)
    period = request.args.get('period', 'monthly')
    c_start1 = request.args.get('c_start1')
    c_end1 = request.args.get('c_end1')
    c_start2 = request.args.get('c_start2')
    c_end2 = request.args.get('c_end2')
    course = request.args.get('course')
    m1_start, m1_end, m2_start, m2_end, _ = get_comparison_dates(offset, period, c_start1, c_end1, c_start2, c_end2)

    raw_data = get_advocate_sheet_data()
    if not raw_data:
        return jsonify({"m1_total": 0, "m2_total": 0, "delta": "+0", "breakdown": [], "recent": []})

    # FILTERING BY COURSE
    if course:
        raw_data = [r for r in raw_data if get_main_course(str(r.get("COURSE", ""))).lower() == course.lower()]

    # Enrich every row with parsed pass-out info
    enriched = []
    for r in raw_data:
        yr, mn, dt, label = _parse_advocate_passout(r)
        r['_passout_year'] = yr
        r['_passout_month_num'] = mn
        r['_passout_date_obj'] = dt.isoformat() if dt else None
        r['_passout_label'] = label
        r['_sort_key'] = dt.isoformat() if dt else f"{yr}-{mn or 0:02d}-00"
        enriched.append(r)

    target_year = m1_start.year
    target_month = m1_start.month
    prev_year = m2_start.year
    prev_month = m2_start.month

    if period == 'all':
        rows_this_month = enriched
        rows_prev_month = []
    else:
        m1_s = m1_start.strftime('%Y-%m-%d')
        m1_e = m1_end.strftime('%Y-%m-%d')
        m2_s = m2_start.strftime('%Y-%m-%d')
        m2_e = m2_end.strftime('%Y-%m-%d')
        
        # We need to elegantly handle rows that have NO exact date, only a Month/Year
        # BUT since user specifically asked for "this date MTD", we will stick to exact date matching.
        rows_this_month = [r for r in enriched if r['_passout_date_obj'] and m1_s <= r['_passout_date_obj'][:10] <= m1_e]
        rows_prev_month = [r for r in enriched if r['_passout_date_obj'] and m2_s <= r['_passout_date_obj'][:10] <= m2_e]

    # --- KPIs for this month ---
    total = len(rows_this_month)
    prev_total = len(rows_prev_month)
    branch_counts = {}
    advocacy_levels = {}
    testimonials_given = 0
    testimonials_not = 0
    course_counts = {}

    for r in rows_this_month:
        branch = r.get("Branch", "Unknown")
        branch_counts[branch] = branch_counts.get(branch, 0) + 1

        level = (r.get("ADVOCACY LEVEL") or "").strip()
        if level:
            advocacy_levels[level] = advocacy_levels.get(level, 0) + 1

        testimonial = (r.get("TESTIMONIALS (GIVEN/NOT)") or "").strip().upper()
        if "GIVEN" in testimonial and "NOT" not in testimonial:
            testimonials_given += 1
        elif testimonial:
            testimonials_not += 1

        course = (r.get("COURSE") or "").strip()
        if course:
            course_counts[course] = course_counts.get(course, 0) + 1

    # --- KPIs for previous month ---
    prev_branch_counts = {}
    prev_advocacy_levels = {}
    prev_testimonials_given = 0
    prev_testimonials_not = 0
    prev_course_counts = {}

    for r in rows_prev_month:
        branch = r.get("Branch", "Unknown")
        prev_branch_counts[branch] = prev_branch_counts.get(branch, 0) + 1

        level = (r.get("ADVOCACY LEVEL") or "").strip()
        if level:
            prev_advocacy_levels[level] = prev_advocacy_levels.get(level, 0) + 1

        testimonial = (r.get("TESTIMONIALS (GIVEN/NOT)") or "").strip().upper()
        if "GIVEN" in testimonial and "NOT" not in testimonial:
            prev_testimonials_given += 1
        elif testimonial:
            prev_testimonials_not += 1

        course = (r.get("COURSE") or "").strip()
        if course:
            prev_course_counts[course] = prev_course_counts.get(course, 0) + 1

    sorted_levels = sorted(advocacy_levels.items(), key=lambda x: x[1], reverse=True)
    sorted_courses = sorted(course_counts.items(), key=lambda x: x[1], reverse=True)

    # Sort rows by date descending within this month
    rows_this_month.sort(key=lambda r: r['_sort_key'], reverse=True)

    # Build month_groups for ALL months (for the full grouped view)
    month_groups_map = {}
    for r in enriched:
        label = r['_passout_label']
        if label not in month_groups_map:
            month_groups_map[label] = {"month": label, "year": r['_passout_year'], "month_num": r['_passout_month_num'] or 0, "rows": []}
        # Strip internal fields for clean JSON
        clean = {k: v for k, v in r.items() if not k.startswith('_')}
        clean['_sort_key'] = r['_sort_key']
        clean['_passout_date_obj'] = r['_passout_date_obj']
        month_groups_map[label]["rows"].append(clean)

    # Sort groups by year desc, month desc
    month_groups = sorted(month_groups_map.values(), key=lambda g: (g['year'], g['month_num']), reverse=True)
    for g in month_groups:
        g['rows'].sort(key=lambda r: r.get('_sort_key', ''), reverse=True)
        g['count'] = len(g['rows'])

    # Clean rows for this month too
    clean_rows = []
    for r in rows_this_month:
        clean = {k: v for k, v in r.items() if not k.startswith('_')}
        clean['_sort_key'] = r['_sort_key']
        clean['_passout_date_obj'] = r['_passout_date_obj']
        clean_rows.append(clean)

    return jsonify({
        "total": total,
        "prev_total": prev_total,
        "branch_counts": [{"name": k, "count": v} for k, v in sorted(branch_counts.items())],
        "prev_branch_counts": [{"name": k, "count": v} for k, v in sorted(prev_branch_counts.items())],
        "advocacy_levels": [{"level": k, "count": v} for k, v in sorted_levels],
        "testimonials_given": testimonials_given,
        "testimonials_not": testimonials_not,
        "prev_testimonials_given": prev_testimonials_given,
        "prev_testimonials_not": prev_testimonials_not,
        "course_breakdown": [{"name": k, "count": v} for k, v in sorted_courses],
        "rows": clean_rows,
        "all_rows": [r2 for g in month_groups for r2 in g['rows']],
        "month_groups": month_groups,
        "target_month_label": f"{calendar.month_name[target_month].upper()} {target_year}",
        "prev_month_label": f"{calendar.month_name[prev_month].upper()} {prev_year}",
    })


@app.route('/api/cache-clear', methods=['POST'])
def api_cache_clear():
    cache.clear()
    logger.info("Cache cleared via refresh button")
    return jsonify({"status": "ok"})


# ---------------- CONTENT CALENDAR ----------------
@cache.memoize(timeout=1800)
def get_content_calendar():
    """Fetch all content calendar rows from Google Apps Script."""
    try:
        resp = requests.get(CONTENT_CALENDAR_URL, timeout=30)
        resp.raise_for_status()
        return resp.json()
    except Exception as e:
        logger.error(f"Content calendar fetch error: {e}")
        return []




@app.route('/api/content-calendar')
@cache.cached(timeout=300, query_string=True)
def api_content_calendar():
    """Return content calendar KPIs filtered to the requested month (full month range)."""
    offset = request.args.get('offset', 0, type=int)
    period = request.args.get('period', 'monthly')
    c_start1 = request.args.get('c_start1')
    c_end1 = request.args.get('c_end1')
    c_start2 = request.args.get('c_start2')
    c_end2 = request.args.get('c_end2')
    course = request.args.get('course')
    ist = pytz.timezone("Asia/Kolkata")
    m1_start, m1_end, m2_start, m2_end, _ = get_comparison_dates(offset, period, c_start1, c_end1, c_start2, c_end2)

    if period == 'monthly' and not c_start1:
        if m1_end.month == 12:
            m1_end = m1_end.replace(month=1, year=m1_end.year+1, day=1) - timedelta(days=1)
        else:
            m1_end = (m1_end.replace(day=1) + timedelta(days=32)).replace(day=1) - timedelta(days=1)

    all_rows = get_content_calendar()

    # FILTERING BY COURSE
    if course:
        all_rows = [r for r in all_rows if get_main_course(str(r.get("Sheet", ""))).lower() == course.lower()]

    STATUS_MAP = {
        "published": "Published",
        "pending":   "Pending",
        "assigned":  "Assigned",
        "":          "Unset",
    }
    FUNNEL_ORDER = ["Awareness", "Consideration", "Conversion", "Retention"]

    rows_this_month = []
    for row in all_rows:
        if not any([row.get("Content Topic"), row.get("Status"), row.get("Scheduled Date"), row.get("Owner/TUTOR")]):
            continue

        date_val = row.get("Published Date") or row.get("Scheduled Date") or ""
        parsed_date = None
        if date_val:
            try:
                date_val_clean = date_val.split('(')[0].strip()
                import re as _re
                date_val_clean = _re.sub(r',?\s*(MON|TUE|WED|THRS|THU|FRI|SAT|SUN)\s*$', '', date_val_clean, flags=_re.IGNORECASE).strip()
                parsed_date = pd.to_datetime(date_val_clean, errors='coerce', dayfirst=False)
                if pd.isna(parsed_date):
                    parsed_date = None
                else:
                    parsed_date = parsed_date.date()
            except Exception:
                parsed_date = None

        if parsed_date and not (m1_start <= parsed_date <= m1_end):
            continue
        
        if not parsed_date:
            continue

        status_raw = (row.get("Status") or "").strip().lower()
        status = STATUS_MAP.get(status_raw, row.get("Status", "Other").strip())

        # [FIX] If status column is empty, but there is a Published Date, infer that it is Published
        if (status == "Unset" or status == "Other") and row.get("Published Date") and str(row.get("Published Date")).strip():
            status = "Published"

        funnel_raw = (row.get("Funnel Stage") or "").strip().lower()
        if "awareness" in funnel_raw:
            funnel = "Awareness"
        elif "consideration" in funnel_raw:
            funnel = "Consideration"
        elif "conversion" in funnel_raw:
            funnel = "Conversion"
        elif "retention" in funnel_raw:
            funnel = "Retention"
        else:
            funnel = "Other"

        rows_this_month.append({
            "sheet":       row.get("Sheet", "").strip(),
            "topic":       row.get("Content Topic", ""),
            "type":        row.get("Content Type", ""),
            "owner":       row.get("Owner/TUTOR", ""),
            "status":      status,
            "funnel":      funnel,
            "date":        parsed_date.isoformat() if parsed_date else "",
            "link_yt":     row.get("Link YT", ""),
            "link_insta":  row.get("Link INSTA", ""),
            "link_fb":     row.get("Link FB", ""),
            "remarks":     row.get("REMARKS", ""),
        })

    total      = len(rows_this_month)
    published  = sum(1 for r in rows_this_month if r["status"] == "Published")
    pending    = sum(1 for r in rows_this_month if r["status"] == "Pending")
    assigned   = sum(1 for r in rows_this_month if r["status"] == "Assigned")
    publish_rate = round(published / total * 100) if total > 0 else 0

    course_summary = {}
    for r in rows_this_month:
        s = r["sheet"]
        if s not in course_summary:
            course_summary[s] = {"total": 0, "published": 0, "pending": 0}
        course_summary[s]["total"] += 1
        if r["status"] == "Published":
            course_summary[s]["published"] += 1
        elif r["status"] == "Pending":
            course_summary[s]["pending"] += 1

    funnel_summary = {}
    for r in rows_this_month:
        f = r["funnel"]
        funnel_summary[f] = funnel_summary.get(f, 0) + 1

    return jsonify({
        "total":        total,
        "published":    published,
        "pending":      pending,
        "assigned":     assigned,
        "publish_rate": publish_rate,
        "courses":      [{"name": k, **v} for k, v in sorted(course_summary.items())],
        "funnel":       [{"stage": f, "count": funnel_summary.get(f, 0)} for f in FUNNEL_ORDER if f in funnel_summary],
        "rows":         rows_this_month,
    })


import threading

def _prewarm_lead_dashboard(m1_start, m1_end, m2_start, m2_end, m2_full_end):
    """Warm the heavy LEAD caches for the default monthly dashboard view."""
    try:
        with concurrent.futures.ThreadPoolExecutor(max_workers=4) as ex:
            f_m1_contacts = ex.submit(get_hubspot_lead_contacts, m1_start, m1_end)
            f_m2_full_contacts = ex.submit(get_hubspot_lead_contacts, m2_start, m2_full_end)
            f_m1_customers = ex.submit(get_hubspot_customer_count, m1_start, m1_end, None)
            f_m2_customers = ex.submit(get_hubspot_customer_count, m2_start, m2_end, None)

            m1_contacts = f_m1_contacts.result()
            m2_full_contacts = f_m2_full_contacts.result()
            f_m1_customers.result()
            f_m2_customers.result()

        _build_hubspot_lead_snapshot(m1_contacts, m1_start, m1_end)
        _build_hubspot_lead_snapshot(m2_full_contacts, m2_start, m2_end)
        _build_hubspot_lead_snapshot(
            m2_full_contacts,
            m2_start,
            m2_full_end,
            include_status_metrics=False
        )
        logger.info("WARMUP: lead dashboard cache ready")
    except Exception as e:
        logger.warning(f"WARMUP lead error (non-fatal): {e}")

def _warmup():
    """Pre-warm all data caches in background so first browser hit is instant."""
    import time
    time.sleep(3)  # Wait for Flask to fully start
    try:
        m1_start, m1_end, m2_start, m2_end, _ = get_comparison_dates(0)
        m2_full_end = m1_start - timedelta(days=1)
        logger.info("WARMUP: pre-fetching all data sources...")
        with concurrent.futures.ThreadPoolExecutor(max_workers=10) as ex:
            # DISCOVER — GA4 (Can be parallel)
            ex.submit(get_discover_metrics, m1_start, m1_end)
            ex.submit(get_discover_metrics, m2_start, m2_end)
            ex.submit(get_daily_new_users, m1_start, m1_end)
            ex.submit(get_daily_new_users, m2_start, m2_full_end)
            ex.submit(get_traffic_sources, m1_start, m1_end)
            
            # USE — Kajabi (Can be parallel)
            ex.submit(get_kajabi_new_customers, m1_start, m1_end)
            ex.submit(get_kajabi_new_customers, m2_start, m2_end)
            ex.submit(get_kajabi_new_customers, m2_start, m2_full_end)
            ex.submit(get_kajabi_sales, m1_start, m1_end)
            ex.submit(get_kajabi_sales, m2_start, m2_end)
            ex.submit(get_kajabi_products)
            ex.submit(get_kajabi_active_customers, m1_start)
            ex.submit(get_kajabi_offers)
            
            # RENEW & Content Calendar (Can be parallel)
            ex.submit(get_renew_sheet_data, m1_start, m1_end)
            ex.submit(get_renew_sheet_data, m2_start, m2_end)
            ex.submit(get_content_calendar)
            ex.submit(_prewarm_lead_dashboard, m1_start, m1_end, m2_start, m2_end, m2_full_end)

        logger.info("WARMUP: base caches ready (LEAD prewarm continues in background if needed)")
    except Exception as e:
        logger.warning(f"WARMUP error (non-fatal): {e}")


if __name__ == '__main__':
    threading.Thread(target=_warmup, daemon=True).start()
    port = int(os.environ.get('PORT', 5001))
    app.run(debug=False, host='0.0.0.0', port=port, use_reloader=False)
