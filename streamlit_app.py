import streamlit as st
import pandas as pd
import requests
import plotly.express as px

# ---------------- CONFIG ----------------
st.set_page_config(
    page_title="Leads Analytics Dashboard",
    layout="wide",
    page_icon="📊"
)

API_URL = "https://script.google.com/macros/s/AKfycbzIarFwdhWeSSZ8_bRDFjSW3KKkd_Dq-utjOsft7Q4waNdCaa06CGd1_IfwHvI6kHcW/exec"

# ---------------- LOAD DATA ----------------
@st.cache_data(ttl=300)
def load_data():
    try:
        r = requests.get(API_URL, timeout=20)
    except Exception:
        st.error("Failed to connect to API.")
        st.stop()

    if r.status_code != 200:
        st.error(f"API Error: {r.status_code}")
        st.text(r.text[:500])
        st.stop()

    try:
        data = r.json()
    except Exception:
        st.error("API did not return JSON.")
        st.text(r.text[:500])
        st.stop()

    if not isinstance(data, list):
        st.error("API response is not a list.")
        st.stop()

    df = pd.DataFrame(data)

    if df.empty:
        st.warning("No data found.")
        st.stop()

    if "Date" in df.columns:
        df["Date"] = pd.to_datetime(df["Date"], errors="coerce")

    return df


df = load_data()

# ---------------- SIDEBAR FILTERS ----------------
st.sidebar.header("🔎 Filters")

owners = st.sidebar.multiselect(
    "Contact Owner",
    sorted(df["Contact owner"].dropna().unique())
)

courses = st.sidebar.multiselect(
    "Course",
    sorted(df["Course"].dropna().unique())
)

date_range = st.sidebar.date_input("Date Range", [])

filtered = df.copy()

if owners:
    filtered = filtered[filtered["Contact owner"].isin(owners)]

if courses:
    filtered = filtered[filtered["Course"].isin(courses)]

if len(date_range) == 2:
    start, end = date_range
    filtered = filtered[
        (filtered["Date"] >= pd.to_datetime(start)) &
        (filtered["Date"] <= pd.to_datetime(end))
    ]

# ---------------- MAIN DASHBOARD ----------------
st.title("📊 Educational Leads Dashboard")
st.caption("Live data from Google Sheets")

st.metric("Total Leads", len(filtered))

# ---- Table 1 ----
st.subheader("📘 Course-wise Lead Count")

course_count = (
    filtered.groupby("Course")["Record ID"]
    .count()
    .reset_index(name="Leads")
    .sort_values("Leads", ascending=False)
)

st.dataframe(course_count, use_container_width=True)

# ---- Visualization 1 ----
st.subheader("📊 Course-wise Leads (Bar Chart)")
fig1 = px.bar(course_count, x="Course", y="Leads", title="Leads per Course")
st.plotly_chart(fig1, use_container_width=True)

# ---- Table 2 ----
st.subheader("📗 Course vs Lead Status")

course_status = pd.pivot_table(
    filtered,
    index="Course",
    columns="Lead Status",
    values="Record ID",
    aggfunc="count",
    fill_value=0
)

# Add Total column (NEW)
course_status["Total"] = course_status.sum(axis=1)

st.dataframe(course_status, use_container_width=True)

# ---- Visualization 2 ----
st.subheader("📊 Course vs Lead Status (Stacked)")
course_status_chart = course_status.drop(columns=["Total"]).reset_index().melt(
    id_vars="Course",
    var_name="Lead Status",
    value_name="Count"
)

fig2 = px.bar(
    course_status_chart,
    x="Course",
    y="Count",
    color="Lead Status",
    title="Lead Status Distribution per Course",
    barmode="stack"
)

st.plotly_chart(fig2, use_container_width=True)

# ---- Table 3 ----
st.subheader("👤 Contact Owner-wise Leads")

owner_count = (
    filtered.groupby("Contact owner")["Record ID"]
    .count()
    .reset_index(name="Leads")
    .sort_values("Leads", ascending=False)
)

st.dataframe(owner_count, use_container_width=True)

# ---- Download ----
st.subheader("⬇ Download Filtered Data")

csv = filtered.to_csv(index=False).encode("utf-8")
st.download_button(
    "Download CSV",
    csv,
    "filtered_leads.csv",
    "text/csv"
)

st.caption("Built with ❤️ using Streamlit")
