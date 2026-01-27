import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import re
from io import BytesIO

# Configuration Constants
MONTHLY_BUDGET = 2_000_000
HEALTH_THRESHOLD_HEALTHY = 5
HEALTH_THRESHOLD_WARNING = 15
CACHE_TTL = 3600

# Global pandas configuration
pd.set_option('future.no_silent_downcasting', True)

# Page configuration
st.set_page_config(
    page_title="RSA Dashboard - Sompo Thailand",
    page_icon="🚗",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS styling for premium look and feel
st.markdown("""
<style>
    html { font-family: 'Inter', 'Segoe UI', -apple-system, sans-serif; }
    .main { padding: 1.5rem 2rem; background-color: #F0F2F6; }
    .stApp { background-color: #F0F2F6; }
    
    /* Sidebar Styling */
    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #1B2838 0%, #2A3F54 100%);
        padding-top: 2rem;
    }
    [data-testid="stSidebar"] label {
        color: #E0E6ED !important;
        font-size: 13px !important;
        margin-top: 1rem;
    }

    /* KPI Metric Cards */
    .metric-card {
        background: linear-gradient(135deg, #1B2838 0%, #2D4A5C 100%);
        padding: 24px 20px;
        border-radius: 12px;
        color: white;
        box-shadow: 0 4px 12px rgba(27, 40, 56, 0.15);
        margin-bottom: 16px;
        position: relative;
        overflow: hidden;
    }
    .metric-card::before {
        content: '';
        position: absolute;
        top: 0; left: 0; right: 0; height: 3px;
        background: linear-gradient(90deg, #4A90D9 0%, #6FB1FF 100%);
    }
    .metric-title { font-size: 12px; font-weight: 600; opacity: 0.85; margin-bottom: 12px; color: #D0DCE8; }
    .metric-value { font-size: 32px; font-weight: 700; color: #FFFFFF; font-variant-numeric: tabular-nums; }
    .positive { color: #27AE60; }
    .negative { color: #E74C3C; }

    /* Health Indicators */
    .health-indicator {
        padding: 24px 28px;
        border-radius: 12px;
        margin: 20px 0;
        border-left: 6px solid;
        background: white;
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.08);
    }
    .health-healthy { border-left-color: #27AE60; background: linear-gradient(135deg, #d4edda 0%, #e8f5e9 100%); color: #155724; }
    .health-warning { border-left-color: #F39C12; background: linear-gradient(135deg, #fff3cd 0%, #fff9e6 100%); color: #856404; }
    .health-critical { border-left-color: #E74C3C; background: linear-gradient(135deg, #f8d7da 0%, #ffe6e8 100%); color: #721c24; }
    
    .health-badge { display: inline-block; padding: 4px 10px; border-radius: 4px; font-size: 11px; font-weight: 700; color: white; }
    .badge-healthy { background-color: #27AE60; }
    .badge-warning { background-color: #F39C12; }
    .badge-critical { background-color: #E74C3C; }

    /* Service Table */
    .service-table-container { background: white; border-radius: 12px; padding: 24px; box-shadow: 0 2px 8px rgba(0,0,0,0.08); overflow-x: auto; }
    .service-table { width: 100%; border-collapse: separate; border-spacing: 0; font-size: 13px; }
    .service-table th { background: #1B2838; color: white; padding: 14px; text-align: left; }
    .service-table td { padding: 12px; border-bottom: 1px solid #E2E8F0; }
    .inline-bar { height: 24px; background: linear-gradient(90deg, #4A90D9 0%, #6FB1FF 100%); border-radius: 4px; transition: width 0.3s ease; }
    
    .section-header { color: #1B2838; font-size: 26px; font-weight: 700; margin-top: 48px; border-bottom: 3px solid #4A90D9; padding-bottom: 12px; }
</style>
""", unsafe_allow_html=True)

# Data processing logic [cite: 165, 166, 167]
@st.cache_data(ttl=CACHE_TTL)
def load_data():
    df = pd.read_excel('(Test) RSA Report.xlsx', header=None)
    header_found = False
    for idx in range(len(df)):
        if 'Policy No.' in df.iloc[idx].values:
            df.columns = df.iloc[idx]
            df = df.iloc[idx+1:].reset_index(drop=True)
            header_found = True
            break
    if not header_found:
        raise ValueError("Policy No. column not found.")

    if 'Fee\n (Baht)' in df.columns:
        df = df.rename(columns={'Fee\n (Baht)': 'Fee (Baht)'})

    # Robust Date Conversion [cite: 167]
    if df['วันที่'].dtype == 'object':
        df['วันที่'] = pd.to_datetime(df['วันที่'], format='%d/%m/%Y', errors='coerce')
    else:
        df['วันที่'] = pd.to_datetime(df['วันที่'], errors='coerce')

    df['Day'] = df['วันที่'].dt.day.astype('float').fillna(0).astype('int64')
    df['Month'] = df['วันที่'].dt.month.astype('float').fillna(0).astype('int64')
    df['Year'] = df['วันที่'].dt.year.astype('float').fillna(0).astype('int64')
    df['LOB'] = df['Policy No.'].str.extract(r'(A[CV]\d)', expand=False).fillna('Unre')
    df['Fee (Baht)'] = pd.to_numeric(df['Fee (Baht)'], errors='coerce').fillna(0)
    
    return df.dropna(how='all').reset_index(drop=True)

# Helper for HTML Table generation [cite: 169-178]
def generate_service_table_html(df, years_in_data):
    lob_order = ['AV1', 'AV5', 'AV3', 'AV9', 'AC3', 'Unre']
    service_types = sorted([str(s) for s in df['ประเภทการบริการ'].dropna().unique()])
    
    max_count = 0
    for lob in lob_order:
        lob_df = df[df['LOB'] == lob]
        for service in service_types:
            for year in years_in_data:
                count = len(lob_df[(lob_df['ประเภทการบริการ'] == service) & (lob_df['Year'] == year)])
                max_count = max(max_count, count)

    html = '<div class="service-table-container"><table class="service-table"><thead><tr><th>LOB</th><th>Service Type</th>'
    for year in years_in_data: html += f'<th>{year}</th>'
    html += '<th>Total</th></tr></thead><tbody>'

    for lob in lob_order:
        lob_df = df[df['LOB'] == lob]
        if lob_df.empty: continue
        for idx, service in enumerate(service_types):
            html += '<tr>'
            if idx == 0: html += f'<td rowspan="{len(service_types)}" style="background:#F7FAFC; font-weight:600;">{lob}</td>'
            html += f'<td>{service}</td>'
            row_total = 0
            for year in years_in_data:
                cnt = len(lob_df[(lob_df['ประเภทการบริการ'] == service) & (lob_df['Year'] == year)])
                row_total += cnt
                bar_w = (cnt / max_count * 100) if max_count > 0 else 0
                html += f'<td><div style="display:flex; align-items:center; gap:8px;"><span style="min-width:20px;">{cnt}</span>'
                html += f'<div style="flex:1; background:#E8EEF7; border-radius:4px;"><div class="inline-bar" style="width:{bar_w}%"></div></div></div></td>'
            html += f'<td style="font-weight:700;">{row_total}</td></tr>'
    html += '</tbody></table></div>'
    return html

# Execution Flow
try:
    df = load_data()
except Exception as e:
    st.error(f"Error: {e}")
    st.stop()

# Sidebar [cite: 181-185]
st.sidebar.markdown("## 📊 Dashboard Filters")
selected_years = st.sidebar.multiselect("Select Year(s)", options=sorted(df['Year'].unique()), default=sorted(df['Year'].unique()))
if not selected_years:
    st.warning("Please select at least one year.")
    st.stop()

filtered_df = df[df['Year'].isin(selected_years)].copy()

# Dashboard Content [cite: 187-191]
st.markdown("# 🚗 RSA Dashboard - Sompo Thailand")
cols = st.columns(4)
total_fee = filtered_df['Fee (Baht)'].sum()
cols[0].markdown(f'<div class="metric-card"><div class="metric-title">Total Cases</div><div class="metric-value">{len(filtered_df):,}</div></div>', unsafe_allow_html=True)
cols[1].markdown(f'<div class="metric-card"><div class="metric-title">Total Fee</div><div class="metric-value">฿{total_fee:,.0f}</div></div>', unsafe_allow_html=True)
cols[2].markdown(f'<div class="metric-card"><div class="metric-title">Monthly Budget</div><div class="metric-value">฿{MONTHLY_BUDGET:,}</div></div>', unsafe_allow_html=True)

# Portfolio Health [cite: 71]
monthly_avg = total_fee / (filtered_df['Month'].nunique() or 1)
variance = (monthly_avg - MONTHLY_BUDGET) / MONTHLY_BUDGET * 100
if abs(variance) <= HEALTH_THRESHOLD_HEALTHY:
    status, h_cls, b_cls = "HEALTHY", "health-healthy", "badge-healthy"
elif abs(variance) <= HEALTH_THRESHOLD_WARNING:
    status, h_cls, b_cls = "WARNING", "health-warning", "badge-warning"
else:
    status, h_cls, b_cls = "CRITICAL", "health-critical", "badge-critical"

st.markdown(f'<div class="health-indicator {h_cls}"><span class="health-badge {b_cls}">{status}</span> Portfolio Status: {status}<br>Monthly Avg: ฿{monthly_avg:,.0f} ({variance:+.1f}%)</div>', unsafe_allow_html=True)

# Visualizations [cite: 73-76]
st.markdown('<div class="section-header">📈 Utilization & Trends</div>', unsafe_allow_html=True)
st.markdown(generate_service_table_html(filtered_df, selected_years), unsafe_allow_html=True)

c1, c2 = st.columns(2)
fig_trend = px.line(filtered_df.groupby('Month')['Fee (Baht)'].sum().reset_index(), x='Month', y='Fee (Baht)', title='Monthly Spending Trend')
c1.plotly_chart(fig_trend, use_container_width=True)

fig_pie = px.pie(filtered_df, names='ประเภทการบริการ', title='Service Distribution', hole=0.4)
c2.plotly_chart(fig_pie, use_container_width=True)

# Export
csv = filtered_df.to_csv(index=False).encode('utf-8-sig')
st.download_button("💾 Download Filtered Data (CSV)", data=csv, file_name="RSA_Export.csv", mime="text/csv")