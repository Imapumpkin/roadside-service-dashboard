import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime

# --- 1. CONFIGURATION ---
MONTHLY_BUDGET = 2_000_000
pd.set_option('future.no_silent_downcasting', True)

st.set_page_config(page_title="Dynamic RSA Dashboard", page_icon="🚗", layout="wide")

# --- 2. DYNAMIC HEADER & UI SETTINGS ---
with st.sidebar:
    st.header("⚙️ Dashboard Settings")
    
    with st.expander("📝 Customize Labels", expanded=True):
        # These change the DISPLAY names globally
        lbl_lob = st.text_input("LOB Header", "Line of Business")
        lbl_serv = st.text_input("Service Header", "Service Type")
        lbl_make = st.text_input("Make Header", "Vehicle Make")
        lbl_model = st.text_input("Model Header", "Vehicle Model")
        lbl_policy = st.text_input("Policy Header", "Policy Type")
    
    header_map = {
        "LOB": lbl_lob,
        "Service": lbl_serv,
        "Make": lbl_make,
        "Model": lbl_model,
        "Policy": lbl_policy
    }

# --- 3. STYLING ---
st.markdown(f"""
<style>
    [data-testid="stSidebar"] {{ background: #1B2838; }}
    .metric-card {{
        background: white; padding: 20px; border-radius: 10px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1); border-top: 5px solid #4A90D9;
    }}
    .metric-val {{ font-size: 24px; font-weight: bold; color: #1B2838; }}
</style>
""", unsafe_allow_html=True)

# --- 4. DATA LOADING ---
@st.cache_data
def load_data():
    try:
        # Load excel and find header row automatically
        df_raw = pd.read_excel('(Test) RSA Report.xlsx', header=None)
        for idx, row in df_raw.iterrows():
            if 'Policy No.' in row.values:
                df = df_raw.iloc[idx+1:].copy()
                df.columns = df_raw.iloc[idx]
                break
        
        # Standardize internal names for logic
        if 'Fee\n (Baht)' in df.columns:
            df = df.rename(columns={'Fee\n (Baht)': 'Fee'})
        
        df['วันที่'] = pd.to_datetime(df['วันที่'], errors='coerce')
        df['Year'] = df['วันที่'].dt.year.fillna(0).astype(int)
        df['Month'] = df['วันที่'].dt.month.fillna(0).astype(int)
        df['Fee'] = pd.to_numeric(df['Fee'], errors='coerce').fillna(0)
        df['LOB_Internal'] = df['Policy No.'].str.extract(r'(A[CV]\d)', expand=False).fillna('Unre')
        
        return df.reset_index(drop=True)
    except Exception as e:
        st.error(f"File Error: {e}")
        return pd.DataFrame()

df = load_data()

if not df.empty:
    # --- 5. DYNAMIC FILTERS ---
    st.sidebar.markdown("---")
    st.sidebar.subheader("🔍 Filters")
    
    # Year Filter
    years = sorted(df['Year'].unique())
    selected_years = st.sidebar.multiselect("Select Year", years, default=years)
    
    # Dynamic LOB Filter (linked to your custom header)
    lobs = sorted(df['LOB_Internal'].unique())
    selected_lobs = st.sidebar.multiselect(f"Filter {header_map['LOB']}", lobs, default=lobs)

    # Filter the dataframe
    mask = (df['Year'].isin(selected_years)) & (df['LOB_Internal'].isin(selected_lobs))
    filtered_df = df[mask].copy()

    # --- 6. DASHBOARD DISPLAY ---
    st.title(f"🚗 {header_map['LOB']} Monitoring")
    
    # KPI Row
    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown(f'<div class="metric-card"><p>Total Cases</p><p class="metric-val">{len(filtered_df):,}</p></div>', unsafe_allow_html=True)
    with c2:
        st.markdown(f'<div class="metric-card"><p>Total Fees</p><p class="metric-val">฿{filtered_df["Fee"].sum():,.0f}</p></div>', unsafe_allow_html=True)
    with c3:
        st.markdown(f'<div class="metric-card"><p>Active {header_map["LOB"]}s</p><p class="metric-val">{filtered_df["LOB_Internal"].nunique()}</p></div>', unsafe_allow_html=True)

    # Dynamic Table
    st.subheader(f"Detailed Breakdown by {header_map['LOB']}")
    
    # Creating a display table that uses your custom headers
    display_table = filtered_df.groupby(['LOB_Internal', 'ประเภทการบริการ']).size().unstack(fill_value=0)
    display_table.index.name = header_map['LOB']
    display_table.columns.name = header_map['Service']
    
    st.dataframe(display_table, use_container_width=True)

    # Chart
    fig = px.bar(filtered_df, x='Month', y='Fee', color='LOB_Internal', 
                 title=f"Monthly Fee Trend by {header_map['LOB']}",
                 labels={'LOB_Internal': header_map['LOB'], 'Fee': 'Cost (Baht)'})
    st.plotly_chart(fig, use_container_width=True)

else:
    st.info("Please ensure '(Test) RSA Report.xlsx' is in the folder.")