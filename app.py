import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import os
import html
from io import BytesIO

# ============================================================================
# CONFIGURATION
# ============================================================================
MONTHLY_BUDGET = 200_000
HEALTH_THRESHOLD_HEALTHY = 5
HEALTH_THRESHOLD_WARNING = 15
CACHE_TTL = 3600
DEFAULT_DATA_FILE = "(Test) RSA Report.xlsx"
UPLOAD_DIR = "uploaded_data"
PASSWORD = "sompo2024"  # Change this or use st.secrets["password"] in production

# Suppress pandas future warning (outside cached functions)
pd.set_option('future.no_silent_downcasting', True)

# Page configuration - must be first Streamlit command
st.set_page_config(
    page_title="RSA Dashboard - Sompo Thailand",
    page_icon="🚗",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ============================================================================
# PASSWORD PROTECTION
# ============================================================================
def check_password():
    """Simple password gate for the dashboard."""
    if "authenticated" not in st.session_state:
        st.session_state.authenticated = False

    if st.session_state.authenticated:
        return True

    st.markdown("""
    <style>
        .login-container {
            max-width: 420px;
            margin: 10vh auto;
            padding: 48px 40px;
            background: white;
            border-radius: 16px;
            box-shadow: 0 8px 32px rgba(27, 40, 56, 0.12);
            text-align: center;
        }
        .login-title {
            font-size: 28px;
            font-weight: 700;
            color: #1B2838;
            margin-bottom: 8px;
        }
        .login-subtitle {
            font-size: 14px;
            color: #718096;
            margin-bottom: 32px;
        }
        .stApp { background-color: #F0F2F6; }
        #MainMenu {visibility: hidden;}
        footer {visibility: hidden;}
        header {visibility: hidden;}
    </style>
    """, unsafe_allow_html=True)

    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.markdown("""
        <div style="text-align:center; margin-top: 10vh;">
            <div style="font-size: 48px; margin-bottom: 16px;">🚗</div>
            <div style="font-size: 28px; font-weight: 700; color: #1B2838; margin-bottom: 8px;">RSA Dashboard</div>
            <div style="font-size: 14px; color: #718096; margin-bottom: 32px;">Sompo Thailand - Roadside Assistance</div>
        </div>
        """, unsafe_allow_html=True)

        password_input = st.text_input("Enter Password", type="password", key="password_input")
        login_btn = st.button("Login", use_container_width=True, type="primary")

        if login_btn:
            # Try st.secrets first, fall back to hardcoded
            try:
                correct_pw = st.secrets["password"]
            except Exception:
                correct_pw = PASSWORD

            if password_input == correct_pw:
                st.session_state.authenticated = True
                st.rerun()
            else:
                st.error("Incorrect password. Please try again.")
    return False


if not check_password():
    st.stop()

# ============================================================================
# CUSTOM CSS
# ============================================================================
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
    * { font-family: 'Inter', 'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif; }
    .main { padding: 1.5rem 2rem; background-color: #F0F2F6; }
    .stApp { background-color: #F0F2F6; }
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}

    /* Sidebar */
    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #1B2838 0%, #2A3F54 100%);
        padding-top: 2rem;
    }
    [data-testid="stSidebar"] .stMarkdown { color: #FFFFFF; }
    [data-testid="stSidebar"] label {
        color: #E0E6ED !important;
        font-weight: 500 !important;
        font-size: 13px !important;
        letter-spacing: 0.03em;
        margin-top: 0.8rem;
    }
    [data-testid="stSidebar"] .stMultiSelect [data-baseweb="select"] {
        background-color: rgba(255, 255, 255, 0.1);
        border: 1px solid rgba(255, 255, 255, 0.2);
        border-radius: 6px;
    }
    [data-testid="stSidebar"] h2 {
        color: #FFFFFF !important; font-size: 18px !important;
        font-weight: 600 !important; margin-bottom: 1.5rem;
    }

    /* KPI Cards */
    .metric-card {
        background: linear-gradient(135deg, #1B2838 0%, #2D4A5C 100%);
        padding: 24px 20px; border-radius: 12px; color: white;
        box-shadow: 0 4px 12px rgba(27,40,56,0.15), 0 1px 3px rgba(0,0,0,0.08);
        margin-bottom: 16px; transition: all 0.3s ease;
        border: 1px solid rgba(255,255,255,0.05);
        position: relative; overflow: hidden;
    }
    .metric-card::before {
        content: ''; position: absolute; top: 0; left: 0; right: 0; height: 3px;
        background: linear-gradient(90deg, #4A90D9 0%, #6FB1FF 100%);
    }
    .metric-card:hover {
        transform: translateY(-2px);
        box-shadow: 0 8px 20px rgba(27,40,56,0.25), 0 2px 6px rgba(0,0,0,0.12);
    }
    .metric-title {
        font-size: 12px; font-weight: 600; margin-bottom: 12px;
        letter-spacing: 0.08em; color: #D0DCE8;
    }
    .metric-value {
        font-size: 28px; font-weight: 700; margin-bottom: 6px;
        line-height: 1.2; color: #FFFFFF; font-variant-numeric: tabular-nums;
    }
    .positive { color: #27AE60; }
    .negative { color: #E74C3C; }
    .warning  { color: #F39C12; }

    /* Health Indicator */
    .health-indicator {
        padding: 24px 28px; border-radius: 12px; margin: 20px 0;
        font-weight: 600; box-shadow: 0 2px 8px rgba(0,0,0,0.08);
        border-left: 6px solid; background: white;
    }
    .health-healthy { background: linear-gradient(135deg, #d4edda 0%, #e8f5e9 100%); color: #155724; border-left-color: #27AE60; }
    .health-warning { background: linear-gradient(135deg, #fff3cd 0%, #fff9e6 100%); color: #856404; border-left-color: #F39C12; }
    .health-critical { background: linear-gradient(135deg, #f8d7da 0%, #ffe6e8 100%); color: #721c24; border-left-color: #E74C3C; }
    .health-indicator .health-title { font-size: 20px; margin-bottom: 16px; font-weight: 700; display: flex; align-items: center; gap: 10px; }
    .health-badge {
        display: inline-block; padding: 4px 12px; border-radius: 4px;
        font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.05em;
    }
    .badge-healthy { background-color: #27AE60; color: white; }
    .badge-warning { background-color: #F39C12; color: white; }
    .badge-critical { background-color: #E74C3C; color: white; }
    .health-indicator .health-stats {
        display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
        gap: 20px; margin-top: 16px;
    }
    .health-stat-item { display: flex; flex-direction: column; }
    .health-stat-label { font-size: 11px; letter-spacing: 0.05em; opacity: 0.8; margin-bottom: 4px; font-weight: 600; }
    .health-stat-value { font-size: 18px; font-weight: 700; font-variant-numeric: tabular-nums; }

    /* Section Headers */
    .section-header {
        color: #1B2838; font-size: 24px; font-weight: 700;
        margin-top: 48px; margin-bottom: 24px; padding-bottom: 12px;
        border-bottom: 3px solid #4A90D9; display: flex; align-items: center; gap: 12px;
    }
    .data-freshness {
        display: inline-block; font-size: 13px; color: #718096;
        margin-left: 16px; padding: 4px 12px; background-color: #EDF2F7;
        border-radius: 6px; font-weight: 500;
    }

    /* Service Table */
    .service-table-container { background: white; border-radius: 12px; padding: 24px; box-shadow: 0 2px 8px rgba(0,0,0,0.08); overflow-x: auto; }
    .service-table { font-size: 13px; width: 100%; border-collapse: separate; border-spacing: 0; }
    .service-table th {
        background: linear-gradient(135deg, #1B2838 0%, #2A3F54 100%);
        color: white; padding: 14px 16px; text-align: left;
        font-weight: 600; font-size: 12px; letter-spacing: 0.05em; border: none;
    }
    .service-table th:first-child { border-top-left-radius: 8px; }
    .service-table th:last-child { border-top-right-radius: 8px; }
    .service-table td { padding: 12px 16px; border-bottom: 1px solid #E2E8F0; background-color: #FFFFFF; }
    .service-table tr:hover td { background-color: #F8FAFC; }
    .subtotal-row td { background-color: #EBF5FB !important; font-weight: 600; color: #1B2838; border-top: 2px solid #4A90D9; border-bottom: 2px solid #4A90D9; }
    .total-row td { background: linear-gradient(135deg, #1B2838 0%, #2A3F54 100%) !important; color: white !important; font-weight: 700; font-size: 14px; padding: 16px; border: none; position: sticky; bottom: 0; z-index: 2; }
    .bar-cell { display: flex; align-items: center; gap: 8px; min-width: 100px; }
    .bar-cell-number { min-width: 40px; font-weight: 600; color: #1B2838; font-variant-numeric: tabular-nums; text-align: right; }
    .bar-cell-number.zero-value { color: #A0AEC0; font-weight: 400; }
    .bar-container { flex: 1; background-color: #E8EEF7; border-radius: 4px; height: 22px; position: relative; overflow: hidden; }
    .inline-bar { height: 100%; background: linear-gradient(90deg, #4A90D9 0%, #6FB1FF 100%); border-radius: 4px; transition: width 0.3s ease; }

    /* Empty State */
    .empty-state { background: white; border-radius: 12px; padding: 48px 24px; text-align: center; box-shadow: 0 2px 8px rgba(0,0,0,0.08); margin: 24px 0; }
    .empty-state-icon { font-size: 64px; margin-bottom: 16px; opacity: 0.5; }
    .empty-state-title { font-size: 20px; font-weight: 600; color: #1B2838; margin-bottom: 8px; }
    .empty-state-message { font-size: 14px; color: #718096; }

    /* Titles */
    h1 { color: #1B2838 !important; font-weight: 700 !important; font-size: 32px !important; margin-bottom: 8px !important; }
    h3 { color: #4A5568 !important; font-weight: 500 !important; font-size: 16px !important; margin-bottom: 32px !important; }

    /* Charts */
    .js-plotly-plot { border-radius: 12px; box-shadow: 0 2px 8px rgba(0,0,0,0.06); background: white; padding: 12px; }

    /* Buttons */
    .stDownloadButton button {
        background: linear-gradient(135deg, #4A90D9 0%, #2D5AA0 100%);
        color: white; border: none; border-radius: 8px; padding: 12px 24px; font-weight: 600;
    }
    .stDownloadButton button:hover { background: linear-gradient(135deg, #2D5AA0 0%, #1B2838 100%); box-shadow: 0 4px 12px rgba(74,144,217,0.3); }

    /* Footer */
    .dashboard-footer { text-align: center; color: #718096; padding: 32px 20px; margin-top: 48px; border-top: 2px solid #E2E8F0; font-size: 13px; background: white; border-radius: 12px; }

    /* Responsive */
    @media (max-width: 768px) {
        .metric-card { margin-bottom: 12px; }
        .metric-value { font-size: 22px; }
        .section-header { font-size: 18px; margin-top: 32px; }
        .health-indicator .health-stats { grid-template-columns: 1fr 1fr; }
        .bar-container { display: none; }
        .bar-cell { min-width: auto; }
    }
    @media (max-width: 1200px) {
        div[style*="grid-template-columns:repeat(6"] { grid-template-columns: repeat(3, 1fr) !important; }
    }
    @media (max-width: 768px) {
        div[style*="grid-template-columns:repeat(6"] { grid-template-columns: repeat(2, 1fr) !important; }
    }
</style>
""", unsafe_allow_html=True)


# ============================================================================
# DATA LOADING
# ============================================================================
@st.cache_data(ttl=CACHE_TTL)
def process_dataframe(df_raw):
    """Process raw dataframe with all cleaning logic."""
    df = df_raw.copy()

    # Find header row dynamically
    header_found = False
    for idx in range(min(len(df), 30)):  # Only check first 30 rows
        if 'Policy No.' in df.iloc[idx].values:
            df.columns = df.iloc[idx]
            df = df.iloc[idx + 1:].reset_index(drop=True)
            header_found = True
            break

    if not header_found:
        raise ValueError("Could not find header row containing 'Policy No.' in the Excel file.")

    # Clean Fee column name (has newline)
    fee_cols = [c for c in df.columns if isinstance(c, str) and 'Fee' in c and 'Exceed' not in c]
    if fee_cols:
        df = df.rename(columns={fee_cols[0]: 'Fee (Baht)'})

    # Process date column - handle both string and datetime types
    if 'วันที่' in df.columns:
        if df['วันที่'].dtype == 'object':
            df['วันที่'] = pd.to_datetime(df['วันที่'], format='%d/%m/%Y', errors='coerce')
        else:
            df['วันที่'] = pd.to_datetime(df['วันที่'], errors='coerce')

        # Drop rows where date parsing failed entirely
        df = df.dropna(subset=['วันที่'])

        # Add date components using standard int (safe after dropna)
        df['Day'] = df['วันที่'].dt.day.astype(int)
        df['Month'] = df['วันที่'].dt.month.astype(int)
        df['Year'] = df['วันที่'].dt.year.astype(int)
        df['วันที่'] = df['วันที่'].dt.date

    # Drop fully empty rows
    df = df.dropna(how='all').reset_index(drop=True)

    # Replace '-' with pd.NA
    df = df.replace('-', pd.NA)

    # Extract Policy Type (LOB)
    if 'Policy No.' in df.columns:
        df['Policy Type'] = df['Policy No.'].str.extract(r'(A[CV]\d)', expand=False)
    df['LOB'] = df['Policy Type'].fillna('Unre') if 'Policy Type' in df.columns else 'Unre'

    # Extract province from license plate
    if 'ทะเบียนรถ' in df.columns:
        # Clean special characters from existing province column
        if 'จังหวัด ทะเบียนรถ' in df.columns:
            df['จังหวัด ทะเบียนรถ'] = df['จังหวัด ทะเบียนรถ'].astype(str).str.replace(r'[.!@#$%^&*\d]', '', regex=True).str.strip()
            df['จังหวัด ทะเบียนรถ'] = df['จังหวัด ทะเบียนรถ'].replace(['', 'nan', 'None', '<NA>'], pd.NA)
        # Extract trailing Thai text after the last digit in the plate
        plate_province = df['ทะเบียนรถ'].astype(str).str.extract(r'(\d)[.\s]*([ก-๙]+)\s*[.!@#$%^&*]*\s*$', expand=True)
        extracted = plate_province[1]
        # Only fill where the existing column is missing
        if 'จังหวัด ทะเบียนรถ' in df.columns:
            df['จังหวัด ทะเบียนรถ'] = df['จังหวัด ทะเบียนรถ'].fillna(extracted)
        else:
            df['จังหวัด ทะเบียนรถ'] = extracted
        df['จังหวัด ทะเบียนรถ'] = df['จังหวัด ทะเบียนรถ'].replace(['กรุงเทพ', 'กทม'], 'กรุงเทพมหานคร')

    # Convert Fee to numeric
    if 'Fee (Baht)' in df.columns:
        df['Fee (Baht)'] = pd.to_numeric(df['Fee (Baht)'], errors='coerce')

    # Normalize vehicle make to uppercase
    if 'ยี่ห้อรถ' in df.columns:
        df['ยี่ห้อรถ'] = df['ยี่ห้อรถ'].astype(str).str.upper().replace('<NA>', pd.NA).replace('NAN', pd.NA)

    return df


def load_default_data():
    """Load data from default Excel file."""
    if os.path.exists(DEFAULT_DATA_FILE):
        return pd.read_excel(DEFAULT_DATA_FILE, header=None)
    return None


def load_persisted_upload():
    """Load previously uploaded file if it exists."""
    os.makedirs(UPLOAD_DIR, exist_ok=True)
    persisted_path = os.path.join(UPLOAD_DIR, "persisted_upload.xlsx")
    if os.path.exists(persisted_path):
        return pd.read_excel(persisted_path, header=None), persisted_path
    return None, None


def persist_uploaded_file(uploaded_file):
    """Save uploaded file to disk for persistence across reruns."""
    os.makedirs(UPLOAD_DIR, exist_ok=True)
    persisted_path = os.path.join(UPLOAD_DIR, "persisted_upload.xlsx")
    with open(persisted_path, "wb") as f:
        f.write(uploaded_file.getbuffer())
    return persisted_path


# ============================================================================
# DATA SOURCE SELECTION (silent – UI for upload moved to bottom)
# ============================================================================
# Check session state for uploaded file first
if 'uploaded_file_bytes' not in st.session_state:
    st.session_state.uploaded_file_bytes = None
    st.session_state.uploaded_file_name = None

df_raw = None
data_source_label = ""

if st.session_state.uploaded_file_bytes is not None:
    df_raw = pd.read_excel(BytesIO(st.session_state.uploaded_file_bytes), header=None)
    data_source_label = f"Uploaded: {st.session_state.uploaded_file_name}"
else:
    persisted_df, persisted_path = load_persisted_upload()
    if persisted_df is not None:
        df_raw = persisted_df
        data_source_label = "Previously uploaded file"
    else:
        df_raw = load_default_data()
        if df_raw is not None:
            data_source_label = DEFAULT_DATA_FILE

if df_raw is None:
    st.markdown("# 🚗 RSA Dashboard - Sompo Thailand")
    st.markdown("""
    <div class="empty-state">
        <div class="empty-state-icon">📂</div>
        <div class="empty-state-title">No Data Available</div>
        <div class="empty-state-message">
            Please upload an RSA Report Excel file using the sidebar, or place the default file in the application directory.
        </div>
    </div>
    """, unsafe_allow_html=True)
    st.stop()

# Process the data
with st.spinner("Processing data..."):
    try:
        df = process_dataframe(df_raw)
    except Exception as e:
        st.error(f"Error processing data: {str(e)}")
        st.stop()

# Validate required columns
required_cols = ['Year', 'Month', 'Fee (Baht)', 'ประเภทการบริการ', 'LOB']
missing = [c for c in required_cols if c not in df.columns]
if missing:
    st.error(f"Missing required columns after processing: {missing}")
    st.stop()

# Session state for stable export timestamp
if 'export_timestamp' not in st.session_state:
    st.session_state.export_timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')


# ============================================================================
# HELPER FUNCTIONS
# ============================================================================
def safe_sorted_unique(series):
    """Get sorted unique string values from a series, handling mixed types."""
    return sorted([str(v) for v in series.dropna().unique()])


def generate_service_table_html(data, years_list):
    """Generate HTML table for service utilization with inline bars."""
    lob_order = ['AV1', 'AV5', 'AV3', 'AV9', 'AC3', 'Unre']
    service_types = safe_sorted_unique(data['ประเภทการบริการ'])

    if not service_types or not years_list:
        return '<div class="empty-state"><div class="empty-state-title">No service data available</div></div>'

    # Calculate max for bar scaling
    max_count = 0
    for lob in lob_order:
        lob_df = data[data['LOB'] == lob]
        for svc in service_types:
            svc_df = lob_df[lob_df['ประเภทการบริการ'] == svc]
            for yr in years_list:
                cnt = len(svc_df[svc_df['Year'] == yr])
                max_count = max(max_count, cnt)

    num_year_cols = len(years_list)
    total_cols = num_year_cols + 3  # LOB + Service Type + years + Total

    tbl = '<div class="service-table-container"><table class="service-table">'
    tbl += '<thead><tr><th>LOB</th><th>Service Type</th>'
    for yr in years_list:
        tbl += f'<th>{yr}</th>'
    tbl += '<th>Total</th></tr></thead><tbody>'

    gt_per_year = {yr: 0 for yr in years_list}
    gt = 0

    for lob in lob_order:
        lob_df = data[data['LOB'] == lob]
        lt_per_year = {yr: 0 for yr in years_list}
        lt = 0

        for idx, svc in enumerate(service_types):
            svc_df = lob_df[lob_df['ประเภทการบริการ'] == svc]
            row_total = 0
            tbl += '<tr>'
            if idx == 0:
                tbl += f'<td rowspan="{len(service_types)}" style="background-color:#F7FAFC;font-weight:600;vertical-align:middle;color:#1B2838;border-right:2px solid #E2E8F0;">{html.escape(str(lob))}</td>'
            tbl += f'<td style="color:#4A5568;">{html.escape(str(svc))}</td>'

            for yr in years_list:
                cnt = len(svc_df[svc_df['Year'] == yr])
                row_total += cnt
                lt_per_year[yr] += cnt
                gt_per_year[yr] += cnt
                bw = (cnt / max_count * 100) if max_count > 0 else 0
                ncls = "bar-cell-number zero-value" if cnt == 0 else "bar-cell-number"
                dv = "—" if cnt == 0 else str(cnt)
                tbl += f'<td><div class="bar-cell"><span class="{ncls}">{dv}</span><div class="bar-container"><div class="inline-bar" style="width:{bw}%;"></div></div></div></td>'

            lt += row_total
            gt += row_total
            tbl += f'<td style="font-weight:700;color:#1B2838;">{row_total}</td></tr>'

        # Subtotal row spans full width
        tbl += f'<tr class="subtotal-row"><td colspan="2"><strong>Subtotal - {html.escape(str(lob))}</strong></td>'
        for yr in years_list:
            tbl += f'<td><strong>{lt_per_year[yr]:,}</strong></td>'
        tbl += f'<td><strong>{lt:,}</strong></td></tr>'

    # Grand total row - sticky bold at bottom
    tbl += f'<tr class="total-row"><td colspan="2"><strong>GRAND TOTAL</strong></td>'
    for yr in years_list:
        tbl += f'<td><strong>{gt_per_year[yr]:,}</strong></td>'
    tbl += f'<td><strong>{gt:,}</strong></td></tr>'
    tbl += '</tbody></table></div>'
    return tbl


PLOTLY_CONFIG = {'displayModeBar': True, 'displaylogo': False}

CHART_LAYOUT_DEFAULTS = dict(
    font=dict(family='Inter, sans-serif', size=12, color='#4A5568'),
    paper_bgcolor='white',
    plot_bgcolor='rgba(0,0,0,0)',
    xaxis=dict(gridcolor='#E2E8F0', showline=True, linecolor='#E2E8F0'),
    yaxis=dict(gridcolor='#E2E8F0', showline=True, linecolor='#E2E8F0'),
)

CHART_TITLE_FONT = {'size': 16, 'color': '#1B2838', 'family': 'Inter, sans-serif'}


@st.cache_data(ttl=CACHE_TTL)
def convert_df_to_csv(_df):
    """Convert dataframe to CSV bytes for download."""
    return _df.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')


# ============================================================================
# FILTERS – compact expander at the top of the main area
# ============================================================================
with st.expander("📊 Filters", expanded=False):
    fc1, fc2, fc3, fc4 = st.columns(4)

    available_years = sorted([int(y) for y in df['Year'].dropna().unique()])
    available_services = safe_sorted_unique(df['ประเภทการบริการ'])
    available_lobs = safe_sorted_unique(df['LOB'])
    month_names = {1:'Jan',2:'Feb',3:'Mar',4:'Apr',5:'May',6:'Jun',7:'Jul',8:'Aug',9:'Sep',10:'Oct',11:'Nov',12:'Dec'}
    available_months = sorted([int(m) for m in df['Month'].dropna().unique()])
    month_options = [f"{m} - {month_names.get(m,'')}" for m in available_months]
    available_channels = safe_sorted_unique(df['รหัสโครงการ']) if 'รหัสโครงการ' in df.columns else []
    available_regions = safe_sorted_unique(df['จังหวัด']) if 'จังหวัด' in df.columns else []
    available_makes = safe_sorted_unique(df['ยี่ห้อรถ']) if 'ยี่ห้อรถ' in df.columns else []
    available_models = safe_sorted_unique(df['รุ่นรถ']) if 'รุ่นรถ' in df.columns else []

    with fc1:
        selected_years = st.multiselect("Year", options=available_years, default=available_years, key="filter_year")
        selected_services = st.multiselect("Service Type", options=['All'] + available_services, default=['All'], key="filter_service")
    with fc2:
        selected_lobs = st.multiselect("LOB", options=['All'] + available_lobs, default=['All'], key="filter_lob")
        selected_month_display = st.multiselect("Month", options=['All'] + month_options, default=['All'], key="filter_month")
    with fc3:
        selected_channels = st.multiselect("Channel", options=['All'] + available_channels, default=['All'], key="filter_channel")
        selected_regions = st.multiselect("Region", options=['All'] + available_regions, default=['All'], key="filter_region")
    with fc4:
        selected_makes = st.multiselect("Vehicle Make", options=['All'] + available_makes, default=['All'], key="filter_make")
        selected_models = st.multiselect("Vehicle Model", options=['All'] + available_models, default=['All'], key="filter_model")

if not selected_years:
    st.warning("Please select at least one year.")
    st.stop()
if 'All' in selected_services:
    selected_services = ['All']
if 'All' in selected_lobs:
    selected_lobs = ['All']
if 'All' in selected_channels:
    selected_channels = ['All']
if 'All' in selected_regions:
    selected_regions = ['All']
if 'All' in selected_makes:
    selected_makes = ['All']
if 'All' in selected_models:
    selected_models = ['All']
if 'All' in selected_month_display:
    selected_months = available_months
else:
    selected_months = [int(m.split(' - ')[0]) for m in selected_month_display]

# ============================================================================
# APPLY FILTERS (single boolean mask)
# ============================================================================
mask = df['Year'].isin(selected_years)
if 'All' not in selected_services:
    mask &= df['ประเภทการบริการ'].isin(selected_services)
if 'All' not in selected_lobs:
    mask &= df['LOB'].isin(selected_lobs)
if 'All' not in selected_channels and 'รหัสโครงการ' in df.columns:
    mask &= df['รหัสโครงการ'].astype(str).isin(selected_channels)
if 'All' not in selected_regions and 'จังหวัด' in df.columns:
    mask &= df['จังหวัด'].astype(str).isin(selected_regions)
if 'All' not in selected_month_display:
    mask &= df['Month'].isin(selected_months)
if 'All' not in selected_makes and 'ยี่ห้อรถ' in df.columns:
    mask &= df['ยี่ห้อรถ'].astype(str).isin(selected_makes)
if 'All' not in selected_models and 'รุ่นรถ' in df.columns:
    mask &= df['รุ่นรถ'].astype(str).isin(selected_models)

filtered_df = df[mask]

# Empty state
if len(filtered_df) == 0:
    st.markdown("# 🚗 RSA Dashboard - Sompo Thailand")
    st.markdown("""
    <div class="empty-state">
        <div class="empty-state-icon">🔍</div>
        <div class="empty-state-title">No Data Found</div>
        <div class="empty-state-message">Current filter selection returned no results. Please adjust your filters.</div>
    </div>
    """, unsafe_allow_html=True)
    st.stop()

# ============================================================================
# DASHBOARD HEADER
# ============================================================================
latest_date = df['วันที่'].max() if 'วันที่' in df.columns else "N/A"
st.markdown("# 🚗 RSA Dashboard - Sompo Thailand")
st.markdown(f"### Roadside Assistance Monitoring <span class='data-freshness'>Data through: {latest_date} | Source: {data_source_label}</span>", unsafe_allow_html=True)

# ============================================================================
# KPIs
# ============================================================================
st.markdown('<div class="section-header">📈 Key Performance Indicators</div>', unsafe_allow_html=True)

current_year = max(selected_years)
prev_year = current_year - 1
cur_df = filtered_df[filtered_df['Year'] == current_year]
prev_df = filtered_df[filtered_df['Year'] == prev_year]

# Current & previous year fee totals (used by portfolio health)
cur_fee = cur_df['Fee (Baht)'].sum()
prev_fee = prev_df['Fee (Baht)'].sum()

current_month = datetime.now().month

# YTD = full year-to-date for current year
ytd_cases = len(cur_df)
ytd_fee = cur_df['Fee (Baht)'].sum()
prev_ytd_cases = len(prev_df[prev_df['Month'] <= current_month])
prev_ytd_fee = prev_df[prev_df['Month'] <= current_month]['Fee (Baht)'].sum()

# MTD = current month only
mtd_fee = cur_df[cur_df['Month'] == current_month]['Fee (Baht)'].sum()
mtd_util = (mtd_fee / MONTHLY_BUDGET * 100) if MONTHLY_BUDGET > 0 else 0
prev_mtd_fee = prev_df[prev_df['Month'] == current_month]['Fee (Baht)'].sum()

# Avg fee current year
cur_avg_raw = cur_df['Fee (Baht)'].mean()
cur_avg = 0.0 if pd.isna(cur_avg_raw) else cur_avg_raw
prev_avg_raw = prev_df['Fee (Baht)'].mean()
prev_avg = 0.0 if pd.isna(prev_avg_raw) else prev_avg_raw

def yoy_html(cur_val, prev_val, compare_year):
    """Generate YoY comparison HTML snippet. Cost increase = bad (red), decrease = good (green)."""
    if prev_val == 0:
        return f'<div style="font-size:11px;margin-top:8px;opacity:0.8;">vs {compare_year}: N/A</div>'
    pct = (cur_val - prev_val) / prev_val * 100
    cls = "negative" if pct > 0 else "positive"
    arrow = "▲" if pct >= 0 else "▼"
    return f'<div style="font-size:11px;margin-top:8px;opacity:0.9;"><span class="{cls}">{arrow} {abs(pct):.1f}%</span> vs {compare_year}</div>'

mc = "negative" if mtd_util > 100 else "positive"

kpi_cards = [
    (f"YTD Total Cases ({current_year})", f"{ytd_cases:,}", yoy_html(ytd_cases, prev_ytd_cases, prev_year)),
    (f"YTD Total Fee ({current_year})", f"฿{ytd_fee:,.0f}", yoy_html(ytd_fee, prev_ytd_fee, prev_year)),
    (f"Avg Fee/Case ({current_year})", f"฿{cur_avg:,.0f}", yoy_html(cur_avg, prev_avg, prev_year)),
    ("Monthly Budget", f"฿{MONTHLY_BUDGET:,}", f'<div style="font-size:11px;margin-top:8px;opacity:0.8;">Annual: ฿{MONTHLY_BUDGET * 12:,}</div>'),
    ("MTD Utilization", f'<span class="{mc}">{mtd_util:.1f}%</span>', f'<div style="font-size:11px;margin-top:8px;opacity:0.8;">฿{mtd_fee:,.0f} / ฿{MONTHLY_BUDGET:,}</div>'),
    (f"MTD Fee ({current_year})", f"฿{mtd_fee:,.0f}", yoy_html(mtd_fee, prev_mtd_fee, prev_year)),
]

kpi_html = '<div style="display:grid;grid-template-columns:repeat(6,1fr);gap:16px;">'
for title, value, extra in kpi_cards:
    kpi_html += f'<div class="metric-card"><div class="metric-title">{title}</div><div class="metric-value">{value}</div>{extra}</div>'
kpi_html += '</div>'
st.markdown(kpi_html, unsafe_allow_html=True)

# ============================================================================
# PORTFOLIO HEALTH
# ============================================================================
st.markdown('<div class="section-header">🏥 Portfolio Health Indicator</div>', unsafe_allow_html=True)

months_in_year = cur_df['Month'].nunique() if len(cur_df) > 0 else 1
run_rate = cur_fee / max(months_in_year, 1)
budget_left_amt = MONTHLY_BUDGET - run_rate
budget_left_pct = (budget_left_amt / MONTHLY_BUDGET * 100) if MONTHLY_BUDGET > 0 else 0
projection = run_rate * 12

# budget_left_pct positive = under budget (good), negative = over budget (bad)
# Use negative of budget_left to check: over budget = bad
over_budget_pct = -budget_left_pct  # positive means over budget
if over_budget_pct <= HEALTH_THRESHOLD_HEALTHY:
    h_status, h_class, h_badge = "HEALTHY", "health-healthy", '<span class="health-badge badge-healthy">Healthy</span>'
elif over_budget_pct <= HEALTH_THRESHOLD_WARNING:
    h_status, h_class, h_badge = "WARNING", "health-warning", '<span class="health-badge badge-warning">Warning</span>'
else:
    h_status, h_class, h_badge = "CRITICAL", "health-critical", '<span class="health-badge badge-critical">Critical</span>'

st.markdown(f"""
<div class="health-indicator {h_class}">
    <div class="health-title">{h_badge} Portfolio Status: {h_status}</div>
    <div class="health-stats">
        <div class="health-stat-item"><div class="health-stat-label">Monthly Run Rate</div><div class="health-stat-value">฿{run_rate:,.0f}</div></div>
        <div class="health-stat-item"><div class="health-stat-label">Budget Left</div><div class="health-stat-value">฿{budget_left_amt:+,.0f} ({budget_left_pct:+.2f}%)</div></div>
        <div class="health-stat-item"><div class="health-stat-label">Year-End Projection</div><div class="health-stat-value">฿{projection:,.0f}</div></div>
        <div class="health-stat-item"><div class="health-stat-label">Annual Budget</div><div class="health-stat-value">฿{MONTHLY_BUDGET * 12:,.0f}</div></div>
    </div>
</div>
""", unsafe_allow_html=True)

# ============================================================================
# SERVICE UTILIZATION TABLE (Interactive Pivot)
# ============================================================================
st.markdown('<div class="section-header">📋 Service Utilization – Interactive Pivot Table</div>', unsafe_allow_html=True)

# --- Pivot table controls ---
pivot_cols_available = [c for c in ['LOB', 'ประเภทการบริการ', 'Year', 'Month', 'จังหวัด', 'ยี่ห้อรถ', 'รุ่นรถ', 'Policy Type', 'รหัสโครงการ', 'แผนก'] if c in filtered_df.columns]
value_cols_available = [c for c in ['Fee (Baht)', 'ลูกค้าจ่ายส่วนต่าง'] if c in filtered_df.columns]

pc1, pc2, pc3, pc4 = st.columns(4)
with pc1:
    pivot_rows = st.multiselect("Rows", options=pivot_cols_available, default=['LOB'], key="pivot_rows")
with pc2:
    pivot_columns = st.multiselect("Columns", options=pivot_cols_available, default=['Year'], key="pivot_columns")
with pc3:
    pivot_value = st.selectbox("Values", options=['Case Count'] + value_cols_available, index=0, key="pivot_value")
with pc4:
    pivot_agg = st.selectbox("Aggregation", options=['Count', 'Sum', 'Mean', 'Median', 'Min', 'Max'], index=0, key="pivot_agg")

if pivot_rows or pivot_columns:
    try:
        agg_map = {'Count': 'count', 'Sum': 'sum', 'Mean': 'mean', 'Median': 'median', 'Min': 'min', 'Max': 'max'}
        agg_func = agg_map[pivot_agg]

        if pivot_value == 'Case Count':
            _pivot_src = filtered_df.copy()
            _pivot_src['_count'] = 1
            val_col = '_count'
            agg_func = 'sum' if pivot_agg == 'Count' else agg_func
        else:
            _pivot_src = filtered_df.copy()
            val_col = pivot_value

        # Ensure all pivot row/column fields are string type to avoid unhashable issues
        all_pivot_fields = list(set(pivot_rows + pivot_columns))
        for col_name in all_pivot_fields:
            _pivot_src[col_name] = _pivot_src[col_name].astype(str).fillna('(blank)')

        # Ensure value column is numeric
        if val_col != '_count':
            _pivot_src[val_col] = pd.to_numeric(_pivot_src[val_col], errors='coerce').fillna(0)

        pivot_kwargs = dict(
            data=_pivot_src,
            values=val_col,
            aggfunc=agg_func,
            fill_value=0,
            margins=True,
            margins_name='Grand Total',
        )
        if pivot_rows:
            pivot_kwargs['index'] = pivot_rows
        if pivot_columns:
            pivot_kwargs['columns'] = pivot_columns

        pivot_result = pd.pivot_table(**pivot_kwargs)

        # Flatten multi-level column names
        if isinstance(pivot_result.columns, pd.MultiIndex):
            pivot_result.columns = [' | '.join(str(c) for c in col).strip(' | ') for col in pivot_result.columns]

        # If index is a MultiIndex, reset it so each row field becomes its own column
        if isinstance(pivot_result.index, pd.MultiIndex):
            pivot_result = pivot_result.reset_index()
        elif pivot_result.index.name:
            pivot_result = pivot_result.reset_index()

        # Format numeric columns
        fmt_pivot = pivot_result.copy()
        numeric_cols = fmt_pivot.select_dtypes(include=['float64', 'float32']).columns
        for col in numeric_cols:
            if pivot_agg in ('Sum', 'Count'):
                fmt_pivot[col] = fmt_pivot[col].astype(int)

        is_int_agg = pivot_agg in ('Sum', 'Count')

        # Separate grand total row from data rows so sorting won't move it
        gt_label = 'Grand Total'
        row_id_cols = [c for c in fmt_pivot.columns if c in pivot_rows]
        if row_id_cols:
            gt_mask = fmt_pivot[row_id_cols[0]].astype(str) == gt_label
        else:
            gt_mask = pd.Series([False] * len(fmt_pivot), index=fmt_pivot.index)

        data_rows = fmt_pivot[~gt_mask].reset_index(drop=True)
        grand_total_rows = fmt_pivot[gt_mask].reset_index(drop=True)

        # Display data rows with comma-formatted numbers via HTML table
        if len(data_rows) > 0:
            num_cols_set = set(data_rows.select_dtypes(include=['int64', 'int32', 'float64', 'float32']).columns)

            pivot_html = '<div class="service-table-container"><table class="service-table"><thead><tr>'
            for col in data_rows.columns:
                pivot_html += f'<th>{html.escape(str(col))}</th>'
            pivot_html += '</tr></thead><tbody>'
            for _, row in data_rows.iterrows():
                pivot_html += '<tr>'
                for col in data_rows.columns:
                    val = row[col]
                    if col in num_cols_set:
                        if is_int_agg:
                            cell = f"{int(val):,}"
                        else:
                            cell = f"{val:,.2f}"
                    else:
                        cell = html.escape(str(val))
                    pivot_html += f'<td>{cell}</td>'
                pivot_html += '</tr>'
            pivot_html += '</tbody></table></div>'
            st.markdown(pivot_html, unsafe_allow_html=True)

        # Display grand total as a compact static row below
        if len(grand_total_rows) > 0:
            gt_html = '<div style="background:linear-gradient(135deg,#1B2838,#2A3F54);border-radius:0 0 8px 8px;padding:0 0;margin-top:0;line-height:1.2;">'
            gt_html += '<table style="width:100%;color:white;font-weight:600;font-size:12px;border:none;border-collapse:collapse;text-align:left;"><tr>'
            for col in grand_total_rows.columns:
                val = grand_total_rows[col].iloc[0]
                if isinstance(val, (int, float)):
                    if is_int_agg:
                        display_val = f"{int(val):,}"
                    else:
                        display_val = f"{val:,.2f}"
                else:
                    display_val = html.escape(str(val))
                gt_html += f'<td style="padding:4px 16px;border:none;text-align:left;">{display_val}</td>'
            gt_html += '</tr></table></div>'
            st.markdown(gt_html, unsafe_allow_html=True)

        # Download pivot (full data including grand total)
        csv_pivot = fmt_pivot.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')
        st.download_button("Download Pivot CSV", data=csv_pivot,
                           file_name="RSA_Pivot_Export.csv", mime="text/csv", key="dl_pivot")
    except Exception as e:
        st.error(f"Pivot table error: {e}")
else:
    st.info("Select at least one Row or Column dimension to build the pivot table.")

# ============================================================================
# COST ANALYSIS
# ============================================================================
st.markdown('<div class="section-header">💰 Cost Analysis Dashboard</div>', unsafe_allow_html=True)

monthly_cost = filtered_df.groupby(['Year', 'Month'])['Fee (Baht)'].sum().reset_index()
if len(monthly_cost) > 0:
    monthly_cost['Year'] = monthly_cost['Year'].astype(int)
    monthly_cost['Month'] = monthly_cost['Month'].astype(int)
    monthly_cost['Date'] = pd.to_datetime(monthly_cost[['Year', 'Month']].assign(Day=1))
    monthly_cost = monthly_cost.sort_values('Date')

    fig_trend = go.Figure()
    year_colors = ['#4A90D9', '#27AE60', '#F39C12', '#2D5AA0', '#1B2838']
    for i, yr in enumerate(sorted(monthly_cost['Year'].unique())):
        yd = monthly_cost[monthly_cost['Year'] == yr]
        c = year_colors[i % len(year_colors)]
        fig_trend.add_trace(go.Scatter(x=yd['Month'], y=yd['Fee (Baht)'], mode='lines+markers', name=f'{yr}', line=dict(width=3, color=c), marker=dict(size=8, color=c)))

    fig_trend.add_trace(go.Scatter(x=list(range(1, 13)), y=[MONTHLY_BUDGET] * 12, mode='lines', name='Budget', line=dict(color='#E74C3C', width=2, dash='dash')))
    fig_trend.update_layout(
        title={'text': 'Monthly Cost Trend with Budget Comparison', 'font': CHART_TITLE_FONT},
        xaxis_title='Month', yaxis_title='Fee (Baht)', hovermode='x unified', height=450,
        xaxis=dict(tickmode='linear', tick0=1, dtick=1, gridcolor='#E2E8F0', showline=True, linecolor='#E2E8F0'),
        yaxis=dict(gridcolor='#E2E8F0', showline=True, linecolor='#E2E8F0'),
        plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='white',
        font=dict(family='Inter, sans-serif', size=12, color='#4A5568'),
        legend=dict(bgcolor='rgba(255,255,255,0.9)', bordercolor='#E2E8F0', borderwidth=1)
    )
    st.plotly_chart(fig_trend, use_container_width=True, config=PLOTLY_CONFIG)

# ============================================================================
# ADDITIONAL ANALYTICS
# ============================================================================
st.markdown('<div class="section-header">📊 Additional Analytics</div>', unsafe_allow_html=True)

c1, c2 = st.columns(2)
with c1:
    svc_dist = filtered_df['ประเภทการบริการ'].value_counts()
    fig_pie = px.pie(values=svc_dist.values, names=svc_dist.index, title='Service Type Distribution', hole=0.4,
                     color_discrete_sequence=['#4A90D9','#27AE60','#F39C12','#E74C3C','#1B2838','#2D5AA0','#6FB1FF'])
    fig_pie.update_traces(textposition='inside', textinfo='percent+label', textfont_size=11)
    fig_pie.update_layout(height=400, title={'font': CHART_TITLE_FONT}, **{k: v for k, v in CHART_LAYOUT_DEFAULTS.items() if k not in ('xaxis', 'yaxis')},
                          legend=dict(bgcolor='rgba(255,255,255,0.9)', bordercolor='#E2E8F0', borderwidth=1))
    st.plotly_chart(fig_pie, use_container_width=True, config=PLOTLY_CONFIG)

with c2:
    lob_counts = filtered_df['LOB'].value_counts().sort_index()
    fig_lob = px.bar(x=lob_counts.index, y=lob_counts.values, title='Cases by LOB',
                     labels={'x':'LOB','y':'Cases'}, color=lob_counts.values,
                     color_continuous_scale=[[0,'#4A90D9'],[0.5,'#2D5AA0'],[1,'#1B2838']])
    fig_lob.update_layout(height=400, showlegend=False, title={'font': CHART_TITLE_FONT}, **CHART_LAYOUT_DEFAULTS)
    st.plotly_chart(fig_lob, use_container_width=True, config=PLOTLY_CONFIG)

c3, c4 = st.columns(2)
with c3:
    top_vol = filtered_df['ประเภทการบริการ'].value_counts().head(10)
    fig_tv = px.bar(x=top_vol.values, y=top_vol.index, orientation='h', title='Top Services by Volume',
                    labels={'x':'Cases','y':'Service'}, color=top_vol.values,
                    color_continuous_scale=[[0,'#27AE60'],[0.5,'#3D8E56'],[1,'#1E7E34']])
    fig_tv.update_layout(height=400, showlegend=False, yaxis={'categoryorder':'total ascending', 'gridcolor':'#E2E8F0', 'showline':True, 'linecolor':'#E2E8F0'}, title={'font': CHART_TITLE_FONT}, **{k: v for k, v in CHART_LAYOUT_DEFAULTS.items() if k != 'yaxis'})
    st.plotly_chart(fig_tv, use_container_width=True, config=PLOTLY_CONFIG)

with c4:
    top_cost = filtered_df.groupby('ประเภทการบริการ')['Fee (Baht)'].sum().sort_values(ascending=False).head(10)
    fig_tc = px.bar(x=top_cost.values, y=top_cost.index, orientation='h', title='Top Services by Cost',
                    labels={'x':'Fee (Baht)','y':'Service'}, color=top_cost.values,
                    color_continuous_scale=[[0,'#F39C12'],[0.5,'#E74C3C'],[1,'#C0392B']])
    fig_tc.update_layout(height=400, showlegend=False, yaxis={'categoryorder':'total ascending', 'gridcolor':'#E2E8F0', 'showline':True, 'linecolor':'#E2E8F0'}, title={'font': CHART_TITLE_FONT}, **{k: v for k, v in CHART_LAYOUT_DEFAULTS.items() if k != 'yaxis'})
    st.plotly_chart(fig_tc, use_container_width=True, config=PLOTLY_CONFIG)

# ============================================================================
# REGIONAL ANALYSIS
# ============================================================================
if 'จังหวัด' in filtered_df.columns:
    st.markdown('<div class="section-header">🗺️ Regional Analysis</div>', unsafe_allow_html=True)
    c5, c6 = st.columns(2)
    with c5:
        rc = filtered_df['จังหวัด'].value_counts().head(15)
        fig_r = px.bar(x=rc.values, y=rc.index, orientation='h', title='Top 15 Regions by Volume',
                       labels={'x':'Cases','y':'Province'}, color=rc.values,
                       color_continuous_scale=[[0,'#4A90D9'],[0.5,'#2D5AA0'],[1,'#1B2838']])
        fig_r.update_layout(height=500, showlegend=False, yaxis={'categoryorder':'total ascending', 'gridcolor':'#E2E8F0', 'showline':True, 'linecolor':'#E2E8F0'}, title={'font': CHART_TITLE_FONT}, **{k: v for k, v in CHART_LAYOUT_DEFAULTS.items() if k != 'yaxis'})
        st.plotly_chart(fig_r, use_container_width=True, config=PLOTLY_CONFIG)
    with c6:
        rcost = filtered_df.groupby('จังหวัด')['Fee (Baht)'].sum().sort_values(ascending=False).head(15)
        fig_rc = px.bar(x=rcost.values, y=rcost.index, orientation='h', title='Top 15 Regions by Cost',
                        labels={'x':'Fee (Baht)','y':'Province'}, color=rcost.values,
                        color_continuous_scale=[[0,'#F39C12'],[0.5,'#E67E22'],[1,'#D35400']])
        fig_rc.update_layout(height=500, showlegend=False, yaxis={'categoryorder':'total ascending', 'gridcolor':'#E2E8F0', 'showline':True, 'linecolor':'#E2E8F0'}, title={'font': CHART_TITLE_FONT}, **{k: v for k, v in CHART_LAYOUT_DEFAULTS.items() if k != 'yaxis'})
        st.plotly_chart(fig_rc, use_container_width=True, config=PLOTLY_CONFIG)

# ============================================================================
# MONTHLY TREND BY SERVICE TYPE
# ============================================================================
st.markdown('<div class="section-header">📈 Monthly Trend by Service Type</div>', unsafe_allow_html=True)

mst = filtered_df.groupby(['Year', 'Month', 'ประเภทการบริการ']).size().reset_index(name='Count')
if len(mst) > 0:
    mst['Year'] = mst['Year'].astype(int)
    mst['Month'] = mst['Month'].astype(int)
    mst['Date'] = pd.to_datetime(mst[['Year', 'Month']].assign(Day=1))
    fig_mst = px.line(mst, x='Date', y='Count', color='ประเภทการบริการ', title='Monthly Case Volume by Service Type', markers=True)
    fig_mst.update_layout(
        xaxis_title='Date', yaxis_title='Cases', hovermode='x unified', height=450,
        title={'font': CHART_TITLE_FONT}, **CHART_LAYOUT_DEFAULTS,
        legend=dict(orientation="v", yanchor="top", y=1, xanchor="left", x=1.02, bgcolor='rgba(255,255,255,0.9)', bordercolor='#E2E8F0', borderwidth=1)
    )
    st.plotly_chart(fig_mst, use_container_width=True, config=PLOTLY_CONFIG)

# ============================================================================
# DATA EXPORT
# ============================================================================
st.markdown('<div class="section-header">💾 Data Export</div>', unsafe_allow_html=True)
ce1, ce2 = st.columns([3, 1])
with ce1:
    st.write(f"Export filtered data ({len(filtered_df):,} records)")
with ce2:
    csv_data = convert_df_to_csv(filtered_df)
    st.download_button("Download CSV", data=csv_data,
                       file_name=f"RSA_Export_{st.session_state.export_timestamp}.csv",
                       mime="text/csv", use_container_width=True)

# ============================================================================
# FILE UPLOAD (bottom of page)
# ============================================================================
st.markdown("---")
st.markdown('<div class="section-header">📂 Data Source</div>', unsafe_allow_html=True)
fu1, fu2 = st.columns([3, 1])
with fu1:
    uploaded_file = st.file_uploader(
        "Upload RSA Report (.xlsx)",
        type=["xlsx"],
        key="file_uploader",
        help="Upload a new Excel file to replace the current data source."
    )
    if uploaded_file is not None:
        new_bytes = uploaded_file.getvalue()
        # Only rerun if this is a genuinely new file (avoid infinite reload loop)
        if st.session_state.uploaded_file_name != uploaded_file.name or st.session_state.uploaded_file_bytes != new_bytes:
            persist_uploaded_file(uploaded_file)
            st.session_state.uploaded_file_bytes = new_bytes
            st.session_state.uploaded_file_name = uploaded_file.name
            st.cache_data.clear()
            st.rerun()
with fu2:
    st.markdown(f"**Current source:** {data_source_label}")
    persisted_path = os.path.join(UPLOAD_DIR, "persisted_upload.xlsx")
    if os.path.exists(persisted_path):
        if st.button("Clear uploaded file", key="clear_upload"):
            os.remove(persisted_path)
            st.session_state.uploaded_file_bytes = None
            st.session_state.uploaded_file_name = None
            st.cache_data.clear()
            st.rerun()

# Footer
st.markdown(f"<div class='dashboard-footer'><strong>RSA Dashboard</strong> - Sompo Thailand<br>Dashboard rendered: {datetime.now().strftime('%Y-%m-%d %H:%M')}</div>", unsafe_allow_html=True)
