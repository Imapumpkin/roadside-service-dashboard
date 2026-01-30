import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import os
import html
import re
from io import BytesIO

try:
    from streamlit_sortables import sort_items
    SORTABLES_AVAILABLE = True
except ImportError:
    SORTABLES_AVAILABLE = False

# ============================================================================
# CONFIGURATION
# ============================================================================
MONTHLY_BUDGET = 200_000
HEALTH_THRESHOLD_HEALTHY = 5
HEALTH_THRESHOLD_WARNING = 15
CACHE_TTL = 3600
DEFAULT_DATA_FILE = "(Test) RSA Report.xlsx"
UPLOAD_DIR = "uploaded_data"
PASSWORD = "sompo2024"

pd.set_option('future.no_silent_downcasting', True)

st.set_page_config(
    page_title="RSA Dashboard - Sompo Thailand",
    page_icon="\U0001f697",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ============================================================================
# PASSWORD PROTECTION
# ============================================================================
def check_password():
    if "authenticated" not in st.session_state:
        st.session_state.authenticated = False
    if st.session_state.authenticated:
        return True

    st.markdown("""
    <style>
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
            <div style="font-size: 48px; margin-bottom: 16px;">\U0001f697</div>
            <div style="font-size: 28px; font-weight: 700; color: #1B2838; margin-bottom: 8px;">RSA Dashboard</div>
            <div style="font-size: 14px; color: #718096; margin-bottom: 32px;">Sompo Thailand - Roadside Assistance</div>
        </div>
        """, unsafe_allow_html=True)
        def try_login():
            try:
                correct_pw = st.secrets["password"]
            except Exception:
                correct_pw = PASSWORD
            if st.session_state.password_input == correct_pw:
                st.session_state.authenticated = True
            else:
                st.session_state.login_error = True

        password_input = st.text_input("Enter Password", type="password", key="password_input", on_change=try_login)
        if st.button("Login", use_container_width=True, type="primary"):
            try_login()
        if st.session_state.authenticated:
            st.rerun()
        if st.session_state.get("login_error"):
            st.session_state.login_error = False
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
    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #1B2838 0%, #2A3F54 100%);
        padding-top: 2rem;
    }
    [data-testid="stSidebar"] .stMarkdown { color: #FFFFFF; }
    [data-testid="stSidebar"] label {
        color: #E0E6ED !important; font-weight: 500 !important;
        font-size: 13px !important; letter-spacing: 0.03em; margin-top: 0.8rem;
    }
    [data-testid="stSidebar"] .stMultiSelect [data-baseweb="select"] {
        background-color: rgba(255, 255, 255, 0.1);
        border: 1px solid rgba(255, 255, 255, 0.2); border-radius: 6px;
    }
    [data-testid="stSidebar"] h2 {
        color: #FFFFFF !important; font-size: 18px !important;
        font-weight: 600 !important; margin-bottom: 1.5rem;
    }
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
    .empty-state { background: white; border-radius: 12px; padding: 48px 24px; text-align: center; box-shadow: 0 2px 8px rgba(0,0,0,0.08); margin: 24px 0; }
    .empty-state-icon { font-size: 64px; margin-bottom: 16px; opacity: 0.5; }
    .empty-state-title { font-size: 20px; font-weight: 600; color: #1B2838; margin-bottom: 8px; }
    .empty-state-message { font-size: 14px; color: #718096; }
    h1 { color: #1B2838 !important; font-weight: 700 !important; font-size: 32px !important; margin-bottom: 8px !important; }
    h3 { color: #4A5568 !important; font-weight: 500 !important; font-size: 16px !important; margin-bottom: 32px !important; }
    .js-plotly-plot { border-radius: 12px; box-shadow: 0 2px 8px rgba(0,0,0,0.06); background: white; padding: 12px; }
    .stDownloadButton button {
        background: linear-gradient(135deg, #4A90D9 0%, #2D5AA0 100%);
        color: white; border: none; border-radius: 8px; padding: 12px 24px; font-weight: 600;
    }
    .stDownloadButton button:hover { background: linear-gradient(135deg, #2D5AA0 0%, #1B2838 100%); box-shadow: 0 4px 12px rgba(74,144,217,0.3); }
    .dashboard-footer { text-align: center; color: #718096; padding: 32px 20px; margin-top: 48px; border-top: 2px solid #E2E8F0; font-size: 13px; background: white; border-radius: 12px; }
    .stExpander { border: none !important; box-shadow: none !important; background: transparent !important; }
    .stExpander > details { border: 1px solid #E2E8F0 !important; border-radius: 8px !important; background: white !important; box-shadow: 0 2px 8px rgba(0,0,0,0.06) !important; }
    .stExpander > details > summary { padding: 12px 16px !important; font-weight: 600 !important; color: #1B2838 !important; font-size: 14px !important; }
    .stExpander > details[open] > summary { border-bottom: 1px solid #E2E8F0 !important; }
    .stExpander > details > div { padding: 16px !important; }
    @media (max-width: 768px) {
        .metric-card { margin-bottom: 12px; }
        .metric-value { font-size: 22px; }
        .section-header { font-size: 18px; margin-top: 32px; }
        .health-indicator .health-stats { grid-template-columns: 1fr 1fr; }
    }
    @media (max-width: 1200px) {
        div[style*="grid-template-columns:repeat(5"] { grid-template-columns: repeat(3, 1fr) !important; }
        div[style*="grid-template-columns:repeat(6"] { grid-template-columns: repeat(3, 1fr) !important; }
    }
    @media (max-width: 768px) {
        div[style*="grid-template-columns:repeat(5"] { grid-template-columns: repeat(2, 1fr) !important; }
        div[style*="grid-template-columns:repeat(6"] { grid-template-columns: repeat(2, 1fr) !important; }
    }
</style>
""", unsafe_allow_html=True)


# ============================================================================
# DATA LOADING
# ============================================================================
@st.cache_data(ttl=CACHE_TTL)
def load_and_process(file_bytes=None, file_path=None):
    """Load from bytes or path and process in one cached step."""
    if file_bytes is not None:
        df_raw = pd.read_excel(BytesIO(file_bytes), header=None)
    elif file_path and os.path.exists(file_path):
        df_raw = pd.read_excel(file_path, header=None)
    else:
        return None

    df = df_raw.copy()
    # Find header row
    for idx in range(min(len(df), 30)):
        if 'Policy No.' in df.iloc[idx].values:
            df.columns = df.iloc[idx]
            df = df.iloc[idx + 1:].reset_index(drop=True)
            break
    else:
        raise ValueError("Could not find header row containing 'Policy No.'")

    # Clean Fee column name
    fee_cols = [c for c in df.columns if isinstance(c, str) and 'Fee' in c and 'Exceed' not in c]
    if fee_cols:
        df = df.rename(columns={fee_cols[0]: 'Fee (Baht)'})

    # Process dates
    if '\u0e27\u0e31\u0e19\u0e17\u0e35\u0e48' in df.columns:
        if df['\u0e27\u0e31\u0e19\u0e17\u0e35\u0e48'].dtype == 'object':
            df['\u0e27\u0e31\u0e19\u0e17\u0e35\u0e48'] = pd.to_datetime(df['\u0e27\u0e31\u0e19\u0e17\u0e35\u0e48'], format='%d/%m/%Y', errors='coerce')
        else:
            df['\u0e27\u0e31\u0e19\u0e17\u0e35\u0e48'] = pd.to_datetime(df['\u0e27\u0e31\u0e19\u0e17\u0e35\u0e48'], errors='coerce')
        df = df.dropna(subset=['\u0e27\u0e31\u0e19\u0e17\u0e35\u0e48'])
        df['Day'] = df['\u0e27\u0e31\u0e19\u0e17\u0e35\u0e48'].dt.day.astype(int)
        df['Month'] = df['\u0e27\u0e31\u0e19\u0e17\u0e35\u0e48'].dt.month.astype(int)
        df['Year'] = df['\u0e27\u0e31\u0e19\u0e17\u0e35\u0e48'].dt.year.astype(int)
        df['\u0e27\u0e31\u0e19\u0e17\u0e35\u0e48'] = df['\u0e27\u0e31\u0e19\u0e17\u0e35\u0e48'].dt.date

    df = df.dropna(how='all').reset_index(drop=True)
    df = df.replace('-', pd.NA)

    # LOB
    if 'Policy No.' in df.columns:
        df['Policy Type'] = df['Policy No.'].str.extract(r'(A[CV]\d)', expand=False)
    df['LOB'] = df['Policy Type'].fillna('Unverify') if 'Policy Type' in df.columns else 'Unre'

    # Province extraction
    if '\u0e17\u0e30\u0e40\u0e1a\u0e35\u0e22\u0e19\u0e23\u0e16' in df.columns:
        if '\u0e08\u0e31\u0e07\u0e2b\u0e27\u0e31\u0e14 \u0e17\u0e30\u0e40\u0e1a\u0e35\u0e22\u0e19\u0e23\u0e16' in df.columns:
            df['\u0e08\u0e31\u0e07\u0e2b\u0e27\u0e31\u0e14 \u0e17\u0e30\u0e40\u0e1a\u0e35\u0e22\u0e19\u0e23\u0e16'] = df['\u0e08\u0e31\u0e07\u0e2b\u0e27\u0e31\u0e14 \u0e17\u0e30\u0e40\u0e1a\u0e35\u0e22\u0e19\u0e23\u0e16'].astype(str).str.replace(r'[.!@#$%^&*\d]', '', regex=True).str.strip()
            df['\u0e08\u0e31\u0e07\u0e2b\u0e27\u0e31\u0e14 \u0e17\u0e30\u0e40\u0e1a\u0e35\u0e22\u0e19\u0e23\u0e16'] = df['\u0e08\u0e31\u0e07\u0e2b\u0e27\u0e31\u0e14 \u0e17\u0e30\u0e40\u0e1a\u0e35\u0e22\u0e19\u0e23\u0e16'].replace(['', 'nan', 'None', '<NA>'], pd.NA)
        plate_province = df['\u0e17\u0e30\u0e40\u0e1a\u0e35\u0e22\u0e19\u0e23\u0e16'].astype(str).str.extract(r'(\d)[.\s]*([ก-๙]+)\s*[.!@#$%^&*]*\s*$', expand=True)
        extracted = plate_province[1]
        if '\u0e08\u0e31\u0e07\u0e2b\u0e27\u0e31\u0e14 \u0e17\u0e30\u0e40\u0e1a\u0e35\u0e22\u0e19\u0e23\u0e16' in df.columns:
            df['\u0e08\u0e31\u0e07\u0e2b\u0e27\u0e31\u0e14 \u0e17\u0e30\u0e40\u0e1a\u0e35\u0e22\u0e19\u0e23\u0e16'] = df['\u0e08\u0e31\u0e07\u0e2b\u0e27\u0e31\u0e14 \u0e17\u0e30\u0e40\u0e1a\u0e35\u0e22\u0e19\u0e23\u0e16'].fillna(extracted)
        else:
            df['\u0e08\u0e31\u0e07\u0e2b\u0e27\u0e31\u0e14 \u0e17\u0e30\u0e40\u0e1a\u0e35\u0e22\u0e19\u0e23\u0e16'] = extracted
        df['\u0e08\u0e31\u0e07\u0e2b\u0e27\u0e31\u0e14 \u0e17\u0e30\u0e40\u0e1a\u0e35\u0e22\u0e19\u0e23\u0e16'] = df['\u0e08\u0e31\u0e07\u0e2b\u0e27\u0e31\u0e14 \u0e17\u0e30\u0e40\u0e1a\u0e35\u0e22\u0e19\u0e23\u0e16'].replace(['\u0e01\u0e23\u0e38\u0e07\u0e40\u0e17\u0e1e', '\u0e01\u0e17\u0e21'], '\u0e01\u0e23\u0e38\u0e07\u0e40\u0e17\u0e1e\u0e21\u0e2b\u0e32\u0e19\u0e04\u0e23')

    if 'Fee (Baht)' in df.columns:
        df['Fee (Baht)'] = pd.to_numeric(df['Fee (Baht)'], errors='coerce')

    if '\u0e22\u0e35\u0e48\u0e2b\u0e49\u0e2d\u0e23\u0e16' in df.columns:
        df['\u0e22\u0e35\u0e48\u0e2b\u0e49\u0e2d\u0e23\u0e16'] = df['\u0e22\u0e35\u0e48\u0e2b\u0e49\u0e2d\u0e23\u0e16'].astype(str).str.upper().replace('<NA>', pd.NA).replace('NAN', pd.NA)

    # Convert key filter columns to category for faster isin()
    for col in ['\u0e1b\u0e23\u0e30\u0e40\u0e20\u0e17\u0e01\u0e32\u0e23\u0e1a\u0e23\u0e34\u0e01\u0e32\u0e23', 'LOB']:
        if col in df.columns:
            df[col] = df[col].astype('category')

    return df


def load_persisted_upload():
    os.makedirs(UPLOAD_DIR, exist_ok=True)
    persisted_path = os.path.join(UPLOAD_DIR, "persisted_upload.xlsx")
    if os.path.exists(persisted_path):
        return persisted_path
    return None


def persist_uploaded_file(uploaded_file):
    os.makedirs(UPLOAD_DIR, exist_ok=True)
    persisted_path = os.path.join(UPLOAD_DIR, "persisted_upload.xlsx")
    with open(persisted_path, "wb") as f:
        f.write(uploaded_file.getbuffer())
    return persisted_path


# ============================================================================
# DATA SOURCE SELECTION
# ============================================================================
if 'uploaded_file_bytes' not in st.session_state:
    st.session_state.uploaded_file_bytes = None
    st.session_state.uploaded_file_name = None

df = None
data_source_label = ""

if st.session_state.uploaded_file_bytes is not None:
    try:
        df = load_and_process(file_bytes=st.session_state.uploaded_file_bytes)
    except Exception:
        df = None
    if df is None:
        st.session_state.uploaded_file_bytes = None
        st.session_state.uploaded_file_name = None
        st.cache_data.clear()
    else:
        data_source_label = f"Uploaded: {st.session_state.uploaded_file_name}"

if df is None and st.session_state.uploaded_file_bytes is None:
    persisted_path = load_persisted_upload()
    if persisted_path:
        try:
            df = load_and_process(file_path=persisted_path)
        except Exception:
            df = None
        if df is None:
            # Bad persisted file — auto-remove it
            os.remove(persisted_path)
        else:
            data_source_label = "Previously uploaded file"
    if df is None:
        df = load_and_process(file_path=DEFAULT_DATA_FILE)
        if df is not None:
            data_source_label = DEFAULT_DATA_FILE

if df is None:
    st.markdown("# \U0001f697 RSA Dashboard - Sompo Thailand")
    st.markdown("""
    <div class="empty-state">
        <div class="empty-state-icon">\U0001f4c2</div>
        <div class="empty-state-title">No Data Available</div>
        <div class="empty-state-message">Please upload an RSA Report Excel file or place the default file in the application directory.</div>
    </div>
    """, unsafe_allow_html=True)
    st.stop()

# Validate required columns
required_cols = ['Year', 'Month', 'Fee (Baht)', '\u0e1b\u0e23\u0e30\u0e40\u0e20\u0e17\u0e01\u0e32\u0e23\u0e1a\u0e23\u0e34\u0e01\u0e32\u0e23', 'LOB']
missing = [c for c in required_cols if c not in df.columns]
if missing:
    st.error(f"Missing required columns after processing: {missing}")
    st.stop()

if 'export_timestamp' not in st.session_state:
    st.session_state.export_timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')


# ============================================================================
# HELPER FUNCTIONS
# ============================================================================
def safe_sorted_unique(series):
    return sorted([str(v) for v in series.dropna().unique()])

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
def convert_df_to_csv(_df_csv):
    return _df_csv.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')


# ============================================================================
# FILTERS - Using multiselect (much faster than individual checkboxes)
# ============================================================================
available_years = sorted([int(y) for y in df['Year'].dropna().unique()])
available_services = safe_sorted_unique(df['\u0e1b\u0e23\u0e30\u0e40\u0e20\u0e17\u0e01\u0e32\u0e23\u0e1a\u0e23\u0e34\u0e01\u0e32\u0e23'])
available_lobs = safe_sorted_unique(df['LOB'])
available_months = sorted([int(m) for m in df['Month'].dropna().unique()])
month_names = {1:'Jan',2:'Feb',3:'Mar',4:'Apr',5:'May',6:'Jun',7:'Jul',8:'Aug',9:'Sep',10:'Oct',11:'Nov',12:'Dec'}
available_channels = safe_sorted_unique(df['\u0e23\u0e2b\u0e31\u0e2a\u0e42\u0e04\u0e23\u0e07\u0e01\u0e32\u0e23']) if '\u0e23\u0e2b\u0e31\u0e2a\u0e42\u0e04\u0e23\u0e07\u0e01\u0e32\u0e23' in df.columns else []
available_regions = safe_sorted_unique(df['\u0e08\u0e31\u0e07\u0e2b\u0e27\u0e31\u0e14']) if '\u0e08\u0e31\u0e07\u0e2b\u0e27\u0e31\u0e14' in df.columns else []
available_makes = safe_sorted_unique(df['\u0e22\u0e35\u0e48\u0e2b\u0e49\u0e2d\u0e23\u0e16']) if '\u0e22\u0e35\u0e48\u0e2b\u0e49\u0e2d\u0e23\u0e16' in df.columns else []
available_models = safe_sorted_unique(df['\u0e23\u0e38\u0e48\u0e19\u0e23\u0e16']) if '\u0e23\u0e38\u0e48\u0e19\u0e23\u0e16' in df.columns else []

# ============================================================================
# DASHBOARD HEADER
# ============================================================================
latest_date = df['\u0e27\u0e31\u0e19\u0e17\u0e35\u0e48'].max() if '\u0e27\u0e31\u0e19\u0e17\u0e35\u0e48' in df.columns else "N/A"
st.markdown("# \U0001f697 RSA Dashboard - Sompo Thailand")
st.markdown(f"### Roadside Assistance Monitoring <span class='data-freshness'>Data through: {latest_date} | Source: {data_source_label}</span>", unsafe_allow_html=True)

# ============================================================================
# KPIs (computed before filters so we always show current year KPIs)
# ============================================================================
st.markdown('<div class="section-header">\U0001f4c8 Key Performance Indicators</div>', unsafe_allow_html=True)

current_year = max(available_years)
prev_year = current_year - 1
cur_df_all = df[df['Year'] == current_year]
prev_df_all = df[df['Year'] == prev_year]
current_month = datetime.now().month

cur_fee = cur_df_all['Fee (Baht)'].sum()
ytd_cases = len(cur_df_all)
ytd_fee = cur_fee
prev_ytd_cases = len(prev_df_all[prev_df_all['Month'] <= current_month])
prev_ytd_fee = prev_df_all[prev_df_all['Month'] <= current_month]['Fee (Baht)'].sum()
mtd_fee = cur_df_all[cur_df_all['Month'] == current_month]['Fee (Baht)'].sum()
mtd_util = (mtd_fee / MONTHLY_BUDGET * 100) if MONTHLY_BUDGET > 0 else 0
prev_mtd_fee = prev_df_all[prev_df_all['Month'] == current_month]['Fee (Baht)'].sum()
cur_avg_raw = cur_df_all['Fee (Baht)'].mean()
cur_avg = 0.0 if pd.isna(cur_avg_raw) else cur_avg_raw
prev_avg_raw = prev_df_all['Fee (Baht)'].mean()
prev_avg = 0.0 if pd.isna(prev_avg_raw) else prev_avg_raw

def yoy_html(cur_val, prev_val, compare_year):
    if prev_val == 0:
        return f'<div style="font-size:11px;margin-top:8px;opacity:0.8;">vs {compare_year}: N/A</div>'
    pct = (cur_val - prev_val) / prev_val * 100
    cls = "negative" if pct > 0 else "positive"
    arrow = "\u25b2" if pct >= 0 else "\u25bc"
    return f'<div style="font-size:11px;margin-top:8px;opacity:0.9;"><span class="{cls}">{arrow} {abs(pct):.1f}%</span> vs {compare_year}</div>'

mc = "negative" if mtd_util > 100 else "positive"
kpi_cards = [
    (f"YTD Total Cases ({current_year})", f"{ytd_cases:,}", yoy_html(ytd_cases, prev_ytd_cases, prev_year)),
    (f"YTD Total Fee ({current_year})", f"\u0e3f{ytd_fee:,.0f}", yoy_html(ytd_fee, prev_ytd_fee, prev_year)),
    (f"Avg Fee/Case ({current_year})", f"\u0e3f{cur_avg:,.0f}", yoy_html(cur_avg, prev_avg, prev_year)),
    (f"MTD Fee ({current_year})", f"\u0e3f{mtd_fee:,.0f}", yoy_html(mtd_fee, prev_mtd_fee, prev_year)),
    ("MTD Utilization", f'<span class="{mc}">{mtd_util:.1f}%</span>', f'<div style="font-size:11px;margin-top:8px;opacity:0.8;">\u0e3f{mtd_fee:,.0f} / \u0e3f{MONTHLY_BUDGET:,}</div>'),
]
kpi_html = '<div style="display:grid;grid-template-columns:repeat(5,1fr);gap:16px;">'
for title, value, extra in kpi_cards:
    kpi_html += f'<div class="metric-card"><div class="metric-title">{title}</div><div class="metric-value">{value}</div>{extra}</div>'
kpi_html += '</div>'
st.markdown(kpi_html, unsafe_allow_html=True)

# ============================================================================
# PORTFOLIO HEALTH
# ============================================================================
st.markdown('<div class="section-header">\U0001f3e5 Portfolio Health Indicator</div>', unsafe_allow_html=True)

months_in_year = cur_df_all['Month'].nunique() if len(cur_df_all) > 0 else 1
run_rate = cur_fee / max(months_in_year, 1)
projection = run_rate * 12
annual_budget = MONTHLY_BUDGET * 12
expected_cost_ytd = MONTHLY_BUDGET * months_in_year
over_budget_pct = ((ytd_fee - expected_cost_ytd) / expected_cost_ytd * 100) if expected_cost_ytd > 0 else 0

if over_budget_pct <= HEALTH_THRESHOLD_HEALTHY:
    h_status, h_class, h_badge = "HEALTHY", "health-healthy", '<span class="health-badge badge-healthy">Healthy</span>'
elif over_budget_pct <= HEALTH_THRESHOLD_WARNING:
    h_status, h_class, h_badge = "WARNING", "health-warning", '<span class="health-badge badge-warning">Warning</span>'
else:
    h_status, h_class, h_badge = "CRITICAL", "health-critical", '<span class="health-badge badge-critical">Critical</span>'

ytd_vs_expected_pct = (ytd_fee / expected_cost_ytd * 100) if expected_cost_ytd > 0 else 0
projection_vs_budget_pct = (projection / annual_budget * 100) if annual_budget > 0 else 0

st.markdown(f"""
<div class="health-indicator {h_class}">
    <div class="health-title">{h_badge} Portfolio Status: {h_status}</div>
    <div class="health-stats">
        <div class="health-stat-item"><div class="health-stat-label">YTD Total Fee vs Expected Cost</div><div class="health-stat-value">\u0e3f{ytd_fee:,.0f} / \u0e3f{expected_cost_ytd:,.0f} ({ytd_vs_expected_pct:.1f}%)</div></div>
        <div class="health-stat-item"><div class="health-stat-label">Monthly Run Rate</div><div class="health-stat-value">\u0e3f{run_rate:,.0f}</div></div>
        <div class="health-stat-item"><div class="health-stat-label">Year-End Projection vs Annual Budget</div><div class="health-stat-value">\u0e3f{projection:,.0f} / \u0e3f{annual_budget:,.0f} ({projection_vs_budget_pct:.1f}%)</div></div>
    </div>
</div>
""", unsafe_allow_html=True)

# ============================================================================
# SERVICE UTILIZATION - Interactive Pivot Table (unified controls)
# ============================================================================
st.markdown('<div class="section-header">\U0001f4cb Service Utilization \u2013 Interactive Pivot Table</div>', unsafe_allow_html=True)

pivot_cols_available = [c for c in ['LOB', '\u0e1b\u0e23\u0e30\u0e40\u0e20\u0e17\u0e01\u0e32\u0e23\u0e1a\u0e23\u0e34\u0e01\u0e32\u0e23', 'Year', 'Month', '\u0e08\u0e31\u0e07\u0e2b\u0e27\u0e31\u0e14', '\u0e22\u0e35\u0e48\u0e2b\u0e49\u0e2d\u0e23\u0e16', '\u0e23\u0e38\u0e48\u0e19\u0e23\u0e16', 'Policy Type', '\u0e23\u0e2b\u0e31\u0e2a\u0e42\u0e04\u0e23\u0e07\u0e01\u0e32\u0e23', '\u0e41\u0e1c\u0e19\u0e01'] if c in df.columns]
value_cols_available = [c for c in ['Fee (Baht)', '\u0e25\u0e39\u0e01\u0e04\u0e49\u0e32\u0e08\u0e48\u0e32\u0e22\u0e2a\u0e48\u0e27\u0e19\u0e15\u0e48\u0e32\u0e07'] if c in df.columns]

with st.container(border=True):
    # Row 1: Data Filter + Columns
    r1c1, r1c2 = st.columns(2)
    with r1c1:
        with st.expander("Data Filter", expanded=False):
            with st.expander("Year", expanded=False):
                selected_years = st.multiselect("Year", available_years, default=available_years, key="sel_years", label_visibility="collapsed")
            with st.expander("Month", expanded=False):
                selected_months = st.multiselect("Month", available_months, default=available_months, key="sel_months", label_visibility="collapsed",
                                                 format_func=lambda m: f"{m} - {month_names.get(m,'')}")
            with st.expander("Service Type", expanded=False):
                selected_services = st.multiselect("Service Type", available_services, default=available_services, key="sel_services", label_visibility="collapsed")
            with st.expander("LOB", expanded=False):
                selected_lobs = st.multiselect("LOB", available_lobs, default=available_lobs, key="sel_lobs", label_visibility="collapsed")
            with st.expander("Channel", expanded=False):
                selected_channels = st.multiselect("Channel", available_channels, default=available_channels, key="sel_channels", label_visibility="collapsed")
            with st.expander("Region", expanded=False):
                selected_regions = st.multiselect("Region", available_regions, default=available_regions, key="sel_regions", label_visibility="collapsed")
            with st.expander("Vehicle Make", expanded=False):
                selected_makes = st.multiselect("Vehicle Make", available_makes, default=available_makes, key="sel_makes", label_visibility="collapsed")
            with st.expander("Vehicle Model", expanded=False):
                selected_models = st.multiselect("Vehicle Model", available_models, default=available_models, key="sel_models", label_visibility="collapsed")
    with r1c2:
        with st.expander("Columns", expanded=False):
            pivot_columns = st.multiselect("Select column fields", options=pivot_cols_available, default=['Year'], key="pivot_columns", label_visibility="collapsed")

    # Row 2: Rows + Values
    r2c1, r2c2 = st.columns(2)
    with r2c1:
        with st.expander("Rows", expanded=False):
            pivot_rows_selected = st.multiselect("Select row fields", options=pivot_cols_available, default=['LOB'], key="pivot_rows", label_visibility="collapsed")
            # Drag-to-reorder inside Rows expander
            pivot_rows = list(pivot_rows_selected) if pivot_rows_selected else []
            if len(pivot_rows_selected) > 1 and SORTABLES_AVAILABLE:
                prev_order = st.session_state.get('_pivot_row_order', [])
                ordered = [x for x in prev_order if x in pivot_rows_selected]
                for x in pivot_rows_selected:
                    if x not in ordered:
                        ordered.append(x)
                st.session_state['_pivot_row_order'] = ordered
                st.caption("Drag to reorder row fields")
                sort_key = "pivot_row_sort_" + "_".join(sorted(ordered))
                pivot_rows = sort_items(ordered, direction="horizontal", key=sort_key)
                st.session_state['_pivot_row_order'] = pivot_rows
            elif len(pivot_rows_selected) > 1:
                st.caption("Row order: " + " \u2192 ".join(pivot_rows_selected))
    with r2c2:
        with st.expander("Values", expanded=False):
            pivot_value = st.selectbox("Select value", options=['Case Count'] + value_cols_available, index=0, key="pivot_value", label_visibility="collapsed")

    # Row 3: Aggregation
    r3c1, r3c2 = st.columns(2)
    with r3c1:
        with st.expander("Aggregation", expanded=False):
            pivot_agg = st.selectbox("Select aggregation", options=['Count', 'Sum', 'Mean', 'Median', 'Min', 'Max', '% of Row Total', '% of Column Total', '% of Grand Total'], index=0, key="pivot_agg", label_visibility="collapsed")

if not selected_years:
    st.warning("Please select at least one year.")
    st.stop()

# ============================================================================
# APPLY FILTERS
# ============================================================================
mask = df['Year'].isin(selected_years)
if len(selected_services) < len(available_services):
    mask &= df['\u0e1b\u0e23\u0e30\u0e40\u0e20\u0e17\u0e01\u0e32\u0e23\u0e1a\u0e23\u0e34\u0e01\u0e32\u0e23'].isin(selected_services)
if len(selected_lobs) < len(available_lobs):
    mask &= df['LOB'].isin(selected_lobs)
if len(selected_months) < len(available_months):
    mask &= df['Month'].isin(selected_months)
if len(selected_channels) < len(available_channels) and '\u0e23\u0e2b\u0e31\u0e2a\u0e42\u0e04\u0e23\u0e07\u0e01\u0e32\u0e23' in df.columns:
    mask &= df['\u0e23\u0e2b\u0e31\u0e2a\u0e42\u0e04\u0e23\u0e07\u0e01\u0e32\u0e23'].astype(str).isin(selected_channels)
if len(selected_regions) < len(available_regions) and '\u0e08\u0e31\u0e07\u0e2b\u0e27\u0e31\u0e14' in df.columns:
    mask &= df['\u0e08\u0e31\u0e07\u0e2b\u0e27\u0e31\u0e14'].astype(str).isin(selected_regions)
if len(selected_makes) < len(available_makes) and '\u0e22\u0e35\u0e48\u0e2b\u0e49\u0e2d\u0e23\u0e16' in df.columns:
    mask &= df['\u0e22\u0e35\u0e48\u0e2b\u0e49\u0e2d\u0e23\u0e16'].astype(str).isin(selected_makes)
if len(selected_models) < len(available_models) and '\u0e23\u0e38\u0e48\u0e19\u0e23\u0e16' in df.columns:
    mask &= df['\u0e23\u0e38\u0e48\u0e19\u0e23\u0e16'].astype(str).isin(selected_models)

filtered_df = df[mask]

if len(filtered_df) == 0:
    st.markdown("""
    <div class="empty-state">
        <div class="empty-state-icon">\U0001f50d</div>
        <div class="empty-state-title">No Data Found</div>
        <div class="empty-state-message">Current filter selection returned no results. Please adjust your filters.</div>
    </div>
    """, unsafe_allow_html=True)
    st.stop()

# ============================================================================
# PIVOT TABLE RENDERING
# ============================================================================
if pivot_rows or pivot_columns:
    try:
        is_pct_agg = pivot_agg in ('% of Row Total', '% of Column Total', '% of Grand Total')
        base_agg = 'Sum' if is_pct_agg else pivot_agg
        agg_map = {'Count': 'count', 'Sum': 'sum', 'Mean': 'mean', 'Median': 'median', 'Min': 'min', 'Max': 'max'}
        agg_func = agg_map[base_agg]

        if pivot_value == 'Case Count':
            _pivot_src = filtered_df.assign(_count=1)
            val_col = '_count'
            if base_agg == 'Count':
                agg_func = 'sum'
        else:
            _pivot_src = filtered_df
            val_col = pivot_value

        all_pivot_fields = list(set(pivot_rows + pivot_columns))
        # Only convert needed columns
        convert_needed = {col: _pivot_src[col].astype(str).fillna('(blank)') for col in all_pivot_fields}
        if convert_needed:
            _pivot_src = _pivot_src.assign(**convert_needed)

        if val_col != '_count':
            _pivot_src = _pivot_src.assign(**{val_col: pd.to_numeric(_pivot_src[val_col], errors='coerce').fillna(0)})

        pivot_kwargs = dict(data=_pivot_src, values=val_col, aggfunc=agg_func, fill_value=0, margins=True, margins_name='Grand Total')
        if pivot_rows:
            pivot_kwargs['index'] = pivot_rows
        if pivot_columns:
            pivot_kwargs['columns'] = pivot_columns

        pivot_result = pd.pivot_table(**pivot_kwargs)

        # Flatten multi-level columns
        if isinstance(pivot_result.columns, pd.MultiIndex):
            pivot_result.columns = [
                'Grand Total' if 'Grand Total' in (parts := [str(c) for c in col]) else ' | '.join(parts).strip(' | ')
                for col in pivot_result.columns
            ]

        # Sort columns
        cols = list(pivot_result.columns)
        gt_col = next((c for c in cols if 'Grand Total' in str(c)), None)
        non_gt_cols = [c for c in cols if c != gt_col]

        def sort_key(x):
            x_str = str(x)
            if x_str.isdigit():
                return (0, int(x_str))
            m = re.match(r'^(\d+)', x_str)
            return (0, int(m.group(1))) if m else (1, x_str)

        try:
            sorted_cols = sorted(non_gt_cols, key=sort_key)
        except Exception:
            sorted_cols = non_gt_cols
        if gt_col:
            sorted_cols.append(gt_col)
        pivot_result = pivot_result[sorted_cols]

        # Reset index
        if isinstance(pivot_result.index, pd.MultiIndex) or pivot_result.index.name:
            pivot_result = pivot_result.reset_index()

        # Apply percentage conversion if needed
        if is_pct_agg:
            row_id_cols_pre = [c for c in pivot_result.columns if c in pivot_rows]
            num_cols_pre = [c for c in pivot_result.columns if c not in row_id_cols_pre]
            gt_col_name = next((c for c in num_cols_pre if 'Grand Total' in str(c)), None)

            # Exclude Grand Total row for percentage base
            gt_mask_pre = pd.Series(False, index=pivot_result.index)
            for rid_col in row_id_cols_pre:
                gt_mask_pre |= pivot_result[rid_col].astype(str).str.contains('Grand Total', na=False)

            if pivot_agg == '% of Row Total' and gt_col_name:
                for idx in pivot_result.index:
                    row_total = pivot_result.loc[idx, gt_col_name]
                    if row_total != 0:
                        for c in num_cols_pre:
                            if c != gt_col_name:
                                pivot_result.loc[idx, c] = pivot_result.loc[idx, c] / row_total * 100
                        pivot_result.loc[idx, gt_col_name] = 100.0
                    else:
                        for c in num_cols_pre:
                            pivot_result.loc[idx, c] = 0.0
            elif pivot_agg == '% of Column Total':
                for c in num_cols_pre:
                    col_total = pivot_result.loc[gt_mask_pre, c].iloc[0] if gt_mask_pre.any() else pivot_result[c].sum()
                    if col_total != 0:
                        pivot_result[c] = pivot_result[c] / col_total * 100
                    else:
                        pivot_result[c] = 0.0
            elif pivot_agg == '% of Grand Total':
                grand_total_val = None
                if gt_col_name and gt_mask_pre.any():
                    grand_total_val = pivot_result.loc[gt_mask_pre, gt_col_name].iloc[0]
                else:
                    grand_total_val = pivot_result[num_cols_pre].values[~gt_mask_pre.values].sum()
                if grand_total_val and grand_total_val != 0:
                    for c in num_cols_pre:
                        pivot_result[c] = pivot_result[c] / grand_total_val * 100

        fmt_pivot = pivot_result
        is_int_agg = pivot_agg in ('Sum', 'Count')
        if is_int_agg:
            numeric_cols = fmt_pivot.select_dtypes(include=['float64', 'float32']).columns
            for col in numeric_cols:
                fmt_pivot[col] = fmt_pivot[col].astype(int)

        # Separate grand total row
        row_id_cols = [c for c in fmt_pivot.columns if c in pivot_rows]
        gt_mask = pd.Series(False, index=fmt_pivot.index)
        for rid_col in row_id_cols:
            gt_mask |= fmt_pivot[rid_col].astype(str).str.contains('Grand Total', na=False)

        data_rows = fmt_pivot[~gt_mask]
        grand_total_rows = fmt_pivot[gt_mask]

        # Compact sort controls inline
        sortable_cols = list(data_rows.columns)
        sc1, sc2, sc3 = st.columns([2, 1, 4])
        with sc1:
            sort_col = st.selectbox("Sort by", options=["(default)"] + sortable_cols, index=0, key="pivot_sort_col")
        with sc2:
            sort_order = st.selectbox("Order", options=["Descending", "Ascending"], index=0, key="pivot_sort_order")
        if sort_col != "(default)" and sort_col in data_rows.columns:
            data_rows = data_rows.sort_values(by=sort_col, ascending=(sort_order == "Ascending"), na_position='last').reset_index(drop=True)

        num_cols_list = list(data_rows.select_dtypes(include=['int64', 'int32', 'float64', 'float32']).columns)
        data_bar_cols = [c for c in num_cols_list if 'Grand Total' not in str(c)]
        global_max = max((data_rows[nc].max() for nc in data_bar_cols), default=1) if data_bar_cols else 1
        if global_max < 1:
            global_max = 1

        if len(grand_total_rows) == 0 and len(data_rows) > 0 and num_cols_list:
            gt_row_data = {}
            for col in data_rows.columns:
                if col in num_cols_list:
                    gt_row_data[col] = data_rows[col].sum()
                elif col in row_id_cols:
                    gt_row_data[col] = "Grand Total" if data_rows.columns.get_loc(col) == 0 else ""
                else:
                    gt_row_data[col] = ""
            grand_total_rows = pd.DataFrame([gt_row_data])

        if len(data_rows) > 0:
            # Build HTML efficiently with list join
            parts = ['<div class="service-table-container" style="max-height:500px;overflow-y:auto;"><table class="service-table"><thead><tr>']
            for col in data_rows.columns:
                parts.append(f'<th>{html.escape(str(col))}</th>')
            parts.append('</tr></thead><tbody>')

            for _, row in data_rows.iterrows():
                parts.append('<tr>')
                for col in data_rows.columns:
                    val = row[col]
                    if col in num_cols_list:
                        cell_text = f"{val:.1f}%" if is_pct_agg else (f"{int(val):,}" if is_int_agg else f"{val:,.2f}")
                        if col in data_bar_cols:
                            bar_pct = (val / global_max * 85) if global_max > 0 else 0
                            parts.append(f'<td style="position:relative;padding:0;"><div style="position:absolute;top:4px;left:4px;bottom:4px;width:{bar_pct:.1f}%;background:linear-gradient(90deg,rgba(74,144,217,0.35),rgba(111,177,255,0.2));z-index:1;border-radius:3px;"></div><div style="position:relative;z-index:2;padding:8px 12px;">{cell_text}</div></td>')
                        else:
                            parts.append(f'<td style="padding:8px 12px;font-weight:600;">{cell_text}</td>')
                    else:
                        parts.append(f'<td style="padding:8px 12px;">{html.escape(str(val))}</td>')
                parts.append('</tr>')

            if len(grand_total_rows) > 0:
                parts.append('<tr>')
                for col in data_rows.columns:
                    if col in grand_total_rows.columns:
                        val = grand_total_rows[col].iloc[0]
                        if isinstance(val, (int, float)) and col in num_cols_list:
                            display_val = f"{val:.1f}%" if is_pct_agg else (f"{int(val):,}" if is_int_agg else f"{val:,.2f}")
                        else:
                            display_val = html.escape(str(val)) if val else ""
                    else:
                        display_val = ""
                    parts.append(f'<td style="background-color:#E2E8F0;color:#1B2838;font-weight:600;padding:8px 12px;border-top:2px solid #CBD5E0;">{display_val}</td>')
                parts.append('</tr>')

            parts.append('</tbody></table></div>')
            st.markdown(''.join(parts), unsafe_allow_html=True)

        csv_pivot = fmt_pivot.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')
        st.download_button("Download Pivot CSV", data=csv_pivot, file_name="RSA_Pivot_Export.csv", mime="text/csv", key="dl_pivot")
    except Exception:
        st.markdown('<div style="background:white;border-radius:12px;padding:48px 24px;text-align:center;box-shadow:0 2px 8px rgba(0,0,0,0.08);color:#A0AEC0;font-size:14px;">No data to display</div>', unsafe_allow_html=True)
else:
    st.info("Select at least one Row or Column dimension to build the pivot table.")

# ============================================================================
# COST ANALYSIS
# ============================================================================
st.markdown('<div class="section-header">\U0001f4b0 Cost Analysis Dashboard</div>', unsafe_allow_html=True)

_BLANK_BOX = '<div style="background:white;border-radius:12px;padding:48px 24px;text-align:center;box-shadow:0 2px 8px rgba(0,0,0,0.08);color:#A0AEC0;font-size:14px;">No data to display</div>'

try:
    monthly_cost = filtered_df.groupby(['Year', 'Month'])['Fee (Baht)'].sum().reset_index()
    if len(monthly_cost) > 0:
        monthly_cost['Year'] = monthly_cost['Year'].astype(int)
        monthly_cost['Month'] = monthly_cost['Month'].astype(int)

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
except Exception:
    st.markdown(_BLANK_BOX, unsafe_allow_html=True)

# ============================================================================
# ADDITIONAL ANALYTICS
# ============================================================================
st.markdown('<div class="section-header">\U0001f4ca Additional Analytics</div>', unsafe_allow_html=True)

c1, c2 = st.columns(2)
with c1:
    try:
        svc_dist = filtered_df['\u0e1b\u0e23\u0e30\u0e40\u0e20\u0e17\u0e01\u0e32\u0e23\u0e1a\u0e23\u0e34\u0e01\u0e32\u0e23'].value_counts()
        fig_pie = px.pie(values=svc_dist.values, names=svc_dist.index, title='Service Type Distribution', hole=0.4,
                         color_discrete_sequence=['#4A90D9','#27AE60','#F39C12','#E74C3C','#1B2838','#2D5AA0','#6FB1FF'])
        fig_pie.update_traces(textposition='inside', textinfo='percent+label', textfont_size=11)
        fig_pie.update_layout(height=400, title={'font': CHART_TITLE_FONT},
                              font=CHART_LAYOUT_DEFAULTS['font'], paper_bgcolor='white', plot_bgcolor='rgba(0,0,0,0)',
                              legend=dict(bgcolor='rgba(255,255,255,0.9)', bordercolor='#E2E8F0', borderwidth=1))
        st.plotly_chart(fig_pie, use_container_width=True, config=PLOTLY_CONFIG)
    except Exception:
        st.markdown(_BLANK_BOX, unsafe_allow_html=True)

with c2:
    try:
        lob_counts = filtered_df['LOB'].value_counts().sort_index()
        fig_lob = px.bar(x=lob_counts.index, y=lob_counts.values, title='Cases by LOB',
                         labels={'x':'LOB','y':'Cases'}, color=lob_counts.values,
                         color_continuous_scale=[[0,'#4A90D9'],[0.5,'#2D5AA0'],[1,'#1B2838']])
        fig_lob.update_layout(height=400, showlegend=False, title={'font': CHART_TITLE_FONT}, **CHART_LAYOUT_DEFAULTS)
        st.plotly_chart(fig_lob, use_container_width=True, config=PLOTLY_CONFIG)
    except Exception:
        st.markdown(_BLANK_BOX, unsafe_allow_html=True)

c3, c4 = st.columns(2)
with c3:
    try:
        top_vol = filtered_df['\u0e1b\u0e23\u0e30\u0e40\u0e20\u0e17\u0e01\u0e32\u0e23\u0e1a\u0e23\u0e34\u0e01\u0e32\u0e23'].value_counts().head(10)
        fig_tv = px.bar(x=top_vol.values, y=top_vol.index, orientation='h', title='Top Services by Volume',
                        labels={'x':'Cases','y':'Service'}, color=top_vol.values,
                        color_continuous_scale=[[0,'#27AE60'],[0.5,'#3D8E56'],[1,'#1E7E34']])
        fig_tv.update_layout(height=400, showlegend=False,
                             yaxis={'categoryorder':'total ascending', 'gridcolor':'#E2E8F0', 'showline':True, 'linecolor':'#E2E8F0'},
                             title={'font': CHART_TITLE_FONT},
                             **{k: v for k, v in CHART_LAYOUT_DEFAULTS.items() if k != 'yaxis'})
        st.plotly_chart(fig_tv, use_container_width=True, config=PLOTLY_CONFIG)
    except Exception:
        st.markdown(_BLANK_BOX, unsafe_allow_html=True)

with c4:
    try:
        top_cost = filtered_df.groupby('\u0e1b\u0e23\u0e30\u0e40\u0e20\u0e17\u0e01\u0e32\u0e23\u0e1a\u0e23\u0e34\u0e01\u0e32\u0e23')['Fee (Baht)'].sum().sort_values(ascending=False).head(10)
        fig_tc = px.bar(x=top_cost.values, y=top_cost.index, orientation='h', title='Top Services by Cost',
                        labels={'x':'Fee (Baht)','y':'Service'}, color=top_cost.values,
                        color_continuous_scale=[[0,'#F39C12'],[0.5,'#E74C3C'],[1,'#C0392B']])
        fig_tc.update_layout(height=400, showlegend=False,
                             yaxis={'categoryorder':'total ascending', 'gridcolor':'#E2E8F0', 'showline':True, 'linecolor':'#E2E8F0'},
                             title={'font': CHART_TITLE_FONT},
                             **{k: v for k, v in CHART_LAYOUT_DEFAULTS.items() if k != 'yaxis'})
        st.plotly_chart(fig_tc, use_container_width=True, config=PLOTLY_CONFIG)
    except Exception:
        st.markdown(_BLANK_BOX, unsafe_allow_html=True)

# ============================================================================
# REGIONAL ANALYSIS
# ============================================================================
if '\u0e08\u0e31\u0e07\u0e2b\u0e27\u0e31\u0e14' in filtered_df.columns:
    st.markdown('<div class="section-header">\U0001f5fa\ufe0f Regional Analysis</div>', unsafe_allow_html=True)
    c5, c6 = st.columns(2)
    with c5:
        try:
            rc = filtered_df['\u0e08\u0e31\u0e07\u0e2b\u0e27\u0e31\u0e14'].value_counts().head(15)
            fig_r = px.bar(x=rc.values, y=rc.index, orientation='h', title='Top 15 Regions by Volume',
                           labels={'x':'Cases','y':'Province'}, color=rc.values,
                           color_continuous_scale=[[0,'#4A90D9'],[0.5,'#2D5AA0'],[1,'#1B2838']])
            fig_r.update_layout(height=500, showlegend=False,
                                yaxis={'categoryorder':'total ascending', 'gridcolor':'#E2E8F0', 'showline':True, 'linecolor':'#E2E8F0'},
                                title={'font': CHART_TITLE_FONT},
                                **{k: v for k, v in CHART_LAYOUT_DEFAULTS.items() if k != 'yaxis'})
            st.plotly_chart(fig_r, use_container_width=True, config=PLOTLY_CONFIG)
        except Exception:
            st.markdown(_BLANK_BOX, unsafe_allow_html=True)
    with c6:
        try:
            rcost = filtered_df.groupby('\u0e08\u0e31\u0e07\u0e2b\u0e27\u0e31\u0e14')['Fee (Baht)'].sum().sort_values(ascending=False).head(15)
            fig_rc = px.bar(x=rcost.values, y=rcost.index, orientation='h', title='Top 15 Regions by Cost',
                            labels={'x':'Fee (Baht)','y':'Province'}, color=rcost.values,
                            color_continuous_scale=[[0,'#F39C12'],[0.5,'#E67E22'],[1,'#D35400']])
            fig_rc.update_layout(height=500, showlegend=False,
                                 yaxis={'categoryorder':'total ascending', 'gridcolor':'#E2E8F0', 'showline':True, 'linecolor':'#E2E8F0'},
                                 title={'font': CHART_TITLE_FONT},
                                 **{k: v for k, v in CHART_LAYOUT_DEFAULTS.items() if k != 'yaxis'})
            st.plotly_chart(fig_rc, use_container_width=True, config=PLOTLY_CONFIG)
        except Exception:
            st.markdown(_BLANK_BOX, unsafe_allow_html=True)

# ============================================================================
# MONTHLY TREND BY SERVICE TYPE
# ============================================================================
st.markdown('<div class="section-header">\U0001f4c8 Monthly Trend by Service Type</div>', unsafe_allow_html=True)

try:
    mst = filtered_df.groupby(['Year', 'Month', '\u0e1b\u0e23\u0e30\u0e40\u0e20\u0e17\u0e01\u0e32\u0e23\u0e1a\u0e23\u0e34\u0e01\u0e32\u0e23']).size().reset_index(name='Count')
    if len(mst) > 0:
        mst['Year'] = mst['Year'].astype(int)
        mst['Month'] = mst['Month'].astype(int)
        mst['Date'] = pd.to_datetime(mst[['Year', 'Month']].assign(Day=1))
        fig_mst = px.line(mst, x='Date', y='Count', color='\u0e1b\u0e23\u0e30\u0e40\u0e20\u0e17\u0e01\u0e32\u0e23\u0e1a\u0e23\u0e34\u0e01\u0e32\u0e23', title='Monthly Case Volume by Service Type', markers=True)
        fig_mst.update_layout(
            xaxis_title='Date', yaxis_title='Cases', hovermode='x unified', height=450,
            title={'font': CHART_TITLE_FONT}, **CHART_LAYOUT_DEFAULTS,
            legend=dict(orientation="v", yanchor="top", y=1, xanchor="left", x=1.02, bgcolor='rgba(255,255,255,0.9)', bordercolor='#E2E8F0', borderwidth=1)
        )
        st.plotly_chart(fig_mst, use_container_width=True, config=PLOTLY_CONFIG)
except Exception:
    st.markdown(_BLANK_BOX, unsafe_allow_html=True)

# ============================================================================
# DATA EXPORT
# ============================================================================
st.markdown('<div class="section-header">\U0001f4be Data Export</div>', unsafe_allow_html=True)
ce1, ce2 = st.columns([3, 1])
with ce1:
    st.write(f"Export filtered data ({len(filtered_df):,} records)")
with ce2:
    csv_data = convert_df_to_csv(filtered_df)
    st.download_button("Download CSV", data=csv_data,
                       file_name=f"RSA_Export_{st.session_state.export_timestamp}.csv",
                       mime="text/csv", use_container_width=True)

# ============================================================================
# FILE UPLOAD
# ============================================================================
st.markdown("---")
st.markdown('<div class="section-header">\U0001f4c2 Data Source</div>', unsafe_allow_html=True)
fu1, fu2 = st.columns([3, 1])
with fu1:
    uploaded_file = st.file_uploader("Upload RSA Report (.xlsx)", type=["xlsx"], key="file_uploader",
                                     help="Upload a new Excel file to replace the current data source.")
    if uploaded_file is not None:
        new_bytes = uploaded_file.getvalue()
        if st.session_state.uploaded_file_name != uploaded_file.name or st.session_state.uploaded_file_bytes != new_bytes:
            # Validate file before persisting
            try:
                test_df = load_and_process(file_bytes=new_bytes)
                if test_df is None:
                    raise ValueError("Could not process file")
                required_check = ['Year', 'Month', 'Fee (Baht)', '\u0e1b\u0e23\u0e30\u0e40\u0e20\u0e17\u0e01\u0e32\u0e23\u0e1a\u0e23\u0e34\u0e01\u0e32\u0e23', 'LOB']
                missing_check = [c for c in required_check if c not in test_df.columns]
                if missing_check:
                    raise ValueError(f"Missing required columns: {missing_check}")
            except Exception:
                st.error("The file is not supported. Please upload a valid RSA Report Excel file.")
            else:
                persist_uploaded_file(uploaded_file)
                st.session_state.uploaded_file_bytes = new_bytes
                st.session_state.uploaded_file_name = uploaded_file.name
                st.cache_data.clear()
                # Clear filter session state so they reset to new data's defaults
                for k in ['sel_years', 'sel_services', 'sel_lobs', 'sel_months',
                           'sel_channels', 'sel_regions', 'sel_makes', 'sel_models']:
                    st.session_state.pop(k, None)
                st.rerun()
with fu2:
    st.markdown(f"**Current source:** {data_source_label}")
    p_path = os.path.join(UPLOAD_DIR, "persisted_upload.xlsx")
    if os.path.exists(p_path):
        if st.button("Clear uploaded file", key="clear_upload"):
            os.remove(p_path)
            st.session_state.uploaded_file_bytes = None
            st.session_state.uploaded_file_name = None
            st.cache_data.clear()
            st.rerun()

st.markdown(f"<div class='dashboard-footer'><strong>RSA Dashboard</strong> - Sompo Thailand<br>Dashboard rendered: {datetime.now().strftime('%Y-%m-%d %H:%M')}</div>", unsafe_allow_html=True)
