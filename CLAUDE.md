# CLAUDE.md - RSA Dashboard Project Context

## Project Overview

This is a Streamlit-based dashboard for Sompo Thailand to monitor Roadside Assistance (RSA) service data. The dashboard displays KPIs, interactive pivot tables, and various charts for cost and utilization analysis.

## Tech Stack

- **Framework:** Streamlit
- **Data Processing:** Pandas
- **Visualization:** Plotly (Express & Graph Objects)
- **Styling:** Custom CSS (style.css)
- **Data Source:** Excel files (.xlsx)

## File Structure

```
Streamlit Dashboard/
├── app.py                 # Main application (single file)
├── style.css              # Custom CSS styles
├── uploaded_data/         # Directory for persisted uploads
│   └── persisted_upload.xlsx
├── (Test) RSA Report.xlsx # Default data file
├── PRD.md                 # Product Requirements Document
└── CLAUDE.md              # This file
```

## Key Configuration (app.py lines 14-20)

```python
MONTHLY_BUDGET = 200_000          # Monthly budget in Baht
HEALTH_THRESHOLD_HEALTHY = 5      # % over budget for healthy status
HEALTH_THRESHOLD_WARNING = 15     # % over budget for warning status
CACHE_TTL = 3600                  # Cache duration in seconds
DEFAULT_DATA_FILE = "(Test) RSA Report.xlsx"
UPLOAD_DIR = "uploaded_data"
```

## Thai Column Name Mapping

| Thai Column | Unicode Escape | English Meaning |
|-------------|----------------|-----------------|
| วันที่ | `\u0e27\u0e31\u0e19\u0e17\u0e35\u0e48` | Date |
| ประเภทการบริการ | `\u0e1b\u0e23\u0e30\u0e40\u0e20\u0e17\u0e01\u0e32\u0e23\u0e1a\u0e23\u0e34\u0e01\u0e32\u0e23` | Service Type |
| จังหวัด | `\u0e08\u0e31\u0e07\u0e2b\u0e27\u0e31\u0e14` | Province |
| ทะเบียนรถ | `\u0e17\u0e30\u0e40\u0e1a\u0e35\u0e22\u0e19\u0e23\u0e16` | License Plate |
| จังหวัด ทะเบียนรถ | `\u0e08\u0e31\u0e07\u0e2b\u0e27\u0e31\u0e14 \u0e17\u0e30\u0e40\u0e1a\u0e35\u0e22\u0e19\u0e23\u0e16` | Plate Province |
| ยี่ห้อรถ | `\u0e22\u0e35\u0e48\u0e2b\u0e49\u0e2d\u0e23\u0e16` | Vehicle Make |
| รุ่นรถ | `\u0e23\u0e38\u0e48\u0e19\u0e23\u0e16` | Vehicle Model |
| รหัสโครงการ | `\u0e23\u0e2b\u0e31\u0e2a\u0e42\u0e04\u0e23\u0e07\u0e01\u0e32\u0e23` | Channel/Project Code |
| แผนก | `\u0e41\u0e1c\u0e19\u0e01` | Department |
| ลูกค้าจ่ายส่วนต่าง | `\u0e25\u0e39\u0e01\u0e04\u0e49\u0e32\u0e08\u0e48\u0e32\u0e22\u0e2a\u0e48\u0e27\u0e19\u0e15\u0e48\u0e32\u0e07` | Customer Excess Payment |

## Important Service Type Values

| Thai Value | Unicode Escape | English |
|------------|----------------|---------|
| ลูกค้าแจ้งยกเลิก | `\u0e25\u0e39\u0e01\u0e04\u0e49\u0e32\u0e41\u0e08\u0e49\u0e07\u0e22\u0e01\u0e40\u0e25\u0e34\u0e01` | Customer Cancellation |
| สอบถามข้อมูล | `\u0e2a\u0e2d\u0e1a\u0e16\u0e32\u0e21\u0e02\u0e49\u0e2d\u0e21\u0e39\u0e25` | Information Inquiry |

## Code Structure

### Main Sections (by line number)

| Lines | Section |
|-------|---------|
| 1-29 | Imports & Page Config |
| 31-70 | Password Authentication |
| 72-77 | CSS Loading |
| 79-179 | Data Loading & Processing (`load_and_process()`) |
| 181-196 | File Persistence Functions |
| 198-255 | Data Source Selection & Validation |
| 257-291 | Helper Functions & Filter Options |
| 293-344 | KPI Cards |
| 346-377 | Portfolio Health Indicator |
| 379-660 | Interactive Pivot Table |
| 662-697 | Cost Analysis Chart |
| 699-760 | Additional Analytics (Pie, Bar Charts) |
| 762-798 | Regional Analysis |
| 800-822 | Monthly Trend by Service Type |
| 824-835 | Data Export |
| 837-882 | File Upload Section |

### Key Functions

#### `load_and_process(file_bytes=None, file_path=None)` (line 83)
- Loads Excel file and processes data
- Auto-detects correct sheet and header row
- Extracts dates, LOB, provinces
- **Fee Rule:** Sets Fee to 100 for cancellation/inquiry (lines 166-169)
- Returns processed DataFrame

#### `check_password()` (line 34)
- Handles authentication
- Uses `st.secrets["password"]` or fallback "sompo2026"

#### `yoy_html(cur_val, prev_val, compare_year)` (line 324)
- Generates YoY comparison HTML with color-coded arrows

### Data Flow

```
Excel File → load_and_process() → df (cached)
                    ↓
            Apply Filters → filtered_df
                    ↓
            Render Dashboard Components
```

## Common Modification Tasks

### Change Monthly Budget
Edit line 15:
```python
MONTHLY_BUDGET = 200_000  # Change this value
```

### Add New Fee Rule
Add condition in `load_and_process()` after line 169:
```python
if 'Fee (Baht)' in df.columns and 'ประเภทการบริการ' in df.columns:
    new_mask = df['ประเภทการบริการ'].isin(['new_service_type'])
    df.loc[new_mask, 'Fee (Baht)'] = new_fee_value
```

### Change Default Pivot Table Row
Edit line 418, change the `default` parameter:
```python
default=['\u0e1b\u0e23\u0e30\u0e40\u0e20\u0e17\u0e01\u0e32\u0e23\u0e1a\u0e23\u0e34\u0e01\u0e32\u0e23']  # Service Type
# or
default=['LOB']  # LOB
```

### Add New Filter Dimension
1. Add to `pivot_cols_available` (line 384)
2. Add filter UI in the Data Filter expander (lines 391-409)
3. Add mask condition in Apply Filters section (lines 440-454)

### Add New Chart
Use `@st.fragment` decorator for performance:
```python
@st.fragment
def render_new_chart():
    st.markdown('<div class="section-header">Chart Title</div>', unsafe_allow_html=True)
    # Chart code here

render_new_chart()
```

### Change Health Thresholds
Edit lines 16-17:
```python
HEALTH_THRESHOLD_HEALTHY = 5   # % for healthy
HEALTH_THRESHOLD_WARNING = 15  # % for warning
```

## Design System (style.css)

The dashboard uses a comprehensive CSS design system with design tokens.

### Design Tokens (CSS Variables)

```css
/* Primary Palette */
--color-primary-900: #0D1B2A;  /* Darkest navy */
--color-primary-800: #1B2838;  /* Dark navy */
--color-accent-500: #4A90D9;   /* Electric blue */

/* Semantic Colors */
--color-success-500: #10B981;  /* Green */
--color-warning-500: #F59E0B;  /* Amber */
--color-danger-500: #EF4444;   /* Red */

/* Spacing Scale */
--space-1 to --space-16       /* 0.25rem to 4rem */

/* Typography */
--font-size-xs to --font-size-4xl  /* 11px to 48px */
```

### Key CSS Classes

| Class | Purpose |
|-------|---------|
| `.metric-card` | KPI cards with glass morphism effect |
| `.metric-title` | Uppercase label in KPI cards |
| `.metric-value` | Large numeric value in KPI cards |
| `.health-indicator` | Portfolio health container |
| `.health-healthy/warning/critical` | Health status variants |
| `.health-badge` | Status pill badge |
| `.service-table-container` | Pivot table wrapper |
| `.service-table` | Data table with sticky headers |
| `.section-header` | Section titles with asymmetric underline |
| `.data-freshness` | Live status badge with pulse animation |
| `.positive/.negative/.warning` | Semantic color classes |
| `.empty-state` | No data placeholder |
| `.dashboard-footer` | Page footer |

### Responsive Breakpoints

| Breakpoint | Target |
|------------|--------|
| 1200px | Large tablets, small desktops |
| 992px | Tablets |
| 768px | Mobile landscape, small tablets |
| 480px | Mobile portrait |

### Accessibility Features
- Focus-visible states for keyboard navigation
- `prefers-reduced-motion` support
- Print-optimized styles
- Custom scrollbar styling

## Caching Strategy

- `@st.cache_data(ttl=CACHE_TTL)` - Used for data loading and CSV conversion
- Cache cleared on file upload via `st.cache_data.clear()`
- `data_version` in session_state triggers filter widget key changes

## Running the App

```bash
cd "D:\Sompo Thailand\Test\RSA Dashboard\Streamlit Dashboard"
streamlit run app.py
```

## Password Configuration

For production, create `.streamlit/secrets.toml`:
```toml
password = "your_secure_password"
```

## Troubleshooting

### "Missing required columns" error
- Ensure Excel has "Policy No." in header row
- Check if fee column exists (any column with "Fee" in name, excluding "Exceed")

### Data not updating after upload
- Clear browser cache
- Click "Clear uploaded file" and re-upload
- Check `uploaded_data/` directory permissions

### Thai text display issues
- Ensure UTF-8 encoding
- Use unicode escapes in code for Thai strings
