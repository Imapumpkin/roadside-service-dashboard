# RSA Dashboard - Product Requirements Document

## Product Overview

**Product Name:** RSA Dashboard
**Organization:** Sompo Thailand
**Purpose:** Monitor and analyze Roadside Assistance (RSA) service utilization, costs, and performance metrics.

## Problem Statement

Sompo Thailand needs to track and analyze roadside assistance claims to:
- Monitor service costs against budget
- Identify trends in service utilization
- Track performance by region, service type, and vehicle
- Make data-driven decisions for cost optimization

## Target Users

- Insurance Operations Team
- Claims Management Team
- Finance/Budget Analysts
- Management (for KPI monitoring)

## Features

### 1. Authentication
- Password-protected access
- Password configurable via Streamlit secrets or default fallback

### 2. Key Performance Indicators (KPIs)
- **YTD Total Cases**: Year-to-date case count with YoY comparison
- **YTD Total Fee**: Year-to-date fee amount with YoY comparison
- **Avg Fee/Case**: Average cost per case with YoY comparison
- **MTD Fee**: Month-to-date fee amount
- **MTD Utilization**: Current month fee vs monthly budget percentage

### 3. Portfolio Health Indicator
- Status levels: HEALTHY, WARNING, CRITICAL
- Based on over-budget percentage thresholds
- Shows YTD vs Expected Cost, Monthly Run Rate, Year-End Projection

### 4. Interactive Pivot Table
- Configurable rows and columns
- Multiple dimension options: LOB, Service Type, Year, Month, Region, Vehicle Make/Model, Channel
- Value options: Case Count, Fee (Baht)
- Aggregation options: Count, Sum, Mean, Median, Min, Max, % of Row/Column/Grand Total
- Sortable columns
- CSV export

### 5. Data Filters
- Year
- Month
- Service Type (ประเภทการบริการ)
- LOB (Line of Business)
- Channel (รหัสโครงการ)
- Region (จังหวัด)
- Vehicle Make (ยี่ห้อรถ)
- Vehicle Model (รุ่นรถ)

### 6. Visualizations
- Monthly Cost Trend with Budget Line
- Service Type Distribution (Donut Chart)
- Cases by LOB (Bar Chart)
- Top Services by Volume
- Top Services by Cost
- Regional Analysis (Volume & Cost)
- Monthly Trend by Service Type

### 7. Data Management
- File upload (.xlsx)
- Persistent storage of uploaded files
- Clear uploaded file option
- CSV export of filtered data

## Business Rules

### Fee Calculation Rules
| Service Type (ประเภทการบริการ) | Fee (Baht) |
|-------------------------------|------------|
| ลูกค้าแจ้งยกเลิก (Customer Cancellation) | 100 |
| สอบถามข้อมูล (Information Inquiry) | 100 |
| Other service types | As per data |

### Budget Configuration
- **Monthly Budget:** 200,000 Baht
- **Annual Budget:** 2,400,000 Baht

### Health Thresholds
- **Healthy:** Over-budget <= 5%
- **Warning:** Over-budget <= 15%
- **Critical:** Over-budget > 15%

### LOB Extraction
- Extracted from Policy Number using pattern `A[CV]\d` (e.g., AC3, AV1, AV5, AV9)
- Unmatched policies marked as "Unverify"

### Province Extraction
- Extracted from license plate field (ทะเบียนรถ)
- Bangkok variants (กรุงเทพ, กทม) normalized to กรุงเทพมหานคร

## Data Requirements

### Required Excel Structure
The Excel file must contain a header row with "Policy No." column. The system auto-detects:
- Sheets containing both "Roadside_Plan" and "Policy Type" headers (preferred)
- Sheets containing "Policy No." header (fallback)

### Required Columns (after processing)
| Column | Description |
|--------|-------------|
| Policy No. | Insurance policy number |
| วันที่ | Service date |
| ประเภทการบริการ | Service type |
| Fee (Baht) | Service fee amount |
| ทะเบียนรถ | Vehicle license plate |
| จังหวัด | Province |
| ยี่ห้อรถ | Vehicle make |
| รุ่นรถ | Vehicle model |
| รหัสโครงการ | Project/Channel code |

### Derived Columns
| Column | Source |
|--------|--------|
| Year | Extracted from วันที่ |
| Month | Extracted from วันที่ |
| Day | Extracted from วันที่ |
| LOB | Extracted from Policy No. |
| Policy Type | Extracted from Policy No. |
| จังหวัด ทะเบียนรถ | Extracted from ทะเบียนรถ |

## Non-Functional Requirements

### Performance
- Data caching with 1-hour TTL
- Efficient filtering using category dtypes
- Fragment-based rendering for charts

### Security
- Password authentication required
- No sensitive data exposure in URL

### Compatibility
- Responsive wide layout
- Modern browsers support
- Thai language support (UTF-8)

## Success Metrics

- Dashboard loads within 3 seconds
- Users can filter and analyze data without training
- All KPIs update correctly with data changes
- Export functionality works reliably
