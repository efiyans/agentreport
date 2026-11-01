# Financial Reporting Web Dashboard

An interactive financial reporting dashboard built with **Streamlit**, **Pandas**, and **openpyxl**. It provides executive-level financial monitoring, P&L analytics, department breakdowns, and exportable formatted Excel workbooks.

---

## 📁 Project Structure

```text
financial-report-agent/
├── app.py                      # Main Streamlit application and dashboard UI
├── requirements.txt            # Project dependencies (streamlit, pandas, openpyxl)
├── .gitignore                  # Git ignore rules for Python & virtual environments
├── README.md                   # Project documentation and setup instructions
├── data/
│   └── sample_financial_data.csv # Template financial transactions dataset
└── utils/
    ├── __init__.py             # Exposes core utility functions
    ├── data_processor.py       # Data loading, validation, and KPI calculations
    └── excel_exporter.py       # Formatted multi-tab Excel export logic using openpyxl
```

---

## 🚀 Quickstart Guide

### 1. Create and Activate a Virtual Environment

```bash
# Create virtual environment
python3 -m venv .venv

# Activate virtual environment
# On Linux / macOS:
source .venv/bin/activate

# On Windows:
# .venv\Scripts\activate
```

### 2. Install Dependencies

```bash
pip install -r requirements.txt
```

### 3. Run the Streamlit Dashboard

```bash
streamlit run app.py
```

The application will start locally (typically at `http://localhost:8501`).

---

## 🌟 Key Features

- **Executive KPI Cards**: Real-time calculation of Total Revenue, Total Expenses, Net Profit / Loss, and Profit Margin %.
- **Interactive Visualizations**:
  - Monthly Revenue vs. Expense comparisons.
  - Net Income trajectories over time.
  - Horizontal breakdown of expense categories.
- **Department P&L Breakdown**: Aggregate revenue, expenses, and net margins segmented by department or business unit.
- **Flexible Data Input**:
  - Built-in synthetic sample data generator with realistic 14-month SaaS/Enterprise transactions.
  - Custom file upload supporting both Excel (`.xlsx`, `.xls`) and CSV files.
- **Dynamic Global Filters**: Filter by date range, department, and transaction type (Revenue / Expense).
- **Executive Excel Export**: Formatted multi-tab `.xlsx` reports generated via `openpyxl`:
  - *Executive Summary*: Styled KPI indicators with currency and percentage formatting.
  - *Monthly Trends*: Aggregated monthly statement table.
  - *Department Breakdown*: Cross-department revenue and cost allocation.
  - *Expense Breakdown*: Category totals and percentage share.
  - *Transactions*: Filtered transaction ledger.

---

## 📋 Data Schema for Custom Uploads

When uploading custom `.xlsx` or `.csv` files, the dataset should include the following columns:

| Column | Type | Description |
| :--- | :--- | :--- |
| `Date` | Date / String | Transaction date (`YYYY-MM-DD` recommended) |
| `Type` | String | Transaction type (`Revenue` or `Expense`) |
| `Category` | String | Category or revenue stream name |
| `Department` | String | Department / Cost Center (*Optional, defaults to 'General'*) |
| `Amount` | Numeric | Financial amount (positive number) |
| `Description` | String | Memo or transaction details (*Optional*) |

---

## 📦 Requirements

- **Streamlit**: Web dashboard interface and reactive state management.
- **Pandas**: Financial aggregations, grouping, and metrics calculations.
- **openpyxl**: Excel workbook reading, creation, and cell styling.
