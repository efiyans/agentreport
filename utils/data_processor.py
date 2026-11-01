"""Data processing and financial analytics utilities."""

import datetime
from typing import Dict, Any, Tuple, Optional
import numpy as np
import pandas as pd


def generate_sample_financial_data() -> pd.DataFrame:
    """Generate a realistic synthetic financial dataset for demonstration."""
    np.random.seed(42)
    start_date = datetime.date(2025, 1, 1)
    
    departments = ["Engineering", "Sales", "Marketing", "Operations", "Customer Success"]
    
    revenue_streams = [
        ("SaaS Subscriptions", "Sales"),
        ("Enterprise Licenses", "Sales"),
        ("Professional Services", "Customer Success"),
        ("Support Contracts", "Customer Success"),
    ]
    
    expense_categories = [
        ("Cloud Infrastructure", "Engineering"),
        ("Software Licenses & Tools", "Engineering"),
        ("Salaries & Benefits", "Operations"),
        ("Paid Advertising", "Marketing"),
        ("Events & Sponsorships", "Marketing"),
        ("Travel & Entertainment", "Sales"),
        ("Office Facilities & Utilities", "Operations"),
        ("Legal & Accounting", "Operations"),
    ]
    
    records = []
    
    # Generate 14 months of sample transactions
    for month_idx in range(14):
        month_date = start_date + datetime.timedelta(days=month_idx * 30)
        
        # Monthly Revenue items
        for stream, dept in revenue_streams:
            num_tx = np.random.randint(2, 6)
            for _ in range(num_tx):
                base_val = 15000 if "Enterprise" in stream else (8000 if "SaaS" in stream else 4000)
                amount = float(np.round(np.random.uniform(base_val * 0.8, base_val * 1.5) * (1 + 0.03 * month_idx), 2))
                day_offset = int(np.random.randint(1, 28))
                tx_date = month_date.replace(day=day_offset)
                records.append({
                    "Date": pd.to_datetime(tx_date),
                    "Type": "Revenue",
                    "Category": stream,
                    "Department": dept,
                    "Amount": amount,
                    "Description": f"{stream} invoice settlement",
                })
        
        # Monthly Expense items
        for cat, dept in expense_categories:
            num_tx = np.random.randint(1, 4)
            for _ in range(num_tx):
                base_exp = 12000 if "Salaries" in cat else (6000 if "Cloud" in cat else 2500)
                amount = float(np.round(np.random.uniform(base_exp * 0.85, base_exp * 1.25), 2))
                day_offset = int(np.random.randint(1, 28))
                tx_date = month_date.replace(day=day_offset)
                records.append({
                    "Date": pd.to_datetime(tx_date),
                    "Type": "Expense",
                    "Category": cat,
                    "Department": dept,
                    "Amount": amount,
                    "Description": f"Monthly {cat} billing",
                })

    df = pd.DataFrame(records)
    df.sort_values(by="Date", inplace=True)
    df.reset_index(drop=True, inplace=True)
    return df


def load_financial_data(file_or_path: Any) -> pd.DataFrame:
    """Load financial data from an Excel (.xlsx, .xls) or CSV file.
    
    Validates required columns: Date, Type, Category, Department, Amount.
    """
    if hasattr(file_or_path, "name"):
        filename = file_or_path.name.lower()
    else:
        filename = str(file_or_path).lower()

    if filename.endswith(".csv"):
        df = pd.read_csv(file_or_path)
    elif filename.endswith((".xlsx", ".xls")):
        df = pd.read_excel(file_or_path, engine="openpyxl")
    else:
        raise ValueError("Unsupported file format. Please upload a .xlsx or .csv file.")

    # Standardize column names (lowercase, stripped)
    column_mapping = {}
    for col in df.columns:
        norm = str(col).strip().lower()
        if norm in ["date", "transaction_date", "tx_date"]:
            column_mapping[col] = "Date"
        elif norm in ["type", "tx_type", "transaction_type"]:
            column_mapping[col] = "Type"
        elif norm in ["category", "cat"]:
            column_mapping[col] = "Category"
        elif norm in ["department", "dept"]:
            column_mapping[col] = "Department"
        elif norm in ["amount", "value", "total"]:
            column_mapping[col] = "Amount"
        elif norm in ["description", "desc", "memo"]:
            column_mapping[col] = "Description"

    df = df.rename(columns=column_mapping)

    required_cols = ["Date", "Type", "Category", "Amount"]
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        raise ValueError(f"Uploaded dataset is missing required columns: {', '.join(missing)}")

    if "Department" not in df.columns:
        df["Department"] = "General"
    if "Description" not in df.columns:
        df["Description"] = ""

    # Parse and clean types
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    df = df.dropna(subset=["Date"])
    df["Amount"] = pd.to_numeric(df["Amount"], errors="coerce").fillna(0.0)
    df["Type"] = df["Type"].astype(str).str.strip().str.capitalize()
    df["Category"] = df["Category"].astype(str).str.strip()
    df["Department"] = df["Department"].astype(str).str.strip()
    
    df.sort_values(by="Date", inplace=True)
    df.reset_index(drop=True, inplace=True)
    return df


def calculate_kpis(df: pd.DataFrame) -> Dict[str, Any]:
    """Calculate core executive financial metrics (Revenue, Expenses, Net Profit, Margin)."""
    if df.empty:
        return {
            "total_revenue": 0.0,
            "total_expenses": 0.0,
            "net_profit": 0.0,
            "profit_margin_pct": 0.0,
            "avg_monthly_revenue": 0.0,
            "avg_monthly_expense": 0.0,
            "transaction_count": 0,
        }

    revenue_mask = df["Type"] == "Revenue"
    expense_mask = df["Type"] == "Expense"

    total_revenue = float(df.loc[revenue_mask, "Amount"].sum())
    total_expenses = float(df.loc[expense_mask, "Amount"].sum())
    net_profit = total_revenue - total_expenses
    profit_margin = (net_profit / total_revenue * 100) if total_revenue > 0 else 0.0

    # Monthly averages
    periods = df["Date"].dt.to_period("M").nunique() or 1
    avg_monthly_rev = total_revenue / periods
    avg_monthly_exp = total_expenses / periods

    return {
        "total_revenue": total_revenue,
        "total_expenses": total_expenses,
        "net_profit": net_profit,
        "profit_margin_pct": profit_margin,
        "avg_monthly_revenue": avg_monthly_rev,
        "avg_monthly_expense": avg_monthly_exp,
        "transaction_count": len(df),
    }


def get_monthly_breakdown(df: pd.DataFrame) -> pd.DataFrame:
    """Aggregate financial data by Year-Month into Revenue, Expenses, Net Income, and Margin."""
    if df.empty:
        return pd.DataFrame(columns=["YearMonth", "Revenue", "Expense", "Net Income", "Margin %"])

    df_copy = df.copy()
    df_copy["YearMonth"] = df_copy["Date"].dt.strftime("%Y-%m")

    grouped = (
        df_copy.groupby(["YearMonth", "Type"])["Amount"]
        .sum()
        .unstack(fill_value=0.0)
        .reset_index()
    )

    if "Revenue" not in grouped.columns:
        grouped["Revenue"] = 0.0
    if "Expense" not in grouped.columns:
        grouped["Expense"] = 0.0

    grouped["Net Income"] = grouped["Revenue"] - grouped["Expense"]
    grouped["Margin %"] = np.where(
        grouped["Revenue"] > 0,
        (grouped["Net Income"] / grouped["Revenue"]) * 100,
        0.0,
    )
    grouped.sort_values(by="YearMonth", inplace=True)
    return grouped


def get_category_breakdown(df: pd.DataFrame, transaction_type: str = "Expense") -> pd.DataFrame:
    """Aggregate totals and percentage breakdown by category for a given transaction type."""
    filtered = df[df["Type"] == transaction_type]
    if filtered.empty:
        return pd.DataFrame(columns=["Category", "Total Amount", "Share %"])

    breakdown = (
        filtered.groupby("Category")["Amount"]
        .sum()
        .reset_index()
        .rename(columns={"Amount": "Total Amount"})
    )
    total_val = breakdown["Total Amount"].sum()
    breakdown["Share %"] = (breakdown["Total Amount"] / total_val * 100) if total_val > 0 else 0.0
    breakdown.sort_values(by="Total Amount", ascending=False, inplace=True)
    breakdown.reset_index(drop=True, inplace=True)
    return breakdown


def get_department_breakdown(df: pd.DataFrame) -> pd.DataFrame:
    """Aggregate financial metrics grouped by department."""
    if df.empty:
        return pd.DataFrame(columns=["Department", "Revenue", "Expense", "Net"])

    pivot = (
        df.groupby(["Department", "Type"])["Amount"]
        .sum()
        .unstack(fill_value=0.0)
        .reset_index()
    )

    if "Revenue" not in pivot.columns:
        pivot["Revenue"] = 0.0
    if "Expense" not in pivot.columns:
        pivot["Expense"] = 0.0

    pivot["Net"] = pivot["Revenue"] - pivot["Expense"]
    pivot.sort_values(by="Expense", ascending=False, inplace=True)
    return pivot

