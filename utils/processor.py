"""Financial & Payroll Report Processor.

Extracts key payroll figures from raw data (CSV or Excel), maps them to
the exact columns and rows of a reference template (.xlsx or .ods), calculates
totals, and outputs a clean, ready-to-download Excel workbook.
"""

import io
import re
import os
import argparse
from pathlib import Path
from typing import Dict, List, Tuple, Any, Optional, Union

import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# Standard payroll semantic aliases for intelligent column mapping
PAYROLL_COLUMN_ALIASES: Dict[str, List[str]] = {
    "employee_id": ["empid", "emp_id", "employee_id", "id", "staff_id", "badge_no", "employee_number", "emp_no"],
    "employee_name": ["name", "employee_name", "full_name", "employee", "staff_name", "worker_name"],
    "department": ["department", "dept", "cost_center", "business_unit", "division", "team", "dept_name"],
    "role": ["role", "position", "designation", "title", "job_title", "occupation"],
    "base_salary": ["base_salary", "basic_salary", "base_pay", "basic_pay", "salary", "basic", "monthly_salary", "base_wage"],
    "overtime": ["overtime", "ot", "overtime_pay", "ot_pay", "ot_amount", "extra_hours_pay"],
    "allowances": ["allowances", "allowance", "housing_allowance", "transport_allowance", "stipend", "benefits_allowance"],
    "bonus": ["bonus", "bonuses", "incentive", "incentives", "commission", "performance_bonus"],
    "gross_pay": ["gross_pay", "gross_salary", "gross_earnings", "total_earnings", "gross_amount", "gross"],
    "tax": ["tax", "income_tax", "paye", "withholding", "withholding_tax", "federal_tax", "state_tax", "tax_deduction"],
    "pension": ["pension", "social_security", "social_security_tax", "401k", "pf", "provident_fund", "fica", "superannuation"],
    "insurance": ["insurance", "health_insurance", "medical_insurance", "dental", "medical", "health_benefit"],
    "other_deductions": ["other_deductions", "misc_deductions", "loan_repayment", "advances", "voluntary_deductions"],
    "total_deductions": ["total_deductions", "deductions", "total_deduction", "all_deductions", "ded", "total_ded", "tot_ded"],
    "net_pay": ["net_pay", "net_salary", "take_home", "net_amount", "take_home_pay", "net"],
    "employer_contributions": ["employer_contributions", "employer_cost", "employer_pension", "employer_taxes", "company_contributions"],
    "total_cost": ["total_cost", "ctc", "company_cost", "total_employment_cost", "total_payroll_cost"],
}


def normalize_text(text: Any) -> str:
    """Normalize strings for tolerant fuzzy comparison (lowercase, alphanumeric)."""
    if text is None:
        return ""
    s = str(text).strip().lower()
    return re.sub(r"[^a-z0-9]", "", s)


def clean_currency_series(series: pd.Series) -> pd.Series:
    """Clean stringified currency amounts (e.g. '$5,400.50' or '(1,200)') to numeric floats."""
    if pd.api.types.is_numeric_dtype(series):
        return series.fillna(0.0).astype(float)
    
    def parse_val(v):
        if pd.isna(v) or v is None:
            return 0.0
        val_str = str(v).strip()
        if not val_str or val_str.lower() in ["-", "n/a", "none", "null"]:
            return 0.0
        # Check negative parentheses format: (123.45) -> -123.45
        is_neg = False
        if val_str.startswith("(") and val_str.endswith(")"):
            is_neg = True
            val_str = val_str[1:-1]
        val_str = re.sub(r"[^\d.]", "", val_str)
        try:
            num = float(val_str)
            return -num if is_neg else num
        except ValueError:
            return 0.0

    return series.apply(parse_val)


def map_column_to_canonical(col_name: str) -> Optional[str]:
    """Map any raw column name to a canonical payroll semantic concept."""
    norm = normalize_text(col_name)
    if not norm:
        return None
    for canonical, aliases in PAYROLL_COLUMN_ALIASES.items():
        if norm == normalize_text(canonical):
            return canonical
        for alias in aliases:
            if norm == normalize_text(alias):
                return canonical
    # Partial substring check
    for canonical, aliases in PAYROLL_COLUMN_ALIASES.items():
        for alias in aliases:
            norm_alias = normalize_text(alias)
            if len(norm_alias) >= 4 and (norm_alias in norm or norm in norm_alias):
                return canonical
    return None


def load_raw_payroll_data(file_or_path: Any) -> pd.DataFrame:
    """Load raw payroll information from CSV or Excel into a cleaned pandas DataFrame.
    
    Automatically handles numeric cleanup and computes standard derived fields
    (Gross Pay, Total Deductions, Net Pay) if they are missing.
    """
    if hasattr(file_or_path, "name"):
        fname = file_or_path.name.lower()
    else:
        fname = str(file_or_path).lower()

    if fname.endswith((".xlsx", ".xls")):
        df = pd.read_excel(file_or_path, engine="openpyxl")
    elif fname.endswith((".csv", ".tsv", ".txt")):
        try:
            df = pd.read_csv(file_or_path)
        except Exception:
            # Fallback for delimiter autodetection
            if hasattr(file_or_path, "seek"):
                file_or_path.seek(0)
            df = pd.read_csv(file_or_path, sep=None, engine="python")
    else:
        # Default try excel then csv
        try:
            df = pd.read_excel(file_or_path, engine="openpyxl")
        except Exception:
            if hasattr(file_or_path, "seek"):
                file_or_path.seek(0)
            df = pd.read_csv(file_or_path)

    df.columns = [str(c).strip() for c in df.columns]

    # Map raw columns to canonical concepts
    col_mapping = {}
    for col in df.columns:
        canonical = map_column_to_canonical(col)
        if canonical:
            col_mapping[col] = canonical

    # Clean numeric fields
    numeric_fields = [
        "base_salary", "overtime", "allowances", "bonus", "gross_pay",
        "tax", "pension", "insurance", "other_deductions", "total_deductions",
        "net_pay", "employer_contributions", "total_cost",
    ]

    for raw_col, canonical in col_mapping.items():
        if canonical in numeric_fields:
            df[raw_col] = clean_currency_series(df[raw_col])

    # Derive missing core calculations
    def get_or_zero(field_key: str) -> pd.Series:
        for raw_c, can in col_mapping.items():
            if can == field_key:
                return df[raw_c]
        return pd.Series(0.0, index=df.index)

    # Derive Gross Pay if not present
    has_gross = any(can == "gross_pay" for can in col_mapping.values())
    if not has_gross:
        base = get_or_zero("base_salary")
        ot = get_or_zero("overtime")
        allow = get_or_zero("allowances")
        bonus = get_or_zero("bonus")
        gross = base + ot + allow + bonus
        if gross.sum() > 0:
            df["Gross Pay"] = gross
            col_mapping["Gross Pay"] = "gross_pay"

    # Derive Total Deductions if not present
    has_deductions = any(can == "total_deductions" for can in col_mapping.values())
    if not has_deductions:
        tax = get_or_zero("tax")
        pension = get_or_zero("pension")
        insurance = get_or_zero("insurance")
        other = get_or_zero("other_deductions")
        total_ded = tax + pension + insurance + other
        if total_ded.sum() > 0:
            df["Total Deductions"] = total_ded
            col_mapping["Total Deductions"] = "total_deductions"

    # Derive Net Pay if not present
    has_net = any(can == "net_pay" for can in col_mapping.values())
    if not has_net:
        gross = get_or_zero("gross_pay")
        ded = get_or_zero("total_deductions")
        net = gross - ded
        if net.sum() != 0:
            df["Net Pay"] = net
            col_mapping["Net Pay"] = "net_pay"

    # Tag canonical mappings on DataFrame for downstream processors
    df.attrs["canonical_columns"] = col_mapping
    return df


def inspect_reference_template(template_file_or_path: Any) -> Dict[str, Any]:
    """Inspect reference template (.xlsx or .ods) to extract structure and styling.
    
    Returns sheet name, header row, column positions, template rows, and openpyxl styles.
    """
    is_path = isinstance(template_file_or_path, (str, Path))
    ext = os.path.splitext(str(template_file_or_path))[1].lower() if is_path else ""
    if hasattr(template_file_or_path, "name"):
        ext = os.path.splitext(template_file_or_path.name)[1].lower()

    # Handle .ods templates via pandas odf engine
    if ext == ".ods":
        try:
            template_df = pd.read_excel(template_file_or_path, engine="odf")
        except ImportError:
            raise ImportError("Reading .ods files requires 'odfpy'. Please install it or use .xlsx templates.")
        
        headers = [str(c).strip() for c in template_df.columns]
        rows = []
        for _, r in template_df.iterrows():
            rows.append([r[c] for c in template_df.columns])
        return {
            "type": "ods",
            "sheet_name": "Sheet1",
            "header_row_idx": 1,
            "columns": [(i + 1, h) for i, h in enumerate(headers)],
            "predefined_rows": rows,
            "workbook": None,
        }

    # Default to openpyxl for .xlsx
    if hasattr(template_file_or_path, "seek"):
        template_file_or_path.seek(0)
    wb = openpyxl.load_workbook(template_file_or_path, data_only=False)
    ws = wb.active

    # Detect header row (search first 15 rows for the best candidate)
    best_row_idx = 1
    best_score = -1
    best_columns: List[Tuple[int, str]] = []

    keyword_hints = ["name", "id", "salary", "pay", "department", "dept", "gross", "net", "deduction", "tax", "overtime", "total"]

    for row_idx in range(1, min(16, ws.max_row + 1)):
        row_vals = [ws.cell(row=row_idx, column=c).value for c in range(1, ws.max_column + 1)]
        non_empty = [(c, str(v).strip()) for c, v in enumerate(row_vals, start=1) if v is not None and str(v).strip() != ""]
        if not non_empty:
            continue
        
        # Score candidate row based on payroll keywords and number of valid column headers
        score = len(non_empty)
        for _, val in non_empty:
            val_norm = val.lower()
            if any(k in val_norm for k in keyword_hints):
                score += 3
        
        if score > best_score:
            best_score = score
            best_row_idx = row_idx
            best_columns = non_empty

    # Check if template has predefined rows (e.g. department names or fixed line items)
    predefined_rows = []
    for r in range(best_row_idx + 1, ws.max_row + 1):
        first_val = ws.cell(row=r, column=best_columns[0][0] if best_columns else 1).value
        # Ignore total / summary rows in template
        if first_val and any(kw in str(first_val).lower() for kw in ["total", "sum", "grand total", "summary"]):
            continue
        row_cells = [ws.cell(row=r, column=c).value for c, _ in best_columns]
        if any(v is not None for v in row_cells):
            predefined_rows.append((r, row_cells))

    return {
        "type": "xlsx",
        "sheet_name": ws.title,
        "header_row_idx": best_row_idx,
        "columns": best_columns,
        "predefined_rows": predefined_rows,
        "workbook": wb,
        "worksheet": ws,
    }


def validate_financial_audit(
    raw_df: pd.DataFrame,
    historical_df: Optional[pd.DataFrame] = None,
    variance_threshold: float = 0.20,
) -> List[Dict[str, Any]]:
    """Automated Financial Auditor.

    Performs audit checks before saving the final report:
    1. Reconciles individual payroll sub-totals against the grand total.
    2. Identifies major variances (> 20%) compared to previous rows,
       department cohort averages, and historical expectations.

    Returns:
        List of structured audit message dictionaries containing:
        'type' ('success', 'warning', 'info'), 'category', and 'message'.
    """
    messages: List[Dict[str, Any]] = []

    if raw_df.empty:
        messages.append({
            "type": "warning",
            "category": "Data Completeness",
            "message": "Audit Warning: Payroll dataset is empty. No financial records to audit.",
        })
        return messages

    col_mappings: Dict[str, str] = raw_df.attrs.get("canonical_columns", {})
    if not col_mappings:
        col_mappings = {}
        for col in raw_df.columns:
            can = map_column_to_canonical(col)
            if can:
                col_mappings[col] = can

    def get_canonical_series(field_key: str) -> pd.Series:
        for raw_c, can in col_mappings.items():
            if can == field_key and raw_c in raw_df.columns:
                return raw_df[raw_c].fillna(0.0).astype(float)
        return pd.Series(0.0, index=raw_df.index, dtype=float)

    # Core payroll component sums
    sum_base = float(get_canonical_series("base_salary").sum())
    sum_ot = float(get_canonical_series("overtime").sum())
    sum_bonus = float(get_canonical_series("bonus").sum())
    sum_allowances = float(get_canonical_series("allowances").sum())
    sum_gross = float(get_canonical_series("gross_pay").sum())

    sum_tax = float(get_canonical_series("tax").sum())
    sum_pension = float(get_canonical_series("pension").sum())
    sum_ins = float(get_canonical_series("insurance").sum())
    sum_other_ded = float(get_canonical_series("other_deductions").sum())
    sum_deductions = float(get_canonical_series("total_deductions").sum())

    sum_net = float(get_canonical_series("net_pay").sum())

    # -------------------------------------------------------------------------
    # Audit Check 1: Sub-totals vs. Grand Total Reconciliation
    # -------------------------------------------------------------------------
    # 1.1 Gross Pay reconciliation: Base + OT + Bonus + Allowances == Gross
    expected_gross = sum_base + sum_ot + sum_bonus + sum_allowances
    gross_discrepancy = abs(sum_gross - expected_gross)

    if sum_gross > 0:
        if gross_discrepancy <= 0.05:
            messages.append({
                "type": "success",
                "category": "Sub-total Reconciliation",
                "message": (
                    f"Gross Pay Balance Verified: Grand total Gross Pay (${sum_gross:,.2f}) "
                    f"perfect reconciles with sum of Base (${sum_base:,.2f}), Overtime (${sum_ot:,.2f}), "
                    f"Bonuses (${sum_bonus:,.2f}), and Allowances (${sum_allowances:,.2f})."
                ),
            })
        else:
            messages.append({
                "type": "warning",
                "category": "Sub-total Reconciliation",
                "message": (
                    f"Sub-total Discrepancy Detected: Grand Total Gross Pay (${sum_gross:,.2f}) "
                    f"does not match sum of individual earnings components (${expected_gross:,.2f}). "
                    f"Unreconciled Variance: ${gross_discrepancy:,.2f}."
                ),
            })

    # 1.2 Deductions reconciliation: Tax + Pension + Insurance + Other == Total Deductions
    expected_deductions = sum_tax + sum_pension + sum_ins + sum_other_ded
    ded_discrepancy = abs(sum_deductions - expected_deductions)

    if sum_deductions > 0:
        if ded_discrepancy <= 0.05:
            messages.append({
                "type": "success",
                "category": "Sub-total Reconciliation",
                "message": (
                    f"Deductions Balance Verified: Grand total Deductions (${sum_deductions:,.2f}) "
                    f"matches sum of Tax (${sum_tax:,.2f}), Pension (${sum_pension:,.2f}), "
                    f"and Insurance (${sum_ins:,.2f})."
                ),
            })
        else:
            messages.append({
                "type": "warning",
                "category": "Sub-total Reconciliation",
                "message": (
                    f"Deductions Discrepancy Detected: Grand Total Deductions (${sum_deductions:,.2f}) "
                    f"differs from sum of itemized deductions (${expected_deductions:,.2f}) "
                    f"by ${ded_discrepancy:,.2f}."
                ),
            })

    # 1.3 Net Pay reconciliation: Gross Pay - Total Deductions == Net Pay
    effective_deductions = sum_deductions if sum_deductions > 0 else expected_deductions
    expected_net = sum_gross - effective_deductions
    net_discrepancy = abs(sum_net - expected_net)

    if sum_net > 0:
        if net_discrepancy <= 0.05:
            messages.append({
                "type": "success",
                "category": "Sub-total Reconciliation",
                "message": (
                    f"Net Pay Reconciliation Verified: Grand total Net Pay (${sum_net:,.2f}) "
                    f"equals Total Gross Pay (${sum_gross:,.2f}) minus Total Deductions (${effective_deductions:,.2f})."
                ),
            })
        else:
            messages.append({
                "type": "warning",
                "category": "Sub-total Reconciliation",
                "message": (
                    f"Net Pay Discrepancy Detected: Grand total Net Pay (${sum_net:,.2f}) "
                    f"differs from (Gross - Deductions) (${expected_net:,.2f}) by ${net_discrepancy:,.2f}."
                ),
            })

    # -------------------------------------------------------------------------
    # Audit Check 2: Major Variance (> 20%) Checks
    # -------------------------------------------------------------------------
    # Identify employee name and department columns
    name_col = None
    dept_col = None
    for raw_c, can in col_mappings.items():
        if can == "employee_name":
            name_col = raw_c
        elif can == "department":
            dept_col = raw_c

    gross_series = get_canonical_series("gross_pay")
    base_series = get_canonical_series("base_salary")
    ot_series = get_canonical_series("overtime")
    bonus_series = get_canonical_series("bonus")

    variance_flags_count = 0

    # 2.1 Sequential (Row-to-Row) Variance (> 20%)
    if len(raw_df) > 1:
        for idx in range(1, len(raw_df)):
            curr_val = float(gross_series.iloc[idx])
            prev_val = float(gross_series.iloc[idx - 1])
            
            curr_name = raw_df[name_col].iloc[idx] if name_col else f"Row {idx+1}"
            prev_name = raw_df[name_col].iloc[idx - 1] if name_col else f"Row {idx}"
            dept_name = f" ({raw_df[dept_col].iloc[idx]})" if dept_col else ""

            if prev_val > 0:
                pct_change = (curr_val - prev_val) / prev_val
                if abs(pct_change) > variance_threshold:
                    variance_flags_count += 1
                    direction = "increase" if pct_change > 0 else "decrease"
                    messages.append({
                        "type": "warning",
                        "category": "Sequential Variance (>20%)",
                        "message": (
                            f"Row-to-Row Variance Alert: {curr_name}{dept_name} Gross Pay (${curr_val:,.2f}) "
                            f"shows a {abs(pct_change)*100:.1f}% {direction} compared to previous record "
                            f"'{prev_name}' (${prev_val:,.2f})."
                        ),
                    })

    # 2.2 Historical / Benchmark Expectations Variance (> 20%)
    if historical_df is not None and not historical_df.empty:
        # Compare department totals against historical baseline
        hist_col_mappings: Dict[str, str] = historical_df.attrs.get("canonical_columns", {})
        hist_dept_col = None
        for raw_c, can in hist_col_mappings.items():
            if can == "department":
                hist_dept_col = raw_c

        if dept_col and hist_dept_col and dept_col in raw_df.columns and hist_dept_col in historical_df.columns:
            curr_dept_totals = raw_df.groupby(dept_col)[raw_df.columns[0]].count() if "gross_pay" not in col_mappings else raw_df.groupby(dept_col)[list(col_mappings.keys())[list(col_mappings.values()).index("gross_pay")]].sum()
            hist_dept_totals = historical_df.groupby(hist_dept_col)[historical_df.columns[0]].count() if "gross_pay" not in hist_col_mappings else historical_df.groupby(hist_dept_col)[list(hist_col_mappings.keys())[list(hist_col_mappings.values()).index("gross_pay")]].sum()

            for d_name, curr_total in curr_dept_totals.items():
                if d_name in hist_dept_totals:
                    hist_total = float(hist_dept_totals[d_name])
                    if hist_total > 0:
                        dept_variance = (float(curr_total) - hist_total) / hist_total
                        if abs(dept_variance) > variance_threshold:
                            variance_flags_count += 1
                            v_dir = "higher" if dept_variance > 0 else "lower"
                            messages.append({
                                "type": "warning",
                                "category": "Historical Variance (>20%)",
                                "message": (
                                    f"Historical Variance Alert: Department '{d_name}' payroll (${curr_total:,.2f}) "
                                    f"is {abs(dept_variance)*100:.1f}% {v_dir} than historical expectation (${hist_total:,.2f})."
                                ),
                            })
    else:
        # Use cohort benchmark (department median/average) as expectation
        if dept_col and dept_col in raw_df.columns and not gross_series.empty:
            df_cohort = raw_df.copy()
            df_cohort["_gross"] = gross_series
            dept_means = df_cohort.groupby(dept_col)["_gross"].mean()

            for idx, r_data in df_cohort.iterrows():
                d_name = r_data[dept_col]
                d_mean = float(dept_means.get(d_name, 0.0))
                g_val = float(r_data["_gross"])
                if d_mean > 0:
                    cohort_diff = (g_val - d_mean) / d_mean
                    if abs(cohort_diff) > variance_threshold:
                        emp_lbl = r_data[name_col] if name_col else f"Record {idx+1}"
                        # Log as informative notice if significantly above cohort average
                        messages.append({
                            "type": "info",
                            "category": "Cohort Variance (>20%)",
                            "message": (
                                f"Cohort Variance Notice: {emp_lbl} ({d_name}) Gross Pay (${g_val:,.2f}) "
                                f"deviates by {cohort_diff*100:+.1f}% from the department average (${d_mean:,.2f})."
                            ),
                        })

    # 2.3 Component Ratio Outlier Checks (> 20% of Base Salary)
    for idx in range(len(raw_df)):
        b_val = float(base_series.iloc[idx])
        ot_val = float(ot_series.iloc[idx])
        bonus_val = float(bonus_series.iloc[idx])
        emp_lbl = raw_df[name_col].iloc[idx] if name_col else f"Row {idx+1}"
        d_lbl = f" ({raw_df[dept_col].iloc[idx]})" if dept_col else ""

        if b_val > 0:
            if (ot_val / b_val) > variance_threshold:
                variance_flags_count += 1
                messages.append({
                    "type": "warning",
                    "category": "Overtime Surge (>20%)",
                    "message": (
                        f"Overtime Surge Alert: {emp_lbl}{d_lbl} overtime pay (${ot_val:,.2f}) "
                        f"represents {(ot_val/b_val)*100:.1f}% of base salary (exceeds 20% threshold)."
                    ),
                })
            if (bonus_val / b_val) > variance_threshold:
                messages.append({
                    "type": "info",
                    "category": "Incentive Spike (>20%)",
                    "message": (
                        f"Incentive Spike Notice: {emp_lbl}{d_lbl} bonus (${bonus_val:,.2f}) "
                        f"represents {(bonus_val/b_val)*100:.1f}% of base salary."
                    ),
                })

    if variance_flags_count == 0:
        messages.append({
            "type": "success",
            "category": "Variance Compliance",
            "message": "Variance Audit Passed: No abnormal volatility or unexpected >20% spikes detected across records.",
        })

    return messages


def find_matching_raw_column(template_col_name: str, raw_df: pd.DataFrame) -> Optional[str]:
    """Find the best matching raw dataframe column for a given template column."""
    # 1. Exact match (case insensitive)
    for c in raw_df.columns:
        if c.strip().lower() == template_col_name.strip().lower():
            return c
        if normalize_text(c) == normalize_text(template_col_name):
            return c

    # 2. Canonical semantic alias matching
    t_canonical = map_column_to_canonical(template_col_name)
    if t_canonical:
        # Check canonical mapping tagged on raw_df
        col_mappings: Dict[str, str] = raw_df.attrs.get("canonical_columns", {})
        if not col_mappings:
            col_mappings = {}
            for col in raw_df.columns:
                can = map_column_to_canonical(col)
                if can:
                    col_mappings[col] = can
        for raw_c, can in col_mappings.items():
            if can == t_canonical:
                return raw_c
        # If not tagged, check raw_df columns
        for c in raw_df.columns:
            if map_column_to_canonical(c) == t_canonical:
                return c

    # 3. Partial substring matching
    norm_t = normalize_text(template_col_name)
    for c in raw_df.columns:
        norm_c = normalize_text(c)
        if (len(norm_t) >= 4 and norm_t in norm_c) or (len(norm_c) >= 4 and norm_c in norm_t):
            return c

    return None


def generate_payroll_report(
    raw_data_input: Any,
    template_input: Any,
    output_path: Optional[Union[str, Path]] = None,
    historical_input: Optional[Any] = None,
    return_audit: bool = False,
    variance_threshold: float = 0.20,
) -> Union[io.BytesIO, Tuple[io.BytesIO, List[Dict[str, Any]]]]:
    """Generate a clean, formatted financial payroll Excel report matching reference template.

    Performs automated financial auditing before saving:
    - Checks that individual payroll sub-totals match the grand total.
    - Checks for major variances (>20%) compared to previous records or historical expectations.

    Args:
        raw_data_input: File path or file-like buffer for raw payroll data (CSV or Excel).
        template_input: File path or file-like buffer for reference template (.xlsx or .ods).
        output_path: Optional file path to save the generated Excel workbook.
        historical_input: Optional prior period raw data for historical variance auditing.
        return_audit: If True, returns (output_buffer, audit_messages).
        variance_threshold: Volatility variance threshold (default 0.20 = 20%).

    Returns:
        io.BytesIO stream (with attached .audit_messages) or tuple (io.BytesIO, audit_messages).
    """
    raw_df = load_raw_payroll_data(raw_data_input)
    t_info = inspect_reference_template(template_input)

    historical_df = None
    if historical_input is not None:
        try:
            historical_df = load_raw_payroll_data(historical_input)
        except Exception:
            historical_df = None

    # Styling constants
    header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    data_font = Font(name="Calibri", size=10, color="111827")
    total_font = Font(name="Calibri", size=11, bold=True, color="000000")
    total_fill = PatternFill(start_color="E5E7EB", end_color="E5E7EB", fill_type="solid")
    zebra_fill = PatternFill(start_color="F9FAFB", end_color="F9FAFB", fill_type="solid")

    thin_border_side = Side(border_style="thin", color="D1D5DB")
    grid_border = Border(
        left=thin_border_side,
        right=thin_border_side,
        top=thin_border_side,
        bottom=thin_border_side,
    )
    total_border = Border(
        top=Side(border_style="thin", color="000000"),
        bottom=Side(border_style="double", color="000000"),
        left=thin_border_side,
        right=thin_border_side,
    )

    currency_fmt = "$#,##0.00"
    count_fmt = "#,##0"

    template_cols: List[Tuple[int, str]] = t_info["columns"]
    header_row_idx: int = t_info["header_row_idx"]

    # Determine report granularity: Predefined rows (e.g. Department Summary) vs Detailed Ledger
    is_predefined_summary = False
    dept_raw_col = None
    for raw_c, can in raw_df.attrs.get("canonical_columns", {}).items():
        if can == "department":
            dept_raw_col = raw_c
            break

    # Check if first column in predefined template corresponds to departments
    first_col_template_label = template_cols[0][1] if template_cols else ""
    first_col_canonical = map_column_to_canonical(first_col_template_label)
    
    if first_col_canonical == "department" and dept_raw_col is not None:
        # Check if template has rows or if raw data should be aggregated by department
        is_predefined_summary = True

    # Prepare openpyxl workbook
    if t_info["type"] == "xlsx" and t_info["workbook"] is not None:
        wb = t_info["workbook"]
        ws = t_info["worksheet"]
        # Clear existing data rows below header row to prevent leftover template placeholders
        max_r = ws.max_row
        if max_r > header_row_idx:
            for r in range(header_row_idx + 1, max_r + 1):
                for c in range(1, ws.max_column + 1):
                    ws.cell(row=r, column=c).value = None
    else:
        # Build fresh workbook matching template columns
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Financial Payroll Report"
        
        # Write headers
        for col_idx, (_, col_name) in enumerate(template_cols, start=1):
            cell = ws.cell(row=header_row_idx, column=col_idx, value=col_name)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = grid_border

    # Build mapping from template column index to raw DataFrame column or formula
    col_mapping: Dict[int, Optional[str]] = {}
    for c_pos, col_name in template_cols:
        matched_raw = find_matching_raw_column(col_name, raw_df)
        col_mapping[c_pos] = matched_raw

    # Track numeric columns for totaling and formatting
    numeric_template_cols: set = set()

    current_write_row = header_row_idx + 1

    if is_predefined_summary and dept_raw_col:
        # Aggregate raw data by department
        departments = sorted(raw_df[dept_raw_col].dropna().unique().tolist())
        for r_offset, dept_name in enumerate(departments):
            dept_mask = raw_df[dept_raw_col] == dept_name
            dept_sub_df = raw_df[dept_mask]

            for c_pos, col_name in template_cols:
                cell = ws.cell(row=current_write_row, column=c_pos)
                raw_c = col_mapping.get(c_pos)
                can_concept = map_column_to_canonical(col_name)

                # Determine value
                if can_concept == "department":
                    cell.value = dept_name
                    cell.alignment = Alignment(horizontal="left", vertical="center")
                elif can_concept in ["employee_id", "employee_name"] or "count" in col_name.lower() or "headcount" in col_name.lower():
                    # Headcount / employee count
                    cell.value = len(dept_sub_df)
                    cell.number_format = count_fmt
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                    numeric_template_cols.add(c_pos)
                elif raw_c and pd.api.types.is_numeric_dtype(dept_sub_df[raw_c]):
                    sum_val = float(dept_sub_df[raw_c].sum())
                    cell.value = sum_val
                    cell.number_format = currency_fmt
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                    numeric_template_cols.add(c_pos)
                else:
                    cell.value = ""

                cell.font = data_font
                cell.border = grid_border
                if r_offset % 2 == 1:
                    cell.fill = zebra_fill

            current_write_row += 1

    else:
        # Tabular Employee Register mapping
        for r_offset, (_, row_data) in enumerate(raw_df.iterrows()):
            for c_pos, col_name in template_cols:
                cell = ws.cell(row=current_write_row, column=c_pos)
                raw_c = col_mapping.get(c_pos)

                if raw_c and raw_c in row_data:
                    val = row_data[raw_c]
                    can_concept = map_column_to_canonical(col_name)
                    
                    if pd.isna(val):
                        cell.value = ""
                    elif isinstance(val, (int, float)):
                        cell.value = float(val)
                        if can_concept in ["base_salary", "overtime", "allowances", "bonus", "gross_pay", "tax", "pension", "insurance", "other_deductions", "total_deductions", "net_pay", "employer_contributions", "total_cost"]:
                            cell.number_format = currency_fmt
                            numeric_template_cols.add(c_pos)
                        else:
                            cell.number_format = count_fmt
                        cell.alignment = Alignment(horizontal="right", vertical="center")
                    else:
                        cell.value = str(val).strip()
                        cell.alignment = Alignment(horizontal="left", vertical="center")
                else:
                    cell.value = ""

                cell.font = data_font
                cell.border = grid_border
                if r_offset % 2 == 1:
                    cell.fill = zebra_fill

            current_write_row += 1

    # -------------------------------------------------------------------------
    # Calculate Totals Row
    # -------------------------------------------------------------------------
    start_data_row = header_row_idx + 1
    end_data_row = current_write_row - 1

    if end_data_row >= start_data_row:
        total_row = current_write_row
        
        # Determine first column label
        first_col_pos = template_cols[0][0]
        ws.cell(row=total_row, column=first_col_pos, value="Total").alignment = Alignment(horizontal="left", vertical="center")
        
        for c_pos, col_name in template_cols:
            cell = ws.cell(row=total_row, column=c_pos)
            cell.font = total_font
            cell.fill = total_fill
            cell.border = total_border

            if c_pos in numeric_template_cols:
                col_letter = get_column_letter(c_pos)
                cell.value = f"=SUM({col_letter}{start_data_row}:{col_letter}{end_data_row})"
                # Determine formatting
                can_concept = map_column_to_canonical(col_name)
                if can_concept in ["employee_id", "employee_name"] or "count" in col_name.lower():
                    cell.number_format = count_fmt
                else:
                    cell.number_format = currency_fmt
                cell.alignment = Alignment(horizontal="right", vertical="center")
            elif c_pos != first_col_pos and cell.value is None:
                cell.value = ""

    # Auto-fit column widths
    ws.views.sheetView[0].showGridLines = True
    for c_pos, _ in template_cols:
        col_letter = get_column_letter(c_pos)
        max_len = 0
        for r in range(header_row_idx, current_write_row + 1):
            val = ws.cell(row=r, column=c_pos).value
            if val is not None:
                val_str = str(val)
                # Ignore Excel formula strings for length
                if val_str.startswith("="):
                    val_str = "$999,999.00"
                max_len = max(max_len, len(val_str))
        ws.column_dimensions[col_letter].width = max(max_len + 4, 14)

    # -------------------------------------------------------------------------
    # Generate Formatted Preview DataFrame for On-Screen Display
    # -------------------------------------------------------------------------
    preview_records = []
    col_names = [name for _, name in template_cols]
    for r in range(header_row_idx + 1, current_write_row):
        row_dict = {}
        for c_pos, name in template_cols:
            val = ws.cell(row=r, column=c_pos).value
            row_dict[name] = val
        preview_records.append(row_dict)

    preview_df = pd.DataFrame(preview_records)

    # Compute actual totals row for preview display
    if not preview_df.empty:
        total_dict = {col_names[0]: "Total"}
        for name in col_names[1:]:
            if name in preview_df.columns:
                num_s = pd.to_numeric(preview_df[name], errors="coerce")
                if num_s.notna().any() and not num_s.dropna().empty:
                    total_dict[name] = float(num_s.sum())
                else:
                    total_dict[name] = ""
            else:
                total_dict[name] = ""
        preview_df_with_totals = pd.concat([preview_df, pd.DataFrame([total_dict])], ignore_index=True)
    else:
        preview_df_with_totals = preview_df

    # -------------------------------------------------------------------------
    # Automated Financial Auditor: Validate sub-totals & major variances (>20%)
    # -------------------------------------------------------------------------
    audit_messages = validate_financial_audit(
        raw_df=raw_df,
        historical_df=historical_df,
        variance_threshold=variance_threshold,
    )

    # Save to buffer
    output_stream = io.BytesIO()
    wb.save(output_stream)
    output_stream.seek(0)
    output_stream.audit_messages = audit_messages
    output_stream.preview_df = preview_df_with_totals

    # Optional save to disk
    if output_path:
        out_p = Path(output_path)
        out_p.parent.mkdir(parents=True, exist_ok=True)
        with open(out_p, "wb") as f:
            f.write(output_stream.getvalue())

    if return_audit:
        return output_stream, audit_messages, preview_df_with_totals
    return output_stream


def create_sample_payroll_files(
    raw_output_path: str = "data/raw_payroll_data.csv",
    template_output_path: str = "data/payroll_report_template.xlsx",
) -> Tuple[str, str]:
    """Create sample raw payroll dataset and a reference template .xlsx for testing."""
    os.makedirs(os.path.dirname(raw_output_path) or ".", exist_ok=True)
    os.makedirs(os.path.dirname(template_output_path) or ".", exist_ok=True)

    # 1. Raw Payroll Data
    raw_df = pd.DataFrame([
        {
            "Emp_ID": "EMP-101",
            "Full_Name": "Alice Johnson",
            "Department": "Engineering",
            "Role": "Lead Architect",
            "Base_Salary": 9500.0,
            "Overtime_Pay": 600.0,
            "Bonus": 1200.0,
            "Tax_Deduction": 2486.0,
            "Pension_Contribution": 570.0,
            "Health_Insurance": 250.0,
        },
        {
            "Emp_ID": "EMP-102",
            "Full_Name": "Bob Smith",
            "Department": "Engineering",
            "Role": "Fullstack Engineer",
            "Base_Salary": 7800.0,
            "Overtime_Pay": 450.0,
            "Bonus": 800.0,
            "Tax_Deduction": 1991.0,
            "Pension_Contribution": 468.0,
            "Health_Insurance": 250.0,
        },
        {
            "Emp_ID": "EMP-103",
            "Full_Name": "Carol Davis",
            "Department": "Sales",
            "Role": "Account Executive",
            "Base_Salary": 6200.0,
            "Overtime_Pay": 0.0,
            "Bonus": 3500.0,
            "Tax_Deduction": 2134.0,
            "Pension_Contribution": 372.0,
            "Health_Insurance": 250.0,
        },
        {
            "Emp_ID": "EMP-104",
            "Full_Name": "David Miller",
            "Department": "Sales",
            "Role": "Sales Manager",
            "Base_Salary": 8200.0,
            "Overtime_Pay": 0.0,
            "Bonus": 4200.0,
            "Tax_Deduction": 2728.0,
            "Pension_Contribution": 492.0,
            "Health_Insurance": 250.0,
        },
        {
            "Emp_ID": "EMP-105",
            "Full_Name": "Elena Rostova",
            "Department": "Marketing",
            "Role": "Growth Lead",
            "Base_Salary": 7100.0,
            "Overtime_Pay": 200.0,
            "Bonus": 1500.0,
            "Tax_Deduction": 1936.0,
            "Pension_Contribution": 426.0,
            "Health_Insurance": 250.0,
        },
        {
            "Emp_ID": "EMP-106",
            "Full_Name": "Frank Wilson",
            "Department": "Operations",
            "Role": "Operations Director",
            "Base_Salary": 8900.0,
            "Overtime_Pay": 0.0,
            "Bonus": 1000.0,
            "Tax_Deduction": 2178.0,
            "Pension_Contribution": 534.0,
            "Health_Insurance": 250.0,
        },
    ])
    raw_df.to_csv(raw_output_path, index=False)

    # 2. Reference Template (.xlsx)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Payroll Summary"
    
    # Styled Template Headers
    headers = [
        "Employee ID",
        "Employee Name",
        "Department",
        "Base Salary",
        "Overtime",
        "Bonus",
        "Gross Pay",
        "Tax",
        "Pension",
        "Total Deductions",
        "Net Pay",
    ]
    
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    
    for c_idx, h in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=c_idx, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.column_dimensions[get_column_letter(c_idx)].width = max(len(h) + 4, 14)
        
    wb.save(template_output_path)
    return raw_output_path, template_output_path


def main():
    """Command-line interface for the payroll financial report generator."""
    parser = argparse.ArgumentParser(
        description="Extract payroll data, map to reference template, calculate totals, and output clean Excel report."
    )
    parser.add_argument("raw_data", nargs="?", help="Path to raw payroll data file (CSV or Excel)")
    parser.add_argument("template", nargs="?", help="Path to reference template file (.xlsx or .ods)")
    parser.add_argument("-o", "--output", default="generated_payroll_report.xlsx", help="Output .xlsx file path")
    parser.add_argument("--create-samples", action="store_true", help="Generate sample raw data and reference template")

    args = parser.parse_args()

    if args.create_samples:
        r_path, t_path = create_sample_payroll_files()
        print(f"Sample raw data created at: {r_path}")
        print(f"Sample template created at: {t_path}")
        return

    if not args.raw_data or not args.template:
        print("Usage: python -m utils.processor <raw_data_file> <reference_template_file> [-o output.xlsx]")
        print("Run with --create-samples to create sample files.")
        return

    print(f"Processing '{args.raw_data}' against template '{args.template}'...")
    res = generate_payroll_report(
        args.raw_data,
        args.template,
        output_path=args.output,
        return_audit=True,
    )
    audit_msgs = res[1]
    print(f"Report successfully generated and saved to: {args.output}\n")
    print("=== AUTOMATED FINANCIAL AUDIT REPORT ===")
    for msg in audit_msgs:
        tag = f"[{msg['type'].upper()}]"
        print(f"{tag:<10} ({msg['category']}): {msg['message']}")


if __name__ == "__main__":
    main()

