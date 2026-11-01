"""Excel report generation utility using openpyxl."""

import io
from datetime import datetime
from typing import Dict, Any
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import pandas as pd


def export_financial_report(
    kpis: Dict[str, Any],
    monthly_df: pd.DataFrame,
    dept_df: pd.DataFrame,
    category_df: pd.DataFrame,
    raw_df: pd.DataFrame,
) -> io.BytesIO:
    """Generate a multi-tab formatted Excel workbook containing complete financial reporting."""
    wb = openpyxl.Workbook()
    
    # Setup styles
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    title_font = Font(name="Calibri", size=16, bold=True, color="1E3A8A")
    subtitle_font = Font(name="Calibri", size=10, italic=True, color="4B5563")
    metric_label_font = Font(name="Calibri", size=10, color="374151")
    metric_val_font = Font(name="Calibri", size=14, bold=True, color="111827")
    
    primary_header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    alt_row_fill = PatternFill(start_color="F9FAFB", end_color="F9FAFB", fill_type="solid")
    kpi_card_fill = PatternFill(start_color="EFF6FF", end_color="EFF6FF", fill_type="solid")

    thin_border_side = Side(border_style="thin", color="D1D5DB")
    grid_border = Border(
        left=thin_border_side,
        right=thin_border_side,
        top=thin_border_side,
        bottom=thin_border_side,
    )
    
    # -------------------------------------------------------------------------
    # Sheet 1: Executive Summary
    # -------------------------------------------------------------------------
    ws_summary = wb.active
    ws_summary.title = "Executive Summary"
    ws_summary.views.sheetView[0].showGridLines = True

    ws_summary["B2"] = "Financial Performance Executive Report"
    ws_summary["B2"].font = title_font
    
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    ws_summary["B3"] = f"Report Generated: {now_str}"
    ws_summary["B3"].font = subtitle_font

    # KPI Table / Cards
    kpi_items = [
        ("Total Revenue", kpis.get("total_revenue", 0.0), "$#,##0.00"),
        ("Total Expenses", kpis.get("total_expenses", 0.0), "$#,##0.00"),
        ("Net Profit / (Loss)", kpis.get("net_profit", 0.0), "$#,##0.00"),
        ("Profit Margin", (kpis.get("profit_margin_pct", 0.0) / 100.0), "0.0%"),
        ("Average Monthly Revenue", kpis.get("avg_monthly_revenue", 0.0), "$#,##0.00"),
        ("Average Monthly Expense", kpis.get("avg_monthly_expense", 0.0), "$#,##0.00"),
        ("Total Transactions", kpis.get("transaction_count", 0), "#,##0"),
    ]

    ws_summary["B5"] = "Key Financial Indicators"
    ws_summary["B5"].font = Font(name="Calibri", size=12, bold=True, color="1E3A8A")

    ws_summary["B6"] = "Metric"
    ws_summary["C6"] = "Value"
    for col_cell in ["B6", "C6"]:
        ws_summary[col_cell].font = header_font
        ws_summary[col_cell].fill = primary_header_fill
        ws_summary[col_cell].alignment = Alignment(horizontal="center", vertical="center")

    for idx, (label, val, num_fmt) in enumerate(kpi_items, start=7):
        cell_lbl = ws_summary[f"B{idx}"]
        cell_val = ws_summary[f"C{idx}"]
        
        cell_lbl.value = label
        cell_lbl.font = metric_label_font
        cell_lbl.border = grid_border
        
        cell_val.value = val
        cell_val.font = metric_val_font if "Net Profit" in label else Font(name="Calibri", size=11, bold=True)
        cell_val.number_format = num_fmt
        cell_val.border = grid_border
        cell_val.alignment = Alignment(horizontal="right")
        
        if idx % 2 == 0:
            cell_lbl.fill = alt_row_fill
            cell_val.fill = alt_row_fill

    # -------------------------------------------------------------------------
    # Helper to append styled DataFrames
    # -------------------------------------------------------------------------
    def append_dataframe_to_sheet(sheet, df: pd.DataFrame, currency_cols=None, pct_cols=None):
        sheet.views.sheetView[0].showGridLines = True
        currency_cols = currency_cols or []
        pct_cols = pct_cols or []

        # Headers
        sheet.append(list(df.columns))
        for cell in sheet[1]:
            cell.font = header_font
            cell.fill = primary_header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center")

        # Rows
        for r_idx, row in enumerate(df.itertuples(index=False), start=2):
            sheet.append(list(row))
            for c_idx, val in enumerate(row, start=1):
                cell = sheet.cell(row=r_idx, column=c_idx)
                cell.border = grid_border
                col_name = df.columns[c_idx - 1]

                if r_idx % 2 == 1:
                    cell.fill = alt_row_fill

                if col_name in currency_cols:
                    cell.number_format = "$#,##0.00"
                    cell.alignment = Alignment(horizontal="right")
                elif col_name in pct_cols:
                    # In case percentage is 0-100, normalize if needed or format as standard
                    cell.number_format = "0.0\"%\""
                    cell.alignment = Alignment(horizontal="right")
                elif isinstance(val, (int, float)):
                    cell.alignment = Alignment(horizontal="right")

        # Adjust column widths
        for col in sheet.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                val_str = str(cell.value or "")
                max_len = max(max_len, len(val_str))
            sheet.column_dimensions[col_letter].width = max(max_len + 4, 14)

    # -------------------------------------------------------------------------
    # Sheet 2: Monthly Trends
    # -------------------------------------------------------------------------
    ws_monthly = wb.create_sheet(title="Monthly Trends")
    append_dataframe_to_sheet(
        ws_monthly,
        monthly_df,
        currency_cols=["Revenue", "Expense", "Net Income"],
        pct_cols=["Margin %"],
    )

    # -------------------------------------------------------------------------
    # Sheet 3: Department Summary
    # -------------------------------------------------------------------------
    ws_dept = wb.create_sheet(title="Department Breakdown")
    append_dataframe_to_sheet(
        ws_dept,
        dept_df,
        currency_cols=["Revenue", "Expense", "Net"],
    )

    # -------------------------------------------------------------------------
    # Sheet 4: Expense Categories
    # -------------------------------------------------------------------------
    ws_cat = wb.create_sheet(title="Expense Breakdown")
    append_dataframe_to_sheet(
        ws_cat,
        category_df,
        currency_cols=["Total Amount"],
        pct_cols=["Share %"],
    )

    # -------------------------------------------------------------------------
    # Sheet 5: Raw Transactions
    # -------------------------------------------------------------------------
    ws_raw = wb.create_sheet(title="Transactions")
    raw_export = raw_df.copy()
    if "Date" in raw_export.columns:
        raw_export["Date"] = raw_export["Date"].dt.strftime("%Y-%m-%d")
    append_dataframe_to_sheet(
        ws_raw,
        raw_export,
        currency_cols=["Amount"],
    )

    # Auto-adjust column widths for Executive Summary
    ws_summary.column_dimensions["A"].width = 4
    ws_summary.column_dimensions["B"].width = 28
    ws_summary.column_dimensions["C"].width = 22

    # Save to buffer
    output_stream = io.BytesIO()
    wb.save(output_stream)
    output_stream.seek(0)
    return output_stream
