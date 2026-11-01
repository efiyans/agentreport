"""Streamlit Financial Reporting Dashboard Application."""

import streamlit as st
import pandas as pd
from datetime import datetime

from utils import (
    generate_sample_financial_data,
    load_financial_data,
    calculate_kpis,
    get_monthly_breakdown,
    get_category_breakdown,
    get_department_breakdown,
    export_financial_report,
    generate_payroll_report,
    create_sample_payroll_files,
)


def format_currency(val: float) -> str:
    """Format floating point numbers as dollar currency."""
    if pd.isna(val):
        return "$0.00"
    if val < 0:
        return f"-${abs(val):,.2f}"
    return f"${val:,.2f}"


def main():
    st.set_page_config(
        page_title="Financial Reporting Dashboard",
        page_icon="📊",
        layout="wide",
        initial_sidebar_state="expanded",
    )

    # Custom styling
    st.markdown(
        """
        <style>
        .main-header {
            font-size: 2.2rem;
            font-weight: 700;
            color: #1E3A8A;
            margin-bottom: 0.2rem;
        }
        .sub-header {
            font-size: 1.05rem;
            color: #4B5563;
            margin-bottom: 1.5rem;
        }
        .metric-container {
            background-color: #F8FAFC;
            border: 1px solid #E2E8F0;
            border-radius: 8px;
            padding: 16px;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    st.markdown('<div class="main-header">📊 Financial Reporting Dashboard</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="sub-header">Executive financial monitoring, budgeting analytics, and Excel report export.</div>',
        unsafe_allow_html=True,
    )

    # -------------------------------------------------------------------------
    # Sidebar: Data Source & Filtering
    # -------------------------------------------------------------------------
    with st.sidebar:
        st.header("⚙️ Data Configuration")
        data_source = st.radio(
            "Select Data Source:",
            options=["Sample Financial Data", "Upload Excel / CSV"],
            index=0,
        )

        raw_df = None
        if data_source == "Upload Excel / CSV":
            uploaded_file = st.file_uploader(
                "Upload Financial File",
                type=["xlsx", "xls", "csv"],
                help="File must contain: Date, Type (Revenue/Expense), Category, and Amount columns.",
            )
            if uploaded_file is not None:
                try:
                    raw_df = load_financial_data(uploaded_file)
                    st.success(f"Loaded {len(raw_df)} records successfully!")
                except Exception as e:
                    st.error(f"Error reading file: {e}")
                    st.stop()
            else:
                st.info("Please upload a .xlsx or .csv financial file, or switch to Sample Data.")
                st.stop()
        else:
            if "sample_data" not in st.session_state:
                st.session_state.sample_data = generate_sample_financial_data()
            raw_df = st.session_state.sample_data
            if st.button("🔄 Regenerate Sample Data"):
                st.session_state.sample_data = generate_sample_financial_data()
                raw_df = st.session_state.sample_data
                st.rerun()

        st.divider()
        st.header("🔍 Global Filters")

        # Date range filter
        min_date = raw_df["Date"].min().date()
        max_date = raw_df["Date"].max().date()
        date_range = st.date_input(
            "Select Reporting Period:",
            value=(min_date, max_date),
            min_value=min_date,
            max_value=max_date,
        )

        # Department filter
        all_departments = sorted(raw_df["Department"].dropna().unique().tolist())
        selected_depts = st.multiselect(
            "Filter by Department:",
            options=all_departments,
            default=all_departments,
        )

        # Transaction type filter
        all_types = ["All"] + sorted(raw_df["Type"].dropna().unique().tolist())
        selected_type = st.selectbox("Transaction Type:", options=all_types, index=0)

    # -------------------------------------------------------------------------
    # Apply Filters
    # -------------------------------------------------------------------------
    filtered_df = raw_df.copy()
    
    if isinstance(date_range, (tuple, list)) and len(date_range) == 2:
        start_d, end_d = date_range
        filtered_df = filtered_df[
            (filtered_df["Date"].dt.date >= start_d) & (filtered_df["Date"].dt.date <= end_d)
        ]
    elif isinstance(date_range, (tuple, list)) and len(date_range) == 1:
        start_d = date_range[0]
        filtered_df = filtered_df[filtered_df["Date"].dt.date >= start_d]

    if selected_depts:
        filtered_df = filtered_df[filtered_df["Department"].isin(selected_depts)]

    if selected_type != "All":
        filtered_df = filtered_df[filtered_df["Type"] == selected_type]

    if filtered_df.empty:
        st.warning("No financial transactions match the selected filters. Please adjust filters.")
        st.stop()

    # Compute Core Analytics
    kpis = calculate_kpis(filtered_df)
    monthly_summary = get_monthly_breakdown(filtered_df)
    category_summary = get_category_breakdown(filtered_df, transaction_type="Expense")
    revenue_cat_summary = get_category_breakdown(filtered_df, transaction_type="Revenue")
    dept_summary = get_department_breakdown(filtered_df)

    # -------------------------------------------------------------------------
    # Executive KPI Metric Cards
    # -------------------------------------------------------------------------
    col1, col2, col3, col4 = st.columns(4)

    with col1:
        st.metric(
            label="Total Revenue",
            value=format_currency(kpis["total_revenue"]),
            help="Total gross revenue earned across filtered periods.",
        )
    with col2:
        st.metric(
            label="Total Expenses",
            value=format_currency(kpis["total_expenses"]),
            help="Total operating expenses recorded.",
        )
    with col3:
        net = kpis["net_profit"]
        st.metric(
            label="Net Profit / (Loss)",
            value=format_currency(net),
            delta=f"{kpis['profit_margin_pct']:.1f}% Margin",
            delta_color="normal" if net >= 0 else "inverse",
        )
    with col4:
        st.metric(
            label="Monthly Avg Revenue",
            value=format_currency(kpis["avg_monthly_revenue"]),
            help=f"Based on {kpis['transaction_count']} filtered transactions.",
        )

    st.markdown("<br>", unsafe_allow_html=True)

    # -------------------------------------------------------------------------
    # Dashboard Tabs
    # -------------------------------------------------------------------------
    tab_overview, tab_monthly, tab_departments, tab_transactions, tab_export, tab_payroll = st.tabs([
        "📈 Overview & Trends",
        "📅 Monthly Performance",
        "🏢 Department & Category Breakdown",
        "📋 Transaction Details",
        "📥 Report Exporter (Excel / CSV)",
        "📑 Payroll & Template Mapping",
    ])

    # Tab 1: Overview & Trends
    with tab_overview:
        col_chart1, col_chart2 = st.columns(2)

        with col_chart1:
            st.subheader("Monthly Revenue vs. Expense")
            if not monthly_summary.empty:
                chart_data = monthly_summary.set_index("YearMonth")[["Revenue", "Expense"]]
                st.bar_chart(chart_data, color=["#10B981", "#EF4444"])
            else:
                st.info("No monthly data available.")

        with col_chart2:
            st.subheader("Net Income Trend")
            if not monthly_summary.empty:
                net_chart_data = monthly_summary.set_index("YearMonth")[["Net Income"]]
                st.line_chart(net_chart_data, color=["#2563EB"])
            else:
                st.info("No monthly data available.")

        st.subheader("Top Expense Categories")
        if not category_summary.empty:
            cat_chart_data = category_summary.set_index("Category")[["Total Amount"]]
            st.bar_chart(cat_chart_data, horizontal=True, color="#F59E0B")
        else:
            st.info("No expense category data available.")

    # Tab 2: Monthly Performance
    with tab_monthly:
        st.subheader("Monthly Financial Statement Summary")
        if not monthly_summary.empty:
            display_monthly = monthly_summary.copy()
            display_monthly["Revenue"] = display_monthly["Revenue"].apply(format_currency)
            display_monthly["Expense"] = display_monthly["Expense"].apply(format_currency)
            display_monthly["Net Income"] = display_monthly["Net Income"].apply(format_currency)
            display_monthly["Margin %"] = display_monthly["Margin %"].apply(lambda x: f"{x:.2f}%")
            st.dataframe(display_monthly, use_container_width=True)
        else:
            st.info("No monthly breakdown available.")

    # Tab 3: Department & Category Breakdown
    with tab_departments:
        col_dept, col_cat = st.columns(2)
        with col_dept:
            st.subheader("Department P&L Contribution")
            if not dept_summary.empty:
                dept_disp = dept_summary.copy()
                dept_disp["Revenue"] = dept_disp["Revenue"].apply(format_currency)
                dept_disp["Expense"] = dept_disp["Expense"].apply(format_currency)
                dept_disp["Net"] = dept_disp["Net"].apply(format_currency)
                st.dataframe(dept_disp, use_container_width=True)
            else:
                st.info("No department data available.")

        with col_cat:
            st.subheader("Expense Categories Breakdown")
            if not category_summary.empty:
                cat_disp = category_summary.copy()
                cat_disp["Total Amount"] = cat_disp["Total Amount"].apply(format_currency)
                cat_disp["Share %"] = cat_disp["Share %"].apply(lambda x: f"{x:.2f}%")
                st.dataframe(cat_disp, use_container_width=True)
            else:
                st.info("No expense breakdown available.")

        st.divider()
        st.subheader("Revenue Streams Breakdown")
        if not revenue_cat_summary.empty:
            rev_disp = revenue_cat_summary.copy()
            rev_disp["Total Amount"] = rev_disp["Total Amount"].apply(format_currency)
            rev_disp["Share %"] = rev_disp["Share %"].apply(lambda x: f"{x:.2f}%")
            st.dataframe(rev_disp, use_container_width=True)
        else:
            st.info("No revenue streams available.")

    # Tab 4: Transaction Details
    with tab_transactions:
        st.subheader("Raw Financial Transaction Ledger")
        
        search_query = st.text_input("Quick Search (Description or Category):", "")
        tx_view = filtered_df.copy()
        if search_query:
            mask = (
                tx_view["Description"].str.contains(search_query, case=False, na=False)
                | tx_view["Category"].str.contains(search_query, case=False, na=False)
            )
            tx_view = tx_view[mask]

        formatted_tx = tx_view.copy()
        formatted_tx["Date"] = formatted_tx["Date"].dt.strftime("%Y-%m-%d")
        formatted_tx["Amount"] = formatted_tx["Amount"].apply(format_currency)
        st.dataframe(formatted_tx, use_container_width=True, height=450)
        st.caption(f"Showing {len(tx_view)} transactions.")

    # Tab 5: Export & Reports
    with tab_export:
        st.subheader("Export Formatted Financial Reports")
        st.write(
            "Generate production-ready Excel workbooks formatted with executive summaries, "
            "styled tables, currency formats, and individual category worksheets using `openpyxl`."
        )

        col_exp1, col_exp2 = st.columns(2)

        with col_exp1:
            st.markdown("#### 📗 Formatted Excel Report")
            try:
                excel_buffer = export_financial_report(
                    kpis=kpis,
                    monthly_df=monthly_summary,
                    dept_df=dept_summary,
                    category_df=category_summary,
                    raw_df=filtered_df,
                )
                today_tag = datetime.now().strftime("%Y%m%d_%H%M")
                st.download_button(
                    label="📥 Download Excel Report (.xlsx)",
                    data=excel_buffer,
                    file_name=f"financial_report_{today_tag}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
            except Exception as ex:
                st.error(f"Failed to generate Excel report: {ex}")

        with col_exp2:
            st.markdown("#### 📄 Filtered Data CSV Export")
            csv_data = filtered_df.to_csv(index=False).encode("utf-8")
            st.download_button(
                label="📥 Download Filtered CSV (.csv)",
                data=csv_data,
                file_name="filtered_transactions.csv",
                mime="text/csv",
                use_container_width=True,
            )

    # Tab 6: Payroll & Template Mapping
    with tab_payroll:
        st.subheader("📑 Employee Payroll & Template Mapping Engine")
        st.write(
            "Extract key payroll figures from raw employee data (CSV/Excel), map them to match "
            "the exact columns and rows of a reference template (.xlsx/.ods), calculate totals, "
            "and generate a clean, ready-to-download Excel report."
        )

        payroll_mode = st.radio(
            "Select Input Source:",
            options=["Use Sample Payroll & Template Files", "Upload Custom Payroll & Template"],
            horizontal=True,
        )

        raw_payroll_input = None
        template_input = None

        if payroll_mode == "Use Sample Payroll & Template Files":
            sample_raw_p = "data/raw_payroll_data.csv"
            sample_tmpl_p = "data/payroll_report_template.xlsx"
            if not Path(sample_raw_p).exists() or not Path(sample_tmpl_p).exists():
                create_sample_payroll_files(sample_raw_p, sample_tmpl_p)
            
            raw_payroll_input = sample_raw_p
            template_input = sample_tmpl_p
            st.info("Using built-in sample raw payroll data and standard `.xlsx` reference template.")
        else:
            col_u1, col_u2 = st.columns(2)
            with col_u1:
                uploaded_raw = st.file_uploader(
                    "1. Upload Raw Payroll Data (.csv, .xlsx, .xls)",
                    type=["csv", "xlsx", "xls"],
                    key="payroll_raw_uploader",
                )
                if uploaded_raw:
                    raw_payroll_input = uploaded_raw

            with col_u2:
                uploaded_template = st.file_uploader(
                    "2. Upload Reference Template (.xlsx, .ods)",
                    type=["xlsx", "ods"],
                    key="payroll_template_uploader",
                )
                if uploaded_template:
                    template_input = uploaded_template

        if raw_payroll_input and template_input:
            if st.button("🚀 Process & Generate Financial Report", type="primary"):
                try:
                    with st.spinner("Mapping columns and generating formatted Excel report..."):
                        report_buffer = generate_payroll_report(raw_payroll_input, template_input)
                    st.success("Financial report generated successfully!")
                    
                    st.download_button(
                        label="📥 Download Mapped Excel Report (.xlsx)",
                        data=report_buffer,
                        file_name=f"payroll_financial_report_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )
                except Exception as ex:
                    st.error(f"Error during report generation: {ex}")
        else:
            st.warning("Please provide both raw payroll data and a reference template to generate the report.")


if __name__ == "__main__":
    main()

