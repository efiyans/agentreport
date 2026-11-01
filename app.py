"""Financial Report Generator & Automated Auditor Web Dashboard.

Streamlit frontend allowing users to upload raw data (CSV) and a past reference
template (Excel), trigger report generation, preview the processed report table,
review automated financial auditor warnings, and download the finished .xlsx file.
"""

from datetime import datetime
from pathlib import Path
from typing import Dict, List, Any, Optional
import hmac

import pandas as pd
import streamlit as st

from utils.processor import (
    generate_payroll_report,
    load_raw_payroll_data,
    create_sample_payroll_files,
    validate_financial_audit,
)


def format_currency(val: Any) -> str:
    """Format numeric values as clean accounting currency strings."""
    if pd.isna(val) or val is None or val == "":
        return ""
    try:
        num = float(val)
        if num < 0:
            return f"-${abs(num):,.2f}"
        return f"${num:,.2f}"
    except (ValueError, TypeError):
        return str(val)


def format_table_for_display(df: pd.DataFrame) -> pd.DataFrame:
    """Format DataFrame for clean display: format monetary columns and counts."""
    display_df = df.copy()
    currency_keywords = [
        "salary", "pay", "bonus", "overtime", "gross", "net", "tax",
        "pension", "deduction", "allowance", "cost", "amount", "total",
    ]
    
    for col in display_df.columns:
        col_lower = str(col).lower()
        # If it's a count/headcount column
        if "count" in col_lower or "headcount" in col_lower:
            display_df[col] = display_df[col].apply(
                lambda v: f"{int(v):,}" if pd.notna(v) and isinstance(v, (int, float)) else str(v or "")
            )
        # If it matches monetary keywords and contains numeric data
        elif any(k in col_lower for k in currency_keywords):
            # Check if column is numeric or contains numbers
            num_series = pd.to_numeric(display_df[col], errors="coerce")
            if num_series.notna().any():
                display_df[col] = display_df[col].apply(
                    lambda v: format_currency(v) if pd.notna(v) and v != "" else ""
                )
    return display_df


def check_password() -> bool:
    """Validate credentials against st.secrets['credentials'].

    Renders a login screen with 'Username' and 'Password' input fields and a 'Log in' button.
    Returns True if the user is authenticated, False otherwise.
    """
    if st.session_state.get("password_correct", False):
        return True

    # Check credentials block in st.secrets
    if "credentials" not in st.secrets:
        st.error(
            "⚠️ Authentication Error: `[credentials]` block is missing from `secrets.toml`. "
            "Please configure `.streamlit/secrets.toml` or Streamlit Cloud Secrets."
        )
        return False

    # Center login form
    _, col_login, _ = st.columns([1, 2, 1])
    with col_login:
        st.markdown(
            """
            <div style="background-color: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 10px; padding: 24px; margin-top: 30px; margin-bottom: 20px;">
                <h3 style="margin-top: 0; color: #1E3A8A; display: flex; align-items: center; gap: 8px;">
                    🔒 Financial Dashboard Login
                </h3>
                <p style="color: #6B7280; font-size: 0.95rem; margin-bottom: 0;">
                    Please enter your authorized credentials to access the financial reporting portal.
                </p>
            </div>
            """,
            unsafe_allow_html=True,
        )

        with st.form("login_form", clear_on_submit=False):
            username_input = st.text_input("Username", key="login_username", placeholder="Enter username")
            password_input = st.text_input("Password", type="password", key="login_password", placeholder="Enter password")
            submit_button = st.form_submit_button("Log in", type="primary", use_container_width=True)

            if submit_button:
                creds = st.secrets["credentials"]
                is_valid = False

                # Handle [credentials] table with username and password
                if "username" in creds and "password" in creds:
                    expected_user = str(creds["username"])
                    expected_pass = str(creds["password"])
                    if hmac.compare_digest(username_input, expected_user) and hmac.compare_digest(password_input, expected_pass):
                        is_valid = True
                # Handle multi-user dict mapping username -> password
                elif username_input in creds and hmac.compare_digest(password_input, str(creds[username_input])):
                    is_valid = True

                if is_valid:
                    st.session_state["password_correct"] = True
                    st.session_state["logged_in_user"] = username_input
                    st.rerun()
                else:
                    st.error("❌ Incorrect username or password. Please try again.")

    return False


def main():
    st.set_page_config(
        page_title="Financial Report Generator & Auditor",
        page_icon="📑",
        layout="wide",
        initial_sidebar_state="expanded",
    )

    # Enforce authentication - stop execution of main app content unless authenticated
    if not check_password():
        st.stop()

    # Styling for professional modern look
    st.markdown(
        """
        <style>
        .main-header {
            font-size: 2.1rem;
            font-weight: 700;
            color: #1E3A8A;
            margin-bottom: 0.2rem;
        }
        .sub-header {
            font-size: 1.0rem;
            color: #4B5563;
            margin-bottom: 1.2rem;
        }
        .metric-card {
            background-color: #F8FAFC;
            border: 1px solid #E2E8F0;
            border-radius: 8px;
            padding: 14px 18px;
        }
        .audit-box {
            background-color: #FFFFFF;
            border: 1px solid #E5E7EB;
            border-radius: 8px;
            padding: 16px;
            margin-top: 10px;
        }
        .stDownloadButton > button {
            background-color: #10B981 !important;
            color: white !important;
            font-weight: 600 !important;
            border: none !important;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    st.markdown('<div class="main-header">📑 Financial Report Generator & Auditor</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="sub-header">'
        'Upload raw financial/payroll data and a reference template to automatically map columns, '
        'calculate grand totals, audit math and variance discrepancies (>20%), and download the finished Excel workbook.'
        '</div>',
        unsafe_allow_html=True,
    )

    # Initialize Session State
    if "report_generated" not in st.session_state:
        st.session_state.report_generated = False
        st.session_state.report_buffer = None
        st.session_state.audit_messages = []
        st.session_state.preview_df = None
        st.session_state.filename = "financial_report.xlsx"

    # -------------------------------------------------------------------------
    # Sidebar: User Session, File Uploaders & Configuration
    # -------------------------------------------------------------------------
    with st.sidebar:
        user_name = st.session_state.get("logged_in_user", "Authorized User")
        st.markdown(f"👤 Signed in as **{user_name}**")
        if st.button("🚪 Log out", key="logout_btn", use_container_width=True):
            st.session_state["password_correct"] = False
            st.session_state.pop("logged_in_user", None)
            st.rerun()

        st.divider()
        st.header("📁 Document Uploads")

        use_demo = st.checkbox(
            "Use Demo Sample Files",
            value=False,
            help="Quick-load pre-configured payroll CSV and template Excel for instant testing.",
        )

        raw_file = None
        template_file = None

        if use_demo:
            raw_path = "data/raw_payroll_data.csv"
            tmpl_path = "data/payroll_report_template.xlsx"
            if not Path(raw_path).exists() or not Path(tmpl_path).exists():
                create_sample_payroll_files(raw_path, tmpl_path)
            raw_file = raw_path
            template_file = tmpl_path
            st.success("✅ Demo sample files loaded.")
        else:
            raw_file = st.file_uploader(
                "1. Upload Raw Data (CSV)",
                type=["csv"],
                help="Raw employee compensation or transaction data in CSV format.",
            )
            template_file = st.file_uploader(
                "2. Upload Reference Template (Excel)",
                type=["xlsx", "ods"],
                help="Reference template workbook (.xlsx or .ods) defining target layout and columns.",
            )

        st.divider()
        st.subheader("⚙️ Auditor Settings")
        variance_pct = st.slider(
            "Variance Alert Threshold",
            min_value=10,
            max_value=50,
            value=20,
            step=5,
            format="%d%%",
            help="Flags rows or departments with compensation shifts exceeding this threshold.",
        )
        variance_threshold = variance_pct / 100.0

        st.divider()
        # Primary Generate Report Button
        can_generate = bool(raw_file and template_file)
        generate_clicked = st.button(
            "🚀 Generate Report",
            type="primary",
            use_container_width=True,
            disabled=not can_generate,
        )

        if not can_generate:
            st.caption("👈 Upload both files or enable Demo Sample Files to generate.")

    # -------------------------------------------------------------------------
    # Processing Logic
    # -------------------------------------------------------------------------
    if generate_clicked and can_generate:
        try:
            with st.spinner("Processing raw data, mapping template columns, and auditing financials..."):
                res = generate_payroll_report(
                    raw_data_input=raw_file,
                    template_input=template_file,
                    return_audit=True,
                    variance_threshold=variance_threshold,
                )
                output_buffer, audit_msgs, preview_df = res[0], res[1], res[2]

            timestamp_str = datetime.now().strftime("%Y%m%d_%H%M")
            st.session_state.report_generated = True
            st.session_state.report_buffer = output_buffer
            st.session_state.audit_messages = audit_msgs
            st.session_state.preview_df = preview_df
            st.session_state.filename = f"financial_payroll_report_{timestamp_str}.xlsx"
            st.toast("Report processed & audited successfully!", icon="✅")

        except Exception as ex:
            st.error(f"❌ Error during report generation: {ex}")
            st.session_state.report_generated = False

    # -------------------------------------------------------------------------
    # Main Content Area
    # -------------------------------------------------------------------------
    if st.session_state.report_generated and st.session_state.preview_df is not None:
        preview_df = st.session_state.preview_df
        audit_msgs = st.session_state.audit_messages

        # Action bar: Download button & Summary Pills
        col_act1, col_act2 = st.columns([2, 1])
        with col_act1:
            st.success("✅ **Report Generated Successfully!** All rows mapped to reference template.")
        with col_act2:
            st.download_button(
                label="📥 Download Final Report (.xlsx)",
                data=st.session_state.report_buffer,
                file_name=st.session_state.filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True,
            )

        # ---------------------------------------------------------------------
        # 1. Executive Summary Metric Cards
        # ---------------------------------------------------------------------
        # Look for gross, net, deductions, and base columns for quick KPI overview
        gross_col = next((c for c in preview_df.columns if "gross" in str(c).lower()), None)
        net_col = next((c for c in preview_df.columns if "net" in str(c).lower()), None)
        ded_col = next((c for c in preview_df.columns if "deduction" in str(c).lower()), None)
        
        # Calculate totals excluding the 'Total' summary row
        data_rows = preview_df[preview_df.iloc[:, 0] != "Total"]
        record_count = len(data_rows)

        total_gross = (
            float(pd.to_numeric(data_rows[gross_col], errors="coerce").sum())
            if gross_col else 0.0
        )
        total_net = (
            float(pd.to_numeric(data_rows[net_col], errors="coerce").sum())
            if net_col else 0.0
        )
        total_ded = (
            float(pd.to_numeric(data_rows[ded_col], errors="coerce").sum())
            if ded_col else 0.0
        )

        col_m1, col_m2, col_m3, col_m4 = st.columns(4)
        with col_m1:
            st.metric("Total Records", f"{record_count:,}")
        with col_m2:
            st.metric("Grand Total Gross", format_currency(total_gross) if total_gross > 0 else "N/A")
        with col_m3:
            st.metric("Total Deductions", format_currency(total_ded) if total_ded > 0 else "N/A")
        with col_m4:
            st.metric("Grand Total Net Pay", format_currency(total_net) if total_net > 0 else "N/A")

        st.markdown("<br>", unsafe_allow_html=True)

        # ---------------------------------------------------------------------
        # 2. Preview Section: Data Table of Processed Report
        # ---------------------------------------------------------------------
        st.subheader("📊 Processed Report Preview")
        st.caption("Live preview of the generated report matching the reference template columns, with calculated Grand Total.")

        formatted_preview = format_table_for_display(preview_df)
        st.dataframe(
            formatted_preview,
            use_container_width=True,
            height=340,
        )

        st.markdown("<br>", unsafe_allow_html=True)

        # ---------------------------------------------------------------------
        # 3. Validation Log Box: Automated Financial Auditor Warnings & Reconciliations
        # ---------------------------------------------------------------------
        st.subheader("🛡️ Automated Financial Auditor Validation Log")
        
        warnings_list = [m for m in audit_msgs if m["type"] == "warning"]
        success_list = [m for m in audit_msgs if m["type"] == "success"]
        info_list = [m for m in audit_msgs if m["type"] == "info"]

        # Status badge banner
        if warnings_list:
            st.warning(
                f"⚠️ **Attention Required**: {len(warnings_list)} potential math or high-variance (> {variance_pct}%) "
                f"issues detected by the automated auditor. Please review below."
            )
        else:
            st.success(
                f"✅ **Audit Verification Passed**: All individual sub-totals reconcile with the Grand Total, "
                f"and no abnormal variances (> {variance_pct}%) were detected."
            )

        with st.container():
            tab_all, tab_warn, tab_recon = st.tabs([
                f"📋 All Findings ({len(audit_msgs)})",
                f"⚠️ Warnings & Variances ({len(warnings_list)})",
                f"✅ Reconciliations Verified ({len(success_list)})",
            ])

            with tab_all:
                for msg in audit_msgs:
                    tag = f"**[{msg['category']}]**"
                    if msg["type"] == "warning":
                        st.warning(f"{tag} {msg['message']}")
                    elif msg["type"] == "success":
                        st.success(f"{tag} {msg['message']}")
                    else:
                        st.info(f"{tag} {msg['message']}")

            with tab_warn:
                if warnings_list:
                    for msg in warnings_list:
                        st.warning(f"**[{msg['category']}]** {msg['message']}")
                else:
                    st.info("🎉 Zero warnings detected! No math discrepancies or major variance spikes found.")

            with tab_recon:
                if success_list:
                    for msg in success_list:
                        st.success(f"**[{msg['category']}]** {msg['message']}")
                else:
                    st.info("No sub-total checks were evaluated.")

        # Bottom secondary download action
        st.divider()
        col_d1, col_d2, col_d3 = st.columns([1, 2, 1])
        with col_d2:
            st.download_button(
                label="📥 Download Final Report (.xlsx)",
                data=st.session_state.report_buffer,
                file_name=st.session_state.filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True,
                key="download_bottom",
            )

    else:
        # ---------------------------------------------------------------------
        # Empty State: Guidance for user
        # ---------------------------------------------------------------------
        st.info("👋 Welcome! Please upload your files in the sidebar to get started.")

        col_g1, col_g2 = st.columns(2)
        with col_g1:
            st.markdown("### 1️⃣ Raw Data (CSV)")
            st.markdown(
                """
                Upload your raw employee compensation or financial records:
                - Supports columns like `Emp_ID`, `Name`, `Department`, `Base_Salary`, `Overtime`, `Bonus`, `Tax`, etc.
                - Automatically calculates derived fields (*Gross Pay, Deductions, Net Pay*) if omitted.
                """
            )
        with col_g2:
            st.markdown("### 2️⃣ Reference Template (Excel)")
            st.markdown(
                """
                Upload a past template report (`.xlsx` or `.ods`):
                - The processor reads your exact target column order, headers, and styles.
                - Reconciles math sub-totals against the Grand Total.
                - Appends an accounting summary row with dynamic `=SUM(...)` formulas.
                """
            )

        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown("#### ⚡ Quick Test:")
        st.markdown(
            "Toggle **'Use Demo Sample Files'** in the sidebar and click **'🚀 Generate Report'** "
            "to preview the workflow immediately with sample payroll datasets."
        )


if __name__ == "__main__":
    main()
