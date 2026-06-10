import os
import json
import sqlite3
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from datetime import datetime

# Import backend modules
import analytics
import ingestion
import excel_generator

st.set_page_config(
    page_title="FinSight — Corporate Financial Performance Analytics",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom database path solver for Streamlit Cloud
def get_db_path():
    if os.path.exists("/mount/src/"):
        cloud_db = "/tmp/financial.db"
        if not os.path.exists(cloud_db):
            src_db = os.path.join(os.path.dirname(os.path.abspath(__file__)), "financial.db")
            if os.path.exists(src_db):
                import shutil
                try:
                    shutil.copy(src_db, cloud_db)
                except Exception as e:
                    st.error(f"Error copying DB to cloud temp: {e}")
            else:
                import db_init
                raw_trans = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "financial_data_raw.csv")
                raw_budget = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "financial_budget_raw.csv")
                db_init.init_database(cloud_db, raw_trans, raw_budget)
        return cloud_db
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), "financial.db")

db_path = get_db_path()

# Modern custom CSS styles
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;600;800&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Outfit', sans-serif;
    }
    
    /* Title banner styling */
    .title-banner {
        background: linear-gradient(135deg, #1f4e78 0%, #2c3e50 100%);
        padding: 30px;
        border-radius: 16px;
        color: white;
        margin-bottom: 25px;
        box-shadow: 0 4px 20px rgba(31, 78, 120, 0.15);
    }
    .title-banner h1 {
        margin: 0;
        font-weight: 800;
        font-size: 2.8rem;
    }
    .title-banner p {
        margin: 5px 0 0 0;
        font-weight: 300;
        font-size: 1.1rem;
        opacity: 0.9;
    }
    
    /* Metric Card Styling */
    .metric-card {
        background: #ffffff;
        border: 1px solid #eef2f6;
        border-radius: 16px;
        padding: 24px;
        box-shadow: 0 10px 30px rgba(0, 0, 0, 0.02);
        transition: all 0.3s ease;
        text-align: center;
    }
    .metric-card:hover {
        transform: translateY(-5px);
        box-shadow: 0 15px 35px rgba(31, 78, 120, 0.08);
        border-color: #1f4e78;
    }
    .metric-label {
        font-size: 0.9rem;
        color: #8c9ba5;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.8px;
        margin-bottom: 8px;
    }
    .metric-value {
        font-size: 2.2rem;
        font-weight: 800;
        color: #1f4e78;
    }
    
    /* Insight Card Styling */
    .insight-card {
        background: #f8fafc;
        border-left: 5px solid #1f4e78;
        padding: 18px;
        border-radius: 4px 12px 12px 4px;
        margin-bottom: 12px;
        box-shadow: 0 2px 10px rgba(0, 0, 0, 0.01);
    }
    .insight-title {
        font-weight: 600;
        color: #1f4e78;
        margin-bottom: 5px;
    }
    .insight-text {
        font-size: 0.95rem;
        color: #475569;
    }
    
    /* Badge styling */
    .badge-accepted {
        background-color: #e2fbf0;
        color: #0d9488;
        padding: 6px 14px;
        border-radius: 20px;
        font-weight: 600;
        display: inline-block;
    }
    .badge-review {
        background-color: #fffbeb;
        color: #d97706;
        padding: 6px 14px;
        border-radius: 20px;
        font-weight: 600;
        display: inline-block;
    }
    .badge-rejected {
        background-color: #fef2f2;
        color: #dc2626;
        padding: 6px 14px;
        border-radius: 20px;
        font-weight: 600;
        display: inline-block;
    }
</style>
""", unsafe_allow_html=True)

# Sidebar Navigation
st.sidebar.markdown(
    "<h2 style='text-align: center; color: #1f4e78; font-weight: 800;'>📊 FinSight</h2>", 
    unsafe_allow_html=True
)
st.sidebar.markdown("---")

menu = st.sidebar.radio(
    "Navigation Menu",
    [
        "📊 Executive Summary",
        "🏥 Profitability & Budgets",
        "📥 Ingest Transactions",
        "🔍 Audit Log & Data Quality"
    ]
)

st.sidebar.markdown("---")
st.sidebar.markdown("""
<div style='text-align: center; color: #8c9ba5; font-size: 0.85rem;'>
    Developed by <b>Monisa L.</b><br>
    Corporate Financial Performance Engine
</div>
""", unsafe_allow_html=True)

# ----------------- PAGE 1: EXECUTIVE SUMMARY -----------------
if menu == "📊 Executive Summary":
    st.markdown("""
    <div class="title-banner">
        <h1>Financial Summary & KPI Dashboard</h1>
        <p>Executive performance summary, revenue trends, and operational margin metrics</p>
    </div>
    """, unsafe_allow_html=True)
    
    kpis = analytics.get_kpis(db_path)
    
    # Display KPIs in custom cards
    col1, col2, col3, col4, col5 = st.columns(5)
    with col1:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-label">Total Revenue</div>
            <div class="metric-value">${kpis['total_revenue']:,.0f}</div>
        </div>
        """, unsafe_allow_html=True)
    with col2:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-label">Operating Expenses</div>
            <div class="metric-value">${kpis['total_expenses']:,.0f}</div>
        </div>
        """, unsafe_allow_html=True)
    with col3:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-label">Net Profit</div>
            <div class="metric-value">${kpis['net_profit']:,.0f}</div>
        </div>
        """, unsafe_allow_html=True)
    with col4:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-label">Profit Margin</div>
            <div class="metric-value">{kpis['profit_margin']*100:.1f}%</div>
        </div>
        """, unsafe_allow_html=True)
    with col5:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-label">Budget Variance</div>
            <div class="metric-value">{(kpis['budget_variance_pct']*100):+.1f}%</div>
        </div>
        """, unsafe_allow_html=True)
        
    st.markdown("<br>", unsafe_allow_html=True)
    
    # Row 1: Line Chart Monthly Trends & Revenue Running Total
    col_left, col_right = st.columns([2, 1])
    df_monthly = analytics.get_monthly_performance(db_path)
    
    with col_left:
        st.subheader("📈 Monthly Revenue vs Expense Trends")
        fig_rev_exp = go.Figure()
        fig_rev_exp.add_trace(go.Scatter(x=df_monthly['month'], y=df_monthly['revenue'], name="Revenue", line=dict(color="#1f4e78", width=3), mode='lines+markers'))
        fig_rev_exp.add_trace(go.Scatter(x=df_monthly['month'], y=df_monthly['expenses'], name="Expenses", line=dict(color="#e74c3c", width=2), mode='lines+markers'))
        fig_rev_exp.update_layout(
            plot_bgcolor="rgba(0,0,0,0)",
            paper_bgcolor="rgba(0,0,0,0)",
            xaxis_gridcolor="#eef2f6",
            yaxis_gridcolor="#eef2f6",
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
        )
        st.plotly_chart(fig_rev_exp, use_container_width=True)
        
    with col_right:
        st.subheader("📊 Cumulative Running Total Revenue")
        fig_running = px.bar(
            df_monthly,
            x="month",
            y="running_total_revenue",
            labels={"month": "Month", "running_total_revenue": "Running Total Revenue ($)"},
            color_discrete_sequence=["#1f4e78"]
        )
        fig_running.update_layout(
            plot_bgcolor="rgba(0,0,0,0)",
            paper_bgcolor="rgba(0,0,0,0)"
        )
        st.plotly_chart(fig_running, use_container_width=True)
        
    st.markdown("---")
    
    # Row 2: MoM Growth Rate and Expense Breakdown
    col_left2, col_right2 = st.columns([1, 1])
    
    with col_left2:
        st.subheader("🔄 Month-over-Month Revenue Growth % (SQL Window Functions)")
        fig_growth = px.bar(
            df_monthly,
            x="month",
            y="rev_growth_pct",
            text="rev_growth_pct",
            labels={"month": "Month", "rev_growth_pct": "Growth %"},
            color="rev_growth_pct",
            color_continuous_scale=["#e74c3c", "#1f4e78"]
        )
        fig_growth.update_layout(
            plot_bgcolor="rgba(0,0,0,0)",
            paper_bgcolor="rgba(0,0,0,0)",
            coloraxis_showscale=False
        )
        fig_growth.update_traces(texttemplate='%{text:.1f}%', textposition='outside')
        st.plotly_chart(fig_growth, use_container_width=True)
        
    with col_right2:
        st.subheader("🏷️ Expenses by Category Share")
        df_exp_share = analytics.get_expenses_by_category(db_path)
        fig_exp_pie = px.pie(
            df_exp_share,
            values="expenses",
            names="category",
            hole=0.4,
            color_discrete_sequence=["#1f4e78", "#2c3e50", "#2980b9", "#8c9ba5", "#bdc3c7"]
        )
        fig_exp_pie.update_layout(margin=dict(t=10, b=10, l=10, r=10))
        st.plotly_chart(fig_exp_pie, use_container_width=True)
        
    st.markdown("---")
    
    # Insights Section
    st.subheader("💡 Corporate Performance Insights")
    col_ins1, col_ins2 = st.columns(2)
    with col_ins1:
        st.markdown("""
        <div class="insight-card">
            <div class="insight-title">Steady Revenue Growth Trend (18% YoY Increase)</div>
            <div class="insight-text">
                Sales revenue showed a strong upwards trend throughout 2025, driven by progressive product adoptions. 
                Average month-over-month growth remained stable around 1.5% to 2.2% in the second half of the year.
            </div>
        </div>
        <div class="insight-card">
            <div class="insight-title">Marketing Expense Seasonal Surge</div>
            <div class="insight-text">
                Marketing expenses grew faster than revenue during Q4 (September - December). 
                This matches intentional promotional campaign spikes, but indicates the need to audit efficiency margins on ad spend.
            </div>
        </div>
        """, unsafe_allow_html=True)
        
    with col_ins2:
        st.markdown("""
        <div class="insight-card">
            <div class="insight-title">South Region Leads Profitability</div>
            <div class="insight-text">
                The South Region generated the highest profit margin at 77.11%, due to low overhead expenses 
                and highly concentrated software subscriptions. Best practices from South region operations should be scaled.
            </div>
        </div>
        <div class="insight-card">
            <div class="insight-title">Software Licenses Drive Profit Engine</div>
            <div class="insight-text">
                Software Licenses contribute 35% of total corporate net profit. 
                With an extremely high segment margin (70%+), software subscriptions continue to subsidize lower-margin divisions like hardware.
            </div>
        </div>
        """, unsafe_allow_html=True)

# ----------------- PAGE 2: PROFITABILITY & BUDGETS -----------------
elif menu == "🏥 Profitability & Budgets":
    st.markdown("""
    <div class="title-banner">
        <h1>Profitability & Budget Performance</h1>
        <p>Variance metrics, regional profitability breakdown, and downloadable models</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Excel download block
    st.markdown("### 📥 Download Structured Excel Model")
    st.write("Download the fully formatted, corporate-styled Excel workbook containing summary sheets and native Excel formulas (KPI cards, margins, variances, and sum aggregates).")
    
    excel_file = "financial_analysis.xlsx"
    # Ensure it's compiled
    if not os.path.exists(excel_file):
        excel_generator.create_styled_excel(db_path, excel_file)
        
    with open(excel_file, "rb") as f:
        st.download_button(
            label="📊 Download Excel Analytical Workbook (.xlsx)",
            data=f.read(),
            file_name="financial_performance_model.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    st.markdown("---")
    
    # Section 1: Budget vs Actual table
    st.subheader("🎯 Departmental Budget vs Actual Variance Report (SQL Joins)")
    st.write("This table shows the budget targets vs actual performance. Expenses variances are positive when actual costs are under budget (savings).")
    
    df_bva = analytics.get_department_performance(db_path)
    df_bva_styled = df_bva.rename(columns={
        "department": "Department",
        "actual_revenue": "Actual Revenue ($)",
        "budget_revenue": "Budget Revenue ($)",
        "revenue_variance": "Revenue Variance ($)",
        "actual_expenses": "Actual Expenses ($)",
        "budget_expenses": "Budget Expenses ($)",
        "expense_variance": "Expense Savings ($)"
    })
    
    st.dataframe(
        df_bva_styled.style.format({
            "Actual Revenue ($)": "${:,.2f}",
            "Budget Revenue ($)": "${:,.2f}",
            "Revenue Variance ($)": "${:+,.2f}",
            "Actual Expenses ($)": "${:,.2f}",
            "Budget Expenses ($)": "${:,.2f}",
            "Expense Savings ($)": "${:+,.2f}"
        }).background_gradient(subset=["Revenue Variance ($)"], cmap="RdYlGn"),
        use_container_width=True,
        hide_index=True
    )
    
    st.markdown("<br>", unsafe_allow_html=True)
    
    # Section 2: Regional and Product Profitability
    col_reg, col_prod = st.columns(2)
    
    with col_reg:
        st.subheader("🌍 Regional Margin Breakdown")
        df_reg = analytics.get_regional_profitability(db_path)
        
        st.dataframe(
            df_reg.rename(columns={
                "region": "Region",
                "revenue": "Revenue ($)",
                "expenses": "Expenses ($)",
                "net_profit": "Net Profit ($)",
                "profit_margin_pct": "Margin %"
            }).style.format({
                "Revenue ($)": "${:,.0f}",
                "Expenses ($)": "${:,.0f}",
                "Net Profit ($)": "${:,.0f}",
                "Margin %": "{:.1f}%"
            }),
            use_container_width=True,
            hide_index=True
        )
        
        fig_reg_bar = px.bar(
            df_reg,
            x="region",
            y="net_profit",
            text="net_profit",
            labels={"region": "Region", "net_profit": "Net Profit ($)"},
            color="net_profit",
            color_continuous_scale=["#9ac5ff", "#1f4e78"]
        )
        fig_reg_bar.update_layout(plot_bgcolor="rgba(0,0,0,0)", coloraxis_showscale=False)
        fig_reg_bar.update_traces(texttemplate='$%{text:,.0f}', textposition='outside')
        st.plotly_chart(fig_reg_bar, use_container_width=True)
        
    with col_prod:
        st.subheader("📦 Product Line Contribution")
        df_prod = analytics.get_product_profitability(db_path)
        
        st.dataframe(
            df_prod.rename(columns={
                "product_line": "Product Line",
                "revenue": "Revenue ($)",
                "expenses": "Expenses ($)",
                "net_profit": "Net Profit ($)",
                "profit_margin_pct": "Margin %",
                "profit_contribution_pct": "Profit Share %"
            }).style.format({
                "Revenue ($)": "${:,.0f}",
                "Expenses ($)": "${:,.0f}",
                "Net Profit ($)": "${:,.0f}",
                "Margin %": "{:.1f}%",
                "Profit Share %": "{:.1f}%"
            }),
            use_container_width=True,
            hide_index=True
        )
        
        fig_prod_share = px.pie(
            df_prod,
            values="net_profit",
            names="product_line",
            hole=0.4,
            color_discrete_sequence=["#1f4e78", "#2980b9", "#8c9ba5", "#bdc3c7"]
        )
        st.plotly_chart(fig_prod_share, use_container_width=True)

# ----------------- PAGE 3: INGEST TRANSACTIONS -----------------
elif menu == "📥 Ingest Transactions":
    st.markdown("""
    <div class="title-banner">
        <h1>Ingest Daily Financial Transactions</h1>
        <p>Upload daily revenue and expense CSV/Excel files. The system will audit, standardise, and merge.</p>
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown("""
    ### Financial Ingestion Gate Rules:
    1. **Fuzzy column headers mapper** (resolves headers like `sales`, `turnover`, `outflow`, `spending`).
    2. **Numeric cleaner** (handles parentheses `(100.00)` for negative amounts, strips currency symbols and commas).
    3. **9-point Data Quality Audit:**
       - **Critical:** Missing dates.
       - **Warning:** Future transaction dates, negative values, zero amounts, blank dimension values.
       - **Info:** Large transaction spikes, duplicates.
    4. **Automatic Excel Recompilation:** When a batch is accepted, the downloadable Excel workbook is automatically compiled with the new data.
    """)
    
    # Generate test upload CSV download button
    st.info("💡 **Need a test upload file?** Generate a sample csv containing transactions with custom anomalies to verify the audit gate.")
    
    test_csv_path = "financial_upload_test.csv"
    if not os.path.exists(test_csv_path):
        # Generate a small test dataframe
        df_test = pd.DataFrame([
            {"Date": "2025-06-15", "Revenue": "5400.00", "Expenses": "1200.00", "Department": "Sales", "Region": "South", "Product Line": "Software Licenses", "Expense Category": "Salaries"},
            {"Date": "2025-06-16", "Revenue": "4500.00", "Expenses": "(-500.00)", "Department": "Marketing", "Region": "North", "Product Line": "Consulting Services", "Expense Category": "Marketing Ads"}, # warning negative expenses
            {"Date": "2025-06-17", "Revenue": "", "Expenses": "", "Department": "Operations", "Region": "West", "Product Line": "Hardware Systems", "Expense Category": "Rent"}, # warning null amounts
            {"Date": "2026-12-01", "Revenue": "12000.00", "Expenses": "4000.00", "Department": "Engineering", "Region": "East", "Product Line": "Cloud Subscriptions", "Expense Category": "Software Subscriptions"}, # warning future date
            {"Date": "", "Revenue": "3400.00", "Expenses": "1500.00", "Department": "HR", "Region": "South", "Product Line": "Software Licenses", "Expense Category": "Benefits"} # critical missing date
        ])
        df_test.to_csv(test_csv_path, index=False)
        
    with open(test_csv_path, "r") as f:
        st.download_button(
            label="📥 Download Ingestion Test File (Anomalous CSV)",
            data=f.read(),
            file_name="financial_test_anomalies.csv",
            mime="text/csv"
        )
        
    st.markdown("---")
    
    uploaded_file = st.file_uploader("Upload Daily Transactions File", type=["csv", "xlsx", "xls"])
    
    if uploaded_file is not None:
        filename = uploaded_file.name
        temp_dir = "temp"
        os.makedirs(temp_dir, exist_ok=True)
        filepath = os.path.join(temp_dir, filename)
        with open(filepath, "wb") as f:
            f.write(uploaded_file.getbuffer())
            
        # Display preview
        st.subheader("📄 Raw Upload Preview (First 5 Rows)")
        try:
            if filename.endswith(".csv"):
                preview_df = pd.read_csv(filepath)
            else:
                preview_df = pd.read_excel(filepath)
            st.write(preview_df.head(5))
        except Exception as e:
            st.error(f"Error reading file: {e}")
            
        # Map Columns preview
        st.subheader("🔍 Auto-Column Mapping Inspector")
        mapped_df = ingestion.map_columns(preview_df)
        mapping_dict = {}
        for canonical in mapped_df.columns:
            found = False
            for col in preview_df.columns:
                mapped_col_tmp = ingestion.map_columns(preview_df[[col]])
                if canonical in mapped_col_tmp.columns and mapped_col_tmp[canonical].dropna().shape[0] == preview_df[col].dropna().shape[0] and preview_df[col].dropna().shape[0] > 0:
                    mapping_dict[canonical] = f"Mapped from: '{col}'"
                    found = True
                    break
            if not found:
                mapping_dict[canonical] = "⚠️ Missing / Blank (Set to Default)"
                
        st.json(mapping_dict)
        
        # Action button
        if st.button("▶️ Process & Audit Ingestion", type="primary"):
            with st.spinner("Executing pipeline: mapping -> cleaning -> auditing -> database merging -> Excel model recompilation..."):
                result = ingestion.process_file_upload(filepath, filename, db_path)
                
                st.markdown("---")
                st.subheader("📊 Ingestion Audit Report")
                
                # Verdict Badge
                verdict = result["status"]
                if verdict == "Accepted":
                    st.markdown("<h4>Verdict: <span class='badge-accepted'>✅ ACCEPTED</span></h4>", unsafe_allow_html=True)
                    st.success(f"Transactions merged successfully! {result['accepted_rows']} records loaded. Excel model recompiled.")
                elif verdict == "Needs Review":
                    st.markdown("<h4>Verdict: <span class='badge-review'>⚠️ NEEDS REVIEW</span></h4>", unsafe_allow_html=True)
                    st.warning(f"Transactions merged with warnings! {result['accepted_rows']} records loaded. Please review alerts.")
                else:
                    st.markdown("<h4>Verdict: <span class='badge-rejected'>🔴 REJECTED</span></h4>", unsafe_allow_html=True)
                    st.error("Transactions rejected! Quality score fell below the 50% threshold. Database rolled back.")
                    
                score = result["health_score"]
                st.metric("Batch Data Quality Score", f"{score*100:.1f}%")
                
                if result["issues"]:
                    st.subheader("🔍 Identified Data Anomalies Log")
                    for rule, entries in result["issues"].items():
                        severity = "🔴 Critical" if rule == "missing_date" else "🟡 Warning" if rule in ["future_dates", "negative_revenue", "negative_expenses", "null_amounts", "blank_fields"] else "⚠️ Info"
                        
                        st.markdown(f"**Rule: `{rule}`** — Severity: **{severity}**")
                        rows_str = ", ".join([f"Row {item['row']}" for item in entries[:10]])
                        if len(entries) > 10:
                            rows_str += f" and {len(entries) - 10} more rows..."
                        st.write(f"- {entries[0]['msg']} (Affected: {rows_str})")
                else:
                    st.success("Perfect Batch! 0 issues identified.")
                    
        # Cleanup
        try:
            os.remove(filepath)
        except:
            pass

# ----------------- PAGE 4: AUDIT LOG & DATA QUALITY -----------------
elif menu == "🔍 Audit Log & Data Quality":
    st.markdown("""
    <div class="title-banner">
        <h1>Financial Audit & Ingestion History</h1>
        <p>Trace data ingestion logs, system updates, and database integrity checks</p>
    </div>
    """, unsafe_allow_html=True)
    
    conn = sqlite3.connect(db_path)
    
    st.subheader("📥 Historical Ingestion Log")
    df_log = pd.read_sql_query("SELECT * FROM ingestion_log ORDER BY batch_id DESC", conn)
    
    if len(df_log) > 0:
        def color_status(val):
            if val == "Accepted":
                return 'color: #0d9488; font-weight: bold;'
            elif val == "Needs Review":
                return 'color: #d97706; font-weight: bold;'
            else:
                return 'color: #dc2626; font-weight: bold;'
                
        df_log_styled = df_log.rename(columns={
            "batch_id": "Batch ID",
            "submitted_at": "Timestamp",
            "filename": "Filename",
            "row_count": "Total Rows",
            "status": "Verdict Status",
            "health_score": "Health Score",
            "accepted_rows": "Accepted Rows",
            "rejected_rows": "Rejected Rows"
        })
        
        st.dataframe(
            df_log_styled.style.map(color_status, subset=["Verdict Status"])
            .format({"Health Score": "{:.1%}"}),
            use_container_width=True,
            hide_index=True
        )
    else:
        st.info("No files ingested yet.")
        
    st.markdown("---")
    
    st.subheader("⚙️ Live Database Quality Audits")
    cursor = conn.cursor()
    
    cursor.execute("SELECT COUNT(*) FROM dim_departments")
    total_depts = cursor.fetchone()[0]
    
    cursor.execute("SELECT COUNT(*) FROM dim_regions")
    total_regions = cursor.fetchone()[0]
    
    cursor.execute("SELECT COUNT(*) FROM fact_transactions")
    total_trans = cursor.fetchone()[0]
    
    cursor.execute("SELECT COUNT(*) FROM fact_transactions WHERE revenue < 0")
    neg_rev_prod = cursor.fetchone()[0]
    
    cursor.execute("SELECT COUNT(*) FROM fact_transactions WHERE expenses < 0")
    neg_exp_prod = cursor.fetchone()[0]
    
    conn.close()
    
    col_a, col_b = st.columns(2)
    with col_a:
        st.markdown(f"""
        <div style="background-color: #f8fafc; padding: 20px; border-radius: 12px; border: 1px solid #e2e8f0; margin-bottom: 15px;">
            <h4 style="margin-top:0; color:#1f4e78;">📋 Database Composition</h4>
            <p><b>Total Departments in Dim:</b> {total_depts}</p>
            <p><b>Total Regions in Dim:</b> {total_regions}</p>
            <p><b>Total Transactions in Fact Table:</b> {total_trans}</p>
        </div>
        """, unsafe_allow_html=True)
        
    with col_b:
        badge_rev = "<span class='badge-accepted'>Clean</span>" if neg_rev_prod == 0 else f"<span class='badge-rejected'>{neg_rev_prod} Bad Rows</span>"
        badge_exp = "<span class='badge-accepted'>Clean</span>" if neg_exp_prod == 0 else f"<span class='badge-rejected'>{neg_exp_prod} Bad Rows</span>"
        
        st.markdown(f"""
        <div style="background-color: #f8fafc; padding: 20px; border-radius: 12px; border: 1px solid #e2e8f0;">
            <h4 style="margin-top:0; color:#1f4e78;">🛡️ Production Integrity Audits</h4>
            <p><b>Negative Revenue Check:</b> {badge_rev}</p>
            <p><b>Negative Expense Check:</b> {badge_exp}</p>
        </div>
        """, unsafe_allow_html=True)
