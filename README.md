# FinSight — Corporate Financial Performance & Profitability Analytics

An end-to-end financial data engineering and business intelligence platform designed to ingest daily corporate transactions, audit data quality, reconcile monthly budgets, and analyze profitability. FinSight bridges raw data engineering (SQL star schemas, Python ETL) with visual storytelling (multi-page Streamlit dashboards, Plotly charts) and advanced financial modeling (custom-formatted Excel workbooks built programmatically).

👉 **[Live Web Application Link](https://financial-performance-analytics-7pgbkyrynm9jbstl2qvgh3.streamlit.app/)**

---

## 🚀 Key Features

1. **Interactive Multi-Page BI Dashboard:**
   - **Executive Summary:** Core financial KPIs (Total Revenue, Operating Expenses, Net Operating Profit, Profit Margin %, ROI %) and trend lines (MoM growth, running totals).
   - **Profitability & Budgets:** Granular breakdowns by region and product segment (identifying high-performing lines) and a department budget-to-actual variance matrix.
   - **Cohorts & Funnels:** Interactive Client Revenue Retention Heatmaps (Net Revenue Retention & Logo Retention) and Accounts Receivable Invoicing conversion funnel.
   - **Ingest Transactions:** Drag-and-drop CSV/Excel portal for daily transactions uploads.
   - **Audit Trail:** logs of files processed and database-wide health diagnostics.

2. **Fuzzy Column Ingestion & Standardisation:**
   - Normalizes incoming column headers dynamically (e.g. mapping `spending`, `outflow`, or `costs` automatically to `expenses`).
   - Automatically cleans currency formatting, strips commas/symbols, and handles accounting negative brackets `(500.00)` -> `-500.00`.

3. **9-Point Data Quality Audit Gate:**
   - Evaluates each batch:
     - **Critical:** Missing dates.
     - **Warning:** Future dates, negative revenue/expense inputs, zero amounts, blank dimension values.
     - **Info:** Large transaction spikes (Revenue > $50K, Expense > $25K), duplicate rows.
   - Automatically computes a **Batch Health Score**. Batches with score `< 50%` are automatically rejected and rolled back to preserve database integrity.

4. **Programmatic Excel Workbook Compiler (`openpyxl`):**
   - Automatically compiles and refines a beautifully styled Excel workbook (`financial_analysis.xlsx`) after every successful data ingestion.
   - **Native Excel Formulas:** Writes actual uppercase formulas (e.g., `=SUM(...)`, `=IF(...)`, `=AVERAGE(...)`) into cells so that the spreadsheet behaves dynamically for users rather than displaying flat, hardcoded values.
   - Designed with a professional corporate theme (Segoe UI, navy fills, double accounting underlines, auto-fitting column widths).

5. **Advanced SQL Analytics Engine:**
   - Operates on a structured SQLite **Star Schema** utilizing Common Table Expressions (CTEs), Joins, and Window Functions (calculating client NRR cohorts, invoice settlement funnels, `LAG` for MoM growth, and running totals).

---

## 📊 Database Architecture (Star Schema)

The database design normalizes daily transaction records and monthly budgets into a high-performance relational structure:

```mermaid
erDiagram
    fact_transactions {
        int transaction_id PK
        string date
        real revenue
        real expenses
        int dept_id FK
        int region_id FK
        int product_id FK
        int category_id FK
        string client_id
    }
    fact_budgets {
        int budget_id PK
        string month
        int dept_id FK
        int region_id FK
        real budgeted_revenue
        real budgeted_expenses
    }
    fact_invoices {
        string invoice_id PK
        string client_id
        string issue_date
        real amount
        string status
    }
    dim_departments {
        int dept_id PK
        string dept_name
    }
    dim_regions {
        int region_id PK
        string region_name
    }
    dim_products {
        int product_id PK
        string product_name
    }
    dim_expense_categories {
        int category_id PK
        string category_name
    }
    fact_transactions }|--|| dim_departments : "references"
    fact_transactions }|--|| dim_regions : "references"
    fact_transactions }|--|| dim_products : "references"
    fact_transactions }|--|| dim_expense_categories : "references"
    fact_budgets }|--|| dim_departments : "references"
    fact_budgets }|--|| dim_regions : "references"
```

---

## 💡 Advanced SQL Query Showcases

### 1. Monthly Performance, Growth & Running Totals (CTEs & Window Functions)
Calculates monthly revenue, expenses, net profit, running revenue total, and Month-over-Month growth rate:

```sql
WITH MonthlyAggs AS (
    SELECT 
        strftime('%Y-%m', date) as month,
        SUM(revenue) as revenue,
        SUM(expenses) as expenses,
        SUM(revenue) - SUM(expenses) as net_profit
    FROM fact_transactions
    GROUP BY month
),
MonthlyGrowth AS (
    SELECT 
        month,
        revenue,
        expenses,
        net_profit,
        LAG(revenue) OVER (ORDER BY month) as prev_month_revenue,
        SUM(revenue) OVER (ORDER BY month ROWS UNBOUNDED PRECEDING) as running_total_revenue
    FROM MonthlyAggs
)
SELECT 
    month,
    ROUND(revenue, 2) as revenue,
    ROUND(expenses, 2) as expenses,
    ROUND(net_profit, 2) as net_profit,
    ROUND(running_total_revenue, 2) as running_total_revenue,
    ROUND(
        CASE 
            WHEN prev_month_revenue IS NULL THEN 0.0 
            ELSE ((revenue - prev_month_revenue) / prev_month_revenue) * 100.0 
        END, 
        2
    ) as rev_growth_pct
FROM MonthlyGrowth
ORDER BY month;
```

### 2. Departmental Budget vs Actual Variance (CTEs & Joins)
Reconciles actual daily transactions against monthly budget targets per department:

```sql
WITH Actuals AS (
    SELECT 
        dept_id,
        SUM(revenue) as actual_revenue,
        SUM(expenses) as actual_expenses
    FROM fact_transactions
    GROUP BY dept_id
),
Budgets AS (
    SELECT 
        dept_id,
        SUM(budgeted_revenue) as budget_revenue,
        SUM(budgeted_expenses) as budget_expenses
    FROM fact_budgets
    GROUP BY dept_id
)
SELECT 
    d.dept_name as department,
    ROUND(a.actual_revenue, 2) as actual_revenue,
    ROUND(b.budget_revenue, 2) as budget_revenue,
    ROUND(a.actual_revenue - b.budget_revenue, 2) as revenue_variance,
    ROUND(a.actual_expenses, 2) as actual_expenses,
    ROUND(b.budget_expenses, 2) as budget_expenses,
    ROUND(b.budget_expenses - a.actual_expenses, 2) as expense_variance
FROM Actuals a
JOIN Budgets b ON a.dept_id = b.dept_id
JOIN dim_departments d ON a.dept_id = d.dept_id
ORDER BY actual_revenue DESC;
```

### 3. Net Revenue Retention (NRR) Cohort Analysis (CTEs & Substr math)
Computes monthly client cohorts and tracks their Net Revenue Retention percentage over time:

```sql
WITH ClientCohort AS (
    SELECT 
        client_id,
        MIN(strftime('%Y-%m', date)) as cohort_month
    FROM fact_transactions
    WHERE revenue > 0 AND client_id IS NOT NULL AND client_id != ''
    GROUP BY client_id
),
ClientMonthlyRevenue AS (
    SELECT 
        client_id,
        strftime('%Y-%m', date) as trans_month,
        SUM(revenue) as monthly_rev
    FROM fact_transactions
    WHERE revenue > 0 AND client_id IS NOT NULL AND client_id != ''
    GROUP BY client_id, trans_month
),
CohortSizes AS (
    SELECT 
        cohort_month,
        COUNT(DISTINCT client_id) as total_clients
    FROM ClientCohort
    GROUP BY cohort_month
),
CohortSpend AS (
    SELECT 
        cc.cohort_month,
        (CAST(substr(cm.trans_month, 1, 4) AS INTEGER) - CAST(substr(cc.cohort_month, 1, 4) AS INTEGER)) * 12 + 
        (CAST(substr(cm.trans_month, 6, 2) AS INTEGER) - CAST(substr(cc.cohort_month, 6, 2) AS INTEGER)) as elapsed_months,
        SUM(cm.monthly_rev) as cohort_revenue,
        COUNT(DISTINCT cm.client_id) as active_clients
    FROM ClientCohort cc
    JOIN ClientMonthlyRevenue cm ON cc.client_id = cm.client_id
    GROUP BY cc.cohort_month, elapsed_months
),
BaseRevenue AS (
    SELECT 
        cohort_month,
        cohort_revenue as base_rev
    FROM CohortSpend
    WHERE elapsed_months = 0
)
SELECT 
    cs.cohort_month,
    cs.total_clients,
    ROUND(br.base_rev, 2) as base_rev,
    c.elapsed_months,
    ROUND(c.cohort_revenue, 2) as cohort_revenue,
    c.active_clients,
    ROUND((c.cohort_revenue / br.base_rev) * 100.0, 2) as revenue_retention_pct,
    ROUND((c.active_clients * 100.0 / cs.total_clients), 2) as client_retention_pct
FROM CohortSpend c
JOIN CohortSizes cs ON c.cohort_month = cs.cohort_month
JOIN BaseRevenue br ON c.cohort_month = br.cohort_month
WHERE c.elapsed_months >= 0 AND c.elapsed_months < 12
ORDER BY cs.cohort_month, c.elapsed_months;
```

### 4. Accounts Receivable Invoicing Funnel (CTEs & Cumulative Union)
Tracks cash collections and invoice statuses cumulatively:

```sql
WITH InvoiceNumericStatus AS (
    SELECT 
        invoice_id,
        amount,
        CASE 
            WHEN status = '1-Created' THEN 1
            WHEN status = '2-Delivered' THEN 2
            WHEN status = '3-Approved' THEN 3
            WHEN status = '4-Pending Payment' THEN 4
            WHEN status = '5-Settled' THEN 5
            ELSE 1
        END as status_num
    FROM fact_invoices
)
SELECT 
    '1. Created' as stage,
    COUNT(*) as invoice_count,
    ROUND(SUM(amount), 2) as total_amount,
    100.0 as pct_conversion
FROM InvoiceNumericStatus
UNION ALL
SELECT 
    '2. Delivered' as stage,
    COUNT(*) as invoice_count,
    ROUND(SUM(amount), 2) as total_amount,
    ROUND(COUNT(*) * 100.0 / (SELECT COUNT(*) FROM InvoiceNumericStatus), 2) as pct_conversion
FROM InvoiceNumericStatus WHERE status_num >= 2
UNION ALL
SELECT 
    '3. Approved' as stage,
    COUNT(*) as invoice_count,
    ROUND(SUM(amount), 2) as total_amount,
    ROUND(COUNT(*) * 100.0 / (SELECT COUNT(*) FROM InvoiceNumericStatus), 2) as pct_conversion
FROM InvoiceNumericStatus WHERE status_num >= 3
UNION ALL
SELECT 
    '4. Pending Payment' as stage,
    COUNT(*) as invoice_count,
    ROUND(SUM(amount), 2) as total_amount,
    ROUND(COUNT(*) * 100.0 / (SELECT COUNT(*) FROM InvoiceNumericStatus), 2) as pct_conversion
FROM InvoiceNumericStatus WHERE status_num >= 4
UNION ALL
SELECT 
    '5. Settled' as stage,
    COUNT(*) as invoice_count,
    ROUND(SUM(amount), 2) as total_amount,
    ROUND(COUNT(*) * 100.0 / (SELECT COUNT(*) FROM InvoiceNumericStatus), 2) as pct_conversion
FROM InvoiceNumericStatus WHERE status_num = 5;
```

---

## 🛠️ How to Run Locally

### 1. Install Dependencies
```bash
pip install -r requirements.txt
```

### 2. Generate Data & Seed Database
Initialize the database tables and load the initial transactions:
```bash
python generate_data.py
python src/db_init.py
```

### 3. Generate Styled Excel Model
```bash
python src/excel_generator.py
```

### 4. Launch Dashboard
```bash
streamlit run src/app.py
```

---

## 📬 Contact & Connections

- **Author:** Monisa L.
- **Email:** [monisa.asi@gmail.com](mailto:monisa.asi@gmail.com)
- **LinkedIn:** [linkedin.com/in/monisa-l-333546366](https://www.linkedin.com/in/monisa-l-333546366)
- **GitHub Profile:** [github.com/Monisa-Analyst](https://github.com/Monisa-Analyst)
