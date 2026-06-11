import sqlite3
import pandas as pd

def get_connection(db_path="src/financial.db"):
    conn = sqlite3.connect(db_path)
    conn.execute("PRAGMA foreign_keys = ON;")
    return conn

def get_kpis(db_path="src/financial.db"):
    conn = get_connection(db_path)
    
    # 1. Total Revenue
    total_rev = pd.read_sql_query("SELECT SUM(revenue) as total FROM fact_transactions", conn).iloc[0]['total']
    total_rev = total_rev if total_rev else 0.0
    
    # 2. Total Expenses
    total_exp = pd.read_sql_query("SELECT SUM(expenses) as total FROM fact_transactions", conn).iloc[0]['total']
    total_exp = total_exp if total_exp else 0.0
    
    # 3. Net Profit
    net_profit = total_rev - total_exp
    
    # 4. Profit Margin %
    margin = (net_profit / total_rev) if total_rev > 0 else 0.0
    
    # 5. ROI % (Net Profit / Expenses)
    roi = (net_profit / total_exp) if total_exp > 0 else 0.0
    
    # 6. Budget Variance (Actual Revenue vs Budgeted Revenue)
    budget_rev = pd.read_sql_query("SELECT SUM(budgeted_revenue) as total FROM fact_budgets", conn).iloc[0]['total']
    budget_rev = budget_rev if budget_rev else 0.0
    
    variance = total_rev - budget_rev
    variance_pct = (variance / budget_rev) if budget_rev > 0 else 0.0
    
    conn.close()
    
    return {
        "total_revenue": total_rev,
        "total_expenses": total_exp,
        "net_profit": net_profit,
        "profit_margin": margin,
        "roi": roi,
        "budget_variance": variance,
        "budget_variance_pct": variance_pct
    }

def get_monthly_performance(db_path="src/financial.db"):
    # SQL query for monthly metrics, growth rates, and running totals
    query = """
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
    """
    conn = get_connection(db_path)
    df = pd.read_sql_query(query, conn)
    conn.close()
    return df

def get_department_performance(db_path="src/financial.db"):
    # Variance of actuals vs budgets by department
    query = """
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
        ROUND(b.budget_expenses - a.actual_expenses, 2) as expense_variance -- positive means under budget
    FROM Actuals a
    JOIN Budgets b ON a.dept_id = b.dept_id
    JOIN dim_departments d ON a.dept_id = d.dept_id
    ORDER BY actual_revenue DESC;
    """
    conn = get_connection(db_path)
    df = pd.read_sql_query(query, conn)
    conn.close()
    return df

def get_regional_profitability(db_path="src/financial.db"):
    query = """
    SELECT 
        r.region_name as region,
        ROUND(SUM(f.revenue), 2) as revenue,
        ROUND(SUM(f.expenses), 2) as expenses,
        ROUND(SUM(f.revenue) - SUM(f.expenses), 2) as net_profit,
        ROUND(((SUM(f.revenue) - SUM(f.expenses)) / SUM(f.revenue)) * 100.0, 2) as profit_margin_pct
    FROM fact_transactions f
    JOIN dim_regions r ON f.region_id = r.region_id
    GROUP BY region
    ORDER BY net_profit DESC;
    """
    conn = get_connection(db_path)
    df = pd.read_sql_query(query, conn)
    conn.close()
    return df

def get_product_profitability(db_path="src/financial.db"):
    query = """
    SELECT 
        p.product_name as product_line,
        ROUND(SUM(f.revenue), 2) as revenue,
        ROUND(SUM(f.expenses), 2) as expenses,
        ROUND(SUM(f.revenue) - SUM(f.expenses), 2) as net_profit,
        ROUND(((SUM(f.revenue) - SUM(f.expenses)) / SUM(f.revenue)) * 100.0, 2) as profit_margin_pct,
        ROUND((SUM(f.revenue) - SUM(f.expenses)) * 100.0 / (SELECT SUM(revenue) - SUM(expenses) FROM fact_transactions), 2) as profit_contribution_pct
    FROM fact_transactions f
    JOIN dim_products p ON f.product_id = p.product_id
    GROUP BY product_line
    ORDER BY net_profit DESC;
    """
    conn = get_connection(db_path)
    df = pd.read_sql_query(query, conn)
    conn.close()
    return df

def get_expenses_by_category(db_path="src/financial.db"):
    query = """
    SELECT 
        c.category_name as category,
        ROUND(SUM(f.expenses), 2) as expenses,
        ROUND(SUM(f.expenses) * 100.0 / (SELECT SUM(expenses) FROM fact_transactions), 2) as expense_share_pct
    FROM fact_transactions f
    JOIN dim_expense_categories c ON f.category_id = c.category_id
    GROUP BY category
    ORDER BY expenses DESC;
    """
    conn = get_connection(db_path)
    df = pd.read_sql_query(query, conn)
    conn.close()
    return df

def get_revenue_cohorts(db_path="src/financial.db"):
    query = """
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
    """
    conn = get_connection(db_path)
    df = pd.read_sql_query(query, conn)
    conn.close()
    return df

def get_invoice_funnel(db_path="src/financial.db"):
    query = """
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
    """
    conn = get_connection(db_path)
    df = pd.read_sql_query(query, conn)
    conn.close()
    return df

if __name__ == "__main__":
    print("Testing financial analytics query engine...")
    print(get_kpis("src/financial.db"))
    print("\nRegional Profitability:")
    print(get_regional_profitability("src/financial.db"))
    print("\nRevenue Cohort sample:")
    print(get_revenue_cohorts("src/financial.db").head())
    print("\nInvoice Funnel:")
    print(get_invoice_funnel("src/financial.db"))

