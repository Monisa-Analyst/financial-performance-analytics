import os
import pandas as pd
import numpy as np

def load_excel_sheet(sheet_name, excel_path="financial_analysis.xlsx"):
    if not os.path.exists(excel_path):
        # Fallback to local root or data folder if path is relative
        base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        excel_path = os.path.join(base_dir, excel_path)
    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"Excel database file not found at: {excel_path}")
    return pd.read_excel(excel_path, sheet_name=sheet_name)

def get_kpis(excel_path="financial_analysis.xlsx"):
    df_trans = load_excel_sheet("fact_transactions", excel_path)
    df_budgets = load_excel_sheet("fact_budgets", excel_path)
    
    total_rev = float(df_trans["revenue"].sum()) if not df_trans.empty else 0.0
    total_exp = float(df_trans["expenses"].sum()) if not df_trans.empty else 0.0
    net_profit = total_rev - total_exp
    
    margin = (net_profit / total_rev) if total_rev > 0 else 0.0
    roi = (net_profit / total_exp) if total_exp > 0 else 0.0
    
    budget_rev = float(df_budgets["budgeted_revenue"].sum()) if not df_budgets.empty else 0.0
    variance = total_rev - budget_rev
    variance_pct = (variance / budget_rev) if budget_rev > 0 else 0.0
    
    return {
        "total_revenue": total_rev,
        "total_expenses": total_exp,
        "net_profit": net_profit,
        "profit_margin": margin,
        "roi": roi,
        "budget_variance": variance,
        "budget_variance_pct": variance_pct
    }

def get_monthly_performance(excel_path="financial_analysis.xlsx"):
    df_trans = load_excel_sheet("fact_transactions", excel_path)
    if df_trans.empty:
        return pd.DataFrame(columns=["month", "revenue", "expenses", "net_profit", "running_total_revenue", "rev_growth_pct"])
        
    df_trans["month"] = df_trans["date"].str[:7]
    df_monthly = df_trans.groupby("month")[["revenue", "expenses"]].sum().reset_index()
    df_monthly = df_monthly.sort_values("month")
    
    df_monthly["net_profit"] = df_monthly["revenue"] - df_monthly["expenses"]
    df_monthly["running_total_revenue"] = df_monthly["revenue"].cumsum()
    
    df_monthly["prev_month_revenue"] = df_monthly["revenue"].shift(1)
    df_monthly["rev_growth_pct"] = np.where(
        df_monthly["prev_month_revenue"].isna(),
        0.0,
        ((df_monthly["revenue"] - df_monthly["prev_month_revenue"]) / df_monthly["prev_month_revenue"]) * 100.0
    )
    df_monthly["rev_growth_pct"] = df_monthly["rev_growth_pct"].round(2)
    df_monthly["net_profit"] = df_monthly["net_profit"].round(2)
    df_monthly["running_total_revenue"] = df_monthly["running_total_revenue"].round(2)
    
    return df_monthly.drop(columns=["prev_month_revenue"])

def get_department_performance(excel_path="financial_analysis.xlsx"):
    df_trans = load_excel_sheet("fact_transactions", excel_path)
    df_budgets = load_excel_sheet("fact_budgets", excel_path)
    df_depts = load_excel_sheet("dim_departments", excel_path)
    
    if df_trans.empty or df_budgets.empty or df_depts.empty:
        return pd.DataFrame(columns=["department", "actual_revenue", "budget_revenue", "revenue_variance", "actual_expenses", "budget_expenses", "expense_variance"])
        
    actuals = df_trans.groupby("dept_id")[["revenue", "expenses"]].sum().reset_index()
    actuals.rename(columns={"revenue": "actual_revenue", "expenses": "actual_expenses"}, inplace=True)
    
    budgets = df_budgets.groupby("dept_id")[["budgeted_revenue", "budgeted_expenses"]].sum().reset_index()
    budgets.rename(columns={"budgeted_revenue": "budget_revenue", "budgeted_expenses": "budget_expenses"}, inplace=True)
    
    merged = pd.merge(actuals, budgets, on="dept_id", how="outer").fillna(0)
    merged = pd.merge(merged, df_depts, on="dept_id", how="left")
    
    merged["revenue_variance"] = merged["actual_revenue"] - merged["budget_revenue"]
    merged["expense_variance"] = merged["budget_expenses"] - merged["actual_expenses"]
    
    # Round metrics
    for col in ["actual_revenue", "budget_revenue", "revenue_variance", "actual_expenses", "budget_expenses", "expense_variance"]:
        merged[col] = merged[col].round(2)
        
    merged.rename(columns={"dept_name": "department"}, inplace=True)
    return merged[["department", "actual_revenue", "budget_revenue", "revenue_variance", "actual_expenses", "budget_expenses", "expense_variance"]].sort_values("actual_revenue", ascending=False)

def get_regional_profitability(excel_path="financial_analysis.xlsx"):
    df_trans = load_excel_sheet("fact_transactions", excel_path)
    df_regions = load_excel_sheet("dim_regions", excel_path)
    
    if df_trans.empty or df_regions.empty:
        return pd.DataFrame(columns=["region", "revenue", "expenses", "net_profit", "profit_margin_pct"])
        
    reg_summary = df_trans.groupby("region_id")[["revenue", "expenses"]].sum().reset_index()
    merged = pd.merge(reg_summary, df_regions, on="region_id", how="left")
    
    merged["net_profit"] = merged["revenue"] - merged["expenses"]
    merged["profit_margin_pct"] = np.where(
        merged["revenue"] > 0,
        (merged["net_profit"] / merged["revenue"]) * 100.0,
        0.0
    )
    
    for col in ["revenue", "expenses", "net_profit", "profit_margin_pct"]:
        merged[col] = merged[col].round(2)
        
    merged.rename(columns={"region_name": "region"}, inplace=True)
    return merged[["region", "revenue", "expenses", "net_profit", "profit_margin_pct"]].sort_values("net_profit", ascending=False)

def get_product_profitability(excel_path="financial_analysis.xlsx"):
    df_trans = load_excel_sheet("fact_transactions", excel_path)
    df_products = load_excel_sheet("dim_products", excel_path)
    
    if df_trans.empty or df_products.empty:
        return pd.DataFrame(columns=["product_line", "revenue", "expenses", "net_profit", "profit_margin_pct", "profit_contribution_pct"])
        
    prod_summary = df_trans.groupby("product_id")[["revenue", "expenses"]].sum().reset_index()
    merged = pd.merge(prod_summary, df_products, on="product_id", how="left")
    
    merged["net_profit"] = merged["revenue"] - merged["expenses"]
    merged["profit_margin_pct"] = np.where(
        merged["revenue"] > 0,
        (merged["net_profit"] / merged["revenue"]) * 100.0,
        0.0
    )
    
    total_net_profit = merged["net_profit"].sum()
    merged["profit_contribution_pct"] = np.where(
        total_net_profit != 0,
        (merged["net_profit"] / total_net_profit) * 100.0,
        0.0
    )
    
    for col in ["revenue", "expenses", "net_profit", "profit_margin_pct", "profit_contribution_pct"]:
        merged[col] = merged[col].round(2)
        
    merged.rename(columns={"product_name": "product_line"}, inplace=True)
    return merged[["product_line", "revenue", "expenses", "net_profit", "profit_margin_pct", "profit_contribution_pct"]].sort_values("net_profit", ascending=False)

def get_expenses_by_category(excel_path="financial_analysis.xlsx"):
    df_trans = load_excel_sheet("fact_transactions", excel_path)
    df_cats = load_excel_sheet("dim_expense_categories", excel_path)
    
    if df_trans.empty or df_cats.empty:
        return pd.DataFrame(columns=["category", "expenses", "expense_share_pct"])
        
    cat_summary = df_trans.groupby("category_id")[["expenses"]].sum().reset_index()
    merged = pd.merge(cat_summary, df_cats, on="category_id", how="left")
    
    total_expenses = merged["expenses"].sum()
    merged["expense_share_pct"] = np.where(
        total_expenses > 0,
        (merged["expenses"] / total_expenses) * 100.0,
        0.0
    )
    
    for col in ["expenses", "expense_share_pct"]:
        merged[col] = merged[col].round(2)
        
    merged.rename(columns={"category_name": "category"}, inplace=True)
    return merged[["category", "expenses", "expense_share_pct"]].sort_values("expenses", ascending=False)

def get_revenue_cohorts(excel_path="financial_analysis.xlsx"):
    df_trans = load_excel_sheet("fact_transactions", excel_path)
    if df_trans.empty:
        return pd.DataFrame(columns=["cohort_month", "total_clients", "base_rev", "elapsed_months", "cohort_revenue", "active_clients", "revenue_retention_pct", "client_retention_pct"])
        
    # Filter valid clients with positive revenue
    df_valid = df_trans[(df_trans["revenue"] > 0) & (df_trans["client_id"].notna()) & (df_trans["client_id"] != "")].copy()
    if df_valid.empty:
        return pd.DataFrame(columns=["cohort_month", "total_clients", "base_rev", "elapsed_months", "cohort_revenue", "active_clients", "revenue_retention_pct", "client_retention_pct"])
        
    df_valid["trans_month"] = df_valid["date"].str[:7]
    
    # Calculate cohort month per client (first purchase month)
    df_cohorts = df_valid.groupby("client_id")["trans_month"].min().reset_index()
    df_cohorts.rename(columns={"trans_month": "cohort_month"}, inplace=True)
    
    # Calculate client monthly revenue
    df_client_monthly = df_valid.groupby(["client_id", "trans_month"])["revenue"].sum().reset_index()
    
    # Merge client revenue with their cohort month
    df_merged = pd.merge(df_client_monthly, df_cohorts, on="client_id", how="left")
    
    # Calculate elapsed months
    def diff_months(row):
        c_yr, c_mo = int(row["cohort_month"][:4]), int(row["cohort_month"][5:])
        t_yr, t_mo = int(row["trans_month"][:4]), int(row["trans_month"][5:])
        return (t_yr - c_yr) * 12 + (t_mo - c_mo)
        
    df_merged["elapsed_months"] = df_merged.apply(diff_months, axis=1)
    
    # Get base cohort sizes (total clients per cohort month)
    df_cohort_sizes = df_cohorts.groupby("cohort_month")["client_id"].nunique().reset_index()
    df_cohort_sizes.rename(columns={"client_id": "total_clients"}, inplace=True)
    
    # Aggregate cohort spending and active clients count by cohort_month and elapsed_months
    df_cohort_spend = df_merged.groupby(["cohort_month", "elapsed_months"]).agg(
        cohort_revenue=("revenue", "sum"),
        active_clients=("client_id", "nunique")
    ).reset_index()
    
    # Get base revenue for each cohort (elapsed_months = 0)
    df_base_rev = df_cohort_spend[df_cohort_spend["elapsed_months"] == 0][["cohort_month", "cohort_revenue"]].copy()
    df_base_rev.rename(columns={"cohort_revenue": "base_rev"}, inplace=True)
    
    # Merge metrics
    df_final = pd.merge(df_cohort_spend, df_cohort_sizes, on="cohort_month", how="left")
    df_final = pd.merge(df_final, df_base_rev, on="cohort_month", how="left")
    
    # Filters only positive elapsed quarters/months up to 12
    df_final = df_final[(df_final["elapsed_months"] >= 0) & (df_final["elapsed_months"] < 12)]
    
    df_final["revenue_retention_pct"] = np.where(
        df_final["base_rev"] > 0,
        (df_final["cohort_revenue"] / df_final["base_rev"]) * 100.0,
        0.0
    ).round(2)
    
    df_final["client_retention_pct"] = np.where(
        df_final["total_clients"] > 0,
        (df_final["active_clients"] / df_final["total_clients"]) * 100.0,
        0.0
    ).round(2)
    
    # Format and sort
    df_final["base_rev"] = df_final["base_rev"].round(2)
    df_final["cohort_revenue"] = df_final["cohort_revenue"].round(2)
    
    return df_final.sort_values(["cohort_month", "elapsed_months"])

def get_invoice_funnel(excel_path="financial_analysis.xlsx"):
    df_inv = load_excel_sheet("fact_invoices", excel_path)
    if df_inv.empty:
        return pd.DataFrame(columns=["stage", "invoice_count", "total_amount", "pct_conversion"])
        
    stages = [
        ("1. Created", 1),
        ("2. Delivered", 2),
        ("3. Approved", 3),
        ("4. Pending Payment", 4),
        ("5. Settled", 5)
    ]
    
    total_count = len(df_inv)
    funnel_data = []
    
    for stage_name, stage_num in stages:
        df_stage = df_inv[df_inv["status_num"] >= stage_num]
        count = len(df_stage)
        amount = float(df_stage["amount"].sum())
        
        pct = (count / total_count * 100.0) if total_count > 0 else 0.0
        funnel_data.append({
            "stage": stage_name,
            "invoice_count": count,
            "total_amount": round(amount, 2),
            "pct_conversion": round(pct, 2)
        })
        
    return pd.DataFrame(funnel_data)

if __name__ == "__main__":
    print("Testing Pandas-Excel financial query engine...")
    kpis = get_kpis("financial_analysis.xlsx")
    print(f"Total Revenue: ${kpis['total_revenue']:,.2f}")
    print(f"Total Expenses: ${kpis['total_expenses']:,.2f}")
    print(f"Net Profit: ${kpis['net_profit']:,.2f}")
    print(f"Profit Margin: {kpis['profit_margin'] * 100:.2f}%")
    print("\nInvoice Funnel:")
    print(get_invoice_funnel("financial_analysis.xlsx"))
