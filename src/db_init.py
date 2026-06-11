import os
import pandas as pd
from datetime import datetime
import openpyxl

def init_excel_database(output_path, trans_csv, budget_csv, invoice_csv=None):
    print(f"Initializing Excel Database at: {output_path}")
    
    # Load raw data files
    df_trans = pd.read_csv(trans_csv)
    df_budget = pd.read_csv(budget_csv)
    
    # Create Dimensions
    depts = sorted(list(set(df_trans["Department"].dropna().unique()).union(set(df_budget["Department"].dropna().unique()))))
    df_depts = pd.DataFrame({
        "dept_id": range(1, len(depts) + 1),
        "dept_name": depts
    })
    dept_lookup = {name: i for i, name in zip(df_depts["dept_id"], df_depts["dept_name"])}
    
    regions = sorted(list(set(df_trans["Region"].dropna().unique()).union(set(df_budget["Region"].dropna().unique()))))
    df_regions = pd.DataFrame({
        "region_id": range(1, len(regions) + 1),
        "region_name": regions
    })
    region_lookup = {name: i for i, name in zip(df_regions["region_id"], df_regions["region_name"])}
    
    products = sorted(list(df_trans["Product Line"].dropna().unique()))
    df_products = pd.DataFrame({
        "product_id": range(1, len(products) + 1),
        "product_name": products
    })
    product_lookup = {name: i for i, name in zip(df_products["product_id"], df_products["product_name"])}
    
    categories = sorted(list(df_trans["Expense Category"].dropna().unique()))
    df_categories = pd.DataFrame({
        "category_id": range(1, len(categories) + 1),
        "category_name": categories
    })
    category_lookup = {name: i for i, name in zip(df_categories["category_id"], df_categories["category_name"])}
    
    # Map budgets
    budget_records = []
    for idx, row in df_budget.iterrows():
        dept_id = dept_lookup.get(row["Department"])
        region_id = region_lookup.get(row["Region"])
        budget_records.append({
            "budget_id": idx + 1,
            "month": row["Month"],
            "dept_id": dept_id,
            "region_id": region_id,
            "budgeted_revenue": float(row["Budgeted Revenue"]),
            "budgeted_expenses": float(row["Budgeted Expenses"])
        })
    df_budgets_sheet = pd.DataFrame(budget_records)
    
    # Deterministic client cohort profiles for auto-generation
    import random
    random.seed(100)
    clients = [f"CL-{i:03d}" for i in range(1, 101)]
    client_cohorts = {}
    for i, client in enumerate(clients):
        r = random.random()
        if r < 0.25: cohort_m = 1
        elif r < 0.40: cohort_m = 2
        elif r < 0.50: cohort_m = 3
        elif r < 0.58: cohort_m = 4
        elif r < 0.65: cohort_m = 5
        elif r < 0.72: cohort_m = 6
        elif r < 0.78: cohort_m = 7
        elif r < 0.84: cohort_m = 8
        elif r < 0.89: cohort_m = 9
        elif r < 0.93: cohort_m = 10
        elif r < 0.97: cohort_m = 11
        else: cohort_m = 12
        client_cohorts[client] = cohort_m
        
    # Map transactions
    trans_records = []
    has_client_col = "Client ID" in df_trans.columns
    
    for idx, row in df_trans.iterrows():
        dept_id = dept_lookup.get(row["Department"])
        region_id = region_lookup.get(row["Region"])
        product_id = product_lookup.get(row["Product Line"])
        category_id = category_lookup.get(row["Expense Category"])
        
        rev_val = float(row["Revenue"]) if pd.notna(row["Revenue"]) else 0.0
        
        if has_client_col and pd.notna(row["Client ID"]):
            client_val = row["Client ID"]
        elif rev_val > 0:
            # Deterministically choose an active client cohort
            dt_str = row["Date"]
            month_val = int(dt_str.split("-")[1])
            available_clients = [c for c, m in client_cohorts.items() if m <= month_val]
            if available_clients:
                seed_val = int(dt_str.replace("-", "")) + int(rev_val * 100)
                state = random.getstate()
                random.seed(seed_val)
                client_val = random.choice(available_clients)
                random.setstate(state)
            else:
                client_val = "CL-001"
        else:
            client_val = ""
            
        trans_records.append({
            "transaction_id": idx + 1,
            "date": row["Date"],
            "revenue": rev_val,
            "expenses": float(row["Expenses"]) if pd.notna(row["Expenses"]) else 0.0,
            "dept_id": dept_id,
            "region_id": region_id,
            "product_id": product_id,
            "category_id": category_id,
            "client_id": client_val
        })
    df_trans_sheet = pd.DataFrame(trans_records)
    
    # Map invoices
    if invoice_csv is None:
        dir_name = os.path.dirname(trans_csv)
        invoice_csv = os.path.join(dir_name, "financial_invoices_raw.csv")
        
    if os.path.exists(invoice_csv):
        print(f"Seeding invoices from: {invoice_csv}")
        df_inv = pd.read_csv(invoice_csv)
        inv_records = []
        for idx, row in df_inv.iterrows():
            status = row["Status"]
            if status == "1-Created": status_num = 1
            elif status == "2-Delivered": status_num = 2
            elif status == "3-Approved": status_num = 3
            elif status == "4-Pending Payment": status_num = 4
            elif status == "5-Settled": status_num = 5
            else: status_num = 1
            
            inv_records.append({
                "invoice_id": row["Invoice ID"],
                "client_id": row["Client ID"],
                "issue_date": row["Issue Date"],
                "amount": float(row["Amount"]) if pd.notna(row["Amount"]) else 0.0,
                "status": status,
                "status_num": status_num
            })
        df_invoices_sheet = pd.DataFrame(inv_records)
    else:
        print("Warning: Invoices CSV not found. Auto-generating invoice sheet...")
        inv_records = []
        state = random.getstate()
        random.seed(100)
        idx_count = 0
        for _, row in df_trans_sheet[df_trans_sheet["revenue"] > 0].iterrows():
            dt_str = row["date"]
            amount = row["revenue"]
            client_id = row["client_id"]
            inv_id = f"INV-{10000 + idx_count}"
            idx_count += 1
            
            month_val = int(dt_str.split("-")[1])
            r = random.random()
            if month_val < 11:
                if r < 0.92: status = "5-Settled"; status_num = 5
                elif r < 0.97: status = "4-Pending Payment"; status_num = 4
                elif r < 0.99: status = "3-Approved"; status_num = 3
                else: status = "2-Delivered"; status_num = 2
            else:
                if r < 0.40: status = "5-Settled"; status_num = 5
                elif r < 0.70: status = "4-Pending Payment"; status_num = 4
                elif r < 0.85: status = "3-Approved"; status_num = 3
                elif r < 0.95: status = "2-Delivered"; status_num = 2
                else: status = "1-Created"; status_num = 1
                
            inv_records.append({
                "invoice_id": inv_id,
                "client_id": client_id,
                "issue_date": dt_str,
                "amount": amount,
                "status": status,
                "status_num": status_num
            })
        random.setstate(state)
        df_invoices_sheet = pd.DataFrame(inv_records)
        
    # Create ingestion log sheet
    df_ingestion_log_sheet = pd.DataFrame([{
        "batch_id": 1,
        "submitted_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "filename": os.path.basename(trans_csv),
        "row_count": len(df_trans),
        "status": "Accepted",
        "health_score": 100.0,
        "issues_json": "[]",
        "accepted_rows": len(df_trans),
        "rejected_rows": 0
    }])
    
    # Save sheets to output path
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df_trans_sheet.to_excel(writer, sheet_name="fact_transactions", index=False)
        df_budgets_sheet.to_excel(writer, sheet_name="fact_budgets", index=False)
        df_invoices_sheet.to_excel(writer, sheet_name="fact_invoices", index=False)
        df_depts.to_excel(writer, sheet_name="dim_departments", index=False)
        df_regions.to_excel(writer, sheet_name="dim_regions", index=False)
        df_products.to_excel(writer, sheet_name="dim_products", index=False)
        df_categories.to_excel(writer, sheet_name="dim_expense_categories", index=False)
        df_ingestion_log_sheet.to_excel(writer, sheet_name="ingestion_log", index=False)
        
    print(f"Excel Database created at {output_path} containing {len(df_trans_sheet)} transactions.")

if __name__ == "__main__":
    out_file = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "financial_analysis.xlsx")
    trans_csv = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "financial_data_raw.csv")
    budget_csv = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "financial_budget_raw.csv")
    invoice_csv = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "financial_invoices_raw.csv")
    
    init_excel_database(out_file, trans_csv, budget_csv, invoice_csv)
