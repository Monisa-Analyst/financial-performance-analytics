import os
import json
import pandas as pd
import numpy as np
from datetime import datetime

# Import excel generator to re-trigger compilation on new ingestion
import excel_generator

COLUMN_MAPPING = {
    "date": ["date", "transaction_date", "trans_date", "day"],
    "revenue": ["revenue", "sales", "income", "turnover", "rev"],
    "expenses": ["expenses", "costs", "expense", "exp", "spending", "outflow"],
    "department": ["department", "dept", "cost_center", "division"],
    "region": ["region", "area", "zone", "location"],
    "product_line": ["product line", "product_line", "product", "segment", "line"],
    "expense_category": ["expense category", "expense_category", "category", "type", "expense_type"],
    "client_id": ["client id", "client_id", "client", "customer id", "customer_id", "customer"]
}

def map_columns(df):
    mapped_cols = {}
    lower_cols = {col.lower().strip().replace("_", " ").replace("  ", " "): col for col in df.columns}
    
    for canonical, variations in COLUMN_MAPPING.items():
        found = False
        for var in variations:
            if var in lower_cols:
                mapped_cols[lower_cols[var]] = canonical
                found = True
                break
        if not found:
            for raw_col_lower, raw_col_orig in lower_cols.items():
                if any(var in raw_col_lower for var in variations):
                    mapped_cols[raw_col_orig] = canonical
                    break
                    
    df_renamed = df.rename(columns=mapped_cols)
    for col in COLUMN_MAPPING.keys():
        if col not in df_renamed.columns:
            df_renamed[col] = np.nan
            
    return df_renamed[list(COLUMN_MAPPING.keys())]

def clean_numeric(val):
    if pd.isna(val) or val == "":
        return 0.0
    val_str = str(val).strip().replace("$", "").replace(",", "")
    if not val_str:
        return 0.0
    if val_str.startswith("(") and val_str.endswith(")"):
        val_str = "-" + val_str[1:-1]
    try:
        return float(val_str)
    except ValueError:
        return 0.0

def parse_date(val):
    if pd.isna(val) or val == "":
        return None
    val_str = str(val).strip()
    
    formats = [
        "%Y-%m-%d", "%Y-%m-%d %H:%M:%S",
        "%m/%d/%Y", "%m/%d/%Y %H:%M", "%m/%d/%y",
        "%d-%m-%Y", "%d/%m/%Y"
    ]
    
    for fmt in formats:
        try:
            dt = datetime.strptime(val_str, fmt)
            return dt.strftime("%Y-%m-%d")
        except ValueError:
            continue
            
    try:
        dt = pd.to_datetime(val_str)
        if pd.notna(dt):
            return dt.strftime("%Y-%m-%d")
    except:
        pass
        
    return None

def run_data_quality_checks(df, excel_path="financial_analysis.xlsx"):
    issues = {
        "missing_date": [],         # Critical
        "future_dates": [],         # Warning
        "negative_revenue": [],     # Warning
        "negative_expenses": [],    # Warning
        "null_amounts": [],         # Warning
        "blank_fields": [],         # Warning
        "overlapping_rows": [],     # Info
        "extreme_revenue": [],      # Info
        "extreme_expenses": []      # Info
    }
    
    n_rows = len(df)
    if n_rows == 0:
        return issues, pd.Series(dtype=bool), 1.0
        
    today_str = datetime.today().strftime("%Y-%m-%d")
    row_failed_health = pd.Series(False, index=df.index)
    
    # Track duplicates within this batch
    duplicates_mask = df.duplicated(keep='first')
    
    for idx, row in df.iterrows():
        date_str = row["date"]
        rev = row["revenue"]
        exp = row["expenses"]
        
        row_num = idx + 1
        row_has_crit_or_warn = False
        
        # 1. Missing Date (Critical)
        if date_str is None:
            issues["missing_date"].append({"row": row_num, "msg": "Transaction date is missing or unparseable."})
            row_has_crit_or_warn = True
            
        # 2. Future Dates (Warning)
        if date_str and date_str > today_str:
            issues["future_dates"].append({"row": row_num, "date": date_str, "msg": f"Date '{date_str}' is in the future."})
            row_has_crit_or_warn = True
            
        # 3. Negative Revenue (Warning)
        if pd.notna(rev) and rev < 0:
            issues["negative_revenue"].append({"row": row_num, "amount": rev, "msg": f"Revenue amount ${rev:.2f} is negative."})
            row_has_crit_or_warn = True
            
        # 4. Negative Expenses (Warning)
        if pd.notna(exp) and exp < 0:
            issues["negative_expenses"].append({"row": row_num, "amount": exp, "msg": f"Expenses amount ${exp:.2f} is negative."})
            row_has_crit_or_warn = True
            
        # 5. Null Amounts Check (Warning if both are 0 or NaN)
        if (pd.isna(rev) or rev == 0) and (pd.isna(exp) or exp == 0):
            issues["null_amounts"].append({"row": row_num, "msg": "Transaction has zero revenue and zero expenses."})
            row_has_crit_or_warn = True
            
        # 6. Blank fields in dimensions (Warning)
        blank_cols = []
        for col in ["department", "region", "product_line", "expense_category"]:
            if pd.isna(row[col]) or str(row[col]).strip() == "":
                blank_cols.append(col)
        if blank_cols:
            issues["blank_fields"].append({"row": row_num, "columns": blank_cols, "msg": f"Blank categorizations in: {', '.join(blank_cols)}"})
            row_has_crit_or_warn = True
            
        # 7. Overlapping Rows (Info)
        if duplicates_mask.iloc[idx]:
            issues["overlapping_rows"].append({"row": row_num, "msg": "Exact duplicate transaction row."})
            
        # 8. Extreme Revenue Outliers (> $50,000) (Info)
        if pd.notna(rev) and rev > 50000:
            issues["extreme_revenue"].append({"row": row_num, "amount": rev, "msg": f"Large revenue spike: ${rev:.2f}."})
            
        # 9. Extreme Expense Outliers (> $25,000) (Info)
        if pd.notna(exp) and exp > 25000:
            issues["extreme_expenses"].append({"row": row_num, "amount": exp, "msg": f"Large expense spike: ${exp:.2f}."})
            
        if row_has_crit_or_warn:
            row_failed_health.iloc[idx] = True
            
    failed_count = row_failed_health.sum()
    health_score = 1.0 - (failed_count / n_rows) if n_rows > 0 else 1.0
    
    return issues, row_failed_health, health_score

def merge_batch_to_excel(df, excel_path="financial_analysis.xlsx"):
    try:
        # Load all existing sheets from Excel
        df_trans_old = pd.read_excel(excel_path, sheet_name="fact_transactions")
        df_depts_old = pd.read_excel(excel_path, sheet_name="dim_departments")
        df_regions_old = pd.read_excel(excel_path, sheet_name="dim_regions")
        df_products_old = pd.read_excel(excel_path, sheet_name="dim_products")
        df_cats_old = pd.read_excel(excel_path, sheet_name="dim_expense_categories")
        df_budgets_old = pd.read_excel(excel_path, sheet_name="fact_budgets")
        df_invoices_old = pd.read_excel(excel_path, sheet_name="fact_invoices")
        df_log_old = pd.read_excel(excel_path, sheet_name="ingestion_log")
        
        # 1. Update dimensions if new values are found in the batch
        depts_new = list(df["department"].dropna().unique())
        dept_names_existing = list(df_depts_old["dept_name"].unique())
        for d in depts_new:
            if d not in dept_names_existing:
                df_depts_old = pd.concat([df_depts_old, pd.DataFrame([{
                    "dept_id": len(df_depts_old) + 1,
                    "dept_name": d
                }])], ignore_index=True)
        dept_lookup = {name: i for i, name in zip(df_depts_old["dept_id"], df_depts_old["dept_name"])}
        
        regions_new = list(df["region"].dropna().unique())
        region_names_existing = list(df_regions_old["region_name"].unique())
        for r in regions_new:
            if r not in region_names_existing:
                df_regions_old = pd.concat([df_regions_old, pd.DataFrame([{
                    "region_id": len(df_regions_old) + 1,
                    "region_name": r
                }])], ignore_index=True)
        region_lookup = {name: i for i, name in zip(df_regions_old["region_id"], df_regions_old["region_name"])}
        
        products_new = list(df["product_line"].dropna().unique())
        product_names_existing = list(df_products_old["product_name"].unique())
        for p in products_new:
            if p not in product_names_existing:
                df_products_old = pd.concat([df_products_old, pd.DataFrame([{
                    "product_id": len(df_products_old) + 1,
                    "product_name": p
                }])], ignore_index=True)
        product_lookup = {name: i for i, name in zip(df_products_old["product_id"], df_products_old["product_name"])}
        
        cats_new = list(df["expense_category"].dropna().unique())
        cat_names_existing = list(df_cats_old["category_name"].unique())
        for c in cats_new:
            if c not in cat_names_existing:
                df_cats_old = pd.concat([df_cats_old, pd.DataFrame([{
                    "category_id": len(df_cats_old) + 1,
                    "category_name": c
                }])], ignore_index=True)
        category_lookup = {name: i for i, name in zip(df_cats_old["category_id"], df_cats_old["category_name"])}
        
        # 2. Map new transaction records to dimension IDs
        transactions_data = []
        start_id = int(df_trans_old["transaction_id"].max()) + 1 if not df_trans_old.empty else 1
        for idx, row in df.iterrows():
            dept_id = dept_lookup.get(row["department"])
            region_id = region_lookup.get(row["region"])
            product_id = product_lookup.get(row["product_line"])
            category_id = category_lookup.get(row["expense_category"])
            client_val = row["client_id"] if "client_id" in df.columns and pd.notna(row["client_id"]) else ""
            
            transactions_data.append({
                "transaction_id": start_id + idx,
                "date": row["date"],
                "revenue": row["revenue"],
                "expenses": row["expenses"],
                "dept_id": dept_id,
                "region_id": region_id,
                "product_id": product_id,
                "category_id": category_id,
                "client_id": client_val
            })
        df_trans_new = pd.DataFrame(transactions_data)
        df_trans_all = pd.concat([df_trans_old, df_trans_new], ignore_index=True)
        
        # 3. Save all sheets back to Excel
        with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
            df_trans_all.to_excel(writer, sheet_name="fact_transactions", index=False)
            df_budgets_old.to_excel(writer, sheet_name="fact_budgets", index=False)
            df_invoices_old.to_excel(writer, sheet_name="fact_invoices", index=False)
            df_depts_old.to_excel(writer, sheet_name="dim_departments", index=False)
            df_regions_old.to_excel(writer, sheet_name="dim_regions", index=False)
            df_products_old.to_excel(writer, sheet_name="dim_products", index=False)
            df_cats_old.to_excel(writer, sheet_name="dim_expense_categories", index=False)
            df_log_old.to_excel(writer, sheet_name="ingestion_log", index=False)
            
        # Re-trigger styled report sheets recompilation
        excel_generator.create_styled_excel(excel_path)
        
        return True, len(df)
        
    except Exception as e:
        print(f"Error merging to Excel: {e}")
        return False, str(e)

def process_file_upload(filepath, filename, excel_path="financial_analysis.xlsx"):
    submitted_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    if not os.path.exists(excel_path):
        base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        excel_path = os.path.join(base_dir, excel_path)
        
    try:
        if filename.endswith(".csv"):
            df_raw = pd.read_csv(filepath)
        elif filename.endswith((".xls", ".xlsx")):
            df_raw = pd.read_excel(filepath)
        else:
            return {"success": False, "msg": "Unsupported file format.", "status": "Rejected"}
    except Exception as e:
        return {"success": False, "msg": f"Failed to read file: {e}", "status": "Rejected"}
        
    n_rows = len(df_raw)
    if n_rows == 0:
        return {"success": False, "msg": "The uploaded file is empty.", "status": "Rejected"}
        
    df_mapped = map_columns(df_raw)
    
    # Clean mappings
    df_clean = df_mapped.copy()
    df_clean["revenue"] = df_mapped["revenue"].apply(clean_numeric)
    df_clean["expenses"] = df_mapped["expenses"].apply(clean_numeric)
    df_clean["date"] = df_mapped["date"].apply(parse_date)
    
    for str_col in ["department", "region", "product_line", "expense_category", "client_id"]:
        df_clean[str_col] = df_clean[str_col].apply(lambda x: str(x).strip() if pd.notna(x) else np.nan)
        
    # Run audits
    issues, row_issues_mask, health_score = run_data_quality_checks(df_clean, excel_path)
    
    if health_score >= 0.8:
        status = "Accepted"
    elif health_score >= 0.5:
        status = "Needs Review"
    else:
        status = "Rejected"
        
    accepted_rows = 0
    rejected_rows = 0
    merge_success = False
    
    issues_summary = {k: v for k, v in issues.items() if len(v) > 0}
    issues_json_str = json.dumps(issues_summary)
    
    if status in ["Accepted", "Needs Review"]:
        merge_success, merge_info = merge_batch_to_excel(df_clean, excel_path)
        if merge_success:
            accepted_rows = n_rows
        else:
            status = "Rejected"
            issues_json_str = json.dumps({"db_error": [f"Database write failure: {merge_info}"]})
            rejected_rows = n_rows
    else:
        rejected_rows = n_rows
        
    # Log to ingestion_log worksheet
    try:
        df_log_old = pd.read_excel(excel_path, sheet_name="ingestion_log")
        start_log_id = int(df_log_old["batch_id"].max()) + 1 if not df_log_old.empty else 1
        
        new_log = pd.DataFrame([{
            "batch_id": start_log_id,
            "submitted_at": submitted_at,
            "filename": filename,
            "row_count": n_rows,
            "status": status,
            "health_score": round(health_score * 100, 2),
            "issues_json": issues_json_str,
            "accepted_rows": accepted_rows,
            "rejected_rows": rejected_rows
        }])
        
        df_log_all = pd.concat([df_log_old, new_log], ignore_index=True)
        
        # Load all updated sheets to save back
        df_trans = pd.read_excel(excel_path, sheet_name="fact_transactions")
        df_budgets = pd.read_excel(excel_path, sheet_name="fact_budgets")
        df_invoices = pd.read_excel(excel_path, sheet_name="fact_invoices")
        df_depts = pd.read_excel(excel_path, sheet_name="dim_departments")
        df_regions = pd.read_excel(excel_path, sheet_name="dim_regions")
        df_products = pd.read_excel(excel_path, sheet_name="dim_products")
        df_cats = pd.read_excel(excel_path, sheet_name="dim_expense_categories")
        
        with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
            df_trans.to_excel(writer, sheet_name="fact_transactions", index=False)
            df_budgets.to_excel(writer, sheet_name="fact_budgets", index=False)
            df_invoices.to_excel(writer, sheet_name="fact_invoices", index=False)
            df_depts.to_excel(writer, sheet_name="dim_departments", index=False)
            df_regions.to_excel(writer, sheet_name="dim_regions", index=False)
            df_products.to_excel(writer, sheet_name="dim_products", index=False)
            df_cats.to_excel(writer, sheet_name="dim_expense_categories", index=False)
            df_log_all.to_excel(writer, sheet_name="ingestion_log", index=False)
            
    except Exception as e:
        print(f"Error logging batch: {e}")
        
    return {
        "success": merge_success or status == "Needs Review",
        "status": status,
        "health_score": health_score,
        "row_count": n_rows,
        "accepted_rows": accepted_rows,
        "rejected_rows": rejected_rows,
        "issues": issues_summary
    }

if __name__ == "__main__":
    print("Testing Ingestion mapper...")
    df_test = pd.DataFrame(columns=["Date", "Rev", "Outflow", "division", "area", "line", "type"])
    print(list(map_columns(df_test).columns))
