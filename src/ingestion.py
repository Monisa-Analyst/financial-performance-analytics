import os
import json
import sqlite3
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

def run_data_quality_checks(df, db_path="src/financial.db"):
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
        dept = row["department"]
        reg = row["region"]
        prod = row["product_line"]
        cat = row["expense_category"]
        
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

def merge_batch_to_db(df, db_path="src/financial.db"):
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    cursor.execute("PRAGMA foreign_keys = ON;")
    
    try:
        cursor.execute("BEGIN TRANSACTION;")
        
        # Insert new dimensions if they don't exist
        for dept in df["department"].dropna().unique():
            cursor.execute("INSERT OR IGNORE INTO dim_departments (dept_name) VALUES (?)", (dept,))
            
        for reg in df["region"].dropna().unique():
            cursor.execute("INSERT OR IGNORE INTO dim_regions (region_name) VALUES (?)", (reg,))
            
        for prod in df["product_line"].dropna().unique():
            cursor.execute("INSERT OR IGNORE INTO dim_products (product_name) VALUES (?)", (prod,))
            
        for cat in df["expense_category"].dropna().unique():
            cursor.execute("INSERT OR IGNORE INTO dim_expense_categories (category_name) VALUES (?)", (cat,))
            
        conn.commit()
        
        # Get lookups
        cursor.execute("SELECT dept_id, dept_name FROM dim_departments")
        dept_lookup = {name: id for id, name in cursor.fetchall()}
        
        cursor.execute("SELECT region_id, region_name FROM dim_regions")
        region_lookup = {name: id for id, name in cursor.fetchall()}
        
        cursor.execute("SELECT product_id, product_name FROM dim_products")
        product_lookup = {name: id for id, name in cursor.fetchall()}
        
        cursor.execute("SELECT category_id, category_name FROM dim_expense_categories")
        category_lookup = {name: id for id, name in cursor.fetchall()}
        
        cursor.execute("BEGIN TRANSACTION;")
        
        transactions_data = []
        for _, row in df.iterrows():
            dept_id = dept_lookup.get(row["department"])
            region_id = region_lookup.get(row["region"])
            product_id = product_lookup.get(row["product_line"])
            category_id = category_lookup.get(row["expense_category"])
            client_val = row["client_id"] if "client_id" in df.columns and pd.notna(row["client_id"]) else ""
            
            transactions_data.append((
                row["date"],
                row["revenue"],
                row["expenses"],
                dept_id,
                region_id,
                product_id,
                category_id,
                client_val
            ))
            
        cursor.executemany("""
        INSERT INTO fact_transactions (date, revenue, expenses, dept_id, region_id, product_id, category_id, client_id)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """, transactions_data)
        
        cursor.execute("COMMIT;")
        conn.close()
        
        # Trigger excel recompilation to include new transactions!
        excel_generator.create_styled_excel(db_path, "financial_analysis.xlsx")
        
        return True, len(df)
        
    except Exception as e:
        cursor.execute("ROLLBACK;")
        conn.close()
        print(f"Error merging: {e}")
        return False, str(e)

def process_file_upload(filepath, filename, db_path="src/financial.db"):
    submitted_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
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
    
    # Clean
    df_clean = df_mapped.copy()
    df_clean["revenue"] = df_mapped["revenue"].apply(clean_numeric)
    df_clean["expenses"] = df_mapped["expenses"].apply(clean_numeric)
    df_clean["date"] = df_mapped["date"].apply(parse_date)
    
    for str_col in ["department", "region", "product_line", "expense_category", "client_id"]:
        df_clean[str_col] = df_clean[str_col].apply(lambda x: str(x).strip() if pd.notna(x) else np.nan)
        
    # Run audits
    issues, row_issues_mask, health_score = run_data_quality_checks(df_clean, db_path)
    
    if health_score >= 0.8:
        status = "Accepted"
    elif health_score >= 0.5:
        status = "Needs Review"
    else:
        status = "Rejected"
        
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    
    accepted_rows = 0
    rejected_rows = 0
    merge_success = False
    
    issues_summary = {k: v for k, v in issues.items() if len(v) > 0}
    issues_json_str = json.dumps(issues_summary)
    
    if status in ["Accepted", "Needs Review"]:
        merge_success, merge_info = merge_batch_to_db(df_clean, db_path)
        if merge_success:
            accepted_rows = n_rows
        else:
            status = "Rejected"
            issues_json_str = json.dumps({"db_error": [f"Database write failure: {merge_info}"]})
            rejected_rows = n_rows
    else:
        rejected_rows = n_rows
        
    cursor.execute("""
    INSERT INTO ingestion_log (submitted_at, filename, row_count, status, health_score, issues_json, accepted_rows, rejected_rows)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    """, (submitted_at, filename, n_rows, status, health_score, issues_json_str, accepted_rows, rejected_rows))
    
    conn.commit()
    conn.close()
    
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
    # Test column mapping on clean data
    print("Testing Ingestion mapper...")
    df_test = pd.DataFrame(columns=["Date", "Rev", "Outflow", "division", "area", "line", "type"])
    print(list(map_columns(df_test).columns))
