import os
import sqlite3
import pandas as pd
from datetime import datetime

def init_database(db_path, trans_csv, budget_csv):
    print(f"Initializing database at: {db_path}")
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    
    cursor.execute("PRAGMA foreign_keys = ON;")
    
    # Drop existing
    cursor.execute("DROP TABLE IF EXISTS fact_transactions;")
    cursor.execute("DROP TABLE IF EXISTS fact_budgets;")
    cursor.execute("DROP TABLE IF EXISTS dim_departments;")
    cursor.execute("DROP TABLE IF EXISTS dim_regions;")
    cursor.execute("DROP TABLE IF EXISTS dim_products;")
    cursor.execute("DROP TABLE IF EXISTS dim_expense_categories;")
    cursor.execute("DROP TABLE IF EXISTS ingestion_log;")
    
    # Create tables
    cursor.execute("""
    CREATE TABLE dim_departments (
        dept_id INTEGER PRIMARY KEY AUTOINCREMENT,
        dept_name TEXT UNIQUE NOT NULL
    );
    """)
    
    cursor.execute("""
    CREATE TABLE dim_regions (
        region_id INTEGER PRIMARY KEY AUTOINCREMENT,
        region_name TEXT UNIQUE NOT NULL
    );
    """)
    
    cursor.execute("""
    CREATE TABLE dim_products (
        product_id INTEGER PRIMARY KEY AUTOINCREMENT,
        product_name TEXT UNIQUE NOT NULL
    );
    """)
    
    cursor.execute("""
    CREATE TABLE dim_expense_categories (
        category_id INTEGER PRIMARY KEY AUTOINCREMENT,
        category_name TEXT UNIQUE NOT NULL
    );
    """)
    
    cursor.execute("""
    CREATE TABLE fact_transactions (
        transaction_id INTEGER PRIMARY KEY AUTOINCREMENT,
        date TEXT NOT NULL,
        revenue REAL DEFAULT 0,
        expenses REAL DEFAULT 0,
        dept_id INTEGER,
        region_id INTEGER,
        product_id INTEGER,
        category_id INTEGER,
        FOREIGN KEY(dept_id) REFERENCES dim_departments(dept_id),
        FOREIGN KEY(region_id) REFERENCES dim_regions(region_id),
        FOREIGN KEY(product_id) REFERENCES dim_products(product_id),
        FOREIGN KEY(category_id) REFERENCES dim_expense_categories(category_id)
    );
    """)
    
    cursor.execute("""
    CREATE TABLE fact_budgets (
        budget_id INTEGER PRIMARY KEY AUTOINCREMENT,
        month TEXT NOT NULL,
        dept_id INTEGER,
        region_id INTEGER,
        budgeted_revenue REAL DEFAULT 0,
        budgeted_expenses REAL DEFAULT 0,
        FOREIGN KEY(dept_id) REFERENCES dim_departments(dept_id),
        FOREIGN KEY(region_id) REFERENCES dim_regions(region_id)
    );
    """)
    
    cursor.execute("""
    CREATE TABLE ingestion_log (
        batch_id INTEGER PRIMARY KEY AUTOINCREMENT,
        submitted_at TEXT NOT NULL,
        filename TEXT NOT NULL,
        row_count INTEGER NOT NULL,
        status TEXT NOT NULL,
        health_score REAL NOT NULL,
        issues_json TEXT NOT NULL,
        accepted_rows INTEGER,
        rejected_rows INTEGER
    );
    """)
    
    conn.commit()
    print("Database tables created.")
    
    # Load and seed dimensions and transactions
    df_trans = pd.read_csv(trans_csv)
    df_budget = pd.read_csv(budget_csv)
    
    # Insert Dimensions
    depts = set(df_trans["Department"].dropna().unique()).union(set(df_budget["Department"].dropna().unique()))
    for dept in depts:
        cursor.execute("INSERT OR IGNORE INTO dim_departments (dept_name) VALUES (?)", (dept,))
        
    regions = set(df_trans["Region"].dropna().unique()).union(set(df_budget["Region"].dropna().unique()))
    for reg in regions:
        cursor.execute("INSERT OR IGNORE INTO dim_regions (region_name) VALUES (?)", (reg,))
        
    products = df_trans["Product Line"].dropna().unique()
    for prod in products:
        cursor.execute("INSERT OR IGNORE INTO dim_products (product_name) VALUES (?)", (prod,))
        
    categories = df_trans["Expense Category"].dropna().unique()
    for cat in categories:
        cursor.execute("INSERT OR IGNORE INTO dim_expense_categories (category_name) VALUES (?)", (cat,))
        
    conn.commit()
    
    # Fetch lookups
    cursor.execute("SELECT dept_id, dept_name FROM dim_departments")
    dept_lookup = {name: id for id, name in cursor.fetchall()}
    
    cursor.execute("SELECT region_id, region_name FROM dim_regions")
    region_lookup = {name: id for id, name in cursor.fetchall()}
    
    cursor.execute("SELECT product_id, product_name FROM dim_products")
    product_lookup = {name: id for id, name in cursor.fetchall()}
    
    cursor.execute("SELECT category_id, category_name FROM dim_expense_categories")
    category_lookup = {name: id for id, name in cursor.fetchall()}
    
    # Insert budgets
    budget_records = []
    for _, row in df_budget.iterrows():
        dept_id = dept_lookup.get(row["Department"])
        region_id = region_lookup.get(row["Region"])
        budget_records.append((
            row["Month"],
            dept_id,
            region_id,
            float(row["Budgeted Revenue"]),
            float(row["Budgeted Expenses"])
        ))
        
    cursor.executemany("""
    INSERT INTO fact_budgets (month, dept_id, region_id, budgeted_revenue, budgeted_expenses)
    VALUES (?, ?, ?, ?, ?)
    """, budget_records)
    
    # Insert transactions
    trans_records = []
    for _, row in df_trans.iterrows():
        dept_id = dept_lookup.get(row["Department"])
        region_id = region_lookup.get(row["Region"])
        product_id = product_lookup.get(row["Product Line"])
        category_id = category_lookup.get(row["Expense Category"])
        
        trans_records.append((
            row["Date"],
            float(row["Revenue"]) if pd.notna(row["Revenue"]) else 0.0,
            float(row["Expenses"]) if pd.notna(row["Expenses"]) else 0.0,
            dept_id,
            region_id,
            product_id,
            category_id
        ))
        
    cursor.executemany("""
    INSERT INTO fact_transactions (date, revenue, expenses, dept_id, region_id, product_id, category_id)
    VALUES (?, ?, ?, ?, ?, ?, ?)
    """, trans_records)
    
    conn.commit()
    
    # Log initial seed batch
    cursor.execute("""
    INSERT INTO ingestion_log (submitted_at, filename, row_count, status, health_score, issues_json, accepted_rows, rejected_rows)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        os.path.basename(trans_csv),
        len(df_trans),
        "Accepted",
        1.0,
        "[]",
        len(df_trans),
        0
    ))
    
    conn.commit()
    print(f"Seeded {len(df_trans)} transactions and {len(df_budget)} monthly budget rows.")
    conn.close()

if __name__ == "__main__":
    db_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "financial.db")
    trans_csv = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "financial_data_raw.csv")
    budget_csv = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "financial_budget_raw.csv")
    
    os.makedirs(os.path.dirname(db_file), exist_ok=True)
    init_database(db_file, trans_csv, budget_csv)
