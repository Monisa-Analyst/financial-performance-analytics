import csv
import random
from datetime import datetime, timedelta

def generate_financial_data(trans_file, budget_file, invoice_file, seed=100):
    random.seed(seed)
    
    departments = ["Sales", "Marketing", "Operations", "Engineering", "HR"]
    regions = ["North", "South", "East", "West"]
    products = ["Software Licenses", "Consulting Services", "Cloud Subscriptions", "Hardware Systems"]
    
    expense_categories = {
        "Sales": ["Travel", "Salaries", "Commissions", "Software Subscriptions"],
        "Marketing": ["Salaries", "Marketing Ads", "Software Subscriptions", "Events"],
        "Operations": ["Salaries", "Rent", "Utilities", "Office Supplies"],
        "Engineering": ["Salaries", "Software Subscriptions", "Hardware Servers", "Training"],
        "HR": ["Salaries", "Recruitment", "Software Subscriptions", "Benefits"]
    }
    
    # Establish regional and product characteristics
    region_revenue_weights = {"North": 1.0, "South": 1.2, "East": 0.9, "West": 1.1}
    region_expense_weights = {"North": 1.0, "South": 0.8, "East": 0.95, "West": 1.05}
    
    product_revenue_weights = {
        "Software Licenses": 1.4,
        "Consulting Services": 1.0,
        "Cloud Subscriptions": 1.2,
        "Hardware Systems": 1.1
    }
    product_expense_weights = {
        "Software Licenses": 0.3, 
        "Consulting Services": 0.7,
        "Cloud Subscriptions": 0.6,
        "Hardware Systems": 0.95 
    }
    
    # Generate 100 client IDs and assign them a cohort acquisition month (1 to 12)
    # Jan gets more signups, decaying through the year.
    clients = [f"CL-{i:03d}" for i in range(1, 101)]
    client_cohorts = {}
    for i, client in enumerate(clients):
        # Weighted distribution of signup months
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
        
    start_date = datetime(2025, 1, 1)
    end_date = datetime(2025, 12, 31)
    
    revenue_records = []
    
    # Generate daily transaction records (~2500 records)
    with open(trans_file, mode='w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        writer.writerow([
            "Date", "Revenue", "Expenses", "Department", 
            "Region", "Product Line", "Expense Category", "Client ID"
        ])
        
        current_date = start_date
        while current_date <= end_date:
            num_trans = random.randint(5, 10)
            month = current_date.month
            growth_factor = 1.0 + (month * 0.015) 
            
            for _ in range(num_trans):
                dept = random.choice(departments)
                region = random.choice(regions)
                prod = random.choice(products)
                cat = random.choice(expense_categories[dept])
                
                trans_type = random.random()
                rev_val = 0.0
                exp_val = 0.0
                client_id = ""
                
                base_rev = random.uniform(1000, 8000)
                base_exp = random.uniform(500, 5000)
                
                if dept == "Marketing" and month >= 9:
                    base_exp *= 1.5
                
                # Check transaction type:
                # 0.4 revenue, 0.4 expense, 0.2 both
                is_revenue = trans_type < 0.4 or trans_type >= 0.8
                is_expense = trans_type >= 0.4
                
                if is_revenue:
                    # Choose a client whose cohort month has started
                    available_clients = [c for c, m in client_cohorts.items() if m <= month]
                    if available_clients:
                        # Humanized Retention Churn Simulation:
                        # Clients are highly likely to transact in their cohort month (Month 0).
                        # In subsequent months, we check if they retain. If they "churn" for this month, 
                        # we either select another client or skip.
                        selected_client = None
                        random.shuffle(available_clients)
                        for c in available_clients:
                            cohort_month = client_cohorts[c]
                            elapsed_months = month - cohort_month
                            
                            # Churn decay formula: Month 0 = 100%, Month 1 = 80%, Month 2 = 70%, Month 3+ = 50%
                            retention_prob = 1.0
                            if elapsed_months == 1: retention_prob = 0.8
                            elif elapsed_months == 2: retention_prob = 0.7
                            elif elapsed_months >= 3: retention_prob = 0.50
                            
                            if random.random() <= retention_prob:
                                selected_client = c
                                break
                        
                        client_id = selected_client if selected_client else random.choice(available_clients)
                    else:
                        client_id = "CL-001" # fallback
                    
                    rev_val = base_rev * region_revenue_weights[region] * product_revenue_weights[prod] * growth_factor
                    rev_val = round(rev_val, 2)
                    
                if is_expense:
                    exp_val = base_exp * region_expense_weights[region] * product_expense_weights[prod]
                    exp_val = round(exp_val, 2)
                
                if is_revenue and is_expense:
                    # Assign same values generated above
                    pass
                    
                date_str = current_date.strftime("%Y-%m-%d")
                writer.writerow([date_str, rev_val, exp_val, dept, region, prod, cat, client_id])
                
                if rev_val > 0:
                    revenue_records.append({
                        "date": date_str,
                        "amount": rev_val,
                        "client_id": client_id
                    })
                
            current_date += timedelta(days=1)
            
    # Generate monthly budgets
    with open(budget_file, mode='w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        writer.writerow([
            "Month", "Department", "Region", "Budgeted Revenue", "Budgeted Expenses"
        ])
        
        for m in range(1, 13):
            month_str = f"2025-{m:02d}"
            growth_factor = 1.0 + (m * 0.012)
            
            for dept in departments:
                for region in regions:
                    budget_rev = random.uniform(15000, 30000) * region_revenue_weights[region] * growth_factor
                    budget_exp = random.uniform(10000, 20000) * region_expense_weights[region]
                    
                    if dept == "Marketing" and m >= 9:
                        budget_exp *= 1.4
                        
                    writer.writerow([
                        month_str, dept, region, 
                        round(budget_rev, 2), round(budget_exp, 2)
                    ])
                    
    # Generate Invoices CSV (Accounts Receivable Funnel)
    # We map invoices directly from the revenue transactions
    with open(invoice_file, mode='w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        writer.writerow(["Invoice ID", "Client ID", "Issue Date", "Amount", "Status"])
        
        for idx, rec in enumerate(revenue_records):
            inv_id = f"INV-{10000 + idx}"
            dt = datetime.strptime(rec["date"], "%Y-%m-%d")
            
            # Funnel stages: 
            # 1-Created -> 2-Delivered -> 3-Approved -> 4-Pending Payment -> 5-Settled
            # Older invoices (before November) are highly likely to be Settled.
            # Recent invoices (November and December) are still in flight.
            r = random.random()
            if dt.month < 11:
                # 92% settled, 5% pending payment, 2% approved, 1% delivered, 0% created
                if r < 0.92: status = "5-Settled"
                elif r < 0.97: status = "4-Pending Payment"
                elif r < 0.99: status = "3-Approved"
                else: status = "2-Delivered"
            else:
                # In flight: 40% settled, 30% pending payment, 15% approved, 10% delivered, 5% created
                if r < 0.40: status = "5-Settled"
                elif r < 0.70: status = "4-Pending Payment"
                elif r < 0.85: status = "3-Approved"
                elif r < 0.95: status = "2-Delivered"
                else: status = "1-Created"
                
            writer.writerow([inv_id, rec["client_id"], rec["date"], rec["amount"], status])

if __name__ == "__main__":
    print("Generating raw financials daily transactions, budgets, and invoices...")
    generate_financial_data("financial_data_raw.csv", "financial_budget_raw.csv", "financial_invoices_raw.csv")
    print("Financial raw datasets generated successfully!")

