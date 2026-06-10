import csv
import random
from datetime import datetime, timedelta

def generate_financial_data(trans_file, budget_file, seed=100):
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
    # South region has high margin (low expenses relative to revenue)
    # Software has high revenue & margin; Hardware has high revenue but very high expenses (low margin)
    region_revenue_weights = {"North": 1.0, "South": 1.2, "East": 0.9, "West": 1.1}
    region_expense_weights = {"North": 1.0, "South": 0.8, "East": 0.95, "West": 1.05}
    
    product_revenue_weights = {
        "Software Licenses": 1.4,
        "Consulting Services": 1.0,
        "Cloud Subscriptions": 1.2,
        "Hardware Systems": 1.1
    }
    product_expense_weights = {
        "Software Licenses": 0.3, # extremely high margin
        "Consulting Services": 0.7,
        "Cloud Subscriptions": 0.6,
        "Hardware Systems": 0.95 # low margin
    }
    
    start_date = datetime(2025, 1, 1)
    end_date = datetime(2025, 12, 31)
    days_count = (end_date - start_date).days + 1
    
    # Generate daily transaction records (~2500 records)
    with open(trans_file, mode='w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        writer.writerow([
            "Date", "Revenue", "Expenses", "Department", 
            "Region", "Product Line", "Expense Category"
        ])
        
        current_date = start_date
        while current_date <= end_date:
            # Number of transactions per day (between 5 and 10)
            num_trans = random.randint(5, 10)
            month = current_date.month
            
            # Apply YoY/MoM growth factor to revenue (increasing over months)
            growth_factor = 1.0 + (month * 0.015) # ~18% growth by end of year
            
            for _ in range(num_trans):
                dept = random.choice(departments)
                region = random.choice(regions)
                prod = random.choice(products)
                cat = random.choice(expense_categories[dept])
                
                # Revenue generation
                # 35% of transactions are pure revenue, 35% pure expenses, 30% have both
                trans_type = random.random()
                
                rev_val = 0.0
                exp_val = 0.0
                
                base_rev = random.uniform(1000, 8000)
                base_exp = random.uniform(500, 5000)
                
                # Marketing expense season skew (grows faster in Q4)
                if dept == "Marketing" and month >= 9:
                    base_exp *= 1.5
                
                if trans_type < 0.4: # Revenue transaction
                    rev_val = base_rev * region_revenue_weights[region] * product_revenue_weights[prod] * growth_factor
                    rev_val = round(rev_val, 2)
                elif trans_type < 0.8: # Expense transaction
                    exp_val = base_exp * region_expense_weights[region] * product_expense_weights[prod]
                    exp_val = round(exp_val, 2)
                else: # Both
                    rev_val = base_rev * region_revenue_weights[region] * product_revenue_weights[prod] * growth_factor
                    exp_val = base_exp * region_expense_weights[region] * product_expense_weights[prod]
                    rev_val = round(rev_val, 2)
                    exp_val = round(exp_val, 2)
                    
                date_str = current_date.strftime("%Y-%m-%d")
                writer.writerow([date_str, rev_val, exp_val, dept, region, prod, cat])
                
            current_date += timedelta(days=1)
            
    # Generate monthly budgets
    # Budgets are mapped per department per region per month (12 months * 5 depts * 4 regions = 240 records)
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
                    # Budget baseline
                    budget_rev = random.uniform(15000, 30000) * region_revenue_weights[region] * growth_factor
                    budget_exp = random.uniform(10000, 20000) * region_expense_weights[region]
                    
                    # Marketing seasonality budget
                    if dept == "Marketing" and m >= 9:
                        budget_exp *= 1.4
                        
                    writer.writerow([
                        month_str, dept, region, 
                        round(budget_rev, 2), round(budget_exp, 2)
                    ])

if __name__ == "__main__":
    print("Generating raw financials daily transactions...")
    generate_financial_data("financial_data_raw.csv", "financial_budget_raw.csv")
    print("Financial raw datasets generated successfully!")
