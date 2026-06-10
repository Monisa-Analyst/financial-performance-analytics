import os
import sqlite3
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def create_styled_excel(db_path="src/financial.db", output_path="financial_analysis.xlsx"):
    print(f"Creating Excel workbook at: {output_path}")
    conn = sqlite3.connect(db_path)
    
    # Fetch aggregates for Excel Sheets
    # 1. Monthly Financials
    df_monthly = pd.read_sql_query("""
        SELECT 
            strftime('%Y-%m', date) as Month,
            ROUND(SUM(revenue), 2) as Revenue,
            ROUND(SUM(expenses), 2) as Expenses
        FROM fact_transactions
        GROUP BY Month
        ORDER BY Month
    """, conn)
    
    # 2. Expenses by Category
    df_expenses = pd.read_sql_query("""
        SELECT 
            c.category_name as Category,
            ROUND(SUM(f.expenses), 2) as Expenses
        FROM fact_transactions f
        JOIN dim_expense_categories c ON f.category_id = c.category_id
        GROUP BY Category
        ORDER BY Expenses DESC
    """, conn)
    
    # 3. Budget vs Actual (by Department for 2025)
    df_budget_vs_actual = pd.read_sql_query("""
        WITH Actuals AS (
            SELECT 
                d.dept_name as Department,
                SUM(f.revenue) as Actual_Revenue,
                SUM(f.expenses) as Actual_Expenses
            FROM fact_transactions f
            JOIN dim_departments d ON f.dept_id = d.dept_id
            GROUP BY Department
        ),
        Budgets AS (
            SELECT 
                d.dept_name as Department,
                SUM(b.budgeted_revenue) as Budget_Revenue,
                SUM(b.budgeted_expenses) as Budget_Expenses
            FROM fact_budgets b
            JOIN dim_departments d ON b.dept_id = d.dept_id
            GROUP BY Department
        )
        SELECT 
            a.Department,
            ROUND(a.Actual_Revenue, 2) as Actual_Revenue,
            ROUND(b.Budget_Revenue, 2) as Budget_Revenue,
            ROUND(a.Actual_Expenses, 2) as Actual_Expenses,
            ROUND(b.Budget_Expenses, 2) as Budget_Expenses
        FROM Actuals a
        JOIN Budgets b ON a.Department = b.Department
    """, conn)
    
    # 4. Regional Profitability
    df_regions = pd.read_sql_query("""
        SELECT 
            r.region_name as Region,
            ROUND(SUM(f.revenue), 2) as Revenue,
            ROUND(SUM(f.expenses), 2) as Expenses
        FROM fact_transactions f
        JOIN dim_regions r ON f.region_id = r.region_id
        GROUP BY Region
        ORDER BY Revenue DESC
    """, conn)
    
    # 5. Product Line Performance
    df_products = pd.read_sql_query("""
        SELECT 
            p.product_name as Product_Line,
            ROUND(SUM(f.revenue), 2) as Revenue,
            ROUND(SUM(f.expenses), 2) as Expenses
        FROM fact_transactions f
        JOIN dim_products p ON f.product_id = p.product_id
        GROUP BY Product_Line
        ORDER BY Revenue DESC
    """, conn)
    
    conn.close()
    
    wb = Workbook()
    
    # Styling configurations
    font_family = "Segoe UI"
    title_font = Font(name=font_family, size=16, bold=True, color="1F4E78")
    section_font = Font(name=font_family, size=12, bold=True, color="2C3E50")
    header_font = Font(name=font_family, size=11, bold=True, color="FFFFFF")
    bold_font = Font(name=font_family, size=10, bold=True)
    normal_font = Font(name=font_family, size=10)
    
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    summary_fill = PatternFill(start_color="EAECEE", end_color="EAECEE", fill_type="solid")
    zebra_fill = PatternFill(start_color="F2F4F4", end_color="F2F4F4", fill_type="solid")
    
    thin_border = Border(
        left=Side(style='thin', color='BDC3C7'),
        right=Side(style='thin', color='BDC3C7'),
        top=Side(style='thin', color='BDC3C7'),
        bottom=Side(style='thin', color='BDC3C7')
    )
    
    double_bottom_border = Border(
        top=Side(style='thin', color='7F8C8D'),
        bottom=Side(style='double', color='2C3E50')
    )
    
    # --- SHEET 1: EXECUTIVE SUMMARY ---
    ws_exec = wb.active
    ws_exec.title = "Executive Summary"
    ws_exec.views.sheetView[0].showGridLines = True
    
    ws_exec.cell(row=2, column=2, value="FinSight Financial Performance Summary").font = title_font
    
    # Helper to create a KPI Card
    def create_kpi_card(ws, start_col, start_row, label, value_formula, number_format):
        ws.merge_cells(start_row=start_row, start_column=start_col, end_row=start_row, end_column=start_col+1)
        ws.merge_cells(start_row=start_row+1, start_column=start_col, end_row=start_row+1, end_column=start_col+1)
        
        lbl_cell = ws.cell(row=start_row, column=start_col, value=label)
        lbl_cell.font = Font(name=font_family, size=9, bold=True, color="7F8C8D")
        lbl_cell.alignment = Alignment(horizontal="center")
        
        val_cell = ws.cell(row=start_row+1, column=start_col, value=value_formula)
        val_cell.font = Font(name=font_family, size=14, bold=True, color="1F4E78")
        val_cell.alignment = Alignment(horizontal="center")
        val_cell.number_format = number_format
        
        # Border box
        for r in range(start_row, start_row+2):
            for c in range(start_col, start_col+2):
                ws.cell(row=r, column=c).border = thin_border
                
    # Add KPI Cards referencing the Monthly Performance sheet
    create_kpi_card(ws_exec, 2, 4, "TOTAL REVENUE", "='Revenue & Margin Analysis'!B15", "$#,##0.00")
    create_kpi_card(ws_exec, 5, 4, "TOTAL NET PROFIT", "='Revenue & Margin Analysis'!D15", "$#,##0.00")
    create_kpi_card(ws_exec, 8, 4, "NET PROFIT MARGIN", "='Revenue & Margin Analysis'!E15", "0.0%")
    
    # Financial indicators table
    ws_exec.cell(row=8, column=2, value="Corporate Performance Overview").font = section_font
    overview_headers = ["Key Metrics", "Current Year Value", "Target Budget", "Variance"]
    for c_idx, h in enumerate(overview_headers, start=2):
        cell = ws_exec.cell(row=9, column=c_idx, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")
        cell.border = thin_border
        
    metrics_list = [
        ("Total Sales Revenue", "='Revenue & Margin Analysis'!B15", "='Budget vs Actual Report'!C8", "=C10-D10"),
        ("Operational Expenses", "='Revenue & Margin Analysis'!C15", "='Budget vs Actual Report'!E8", "=D11-C11"),
        ("Net Operating Profit", "='Revenue & Margin Analysis'!D15", "=C10-C11", "=C12-D12"),
    ]
    
    for r_idx, (m_name, act_f, bud_f, var_f) in enumerate(metrics_list, start=10):
        c1 = ws_exec.cell(row=r_idx, column=2, value=m_name)
        c2 = ws_exec.cell(row=r_idx, column=3, value=act_f)
        c3 = ws_exec.cell(row=r_idx, column=4, value=bud_f)
        c4 = ws_exec.cell(row=r_idx, column=5, value=var_f)
        
        c1.font = bold_font
        c1.border = thin_border
        
        for c in [c2, c3, c4]:
            c.font = normal_font
            c.border = thin_border
            c.number_format = "$#,##0.00"
            c.alignment = Alignment(horizontal="right")
            
    # Add border totals
    for c in range(2, 6):
        ws_exec.cell(row=13, column=c).border = Border(bottom=Side(style='double', color='1F4E78'))
        
    # --- SHEET 2: REVENUE & MARGIN ANALYSIS ---
    ws_rev = wb.create_sheet(title="Revenue & Margin Analysis")
    ws_rev.views.sheetView[0].showGridLines = True
    
    ws_rev.cell(row=1, column=1, value="Monthly Revenue & Profitability Analysis").font = title_font
    
    rev_headers = ["Month", "Revenue", "Expenses", "Net Profit", "Profit Margin", "Running Total Revenue"]
    for c_idx, h in enumerate(rev_headers, start=1):
        cell = ws_rev.cell(row=2, column=c_idx, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")
        cell.border = thin_border
        
    for r_idx, row in df_monthly.iterrows():
        excel_row = r_idx + 3
        # Add month
        ws_rev.cell(row=excel_row, column=1, value=row["Month"]).font = normal_font
        
        # Add Raw numbers
        ws_rev.cell(row=excel_row, column=2, value=row["Revenue"])
        ws_rev.cell(row=excel_row, column=3, value=row["Expenses"])
        
        # Add Excel Formulas
        # Net Profit = Revenue - Expenses
        ws_rev.cell(row=excel_row, column=4, value=f"=B{excel_row}-C{excel_row}")
        # Profit Margin = Net Profit / Revenue
        ws_rev.cell(row=excel_row, column=5, value=f"=IF(B{excel_row}=0, 0, D{excel_row}/B{excel_row})")
        # Running Total = SUM(B$3:B{current})
        ws_rev.cell(row=excel_row, column=6, value=f"=SUM(B$3:B{excel_row})")
        
        # Style row cells
        for c_idx in range(1, 7):
            cell = ws_rev.cell(row=excel_row, column=c_idx)
            cell.font = normal_font
            cell.border = thin_border
            if r_idx % 2 == 1:
                cell.fill = zebra_fill
                
            if c_idx in [2, 3, 4, 6]:
                cell.number_format = "$#,##0.00"
                cell.alignment = Alignment(horizontal="right")
            elif c_idx == 5:
                cell.number_format = "0.0%"
                cell.alignment = Alignment(horizontal="right")
            else:
                cell.alignment = Alignment(horizontal="center")
                
    # Add Totals Row
    tot_row = len(df_monthly) + 3
    ws_rev.cell(row=tot_row, column=1, value="Total").font = bold_font
    ws_rev.cell(row=tot_row, column=2, value=f"=SUM(B3:B{tot_row-1})")
    ws_rev.cell(row=tot_row, column=3, value=f"=SUM(C3:C{tot_row-1})")
    ws_rev.cell(row=tot_row, column=4, value=f"=SUM(D3:D{tot_row-1})")
    ws_rev.cell(row=tot_row, column=5, value=f"=IF(B{tot_row}=0, 0, D{tot_row}/B{tot_row})")
    ws_rev.cell(row=tot_row, column=6, value="") # running total blank
    
    for c_idx in range(1, 7):
        cell = ws_rev.cell(row=tot_row, column=c_idx)
        cell.font = bold_font
        cell.fill = summary_fill
        cell.border = double_bottom_border
        if c_idx in [2, 3, 4]:
            cell.number_format = "$#,##0.00"
            cell.alignment = Alignment(horizontal="right")
        elif c_idx == 5:
            cell.number_format = "0.0%"
            cell.alignment = Alignment(horizontal="right")
            
    # --- SHEET 3: BUDGET VS ACTUAL REPORT ---
    ws_bva = wb.create_sheet(title="Budget vs Actual Report")
    ws_bva.views.sheetView[0].showGridLines = True
    
    ws_bva.cell(row=1, column=1, value="Budget vs Actual Variance Report (by Department)").font = title_font
    
    bva_headers = ["Department", "Actual Revenue", "Budget Revenue", "Revenue Variance", "Actual Expenses", "Budget Expenses", "Expense Variance"]
    for c_idx, h in enumerate(bva_headers, start=1):
        cell = ws_bva.cell(row=2, column=c_idx, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")
        cell.border = thin_border
        
    for r_idx, row in df_budget_vs_actual.iterrows():
        excel_row = r_idx + 3
        
        ws_bva.cell(row=excel_row, column=1, value=row["Department"]).font = bold_font
        ws_bva.cell(row=excel_row, column=2, value=row["Actual_Revenue"])
        ws_bva.cell(row=excel_row, column=3, value=row["Budget_Revenue"])
        # Revenue Variance = Actual - Budget
        ws_bva.cell(row=excel_row, column=4, value=f"=B{excel_row}-C{excel_row}")
        
        ws_bva.cell(row=excel_row, column=5, value=row["Actual_Expenses"])
        ws_bva.cell(row=excel_row, column=6, value=row["Budget_Expenses"])
        # Expense Variance = Budget - Actual (positive is savings)
        ws_bva.cell(row=excel_row, column=7, value=f"=F{excel_row}-E{excel_row}")
        
        for c_idx in range(1, 8):
            cell = ws_bva.cell(row=excel_row, column=c_idx)
            cell.border = thin_border
            if r_idx % 2 == 1:
                cell.fill = zebra_fill
                
            if c_idx > 1:
                cell.number_format = "$#,##0.00"
                cell.alignment = Alignment(horizontal="right")
                cell.font = normal_font
                
    # Add Totals Row
    tot_row = len(df_budget_vs_actual) + 3
    ws_bva.cell(row=tot_row, column=1, value="Total").font = bold_font
    ws_bva.cell(row=tot_row, column=2, value=f"=SUM(B3:B{tot_row-1})")
    ws_bva.cell(row=tot_row, column=3, value=f"=SUM(C3:C{tot_row-1})")
    ws_bva.cell(row=tot_row, column=4, value=f"=B{tot_row}-C{tot_row}")
    ws_bva.cell(row=tot_row, column=5, value=f"=SUM(E3:E{tot_row-1})")
    ws_bva.cell(row=tot_row, column=6, value=f"=SUM(F3:F{tot_row-1})")
    ws_bva.cell(row=tot_row, column=7, value=f"=F{tot_row}-E{tot_row}")
    
    for c_idx in range(1, 8):
        cell = ws_bva.cell(row=tot_row, column=c_idx)
        cell.font = bold_font
        cell.fill = summary_fill
        cell.border = double_bottom_border
        if c_idx > 1:
            cell.number_format = "$#,##0.00"
            cell.alignment = Alignment(horizontal="right")
            
    # --- SHEET 4: PROFITABILITY BY SEGMENT ---
    ws_seg = wb.create_sheet(title="Segment Profitability")
    ws_seg.views.sheetView[0].showGridLines = True
    
    # Regional table
    ws_seg.cell(row=1, column=1, value="Profitability by Region").font = section_font
    reg_headers = ["Region", "Revenue", "Expenses", "Net Profit", "Margin"]
    for c_idx, h in enumerate(reg_headers, start=1):
        cell = ws_seg.cell(row=2, column=c_idx, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")
        cell.border = thin_border
        
    for r_idx, row in df_regions.iterrows():
        excel_row = r_idx + 3
        ws_seg.cell(row=excel_row, column=1, value=row["Region"]).font = bold_font
        ws_seg.cell(row=excel_row, column=2, value=row["Revenue"])
        ws_seg.cell(row=excel_row, column=3, value=row["Expenses"])
        ws_seg.cell(row=excel_row, column=4, value=f"=B{excel_row}-C{excel_row}")
        ws_seg.cell(row=excel_row, column=5, value=f"=IF(B{excel_row}=0,0,D{excel_row}/B{excel_row})")
        
        for c_idx in range(1, 6):
            cell = ws_seg.cell(row=excel_row, column=c_idx)
            cell.border = thin_border
            if c_idx in [2, 3, 4]:
                cell.number_format = "$#,##0.00"
                cell.alignment = Alignment(horizontal="right")
            elif c_idx == 5:
                cell.number_format = "0.0%"
                cell.alignment = Alignment(horizontal="right")
                
    # Regional Total
    reg_tot = len(df_regions) + 3
    ws_seg.cell(row=reg_tot, column=1, value="Total").font = bold_font
    ws_seg.cell(row=reg_tot, column=2, value=f"=SUM(B3:B{reg_tot-1})")
    ws_seg.cell(row=reg_tot, column=3, value=f"=SUM(C3:C{reg_tot-1})")
    ws_seg.cell(row=reg_tot, column=4, value=f"=B{reg_tot}-C{reg_tot}")
    ws_seg.cell(row=reg_tot, column=5, value=f"=IF(B{reg_tot}=0,0,D{reg_tot}/B{reg_tot})")
    
    for c_idx in range(1, 6):
        cell = ws_seg.cell(row=reg_tot, column=c_idx)
        cell.font = bold_font
        cell.fill = summary_fill
        cell.border = double_bottom_border
        if c_idx in [2, 3, 4]:
            cell.number_format = "$#,##0.00"
        elif c_idx == 5:
            cell.number_format = "0.0%"
            
    # Product Line table (offset by 2 columns)
    ws_seg.cell(row=1, column=7, value="Profitability by Product Line").font = section_font
    prod_headers = ["Product Line", "Revenue", "Expenses", "Net Profit", "Margin"]
    for c_idx, h in enumerate(prod_headers, start=7):
        cell = ws_seg.cell(row=2, column=c_idx, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")
        cell.border = thin_border
        
    for r_idx, row in df_products.iterrows():
        excel_row = r_idx + 3
        ws_seg.cell(row=excel_row, column=7, value=row["Product_Line"]).font = bold_font
        ws_seg.cell(row=excel_row, column=8, value=row["Revenue"])
        ws_seg.cell(row=excel_row, column=9, value=row["Expenses"])
        ws_seg.cell(row=excel_row, column=10, value=f"=H{excel_row}-I{excel_row}")
        ws_seg.cell(row=excel_row, column=11, value=f"=IF(H{excel_row}=0,0,J{excel_row}/H{excel_row})")
        
        for c_idx in range(7, 12):
            cell = ws_seg.cell(row=excel_row, column=c_idx)
            cell.border = thin_border
            if c_idx in [8, 9, 10]:
                cell.number_format = "$#,##0.00"
                cell.alignment = Alignment(horizontal="right")
            elif c_idx == 11:
                cell.number_format = "0.0%"
                cell.alignment = Alignment(horizontal="right")
                
    # Product Total
    prod_tot = len(df_products) + 3
    ws_seg.cell(row=prod_tot, column=7, value="Total").font = bold_font
    ws_seg.cell(row=prod_tot, column=8, value=f"=SUM(H3:H{prod_tot-1})")
    ws_seg.cell(row=prod_tot, column=9, value=f"=SUM(I3:I{prod_tot-1})")
    ws_seg.cell(row=prod_tot, column=10, value=f"=H{prod_tot}-I{prod_tot}")
    ws_seg.cell(row=prod_tot, column=11, value=f"=IF(H{prod_tot}=0,0,J{prod_tot}/H{prod_tot})")
    
    for c_idx in range(7, 12):
        cell = ws_seg.cell(row=prod_tot, column=c_idx)
        cell.font = bold_font
        cell.fill = summary_fill
        cell.border = double_bottom_border
        if c_idx in [8, 9, 10]:
            cell.number_format = "$#,##0.00"
        elif c_idx == 11:
            cell.number_format = "0.0%"
            
    # Auto-fit column widths across all sheets
    for ws in wb.worksheets:
        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                val = str(cell.value or '')
                if val.startswith('='): # skip length calculation for formulas
                    val = " $1,000,000.00 "
                max_len = max(max_len, len(val))
            ws.column_dimensions[col_letter].width = max(max_len + 3, 11)
            
    wb.save(output_path)
    print("Excel workbook created successfully!")

if __name__ == "__main__":
    db_file = "src/financial.db"
    out_file = "financial_analysis.xlsx"
    create_styled_excel(db_file, out_file)
