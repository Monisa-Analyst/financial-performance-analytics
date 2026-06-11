# Power BI Data Modeling & DAX Reference Guide — FinSight

This guide documents the data model structure, ETL transformations, and analytical metrics configured in Power BI for the **FinSight** corporate performance platform.

---

## 📐 Star Schema Architecture

FinSight uses a clean star schema to ensure high query performance, clean filtering, and easy-to-read DAX calculations:

```mermaid
classDiagram
    class dim_departments {
        Int dept_id (PK)
        String dept_name
    }
    class dim_regions {
        Int region_id (PK)
        String region_name
    }
    class dim_products {
        Int product_id (PK)
        String product_name
    }
    class dim_expense_categories {
        Int category_id (PK)
        String category_name
    }
    class dim_calendar {
        Date Date (PK)
        Int Year
        Int Quarter
        String Month
        Int MonthNo
    }
    class fact_transactions {
        Int transaction_id (PK)
        Date Date (FK)
        Real revenue
        Real expenses
        Int dept_id (FK)
        Int region_id (FK)
        Int product_id (FK)
        Int category_id (FK)
        String client_id
    }
    class fact_budgets {
        Int budget_id (PK)
        String month (FK)
        Int dept_id (FK)
        Int region_id (FK)
        Real budgeted_revenue
        Real budgeted_expenses
    }
    class fact_invoices {
        String invoice_id (PK)
        String client_id
        Date issue_date (FK)
        Real amount
        String status
    }

    fact_transactions }--|| dim_departments : "dept_id"
    fact_transactions }--|| dim_regions : "region_id"
    fact_transactions }--|| dim_products : "product_id"
    fact_transactions }--|| dim_expense_categories : "category_id"
    fact_transactions }--|| dim_calendar : "Date"
    fact_budgets }--|| dim_departments : "dept_id"
    fact_budgets }--|| dim_regions : "region_id"
    fact_invoices }--|| dim_calendar : "issue_date"
```

### Relationships Settings
- **`fact_transactions ➔ dim_calendar`**: Many-to-One (`* : 1`), Single Direction, Active on `Date`.
- **`fact_transactions ➔ dim_departments`**: Many-to-One (`* : 1`), Single Direction.
- **`fact_transactions ➔ dim_regions`**: Many-to-One (`* : 1`), Single Direction.
- **`fact_transactions ➔ dim_products`**: Many-to-One (`* : 1`), Single Direction.
- **`fact_transactions ➔ dim_expense_categories`**: Many-to-One (`* : 1`), Single Direction.
- **`fact_budgets ➔ dim_departments`**: Many-to-One (`* : 1`), Single Direction.
- **`fact_budgets ➔ dim_regions`**: Many-to-One (`* : 1`), Single Direction.

---

## 🛠️ Power Query ETL (M Code)

To transform and standardize column headers dynamically, the following M query script is used when loading transactional batches:

```powerquery
let
    Source = Excel.Workbook(File.Contents("C:\Users\HP\Desktop\git\financial-performance-analytics\financial_analysis.xlsx"), null, true),
    Transactions_Sheet = Source{[Item="fact_transactions",Kind="Sheet"]}[Data],
    PromoteHeaders = Table.PromoteHeaders(Transactions_Sheet, [PromoteAllScalars=true]),
    
    // Auto-detect and map columns dynamically
    RenameCols = Table.RenameColumns(PromoteHeaders, {
        {"spending", "Expenses"}, {"costs", "Expenses"}, {"outflow", "Expenses"},
        {"income", "Revenue"}, {"sales", "Revenue"}, {"earnings", "Revenue"}
    }, MissingField.UseNull),
    
    // Standardize currency formats and strip symbols
    CleanRevenue = Table.TransformColumns(RenameCols, {{"Revenue", each if _ is text then Value.FromText(Text.Select(_, {"0".."9", ".", "-"})) else _, type number}}),
    CleanExpenses = Table.TransformColumns(CleanRevenue, {{"Expenses", each if _ is text then Value.FromText(Text.Select(_, {"0".."9", ".", "-"})) else _, type number}}),
    
    // Standardize Date formats
    ParseDate = Table.TransformColumnTypes(CleanExpenses,{{"Date", type date}})
in
    ParseDate
```

---

## 📊 DAX Measure Library

Recruiters look for clean, structured, and advanced DAX formulas that demonstrate deep BI dashboard expertise.

### 1. Core Financial Measures

```dax
// Total Revenue
Total Revenue = SUM('fact_transactions'[revenue])

// Total Operating Expenses
Total Expenses = SUM('fact_transactions'[expenses])

// Net Operating Profit
Net Profit = [Total Revenue] - [Total Expenses]

// Profit Margin %
Profit Margin % = DIVIDE([Net Profit], [Total Revenue], 0)

// Return on Investment (ROI %)
ROI % = DIVIDE([Net Profit], [Total Expenses], 0)
```

### 2. Variance & Budgeting Calculations

```dax
// Budgeted Revenue
Budget Revenue = SUM('fact_budgets'[budgeted_revenue])

// Budgeted Expenses
Budget Expenses = SUM('fact_budgets'[budgeted_expenses])

// Revenue Variance (Actual vs Budget)
Revenue Variance = [Total Revenue] - [Budget Revenue]

// Revenue Variance %
Revenue Variance % = DIVIDE([Revenue Variance], [Budget Revenue], 0)

// Expense Savings Variance (Positive means under budget)
Expense Variance = [Budget Expenses] - [Total Expenses]
```

### 3. Time Intelligence (MoM Growth & Running Totals)

```dax
// Month-over-Month Revenue Growth %
Revenue MoM Growth % = 
VAR CurrentRev = [Total Revenue]
VAR PrevMonthRev = 
    CALCULATE(
        [Total Revenue],
        DATEADD('dim_calendar'[Date], -1, MONTH)
    )
RETURN
    DIVIDE(CurrentRev - PrevMonthRev, PrevMonthRev, 0)

// Running Year-to-Date Revenue Total
YTD Revenue = 
TOTALYTD(
    [Total Revenue],
    'dim_calendar'[Date]
)
```

### 4. Client Lifetime Net Revenue Retention (NRR)
This measure calculates retention by tracking revenue from a cohort of clients acquired in a specific month:

```dax
Net Revenue Retention % = 
VAR CohortMonth = SELECTEDVALUE('dim_calendar'[CohortMonth])
VAR SelectedMonth = SELECTEDVALUE('dim_calendar'[Month])
VAR BaseRevenue = 
    CALCULATE(
        [Total Revenue],
        FILTER(
            ALL('dim_calendar'),
            'dim_calendar'[CohortMonth] = CohortMonth && 
            'dim_calendar'[Month] = CohortMonth
        )
    )
VAR CurrentCohortRevenue = 
    CALCULATE(
        [Total Revenue],
        FILTER(
            ALL('dim_calendar'),
            'dim_calendar'[CohortMonth] = CohortMonth && 
            'dim_calendar'[Month] = SelectedMonth
        )
    )
RETURN
    DIVIDE(CurrentCohortRevenue, BaseRevenue, 0)
```

### 5. Accounts Receivable Invoicing Funnel
Calculates the settlement conversion rate through the invoicing collection pipeline:

```dax
Settlement Conversion Rate % = 
VAR ActiveStage = SELECTEDVALUE('dim_invoices_stages'[StageOrder])
VAR TotalInvoicesCreated = COUNTROWS('fact_invoices')
VAR CurrentStageInvoices = 
    CALCULATE(
        COUNTROWS('fact_invoices'),
        'fact_invoices'[status_num] >= ActiveStage
    )
RETURN
    DIVIDE(CurrentStageInvoices, TotalInvoicesCreated, 0)
```
