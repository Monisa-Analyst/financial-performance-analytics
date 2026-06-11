# FinSight — Corporate Financial Performance, Excel Modeling & Power BI Analytics

An end-to-end corporate financial analytics and business intelligence platform designed to ingest daily corporate transactions, audit data quality, reconcile monthly budgets, and analyze profitability. FinSight bridges automated data engineering (Power Query, Python ETL) with dynamic financial modeling (formula-driven multi-sheet Excel workbooks) and interactive visual reporting (Power BI Star Schema, DAX, multi-page Streamlit dashboards).

👉 **[Live Web Application Link](https://share.streamlit.io/monisa-analyst/financial-performance-analytics/main/src/app.py)**

---

## 🚀 Key Features

1. **Interactive Multi-Page BI Dashboard (Streamlit & Plotly):**
   - **Executive Summary:** Core financial KPIs (Total Revenue, Operating Expenses, Net Operating Profit, Profit Margin %, ROI %, Budget Variance %) and trend lines (MoM growth, running totals).
   - **Profitability & Budgets:** Granular breakdowns by region and product segment (identifying high-performing lines) and a department budget-to-actual variance matrix.
   - **Cohorts & Funnels:** Interactive Client Revenue Retention Heatmaps (Net Revenue Retention & Logo Retention) and Accounts Receivable Invoicing conversion funnel.
   - **Ingest Transactions:** Drag-and-drop CSV/Excel portal for daily transactions uploads.
   - **Audit Trail:** logs of files processed and database-wide health diagnostics.
   - **Power BI & DAX Report:** Dedicated showcase detailing Star Schema connections, Power Query M code transformations, and the DAX measure library.

2. **Fuzzy Column Ingestion & Standardisation:**
   - Normalizes incoming column headers dynamically (e.g. mapping `spending`, `outflow`, or `costs` automatically to `expenses`).
   - Automatically cleans currency formatting, strips commas/symbols, and handles accounting negative brackets `(500.00)` -> `-500.00`.

3. **9-Point Data Quality Audit Gate:**
   - Evaluates each batch:
     - **Critical:** Missing dates.
     - **Warning:** Future dates, negative revenue/expense inputs, zero amounts, blank dimension values.
     - **Info:** Large transaction spikes (Revenue > $50K, Expense > $25K), duplicate rows.
   - Automatically computes a **Batch Data Quality Score**. Batches with score `< 50%` are automatically rejected and rolled back to preserve database integrity.

4. **Dynamic Excel Workbook Compiler (`openpyxl`):**
   - Automatically compiles and refines a beautifully styled Excel workbook (`financial_analysis.xlsx`) after every successful data ingestion.
   - **Recruiter-Friendly Native Formulas:** Writes actual uppercase formulas (e.g., `=SUMIFS(...)`, `=XLOOKUP(...)`, `=INDEX(...)`, `=MATCH(...)`) into cells so that the spreadsheet behaves dynamically for users rather than displaying flat, hardcoded values.
   - Designed with a professional corporate theme (Segoe UI, navy fills, double accounting underlines, auto-fitting column widths).

5. **Power BI Star Schema & DAX Integration:**
   - Fully documented Star Schema structure linking dimensional lookup sheets to transactional facts.
   - DAX query library configured to calculate cohort net revenue retention (NRR), time-series MoM variances, and cumulative invoice collections.

---

## 📐 Star Schema Architecture (Power BI)

The Excel workbook worksheets are structured to serve as clean data sources for a standard Star Schema model:

```
                  ┌─────────────────┐
                  │ dim_departments │
                  └────────┬────────┘
                           │ 1:N
 ┌─────────────┐     ┌─────┴──────────┐     ┌─────────────┐
 │ dim_regions ├─────┤fact_transactions├─────┤dim_products │
 └─────────────┘ 1:N └─────┬──────────┘ N:1 └─────────────┘
                           │ N:1
                  ┌────────┴─────────────┐
                  │dim_expense_categories│
                  └──────────────────────┘
```

- **Fact Tables:** `fact_transactions`, `fact_budgets`, `fact_invoices`
- **Dimension Tables:** `dim_departments`, `dim_regions`, `dim_products`, `dim_expense_categories`, `dim_calendar`

---

## 💡 Advanced Showcases

Finance recruiters look for clean, standardized, and advanced models. FinSight highlights these directly in the Excel and Power BI layers:

### 1. Advanced Excel Formulas (Written Programmatically)
Instead of writing flat values, the Excel compiler writes dynamic relational formulas to calculate summaries:

*   **SUMIFS with XLOOKUP (Departmental Budget vs Actual Actuals):**
    Queries transactions based on the department ID mapped dynamically via `XLOOKUP`:
    ```excel
    =SUMIFS(fact_transactions!C:C, fact_transactions!E:E, XLOOKUP(A3, dim_departments!B:B, dim_departments!A:A))
    ```
*   **SUMIFS with Month Wildcards (Monthly Revenue):**
    Aggregates revenue by month using wildcard prefix matching on date text strings:
    ```excel
    =SUMIFS(fact_transactions!C:C, fact_transactions!B:B, A3&"-*")
    ```
*   **Nested IF with Division Safeguards (Margins):**
    Calculates profit margins dynamically while preventing `#DIV/0!` errors:
    ```excel
    =IF(B3=0, 0, D3/B3)
    ```

---

### 2. Advanced Power BI DAX Measures

*   **Month-over-Month Revenue Growth % (Time Intelligence):**
    Uses variables and contextual calculation modifications to compare sales MoM:
    ```dax
    Revenue MoM Growth % = 
    VAR CurrentRev = [Total Revenue]
    VAR PrevMonthRev = 
        CALCULATE(
            [Total Revenue],
            DATEADD('dim_calendar'[Date], -1, MONTH)
        )
    RETURN
        DIVIDE(CurrentRev - PrevMonthRev, PrevMonthRev, 0)
    ```

*   **Net Revenue Retention (NRR) Cohorts:**
    Aggregates repeat transaction values from a customer signup cohort over elapsed months:
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

*   **Accounts Receivable Invoicing funnel collection rate:**
    Evaluates cumulative settlement conversion rates through the collection stages:
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

---

## 🛠️ How to Run Locally

### 1. Install Dependencies
```bash
pip install -r requirements.txt
```

### 2. Generate Data & Seed Excel Model
Initialize the worksheets and load the raw transactional tables:
```bash
python generate_data.py
python src/db_init.py
```

### 3. Compile Styled Corporate Report Sheets
```bash
python src/excel_generator.py
```
This builds the styled charts and active formulas inside `financial_analysis.xlsx`.

### 4. Run Automated Pipeline Tests
```bash
python test_finance.py
```

### 5. Launch Dashboard
```bash
streamlit run src/app.py
```
Open `http://localhost:8501` in your browser.

---

## 📬 Contact & Connections

- **Author:** Monisa L.
- **Email:** [monisa.asi@gmail.com](mailto:monisa.asi@gmail.com)
- **LinkedIn:** [linkedin.com/in/monisa-l-333546366](https://www.linkedin.com/in/monisa-l-333546366)
- **GitHub Profile:** [github.com/Monisa-Analyst](https://github.com/Monisa-Analyst)
