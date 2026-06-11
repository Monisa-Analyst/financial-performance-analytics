import os
import sys
import unittest
import pandas as pd

# Add src to path
sys.path.append(os.path.join(os.path.dirname(os.path.abspath(__file__)), "src"))

import analytics
import ingestion
import excel_generator

class TestFinancialPipeline(unittest.TestCase):
    def setUp(self):
        self.excel_file = "financial_analysis.xlsx"
        self.test_csv = "financial_upload_test.csv"
        
        # Create test CSV file if it doesn't exist
        df_test = pd.DataFrame([
            {"Date": "2025-06-15", "Revenue": "5400.00", "Expenses": "1200.00", "Department": "Sales", "Region": "South", "Product Line": "Software Licenses", "Expense Category": "Salaries"},
            {"Date": "2025-06-16", "Revenue": "4500.00", "Expenses": "-500.00", "Department": "Marketing", "Region": "North", "Product Line": "Consulting Services", "Expense Category": "Marketing Ads"},
            {"Date": "2025-06-17", "Revenue": "", "Expenses": "", "Department": "Operations", "Region": "West", "Product Line": "Hardware Systems", "Expense Category": "Rent"},
            {"Date": "2026-12-01", "Revenue": "12000.00", "Expenses": "4000.00", "Department": "Engineering", "Region": "East", "Product Line": "Cloud Subscriptions", "Expense Category": "Software Subscriptions"},
            {"Date": "", "Revenue": "3400.00", "Expenses": "1500.00", "Department": "HR", "Region": "South", "Product Line": "Software Licenses", "Expense Category": "Benefits"}
        ])
        df_test.to_csv(self.test_csv, index=False)
        
    def test_excel_database_exists(self):
        self.assertTrue(os.path.exists(self.excel_file), "Excel database file should exist.")
        
    def test_kpi_queries(self):
        kpis = analytics.get_kpis(self.excel_file)
        self.assertIn("total_revenue", kpis)
        self.assertIn("net_profit", kpis)
        self.assertIn("roi", kpis)
        self.assertGreater(kpis["total_revenue"], 0, "Revenue should be positive.")
        
    def test_excel_generation(self):
        # Generate reports sheets
        excel_generator.create_styled_excel(self.excel_file)
        self.assertTrue(os.path.exists(self.excel_file), "Excel workbook should remain created successfully.")
        
    def test_dirty_ingestion_anomalies(self):
        result = ingestion.process_file_upload(self.test_csv, "financial_upload_test.csv", self.excel_file)
        self.assertFalse(result["success"], "Highly anomalous batch should fail merge.")
        self.assertEqual(result["status"], "Rejected", "Status should be Rejected due to anomalous rows.")
        self.assertIn("missing_date", result["issues"], "Should catch missing date.")
        self.assertIn("future_dates", result["issues"], "Should catch future date.")
        self.assertIn("negative_expenses", result["issues"], "Should catch negative expense value.")
        
    def test_column_mapping(self):
        df_raw = pd.DataFrame(columns=["transaction_date", "sales", "spending", "cost_center", "area", "segment", "type"])
        df_mapped = ingestion.map_columns(df_raw)
        self.assertEqual(list(df_mapped.columns), list(ingestion.COLUMN_MAPPING.keys()))

if __name__ == "__main__":
    unittest.main()
