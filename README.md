# Retail & Supply Chain Profitability Analyzer

An end-to-end data engineering and financial modeling pipeline combining **pure-Python ETL** (`csv`, `openpyxl`) with **dynamic Microsoft Excel BI modeling** and automated Markdown reporting.

## Architecture & Workflow

1. **Raw Data Simulation (`generate_raw_data.py`)**
   * Generates a realistic 2,535-row supply chain order dataset (`raw_supply_chain_orders.csv`) containing intentional data quality issues: duplicate order IDs, inconsistent date formats, unstandardized region names, and missing unit costs.

2. **Pure-Python ETL & Imputation (`clean_and_export.py`)**
   * Deduplicates records to 2,500 unique orders using standard library data structures (`csv`, `collections`, `datetime`).
   * Normalizes dates and regional labels, imputes missing unit costs using category-level cost-to-price ratios, and flags imputed rows (`Cost_Imputed`).
   * Exports a multi-tab Excel workbook (`supply_chain_analysis.xlsx`) with formatted headers and structured sheets (`Clean_Orders` and `Executive_Dashboard`).

3. **Financial Modeling & Business Intelligence (`supply_chain_analysis.xlsx`)**
   * **Order-Level Calculations:** Dynamic `Total_COGS`, `Total_Revenue`, `Gross_Profit`, and `Gross_Margin_%` columns.
   * **Category Scorecard:** `SUMIFS`-driven KPI summary paired with a Revenue vs. Gross Profit visual chart.
   * **Interactive Drill-Down:** Regional and supplier Pivot Table controlled by live `Category` and `Cost_Imputed` Slicers.
   * **Supplier Margin Leaderboard:** `INDEX/MATCH`, `COUNTIFS`, and `SUMIFS` ranking table with conditional formatting color scales.

4. **Automated Audit & Reporting (`generate_report.py`)**
   * Reads the completed Excel workbook back into Python (`openpyxl`), reconciles KPIs against raw order totals, and outputs an executive brief (`executive_summary.md`).

## Tech Stack
* **Python 3** (`csv`, `openpyxl`, standard library—zero C-extension/DLL dependencies)
* **Microsoft Excel** (`SUMIFS`, `COUNTIFS`, `INDEX/MATCH`, `IFERROR`, Pivot Tables, Slicers, Conditional Formatting)

## Quickstart

```bash
# 1. Create and activate virtual environment
python -m venv venv
.\venv\Scripts\Activate.ps1

# 2. Install openpyxl
python -m pip install openpyxl

# 3. Generate raw dataset and run ETL pipeline
python generate_raw_data.py
python clean_and_export.py

# 4. Generate executive Markdown report from the Excel workbook
python generate_report.py