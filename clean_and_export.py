import csv
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

# 1. Extract raw records
with open("raw_supply_chain_orders.csv", mode="r", encoding="utf-8") as f:
    raw_rows = list(csv.DictReader(f))

# 2. Deduplicate by Order_ID and compute category cost ratios for imputation
seen_ids = set()
unique_rows = []
category_ratios = {}

for row in raw_rows:
    oid = row["Order_ID"].strip()
    if oid in seen_ids:
        continue
    seen_ids.add(oid)
    unique_rows.append(row)

    if row["Unit_Cost_USD"]:
        cat = row["Category"].strip()
        cost = float(row["Unit_Cost_USD"])
        price = float(row["Unit_Price_USD"])
        category_ratios.setdefault(cat, []).append(cost / price)

avg_ratio_by_cat = {
    cat: sum(ratios) / len(ratios) for cat, ratios in category_ratios.items()
}

# 3. Transform and sanitize fields
cleaned_rows = []
for row in unique_rows:
    # Standardize mixed date formats
    raw_date = row["Order_Date"].strip()
    try:
        parsed_date = datetime.strptime(raw_date, "%Y-%m-%d")
    except ValueError:
        parsed_date = datetime.strptime(raw_date, "%d/%m/%Y")

    # Normalize Region casing and whitespace
    region = row["Region"].strip().title() if row["Region"].strip() else "Unassigned"
    category = row["Category"].strip()
    units = int(row["Units_Sold"])
    unit_price = round(float(row["Unit_Price_USD"]), 2)

    # Impute missing Unit_Cost_USD using category average ratio
    if row["Unit_Cost_USD"]:
        unit_cost = round(float(row["Unit_Cost_USD"]), 2)
        cost_imputed = "No"
    else:
        unit_cost = round(unit_price * avg_ratio_by_cat[category], 2)
        cost_imputed = "Yes"

    cleaned_rows.append([
        row["Order_ID"].strip(),
        parsed_date.strftime("%Y-%m-%d"),
        parsed_date.strftime("%Y-%m"),
        row["Supplier_ID"].strip(),
        row["Supplier_Name"].strip(),
        region,
        category,
        units,
        unit_cost,
        unit_price,
        cost_imputed,
    ])

# Sort chronologically by Order_Date
cleaned_rows.sort(key=lambda x: x[1])

# 4. Load into a structured Excel workbook
wb = Workbook()

# Sheet 1: Clean_Orders
ws_data = wb.active
ws_data.title = "Clean_Orders"

headers = [
    "Order_ID",
    "Order_Date",
    "Year_Month",
    "Supplier_ID",
    "Supplier_Name",
    "Region",
    "Category",
    "Units_Sold",
    "Unit_Cost_USD",
    "Unit_Price_USD",
    "Cost_Imputed",
]
ws_data.append(headers)

header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
header_font = Font(color="FFFFFF", bold=True)

for cell in ws_data[1]:
    cell.fill = header_fill
    cell.font = header_font
    cell.alignment = Alignment(horizontal="center")

for r in cleaned_rows:
    ws_data.append(r)

# Apply currency formatting to cost and price columns
for row_idx in range(2, len(cleaned_rows) + 2):
    ws_data[f"I{row_idx}"].number_format = "$#,##0.00"
    ws_data[f"J{row_idx}"].number_format = "$#,##0.00"

# Sheet 2: Executive_Dashboard (template for manual Excel modeling)
ws_dash = wb.create_sheet(title="Executive_Dashboard")
ws_dash["A1"] = "SUPPLY CHAIN PROFITABILITY & SUPPLIER SCORECARD"
ws_dash["A1"].font = Font(size=14, bold=True, color="1F4E78")

output_file = "supply_chain_analysis.xlsx"
wb.save(output_file)
print(f"Cleaned {len(cleaned_rows)} unique orders and exported to {output_file}.")