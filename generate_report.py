from openpyxl import load_workbook

wb = load_workbook("supply_chain_analysis.xlsx", data_only=True)
ws_dash = wb["Executive_Dashboard"]
ws_data = wb["Clean_Orders"]

# Build fallback aggregations directly from Clean_Orders in case Excel didn't cache values
cat_fallback = {}
sup_fallback = {}

for row in ws_data.iter_rows(min_row=2, values_only=True):
    if not row[0]:
        continue
    sup_id, sup_name, region, cat = row[3], row[4], row[5], row[6]
    units, cost, price = float(row[7]), float(row[8]), float(row[9])
    rev = units * price
    profit = rev - (units * cost)

    c = cat_fallback.setdefault(cat, {"units": 0, "rev": 0.0, "profit": 0.0})
    c["units"] += units
    c["rev"] += rev
    c["profit"] += profit

    s = sup_fallback.setdefault(sup_id, {"name": sup_name, "orders": 0, "rev": 0.0, "profit": 0.0})
    s["orders"] += 1
    s["rev"] += rev
    s["profit"] += profit

# Locate Category and Supplier tables dynamically on Executive_Dashboard
categories = []
suppliers = []

for r_idx in range(1, ws_dash.max_row + 1):
    first_val = str(ws_dash.cell(r_idx, 1).value or "").strip()
    if first_val in cat_fallback:
        units = ws_dash.cell(r_idx, 2).value or cat_fallback[first_val]["units"]
        rev = ws_dash.cell(r_idx, 3).value or cat_fallback[first_val]["rev"]
        profit = ws_dash.cell(r_idx, 4).value or cat_fallback[first_val]["profit"]
        margin = ws_dash.cell(r_idx, 5).value or (profit / rev if rev else 0)
        categories.append((first_val, units, rev, profit, margin))
    elif first_val in sup_fallback:
        sup_name = ws_dash.cell(r_idx, 2).value or sup_fallback[first_val]["name"]
        orders = ws_dash.cell(r_idx, 3).value or sup_fallback[first_val]["orders"]
        profit = ws_dash.cell(r_idx, 4).value or sup_fallback[first_val]["profit"]
        rev = sup_fallback[first_val]["rev"]
        margin = ws_dash.cell(r_idx, 5).value or (profit / rev if rev else 0)
        suppliers.append((first_val, sup_name, orders, profit, margin))

# Build Markdown executive summary
report_lines = [
    "# Supply Chain Profitability & Supplier Executive Report",
    "",
    "## 1. Category Performance Summary",
    "| Category | Total Units | Total Revenue (USD) | Gross Profit (USD) | Avg Margin |",
    "| :--- | ---: | ---: | ---: | ---: |",
]

for cat_name, units, rev, profit, margin in categories:
    report_lines.append(
        f"| {cat_name} | {int(units):,} | ${float(rev):,.2f} | ${float(profit):,.2f} | {float(margin):.1%} |"
    )

report_lines.extend([
    "",
    "## 2. Supplier Margin Leaderboard",
    "| Supplier ID | Supplier Name | Orders Fulfilled | Total Profit (USD) | Effective Margin |",
    "| :--- | :--- | ---: | ---: | ---: |",
])

for sup_id, sup_name, orders, profit, margin in suppliers:
    report_lines.append(
        f"| {sup_id} | {sup_name} | {int(orders):,} | ${float(profit):,.2f} | {float(margin):.1%} |"
    )

with open("executive_summary.md", "w", encoding="utf-8") as f:
    f.write("\n".join(report_lines))

print(f"Successfully generated executive_summary.md ({len(categories)} categories, {len(suppliers)} suppliers)!")

