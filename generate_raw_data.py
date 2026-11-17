import csv
import random
from datetime import datetime, timedelta

random.seed(42)

regions = ["Nairobi", "nairobi ", "Mombasa", "Kisumu", "Nakuru", "Eldoret", ""]
categories = ["Electronics", "Packaging", "Raw Materials", "Logistics", "Hardware"]
suppliers = {
    "SUP-01": ("Apex Industrial", 1.15),
    "SUP-02": ("Rift Valley Traders", 1.05),
    "SUP-03": ("Coastal Freight Ltd", 1.25),
    "SUP-04": ("Savannah Components", 0.95),
}

rows = []
start_date = datetime(2026, 1, 1)

for i in range(1, 2501):
    order_date = start_date + timedelta(days=random.randint(0, 300))
    date_str = (
        order_date.strftime("%Y-%m-%d")
        if i % 7 != 0
        else order_date.strftime("%d/%m/%Y")
    )

    sup_id = random.choice(list(suppliers.keys()))
    sup_name, cost_multiplier = suppliers[sup_id]
    category = random.choice(categories)
    region = random.choice(regions)

    units_sold = random.randint(5, 150)
    unit_cost = round(random.uniform(12.0, 240.0) * cost_multiplier, 2)
    unit_price = round(unit_cost * random.uniform(1.15, 1.45), 2)

    if i % 45 == 0:
        unit_cost = ""

    rows.append([
        f"ORD-{10000 + i}",
        date_str,
        sup_id,
        sup_name,
        region,
        category,
        units_sold,
        unit_cost,
        unit_price,
    ])

# Inject 35 duplicate rows to simulate a dirty export
duplicates = random.sample(rows, 35)
dirty_rows = rows + duplicates
random.shuffle(dirty_rows)

headers = [
    "Order_ID",
    "Order_Date",
    "Supplier_ID",
    "Supplier_Name",
    "Region",
    "Category",
    "Units_Sold",
    "Unit_Cost_USD",
    "Unit_Price_USD",
]

with open("raw_supply_chain_orders.csv", mode="w", newline="", encoding="utf-8") as f:
    writer = csv.writer(f)
    writer.writerow(headers)
    writer.writerows(dirty_rows)

print(f"Generated raw_supply_chain_orders.csv with {len(dirty_rows)} rows.")