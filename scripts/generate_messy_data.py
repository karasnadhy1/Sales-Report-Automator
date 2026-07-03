"""
Generates a fake but realistic MESSY sales CSV — the kind of export
a small business would pull from a POS system or Shopify.

Messiness baked in on purpose (this is what we'll clean on Day 1):
- Inconsistent date formats
- Inconsistent text casing (category, region)
- Currency symbols / commas mixed into numeric columns
- Missing values scattered around
- Duplicate rows
- Extra whitespace around text
- A few negative quantities (returns) that look like data errors
"""
import csv
import random

random.seed(42)  # reproducible output

CUSTOMERS = [
    "Maria Santos", "James O'Connor", "Liu Wei", "Fatima Al-Rashid",
    "Carlos Mendez", "Aisha Bello", "Tom Becker", "Priya Nair",
    "John Smith", "Emma Wilson", "David Kim", "Sofia Rossi",
    "Ahmed Hassan", "Grace Park", "Lucas Ferreira"
]
PRODUCTS = {
    "Laptop": ("Electronics", 899.99),
    "Headphones": ("Electronics", 59.99),
    "Phone Case": ("Accessories", 14.99),
    "Desk Lamp": ("Home", 24.50),
    "Coffee Mug": ("Home", 9.99),
    "Backpack": ("Accessories", 45.00),
    "Monitor": ("Electronics", 219.00),
    "Notebook": ("Office", 4.50),
    "Wireless Mouse": ("Electronics", 19.99),
    "Water Bottle": ("Home", 12.00),
}
REGIONS = ["North", "South", "East", "West"]
REGION_VARIANTS = {  # inconsistent labeling as it'd really appear
    "North": ["North", "north", "NORTH", "N."],
    "South": ["South", "south", "SOUTH"],
    "East": ["East", "east", "EAST"],
    "West": ["West", "west", " West"],
}
SALES_REPS = ["J. Patel", "M. Owens", "R. Diaz", "K. Chen"]

DATE_FORMATS = ["{m}/{d}/{y}", "{y}-{m:02d}-{d:02d}", "{d}-{m}-{y}", "{m}/{d}/{y}"]

def random_date_str():
    m, d, y = random.randint(1, 12), random.randint(1, 28), 2026
    fmt = random.choice(DATE_FORMATS)
    return fmt.format(m=m, d=d, y=y)

def messy_price(price):
    roll = random.random()
    if roll < 0.2:
        return f"${price:,.2f}"
    if roll < 0.3:
        return f"{price:.2f} "
    if roll < 0.35:
        return ""  # missing
    return str(price)

rows = []
for i in range(1, 181):
    product = random.choice(list(PRODUCTS.keys()))
    category, base_price = PRODUCTS[product]
    region_clean = random.choice(REGIONS)
    region = random.choice(REGION_VARIANTS[region_clean])
    customer = random.choice(CUSTOMERS)
    if random.random() < 0.1:
        customer = f"  {customer}  "  # stray whitespace
    cat_display = random.choice([category, category.lower(), category.upper()])
    qty = random.randint(1, 6)
    if random.random() < 0.05:
        qty = -qty  # return, looks like an error
    if random.random() < 0.04:
        qty = ""  # missing

    rows.append({
        "order_id": 1000 + i,
        "order_date": random_date_str() if random.random() > 0.03 else "",
        "customer_name": customer,
        "product": product,
        "category": cat_display,
        "quantity": qty,
        "unit_price": messy_price(base_price),
        "region": region,
        "sales_rep": random.choice(SALES_REPS) if random.random() > 0.05 else "",
    })

# inject a handful of exact duplicate rows, like a double-export glitch
for _ in range(6):
    rows.append(random.choice(rows[:50]).copy())

random.shuffle(rows)

with open("raw_sales_data.csv", "w", newline="") as f:
    writer = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
    writer.writeheader()
    writer.writerows(rows)

print(f"Generated {len(rows)} rows -> raw_sales_data.csv")
