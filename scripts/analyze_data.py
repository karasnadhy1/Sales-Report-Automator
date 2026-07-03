"""
Day 2: Turn clean_sales_data.csv into the metrics a business owner
actually wants to see. This is the "so what" layer — clean data on
its own doesn't sell a freelance gig, insight does.
"""
import pandas as pd

CLEAN_PATH = "clean_sales_data.csv"

df = pd.read_csv(CLEAN_PATH, parse_dates=["order_date"])
print(f"Loaded {len(df)} clean rows spanning {df['order_date'].min().date()} to {df['order_date'].max().date()}\n")

# Exclude returns from revenue metrics — a return isn't a sale
sales = df[~df["is_return"]].copy()

# --- Metric 1: Total revenue ---
total_revenue = sales["revenue"].sum()
print(f"Total revenue: ${total_revenue:,.2f}")

# --- Metric 2: Revenue by category ---
by_category = (
    sales.groupby("category")["revenue"]
    .sum()
    .sort_values(ascending=False)
    .round(2)
)
print("\nRevenue by category:")
print(by_category)

# --- Metric 3: Revenue by region ---
by_region = (
    sales.groupby("region")["revenue"]
    .sum()
    .sort_values(ascending=False)
    .round(2)
)
print("\nRevenue by region:")
print(by_region)

# --- Metric 4: Top 5 products by revenue ---
top_products = (
    sales.groupby("product")["revenue"]
    .sum()
    .sort_values(ascending=False)
    .head(5)
    .round(2)
)
print("\nTop 5 products by revenue:")
print(top_products)

# --- Metric 5: Top sales reps by revenue ---
by_rep = (
    sales.groupby("sales_rep")["revenue"]
    .sum()
    .sort_values(ascending=False)
    .round(2)
)
print("\nRevenue by sales rep:")
print(by_rep)

# --- Metric 6: Month-over-month revenue trend ---
sales["month"] = sales["order_date"].dt.to_period("M")
monthly = sales.groupby("month")["revenue"].sum().round(2)
monthly_growth = monthly.pct_change().round(3) * 100
print("\nMonthly revenue:")
print(monthly)
print("\nMonth-over-month growth (%):")
print(monthly_growth)

# --- Metric 7: Returns summary ---
returns = df[df["is_return"]]
print(f"\nReturns: {len(returns)} orders, ${returns['revenue'].sum():,.2f} in returned value")

# --- Save a summary CSV for use in the Excel report on Day 3 ---
summary_rows = []
for cat, val in by_category.items():
    summary_rows.append({"metric": "revenue_by_category", "label": cat, "value": val})
for reg, val in by_region.items():
    summary_rows.append({"metric": "revenue_by_region", "label": reg, "value": val})
for prod, val in top_products.items():
    summary_rows.append({"metric": "top_product", "label": prod, "value": val})
for rep, val in by_rep.items():
    summary_rows.append({"metric": "revenue_by_rep", "label": rep, "value": val})
for month, val in monthly.items():
    summary_rows.append({"metric": "monthly_revenue", "label": str(month), "value": val})

summary_df = pd.DataFrame(summary_rows)
summary_df.to_csv("summary_metrics.csv", index=False)
print(f"\nSaved {len(summary_df)} summary metrics -> summary_metrics.csv")
