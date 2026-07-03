"""
Day 1: Clean the raw messy sales export.

Each step below fixes one specific problem. Run this and read the
printed before/after counts — that's the part you'll explain to a
client when you pitch this kind of project.
"""
import pandas as pd

RAW_PATH = "raw_sales_data.csv"
CLEAN_PATH = "clean_sales_data.csv"

df = pd.read_csv(RAW_PATH, dtype=str)  # read everything as text first, we control conversion ourselves
print(f"Loaded {len(df)} raw rows")

# --- Step 1: strip stray whitespace from every text column ---
# "  Maria Santos  " -> "Maria Santos"
for col in df.select_dtypes(include=["object", "string"]).columns:
    df[col] = df[col].str.strip()

# --- Step 2: standardize text casing (category, region) ---
# "electronics" / "ELECTRONICS" / "Electronics" -> "Electronics"
df["category"] = df["category"].str.title()
df["region"] = df["region"].str.title()
# fix the "N." abbreviation we saw in region
df["region"] = df["region"].replace({"N.": "North"})

# --- Step 3: parse mixed date formats into one real date type ---
# Source data mixes M/D/Y, Y-M-D, and D-M-Y. We try each format in turn.
def parse_messy_date(value):
    if pd.isna(value) or value == "":
        return pd.NaT
    for fmt in ("%m/%d/%Y", "%Y-%m-%d", "%d-%m-%Y"):
        try:
            return pd.to_datetime(value, format=fmt)
        except ValueError:
            continue
    return pd.NaT  # couldn't parse -> flag as missing rather than guess

df["order_date"] = df["order_date"].apply(parse_messy_date)

# --- Step 4: clean unit_price (strip $ and commas, convert to float) ---
df["unit_price"] = (
    df["unit_price"]
    .str.replace("$", "", regex=False)
    .str.replace(",", "", regex=False)
    .str.strip()
    .replace("", pd.NA)
    .astype(float)
)

# --- Step 5: clean quantity (convert to numeric, flag negatives as returns) ---
df["quantity"] = pd.to_numeric(df["quantity"], errors="coerce")
df["is_return"] = df["quantity"] < 0
df["quantity"] = df["quantity"].abs()  # keep magnitude, return status now tracked separately

# --- Step 6: remove exact duplicate rows (the double-export glitch) ---
before = len(df)
df = df.drop_duplicates()
print(f"Removed {before - len(df)} duplicate rows")

# --- Step 7: handle remaining missing values ---
missing_dates = df["order_date"].isna().sum()
missing_prices = df["unit_price"].isna().sum()
missing_qty = df["quantity"].isna().sum()
missing_reps = (df["sales_rep"] == "").sum() if "sales_rep" in df else 0
print(f"Missing after cleaning -> dates: {missing_dates}, prices: {missing_prices}, "
      f"quantities: {missing_qty}, sales reps: {missing_reps}")

# Drop rows where we genuinely can't compute revenue (no price or qty or date)
# but keep a record of how many we dropped, since a client will ask.
before = len(df)
df = df.dropna(subset=["order_date", "unit_price", "quantity"])
print(f"Dropped {before - len(df)} rows with unrecoverable missing data")

# Fill missing sales rep with "Unassigned" rather than dropping the row
df["sales_rep"] = df["sales_rep"].replace("", "Unassigned")

# --- Step 8: add a computed revenue column for later analysis ---
df["revenue"] = df["quantity"] * df["unit_price"]

# --- Step 9: sort and save ---
df = df.sort_values("order_date").reset_index(drop=True)
df.to_csv(CLEAN_PATH, index=False)
print(f"\nSaved {len(df)} clean rows -> {CLEAN_PATH}")
print(df.head())
