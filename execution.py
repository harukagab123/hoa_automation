# execution.py
import os
import pandas as pd
from datetime import datetime
from calendar import monthrange

# === CONFIG ===
base_dir    = r"C:\Users\haruk\OneDrive\Desktop\Projects\hoa_automation"
emails_path = os.path.join(base_dir, "emails.csv")          # MUST have 'Association Name' + an email column
output_dir  = os.path.join(base_dir, "output")
output_csv  = os.path.join(output_dir, "letters_export.csv")

# === HELPERS ===
def coalesce(*vals):
    for v in vals:
        if v is None:
            continue
        s = str(v).strip()
        if s:
            return s
    return ""

def clean_street_number(val: object) -> str:
    """
    Fix 'Street #' values that come in as 123.0, 45.0, etc.
    If it's an integer-looking float, drop the .0; otherwise return as string.
    """
    s = coalesce(val)
    if not s:
        return ""
    if s.endswith(".0"):
        s = s[:-2]
    try:
        f = float(s)
        if f.is_integer():
            return str(int(f))
        return s
    except ValueError:
        return s

def build_full_address(row) -> str:
    """
    Build: "<Street #> <Address 1>, <City> <State> <Zip>"
    (comma after Address 1 only if both halves exist)
    """
    street = clean_street_number(row.get('Street #'))
    addr1  = coalesce(row.get('Address 1'))
    city   = coalesce(row.get('City'))
    state  = coalesce(row.get('State'))
    zipc   = coalesce(row.get('Zip Code'))

    first_line = " ".join(p for p in [street, addr1] if p)
    city_state_zip = " ".join(p for p in [city, state, zipc] if p)

    if first_line and city_state_zip:
        return f"{first_line}, {city_state_zip}"
    return first_line or city_state_zip

def pick_assoc_column(columns):
    # Be flexible if the CSV header varies
    for c in ["Association Name", "HOA Name", "Association"]:
        if c in columns:
            return c
    for c in columns:
        if 'assoc' in c.lower():
            return c
    return None

def find_email_column(columns):
    # Choose the first column whose name contains 'email' (case-insensitive)
    for c in columns:
        if 'email' in c.lower():
            return c
    return None

def normalize_key(s: str) -> str:
    # Normalize association keys for robust matching
    return " ".join(coalesce(s).lower().split())

def money(x) -> str:
    """Format numeric/str to $#,###.##; blank on errors."""
    try:
        val = float(str(x).replace(',', '').replace('$', ''))
        return f"${val:,.2f}"
    except Exception:
        return ""

# === FIND LATEST converted_*.csv IN base_dir ===
converted_files = [f for f in os.listdir(base_dir) if f.startswith("converted") and f.endswith(".csv")]
if not converted_files:
    raise SystemExit("No converted CSV files found in the directory.")
converted_files.sort(key=lambda f: os.path.getmtime(os.path.join(base_dir, f)), reverse=True)
latest_converted = os.path.join(base_dir, converted_files[0])
print(f"Opening: {latest_converted}")

# === LOAD MAIN DATA ===
df = pd.read_csv(latest_converted, dtype=str)

# Schema guards
required_cols = ['Balance', 'Address Type', 'Account #']
for col in required_cols:
    if col not in df.columns:
        raise SystemExit(f"CSV missing '{col}' column.")

assoc_col = pick_assoc_column(df.columns)
if assoc_col is None:
    raise SystemExit("Could not find an Association Name column in the converted CSV.")

# Normalize Balance & FILTERS
df['Balance'] = (
    df['Balance'].astype(str)
      .str.replace(r'[\$,]', '', regex=True)
      .replace('', '0')
      .astype(float)
)

# 1) Drop rows with Balance < 10.00
df = df[df['Balance'] >= 10.00].copy()

# 2) Keep only Property Address and Owner's Offsite Address
df = df[df['Address Type'].isin(['Property Address', "Owner's Offsite Address"])].copy()

# Sort & group
df = df.sort_values(by='Account #', ascending=True)
grouped = df.groupby('Account #', sort=True)

# === LOAD emails.csv for Association -> Email mapping (STRICT) ===
if not os.path.exists(emails_path):
    raise SystemExit(f"emails.csv not found: {emails_path}")

em_df = pd.read_csv(emails_path, dtype=str)
if 'Association Name' not in em_df.columns:
    raise SystemExit("emails.csv must include 'Association Name' column.")
email_col = find_email_column(em_df.columns)
if email_col is None:
    raise SystemExit("emails.csv must include an email column (name contains 'email').")

assoc_to_email = {}
for _, r in em_df.iterrows():
    assoc_key = normalize_key(r.get('Association Name'))
    email_val = coalesce(r.get(email_col))
    if assoc_key and email_val:
        assoc_to_email.setdefault(assoc_key, email_val)

# === DATES ===
today = datetime.now()
today_str = today.strftime("%B %d, %Y")
last_day = monthrange(today.year, today.month)[1]
last_day_of_month = today.replace(day=last_day).strftime("%B %d, %Y")

# === OUTPUT HEADERS (duplicates included) ===
headers = [
    "{{date}}",
    "{{ownersName}}",
    "{{propertyAddress}}",
    "{{associationName}}",
    "{{accNum}}",
    "{{propertyAddress}}",
    "{{last_day_of_month}}",
    "{{emailAddress}}",   # strictly from emails.csv by Association Name
    "{{accNum}}",
    "{{amount}}",
]

# === BUILD ROWS (NO DOC GENERATION, NO PER-ASSOCIATION FOLDERS) ===
rows_matrix = []

for acc_num, group in grouped:
    # Split by address type (kept so you can reuse for L1/L2 later if needed)
    prop_rows    = group[group['Address Type'] == 'Property Address']
    offsite_rows = group[group['Address Type'] == "Owner's Offsite Address"]

    # Choose a source row (prefer property)
    source_row = prop_rows.iloc[0] if len(prop_rows) else group.iloc[0]

    owners_name = f"{coalesce(source_row.get('First Name'))} {coalesce(source_row.get('Last Name'))}".strip()
    association = coalesce(source_row.get(assoc_col))
    acc_num_str = coalesce(acc_num)

    # Addresses
    property_address = build_full_address(prop_rows.iloc[0] if len(prop_rows) else source_row)

    # Email strictly from emails.csv by Association Name
    email_addr = assoc_to_email.get(normalize_key(association), "")

    # Amount (from the chosen source row’s Balance)
    amount_str = money(source_row.get('Balance'))

    # Append one ordered row matching headers exactly (including duplicates)
    rows_matrix.append([
        today_str,           # {{date}}
        owners_name,         # {{ownersName}}
        property_address,    # {{propertyAddress}}
        association,         # {{associationName}}
        acc_num_str,         # {{accNum}}
        property_address,    # {{propertyAddress}} (duplicate header)
        last_day_of_month,   # {{last_day_of_month}}
        email_addr,          # {{emailAddress}}
        acc_num_str,         # {{accNum}} (duplicate header)
        amount_str,          # {{amount}}
    ])

# === WRITE CSV ONLY ===
os.makedirs(output_dir, exist_ok=True)
final_df = pd.DataFrame(rows_matrix, columns=headers)  # pandas supports duplicate headers here
final_df.to_csv(output_csv, index=False, encoding="utf-8-sig")

print(f"Wrote CSV: {output_csv}")
print(f"Rows exported: {len(rows_matrix)}")
