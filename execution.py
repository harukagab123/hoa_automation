import os
import pandas as pd
from datetime import datetime
from docxtpl import DocxTemplate
from calendar import monthrange

# === CONFIG ===
directory       = r"C:\Users\haruk\OneDrive\Desktop\Projects\hoa_automation"
template1_path  = os.path.join(directory, "template", "Letter 1.docx")  # {{...}} placeholders
template2_path  = os.path.join(directory, "template", "Letter 2.docx")  # {{...}} placeholders
emails_path     = os.path.join(directory, "emails.csv")     # map by Association Name -> Email
output_dir      = os.path.join(directory, "output")

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

def build_st_unit(row) -> str:
    street = clean_street_number(row.get('Street #'))
    addr1  = coalesce(row.get('Address 1'))
    return " ".join(p for p in [street, addr1] if p)

def build_city_state_zip(row) -> str:
    city  = coalesce(row.get('City'))
    state = coalesce(row.get('State'))
    zipc  = coalesce(row.get('Zip Code'))
    left = ", ".join([p for p in [city, state] if p])
    return " ".join([p for p in [left, zipc] if p])

def build_full_address(row) -> str:
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

def money(x) -> str:
    try:
        return f"${float(str(x).replace(',', '').replace('$', '')):,.2f}"
    except Exception:
        return ""

def pick_assoc_column(columns):
    candidates = ["Association Name", "HOA Name", "Association"]
    for c in candidates:
        if c in columns:
            return c
    for c in columns:
        if 'assoc' in c.lower():
            return c
    return None

def first_email_col(columns):
    for c in columns:
        if 'email' in c.lower():
            return c
    return None

# === FIND LATEST CSV ===
converted_files = [f for f in os.listdir(directory) if f.startswith("converted") and f.endswith(".csv")]
if not converted_files:
    raise SystemExit("No converted CSV files found in the directory.")

converted_files.sort(key=lambda f: os.path.getmtime(os.path.join(directory, f)), reverse=True)
latest_converted = os.path.join(directory, converted_files[0])
print(f"Opening: {latest_converted}")

# === LOAD DATA ===
df = pd.read_csv(latest_converted, dtype=str)
if 'Balance' not in df.columns:
    raise SystemExit("CSV missing 'Balance' column.")

df['Balance'] = (
    df['Balance'].astype(str)
    .str.replace(r'[\$,]', '', regex=True)
    .replace('', '0')
    .astype(float)
)

df = df[df['Balance'] >= 10.00].copy()

allowed_addr_types = ['Property Address', "Owner's Offsite Address"]
if 'Address Type' not in df.columns:
    raise SystemExit("CSV missing 'Address Type' column.")
df = df[df['Address Type'].isin(allowed_addr_types)].copy()

assoc_col = pick_assoc_column(df.columns)
if assoc_col is None:
    raise SystemExit("Could not find an Association Name column.")

if 'Account #' not in df.columns:
    raise SystemExit("CSV missing 'Account #' column.")
df = df.sort_values(by='Account #', ascending=True)

# === LOAD EMAILS ===
assoc_to_email = {}
if os.path.exists(emails_path):
    em_df = pd.read_csv(emails_path, dtype=str)
    if 'Association Name' not in em_df.columns:
        em_assoc_col = pick_assoc_column(em_df.columns)
    else:
        em_assoc_col = 'Association Name'
    em_email_col = first_email_col(em_df.columns)
    if em_assoc_col and em_email_col:
        for _, r in em_df.iterrows():
            assoc_key = coalesce(r.get(em_assoc_col)).lower()
            email_val = coalesce(r.get(em_email_col))
            if assoc_key and email_val:
                assoc_to_email[assoc_key] = email_val
    else:
        print("emails.csv missing required columns.")
else:
    print("emails.csv not found — emailAddress will be blank.")

# === BUILD EXPORTS ===
os.makedirs(output_dir, exist_ok=True)
date_folder = datetime.now().strftime("%m-%d_%y")
export_dir  = os.path.join(output_dir, date_folder)
os.makedirs(export_dir, exist_ok=True)

today = datetime.now()
today_str = today.strftime("%B %d, %Y")
last_day = monthrange(today.year, today.month)[1]
last_day_of_month = today.replace(day=last_day).strftime("%B %d, %Y")

grouped = df.groupby('Account #', sort=True)

headers_l1 = [
    "{{date}}","{{ownersName}}","{{propertyAddress}}","{{associationName}}","{{accNum}}",
    "{{last_day_of_month}}","{{emailAddress}}","{{amount}}",
    "{{propertyAddress_st_unit}}","{{propertyAddress_city_state_zip}}",
]

headers_l2 = [
    "{{date}}","{{ownersName}}","{{propertyAddress}}","{{associationName}}","{{accNum}}",
    "{{last_day_of_month}}","{{emailAddress}}","{{amount}}","{{ownersOffsiteAddress}}",
    "{{ownersOffsiteAddress_st_unit}}","{{ownersOffsiteAddress_city_state_zip}}",
]

rows_l1, rows_l2 = [], []

for acc_num, group in grouped:
    prop_rows    = group[group['Address Type'] == 'Property Address']
    offsite_rows = group[group['Address Type'] == "Owner's Offsite Address"]
    source_row = prop_rows.iloc[0] if len(prop_rows) else group.iloc[0]

    owners_name = f"{coalesce(source_row.get('First Name'))} {coalesce(source_row.get('Last Name'))}".strip()
    association = coalesce(source_row.get(assoc_col))
    amount_str  = money(source_row.get('Balance'))
    acc_num_str = coalesce(acc_num)

    property_address = build_full_address(prop_rows.iloc[0] if len(prop_rows) else source_row)
    offsite_address  = build_full_address(offsite_rows.iloc[0]) if len(offsite_rows) else ""

    prop_st_unit = build_st_unit(prop_rows.iloc[0]) if len(prop_rows) else ""
    prop_csz     = build_city_state_zip(prop_rows.iloc[0]) if len(prop_rows) else ""

    off_st_unit  = build_st_unit(offsite_rows.iloc[0]) if len(offsite_rows) else ""
    off_csz      = build_city_state_zip(offsite_rows.iloc[0]) if len(offsite_rows) else ""

    email_addr = assoc_to_email.get(association.lower(), "") if association else ""

    if len(prop_rows) and len(offsite_rows):
        rows_l2.append([
            today_str, owners_name, property_address, association, acc_num_str,
            last_day_of_month, email_addr, amount_str,
            offsite_address, off_st_unit, off_csz,
        ])
    else:
        rows_l1.append([
            today_str, owners_name, property_address, association, acc_num_str,
            last_day_of_month, email_addr, amount_str,
            prop_st_unit, prop_csz,
        ])

# Write only Letter1 and Letter2 CSVs
letter1_csv = os.path.join(export_dir, "letter1_exports.csv")
letter2_csv = os.path.join(export_dir, "letter2_exports.csv")

pd.DataFrame(rows_l1, columns=headers_l1).to_csv(letter1_csv, index=False, encoding="utf-8-sig")
pd.DataFrame(rows_l2, columns=headers_l2).to_csv(letter2_csv, index=False, encoding="utf-8-sig")

print(f"Wrote CSVs:\n - {letter1_csv}\n - {letter2_csv}")
print(f"Rows exported -> Letter1: {len(rows_l1)} | Letter2: {len(rows_l2)}")
