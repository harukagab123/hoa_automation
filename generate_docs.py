# generate_docs.py
import os
import re
from datetime import datetime
import pandas as pd
from docxtpl import DocxTemplate

# === CONFIG ===
BASE_DIR        = r"C:\Users\haruk\OneDrive\Desktop\Projects\hoa_automation"
INPUT_CSV       = os.path.join(BASE_DIR, "output", "letters_export.csv")
TEMPLATE1_PATH  = os.path.join(BASE_DIR, "template", "Letter 1.docx")  # placeholders: see CONTEXT below
TEMPLATE2_PATH  = os.path.join(BASE_DIR, "template", "Letter 2.docx")  # adds ownersOffsiteAddress
OUTPUT_DIR      = os.path.join(BASE_DIR, "output")

# === HELPERS ===
def clean_value(v: object) -> str:
    """Normalize CSV cell values: treat NaN/None/whitespace as empty; return stripped string."""
    if v is None:
        return ""
    s = str(v).strip()
    return "" if s.lower() in ("nan", "none") else s

def parse_month_from_date(date_str: str) -> str:
    """
    'September 17, 2025' -> 'September'
    Fallback to current month if parsing fails.
    """
    try:
        dt = datetime.strptime(date_str, "%B %d, %Y")
        return dt.strftime("%B")
    except Exception:
        return datetime.now().strftime("%B")

def resolve_cols(columns, base_name: str):
    """
    Given DataFrame columns and a base header (e.g., '{{accNum}}'),
    return all matching column names in a stable order:
      ['{{accNum}}', '{{accNum}}.1', '{{accNum}}.2', ...]
    """
    matches = [c for c in columns if c == base_name or c.startswith(base_name + ".")]
    # ensure base comes first, then suffixes in natural order
    return sorted(matches, key=lambda c: (0, 0) if c == base_name else (1, c))

def get_first_nonempty(row, col_list):
    """Return first non-empty value across possible duplicate/suffixed columns."""
    for col in col_list:
        if col in row:
            val = clean_value(row[col])
            if val:
                return val
    return ""

def safe_folder_name(name: str) -> str:
    """Sanitize folder/file parts to avoid invalid path characters on Windows."""
    name = clean_value(name) or "Unknown"
    return re.sub(r'[<>:"/\\|?*]+', "_", name)

# === CHECK INPUTS ===
if not os.path.exists(INPUT_CSV):
    raise SystemExit(f"CSV not found: {INPUT_CSV}")

template1_exists = os.path.exists(TEMPLATE1_PATH)
template2_exists = os.path.exists(TEMPLATE2_PATH)
if not template1_exists and not template2_exists:
    raise SystemExit(
        "No templates found.\n"
        f"Missing: {TEMPLATE1_PATH}\n"
        f"Missing: {TEMPLATE2_PATH}"
    )
if not template1_exists:
    print(f"WARNING: Missing template: {TEMPLATE1_PATH} — Letter 1 docs will be skipped.")
if not template2_exists:
    print(f"WARNING: Missing template: {TEMPLATE2_PATH} — Letter 2 docs will be skipped.")

# === READ CSV WITH PANDAS (handles duplicate headers by auto-suffixing .1, .2, ...) ===
df = pd.read_csv(INPUT_CSV, dtype=str, encoding="utf-8-sig")

# Pre-resolve columns for all fields we care about
wanted_headers = [
    "{{date}}",
    "{{ownersName}}",
    "{{propertyAddress}}",
    "{{associationName}}",
    "{{accNum}}",
    "{{last_day_of_month}}",
    "{{emailAddress}}",
    "{{amount}}",
    # Optional (if present, we’ll use Letter 2):
    "{{ownersOffsiteAddress}}",
]
colmap = {h: resolve_cols(df.columns, h) for h in wanted_headers}

# === MAIN ===
os.makedirs(OUTPUT_DIR, exist_ok=True)
generated = 0

for _, row in df.iterrows():
    # Extract fields (first non-empty across duplicates/suffixes)
    date_str          = get_first_nonempty(row, colmap["{{date}}"]) or datetime.now().strftime("%B %d, %Y")
    owners_name       = get_first_nonempty(row, colmap["{{ownersName}}"])
    property_address  = get_first_nonempty(row, colmap["{{propertyAddress}}"])
    association_name  = get_first_nonempty(row, colmap["{{associationName}}"])
    acc_num           = get_first_nonempty(row, colmap["{{accNum}}"])
    last_day_of_month = get_first_nonempty(row, colmap["{{last_day_of_month}}"])
    email_address     = get_first_nonempty(row, colmap["{{emailAddress}}"])
    amount_str        = get_first_nonempty(row, colmap["{{amount}}"])
    offsite_address   = get_first_nonempty(row, colmap.get("{{ownersOffsiteAddress}}", []))

    month_str = parse_month_from_date(date_str)
    acc_num_safe = safe_folder_name(acc_num or "Unknown_Account")
    assoc_safe   = safe_folder_name(association_name or "Unknown_Association")

    # Output folder per association
    assoc_folder = os.path.join(OUTPUT_DIR, assoc_safe)
    os.makedirs(assoc_folder, exist_ok=True)

    # Context for templates (placeholders WITHOUT braces)
    base_context = {
        "date": date_str,
        "ownersName": owners_name,
        "propertyAddress": property_address,
        "associationName": association_name,
        "accNum": acc_num,
        "last_day_of_month": last_day_of_month,
        "emailAddress": email_address,
        "amount": amount_str,
    }

    # Pick template:
    # - If {{ownersOffsiteAddress}} exists & non-empty AND Letter 2 template exists -> Letter 2
    # - Else -> Letter 1 (if available)
    use_letter2 = bool(offsite_address) and template2_exists

    if use_letter2:
        context = {**base_context, "ownersOffsiteAddress": offsite_address}
        tpl_path = TEMPLATE2_PATH
        out_name = f"{acc_num_safe} - {month_str} - Letter 2.docx"
    else:
        if not template1_exists:
            print(f"SKIPPED (no Letter 1 template): Account #{acc_num_safe} ({assoc_safe})")
            continue
        context = base_context
        tpl_path = TEMPLATE1_PATH
        out_name = f"{acc_num_safe} - {month_str} - Letter 1.docx"

    # Render and save
    try:
        doc = DocxTemplate(tpl_path)
        doc.render(context)
        out_path = os.path.join(assoc_folder, out_name)
        doc.save(out_path)
        print(f"Generated: {out_path}")
        generated += 1
    except Exception as e:
        print(f"ERROR generating for Account #{acc_num_safe} ({assoc_safe}): {e}")

print(f"\nDone. Generated {generated} document(s).")
