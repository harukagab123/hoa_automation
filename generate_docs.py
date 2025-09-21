# generate_docs.py
import os
import re
from datetime import datetime
import pandas as pd
from docxtpl import DocxTemplate

# === CONFIG ===
BASE_DIR        = r"C:\Users\haruk\OneDrive\Desktop\Projects\hoa_automation"
OUTPUT_DIR      = os.path.join(BASE_DIR, "output")
# Primary expected filenames (per your request)
INPUT_L1_NAME   = "letter1_exports.csv"
INPUT_L2_NAME   = "letter2_exports.csv"  # if you still use "letter_exports.csv", see fallback below

TEMPLATE1_PATH  = os.path.join(BASE_DIR, "template", "Letter 1.docx")  # placeholders listed in CONTEXT below
TEMPLATE2_PATH  = os.path.join(BASE_DIR, "template", "Letter 2.docx")  # includes ownersOffsiteAddress

# === HELPERS ===
def clean_value(v: object) -> str:
    """Normalize CSV cell values: treat NaN/None/whitespace as empty; return stripped string."""
    if v is None:
        return ""
    s = str(v).strip()
    return "" if s.lower() in ("nan", "none") else s

def parse_month_from_date(date_str: str) -> str:
    """'September 17, 2025' -> 'September'; fallback to current month if parsing fails."""
    try:
        dt = datetime.strptime(date_str, "%B %d, %Y")
        return dt.strftime("%B")
    except Exception:
        return datetime.now().strftime("%B")

def safe_folder_name(name: str) -> str:
    """Sanitize folder/file parts to avoid invalid path characters on Windows."""
    name = clean_value(name) or "Unknown"
    return re.sub(r'[<>:\"/\\|?*]+', "_", name)

def get_cell(row: pd.Series, col_name: str) -> str:
    """Safe cell accessor that returns '' if the column is missing."""
    try:
        return clean_value(row.get(col_name))
    except Exception:
        return ""

def find_input_csv(primary_path: str, basename: str, fallback_basename: str | None = None) -> str:
    """
    Find CSV in OUTPUT_DIR. Try primary_path first.
    If missing, search the most recent subfolder under OUTPUT_DIR that contains `basename`.
    Optionally try a fallback_basename (e.g., 'letter_exports.csv' for legacy).
    """
    if os.path.exists(primary_path):
        return primary_path

    # Try dated subfolders (pick the most recently modified one that contains the file)
    candidates = []
    if os.path.isdir(OUTPUT_DIR):
        for entry in os.listdir(OUTPUT_DIR):
            sub = os.path.join(OUTPUT_DIR, entry)
            if os.path.isdir(sub):
                cand = os.path.join(sub, basename)
                if os.path.exists(cand):
                    candidates.append((os.path.getmtime(cand), cand))
                if fallback_basename:
                    cand_fb = os.path.join(sub, fallback_basename)
                    if os.path.exists(cand_fb):
                        candidates.append((os.path.getmtime(cand_fb), cand_fb))
    if candidates:
        candidates.sort(reverse=True)
        return candidates[0][1]

    # Last resort: try OUTPUT_DIR/fallback_basename at root
    if fallback_basename:
        fb_root = os.path.join(OUTPUT_DIR, fallback_basename)
        if os.path.exists(fb_root):
            return fb_root

    raise SystemExit(f"CSV not found: expected '{basename}' under {OUTPUT_DIR} (checked subfolders too).")

# === CHECK TEMPLATES ===
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

# === LOCATE INPUT CSVs ===
INPUT_L1_PATH = find_input_csv(
    primary_path=os.path.join(OUTPUT_DIR, INPUT_L1_NAME),
    basename=INPUT_L1_NAME
)
# For Letter 2, also accept legacy filename 'letter_exports.csv' as a fallback.
INPUT_L2_PATH = find_input_csv(
    primary_path=os.path.join(OUTPUT_DIR, INPUT_L2_NAME),
    basename=INPUT_L2_NAME,
    fallback_basename="letter_exports.csv"  # remove this if you never use the old name
)

# === READ CSVs ===
# Letter 1 schema (now includes split address fields):
# {{date}}, {{ownersName}}, {{propertyAddress}}, {{associationName}}, {{accNum}},
# {{last_day_of_month}}, {{emailAddress}}, {{amount}},
# {{propertyAddress_st_unit}}, {{propertyAddress_city_state_zip}}
df_l1 = pd.read_csv(INPUT_L1_PATH, dtype=str, encoding="utf-8-sig").fillna("")

# Letter 2 schema (now includes offsite split fields):
# all of the above + {{ownersOffsiteAddress}} + {{ownersOffsiteAddress_st_unit}}, {{ownersOffsiteAddress_city_state_zip}}
df_l2 = pd.read_csv(INPUT_L2_PATH, dtype=str, encoding="utf-8-sig").fillna("")

# === ENSURE OUTPUT DIR EXISTS ===
os.makedirs(OUTPUT_DIR, exist_ok=True)

generated_l1 = 0
generated_l2 = 0

# === GENERATE LETTER 1 DOCS ===
if template1_exists:
    for _, row in df_l1.iterrows():
        date_str          = get_cell(row, "{{date}}")
        owners_name       = get_cell(row, "{{ownersName}}")
        property_address  = get_cell(row, "{{propertyAddress}}")
        association_name  = get_cell(row, "{{associationName}}")
        acc_num           = get_cell(row, "{{accNum}}")
        last_day_of_month = get_cell(row, "{{last_day_of_month}}")
        email_address     = get_cell(row, "{{emailAddress}}")
        amount_str        = get_cell(row, "{{amount}}")

        # NEW fields for Letter 1:
        prop_st_unit      = get_cell(row, "{{propertyAddress_st_unit}}")
        prop_city_state_zip = get_cell(row, "{{propertyAddress_city_state_zip}}")

        if not date_str:
            date_str = datetime.now().strftime("%B %d, %Y")
        month_str   = parse_month_from_date(date_str)
        acc_safe    = safe_folder_name(acc_num or "Unknown_Account")
        assoc_safe  = safe_folder_name(association_name or "Unknown_Association")

        assoc_folder = os.path.join(OUTPUT_DIR, assoc_safe)
        os.makedirs(assoc_folder, exist_ok=True)

        context = {
            "date": date_str,
            "ownersName": owners_name,
            "propertyAddress": property_address,
            "associationName": association_name,
            "accNum": acc_num,
            "last_day_of_month": last_day_of_month,
            "emailAddress": email_address,
            "amount": amount_str,
            # NEW in context:
            "propertyAddress_st_unit": prop_st_unit,
            "propertyAddress_city_state_zip": prop_city_state_zip,
        }

        try:
            tpl = DocxTemplate(TEMPLATE1_PATH)
            out_name = f"{acc_safe} - {month_str} - Letter 1.docx"
            tpl.render(context)
            tpl.save(os.path.join(assoc_folder, out_name))
            print(f"Generated (L1): {os.path.join(assoc_folder, out_name)}")
            generated_l1 += 1
        except Exception as e:
            print(f"ERROR (L1) for Account #{acc_safe} ({assoc_safe}): {e}")
else:
    print("Skipping Letter 1 generation (template missing).")

# === GENERATE LETTER 2 DOCS ===
if template2_exists:
    for _, row in df_l2.iterrows():
        date_str          = get_cell(row, "{{date}}")
        owners_name       = get_cell(row, "{{ownersName}}")
        property_address  = get_cell(row, "{{propertyAddress}}")
        association_name  = get_cell(row, "{{associationName}}")
        acc_num           = get_cell(row, "{{accNum}}")
        last_day_of_month = get_cell(row, "{{last_day_of_month}}")
        email_address     = get_cell(row, "{{emailAddress}}")
        amount_str        = get_cell(row, "{{amount}}")
        offsite_address   = get_cell(row, "{{ownersOffsiteAddress}}")

        # NEW fields for Letter 2:
        off_st_unit         = get_cell(row, "{{ownersOffsiteAddress_st_unit}}")
        off_city_state_zip  = get_cell(row, "{{ownersOffsiteAddress_city_state_zip}}")

        if not date_str:
            date_str = datetime.now().strftime("%B %d, %Y")
        month_str   = parse_month_from_date(date_str)
        acc_safe    = safe_folder_name(acc_num or "Unknown_Account")
        assoc_safe  = safe_folder_name(association_name or "Unknown_Association")

        assoc_folder = os.path.join(OUTPUT_DIR, assoc_safe)
        os.makedirs(assoc_folder, exist_ok=True)

        context = {
            "date": date_str,
            "ownersName": owners_name,
            "propertyAddress": property_address,
            "associationName": association_name,
            "accNum": acc_num,
            "last_day_of_month": last_day_of_month,
            "emailAddress": email_address,
            "amount": amount_str,
            "ownersOffsiteAddress": offsite_address,
            # NEW in context:
            "ownersOffsiteAddress_st_unit": off_st_unit,
            "ownersOffsiteAddress_city_state_zip": off_city_state_zip,
        }

        try:
            tpl = DocxTemplate(TEMPLATE2_PATH)
            out_name = f"{acc_safe} - {month_str} - Letter 2.docx"
            tpl.render(context)
            tpl.save(os.path.join(assoc_folder, out_name))
            print(f"Generated (L2): {os.path.join(assoc_folder, out_name)}")
            generated_l2 += 1
        except Exception as e:
            print(f"ERROR (L2) for Account #{acc_safe} ({assoc_safe}): {e}")
else:
    print("Skipping Letter 2 generation (template missing).")

print(f"\nDone. Generated documents — Letter 1: {generated_l1} | Letter 2: {generated_l2}")
