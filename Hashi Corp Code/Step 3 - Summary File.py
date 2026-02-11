import pandas as pd
import os
import shutil
from datetime import date
from openpyxl import load_workbook

# ---------- CONFIG ----------
CSV_DIR = os.path.expanduser("~/Downloads")   # folder containing CSV files
TEMPLATE_FILE = os.path.expanduser("~/Documents/Office Docs/Massload Files/Reference File/Hashicorp_production_load_summary_file.xlsx")
# ----------------------------

today = date.today().strftime("%Y-%m-%d")
OUTPUT_FILE = os.path.expanduser(f"~/Downloads/Hashi_production_load_summary_{today}.xlsx")

# ---------- COPY TEMPLATE ----------
shutil.copy(TEMPLATE_FILE, OUTPUT_FILE)
print(f"📄 Template copied → {OUTPUT_FILE}")

# ---------- FILE DEFINITIONS ----------
FILES = {
    "opp_success": {
        "sheet": "Opportunity success",
        "keywords": ["opportunity", "upsert", "success"]
    },
    "opp_error": {
        "sheet": "Opportunity failures",
        "keywords": ["opportunity", "upsert", "error"]
    },
    "prod_success": {
        "sheet": "Product success",
        "keywords": ["opportunity", "product", "upsert", "success"]
    },
    "prod_error": {
        "sheet": "Product Failures",
        "keywords": ["opportunity", "product", "upsert", "error"]
    }
}

SUMMARY_CELLS = {
    "opp_success": "E5",
    "opp_error": "F5",
    "opp_total": "G5",
    "prod_success": "E6",
    "prod_error": "F6",
    "prod_total": "G6"
}


def find_csv(keywords):
    for file in os.listdir(CSV_DIR):
        name = file.lower()
        if name.endswith(".csv") and all(k in name for k in keywords):
            return os.path.join(CSV_DIR, file)
    return None


row_counts = {}

# ---------- WRITE DATA TO COPIED FILE ----------
with pd.ExcelWriter(
    OUTPUT_FILE,
    engine="openpyxl",
    mode="a",
    if_sheet_exists="replace"
) as writer:

    for key, info in FILES.items():
        csv_path = find_csv(info["keywords"])

        if csv_path:
            df = pd.read_csv(csv_path)
            df.to_excel(writer, sheet_name=info["sheet"], index=False)
            row_counts[key] = len(df)
            print(f"✅ {info['sheet']} → {len(df)} rows")
        else:
            # Empty sheet if missing
            pd.DataFrame().to_excel(writer, sheet_name=info["sheet"], index=False)
            row_counts[key] = 0
            print(f"⚠️ {info['sheet']} missing → count 0")

# ---------- UPDATE SUMMARY ----------
workbook = load_workbook(OUTPUT_FILE)
summary = workbook["Summary"]

summary[SUMMARY_CELLS["opp_success"]] = row_counts["opp_success"]
summary[SUMMARY_CELLS["opp_error"]] = row_counts["opp_error"]
summary[SUMMARY_CELLS["prod_success"]] = row_counts["prod_success"]
summary[SUMMARY_CELLS["prod_error"]] = row_counts["prod_error"]

summary[SUMMARY_CELLS["opp_total"]] = (
    row_counts["opp_success"] + row_counts["opp_error"]
)
summary[SUMMARY_CELLS["prod_total"]] = (
    row_counts["prod_success"] + row_counts["prod_error"]
)

workbook.save(OUTPUT_FILE)

print("\n🎯 Summary file generated successfully")
