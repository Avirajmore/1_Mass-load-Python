import pandas as pd
import os
import time

# ---------- CONFIG ----------
WORKING_DIR = os.path.expanduser("~/Downloads")
HASHI_FILE = os.path.expanduser("~/Downloads/Hashi lineitem.csv")
RENAMED_FILE = os.path.expanduser("~/Downloads/Source id file.csv")
OUTPUT_FILE = os.path.expanduser("~/Downloads/Hashi lineitem .csv")
WAIT_TIME = 3
# ----------------------------

# -------- STEP 1: WAIT & RENAME --------
print(f"⏳ Waiting {WAIT_TIME} seconds...")
time.sleep(WAIT_TIME)

bulk_files = [
    f for f in os.listdir(WORKING_DIR)
    if f.lower().endswith(".csv") and "bulkquery_result_" in f.lower()
]

if not bulk_files:
    raise FileNotFoundError("❌ No bulkQuery_result_ CSV file found")

latest_bulk_file = max(
    bulk_files,
    key=lambda f: os.path.getmtime(os.path.join(WORKING_DIR, f))
)

old_path = os.path.join(WORKING_DIR, latest_bulk_file)
new_path = os.path.join(WORKING_DIR, RENAMED_FILE)

os.rename(old_path, new_path)
print(f"✅ Renamed: {latest_bulk_file} → {RENAMED_FILE}")

# -------- STEP 2: READ FILES --------
source_df = pd.read_csv(RENAMED_FILE)
hashi_df = pd.read_csv(HASHI_FILE)

# -------- STEP 3: STANDARDIZE CASE --------
source_df["Source_ID__c_std"] = (
    source_df["Source_ID__c"]
    .astype(str)
    .str.strip()
    .str.upper()
)

hashi_df["SOURCE_ID__C_std"] = (
    hashi_df["SOURCE_ID__C"]
    .astype(str)
    .str.strip()
    .str.upper()
)

# -------- STEP 4: MERGE & COPY ID --------
merged_df = hashi_df.merge(
    source_df[["Source_ID__c_std", "Id"]],
    left_on="SOURCE_ID__C_std",
    right_on="Source_ID__c_std",
    how="left"
)

# Rename Id column to Opportunityid
merged_df.rename(columns={"Id": "Opportunityid"}, inplace=True)

# -------- STEP 5: FORMAT DATE --------
merged_df["EXPIRATION_DATE__C"] = pd.to_datetime(
    merged_df["EXPIRATION_DATE__C"],
    errors="coerce"
).dt.strftime("%Y-%m-%d")

# -------- CLEANUP --------
merged_df.drop(columns=["SOURCE_ID__C_std", "Source_ID__c_std"], inplace=True)

# -------- STEP 6: SAVE OUTPUT --------

os.remove(HASHI_FILE)
merged_df.to_csv(OUTPUT_FILE, index=False)

print("✅ Comparison completed successfully")
print(f"📄 Output file created: {OUTPUT_FILE}")
