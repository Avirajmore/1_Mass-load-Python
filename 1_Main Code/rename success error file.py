import os

# -------- CONFIG --------
FOLDER_PATH = "/Users/avirajmore/Downloads"   # <-- change this
# ------------------------

# -------- RENAME RULES --------
RENAME_RULES = {
    "opportunity insert": {
        "success": "opptysuccess",
        "error": "opptyerror"
    },
    "opportunity product insert": {
        "success": "productsuccess",
        "error": "producterror"
    },
    "feed item insert": {
        "success": "feedsuccess",
        "error": "feederror"
    },
    "opportunity strategy insert":{
        "success":"tagssuccess",
        "error":"tagserror"
    }
}
# -------------------------------

print("🔍 Scanning folder for CSV files...\n")

for file in os.listdir(FOLDER_PATH):

    if not file.lower().endswith(".csv"):
        continue

    file_lower = file.lower()
    new_name = None

    # Find matching rule
    for pattern, outcomes in RENAME_RULES.items():
        if pattern in file_lower:
            for status, base_name in outcomes.items():
                if status in file_lower:
                    new_name = f"{base_name}.csv"
                    break
        if new_name:
            break

    if new_name:
        old_path = os.path.join(FOLDER_PATH, file)
        new_path = os.path.join(FOLDER_PATH, new_name)

        # Prevent overwrite
        counter = 1
        base, ext = os.path.splitext(new_name)
        while os.path.exists(new_path):
            new_path = os.path.join(FOLDER_PATH, f"{base}_{counter}{ext}")
            counter += 1

        os.rename(old_path, new_path)
        print(f"✅ Renamed: {file} → {os.path.basename(new_path)}")

    else:
        print(f"⚠️ Skipped (no rule matched): {file}")

print("\n🎯 Renaming completed!")
