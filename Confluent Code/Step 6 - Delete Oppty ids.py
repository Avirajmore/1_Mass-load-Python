import os
from datetime import datetime
import pandas as pd
import pyperclip
import shutil


# ============================
# UI
# ============================
def show_title(title):
    line_width = 100
    line = "=" * line_width
    print(f"\n{line}")
    print(title.center(line_width))
    print(f"{line}\n")


# ============================
# CHUNKING FUNCTIONS
# ============================
def chunk_ids(ids, max_chars=90000, base_query_len=200):
    chunks = []
    current_chunk = []
    current_len = base_query_len

    for val in ids:
        val_str = f"'{val}'"
        if current_len + len(val_str) + 1 > max_chars:
            chunks.append(current_chunk)
            current_chunk = [val]
            current_len = base_query_len + len(val_str)
        else:
            current_chunk.append(val)
            current_len += len(val_str) + 1

    if current_chunk:
        chunks.append(current_chunk)

    return chunks


def build_queries(chunks, field_name, object_name):
    queries = []
    for chunk in chunks:
        formatted = ",".join(f"'{v}'" for v in chunk)
        query = f"""
SELECT Id,{field_name}
FROM {object_name}
WHERE {field_name} IN ({formatted})
"""
        queries.append(query.strip())
    return queries


def rename_latest_file(part_name):
    DOWNLOAD_FOLDER = os.path.expanduser("~/Downloads")

    matching_files = [
        f for f in os.listdir(DOWNLOAD_FOLDER)
        if f.lower().endswith(".csv") and "bulkquery_result_" in f.lower()
    ]

    if not matching_files:
        print("❌ No matching bulkQuery_result_ CSV file found.")
        return None

    latest_file = max(
        matching_files,
        key=lambda f: os.path.getmtime(os.path.join(DOWNLOAD_FOLDER, f))
    )

    old_path = os.path.join(DOWNLOAD_FOLDER, latest_file)
    new_path = os.path.join(DOWNLOAD_FOLDER, part_name)
    os.rename(old_path, new_path)

    return new_path


def merge_files(prefix, final_name):
    folder = os.path.expanduser("~/Downloads")

    files = sorted([
        f for f in os.listdir(folder)
        if f.startswith(prefix) and f.endswith(".csv")
    ])

    if not files:
        print(f"❌ No files found for {prefix}")
        return

    df_list = [pd.read_csv(os.path.join(folder, f)) for f in files]
    final_df = pd.concat(df_list, ignore_index=True)

    final_df.to_csv(os.path.join(folder, final_name), index=False)

    print(f"✅ Merged file created: {final_name}")


# ============================
# START
# ============================
title = "📂 Create delete Files 📂"
show_title(title)

DOWNLOAD_FOLDER = os.path.expanduser("~/Downloads")

# ---------- PRODUCT ----------
product_df = pd.read_csv(os.path.expanduser("~/Downloads/Confluent Load/Duplicate Files/product_Record_Mismatch.csv"))

product_delete = (
    product_df[product_df["_merge"] == "Not_in_ISCED"][["Id"]]
    .rename(columns={"Id": "Delete Product"})
)

product_delete.to_csv(os.path.expanduser("~/Downloads/Confluent Load/Duplicate Files/PreDelete_Product.csv"), index=False)


# ---------- OPPORTUNITY ----------
oppty_df = pd.read_csv(os.path.expanduser("~/Downloads/Confluent Load/Duplicate Files/oppty_Record_Mismatch.csv"))

oppty_delete = (
    oppty_df[oppty_df["_merge"] == "Not_in_ISCED"][["Id"]]
    .rename(columns={"Id": "Delete Oppty"})
)

oppty_delete.to_csv(os.path.expanduser("~/Downloads/Confluent Load/Duplicate Files/PreDelete_Oppty.csv"), index=False)


oppty_file_path = os.path.expanduser("~/Downloads/Confluent Load/Duplicate Files/PreDelete_Oppty.csv")
product_file_path = os.path.expanduser("~/Downloads/Confluent Load/Duplicate Files/PreDelete_Product.csv")


# ============================
# OPPORTUNITY CHUNK PROCESS
# ============================
df = pd.read_csv(oppty_file_path)
values = df["Delete Oppty"].dropna().astype(str).tolist()

chunks = chunk_ids(values)
queries = build_queries(chunks, "Source_ID__c", "Opportunity")

for i, query in enumerate(queries, 1):
    pyperclip.copy(query)
    print(f"\n✅ Oppty Query {i}/{len(queries)} copied.")

    input("Run in Workbench → Download → Press Enter...")

    rename_latest_file(f"DELETE OPPTY_part{i}.csv")

merge_files("DELETE OPPTY_part", "DELETE OPPTY.csv")


# ============================
# PRODUCT CHUNK PROCESS
# ============================
df = pd.read_csv(product_file_path)
values = df["Delete Product"].dropna().astype(str).tolist()

chunks = chunk_ids(values)
queries = build_queries(chunks, "Lineitem_Legacy_Id__c", "OpportunityLineitem")

for i, query in enumerate(queries, 1):
    pyperclip.copy(query)
    print(f"\n✅ Product Query {i}/{len(queries)} copied.")

    input("Run in Workbench → Download → Press Enter...")

    rename_latest_file(f"DELETE PRODUCT_part{i}.csv")

merge_files("DELETE PRODUCT_part", "DELETE PRODUCT.csv")


# ============================
# MOVE FINAL FILES
# ============================
shutil.move(
    os.path.expanduser("~/Downloads/DELETE OPPTY.csv"),
    os.path.expanduser("~/Downloads/Confluent Load/Main Files/DELETE OPPTY.csv")
)

shutil.move(
    os.path.expanduser("~/Downloads/DELETE PRODUCT.csv"),
    os.path.expanduser("~/Downloads/Confluent Load/Main Files/DELETE PRODUCT.csv")
)


# ============================
# SUMMARY FILE UPDATE
# ============================
summary_folder = os.path.expanduser("~/Downloads/Confluent Load")

print(f"\n\n SUMMARY FOLDER:-{summary_folder}\n\n")

oppty_file = os.path.expanduser("~/Downloads/Confluent Load/Main Files/DELETE OPPTY.csv")
product_file = os.path.expanduser("~/Downloads/Confluent Load/Main Files/DELETE PRODUCT.csv")

summary_file = os.path.join(summary_folder, "SUMMARY FILE - CONFLUENT PROD.xlsx")

if not os.path.exists(summary_file):
    print("❌ Summary file not found!")
    exit()

df_oppty = pd.read_csv(oppty_file)
df_product = pd.read_csv(product_file)

with pd.ExcelWriter(summary_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    df_oppty.to_excel(writer, sheet_name="DELETE OPPTY", index=False)
    df_product.to_excel(writer, sheet_name="DELETE PRODUCT", index=False)


# ============================
# RENAME SUMMARY FILE
# ============================
today_date = datetime.today().strftime("%Y-%m-%d")
name, ext = os.path.splitext("SUMMARY FILE - CONFLUENT PROD.xlsx")
new_filename = f"{name} ({today_date}){ext}"

os.rename(
    os.path.join(summary_folder, "SUMMARY FILE - CONFLUENT PROD.xlsx"),
    os.path.join(summary_folder, new_filename)
)


# ============================
# RENAME FOLDER
# ============================
def get_day_suffix(day):
    if 11 <= day <= 13:
        return "th"
    return {1: "st", 2: "nd", 3: "rd"}.get(day % 10, "th")


now = datetime.now()
day = now.day
month = now.strftime("%B")

suffix = get_day_suffix(day)

new_folder_name = f"Confluent Load ({day}{suffix} {month})"

base_path = os.path.expanduser("~/Downloads")

old_path = os.path.join(base_path, "Confluent Load")
new_path = os.path.join(base_path, new_folder_name)

if os.path.exists(old_path):
    os.rename(old_path, new_path)
else:
    print("\n❌ Folder not found!")


# ============================
# END
# ============================
title = "📂 Confluent Load Done 📂"
show_title(title)