import os
import sys
import shutil
from datetime import datetime, date

import pandas as pd
import pyperclip
from openpyxl import load_workbook


DOWNLOAD_FOLDER = os.path.expanduser("~/Downloads")
BASE_FOLDER = os.path.join(DOWNLOAD_FOLDER, "Hashi Load")
MAIN_FILES_FOLDER = os.path.join(BASE_FOLDER, "Main Files")
UNIMPORTANT_FOLDER = os.path.join(BASE_FOLDER, "Unimportant")
SUCCESS_ERROR_FOLDER = os.path.join(BASE_FOLDER, "Success and Error files")
DUPLICATE_FOLDER = os.path.join(BASE_FOLDER, "Duplicate Files")

HASHI_OPPTY_FILE = os.path.join(DOWNLOAD_FOLDER, "Hashi oppty.csv")
HASHI_LINEITEM_FILE = os.path.join(DOWNLOAD_FOLDER, "hashi lineitem.csv")

ACCOUNT_EXPORT_FILE = os.path.join(DOWNLOAD_FOLDER, "Account export.csv")
ACCOUNTS_TO_IMPORT_FILE = os.path.join(DOWNLOAD_FOLDER, "Accounts to import.xlsx")
SOURCE_ID_FILE = os.path.join(DOWNLOAD_FOLDER, "Source id file.csv")

SUMMARY_TEMPLATE_FILE = os.path.expanduser(
    "~/Documents/Office Docs/Massload Files/Reference File/SUMMARY_FILE_Hashicorp_production.xlsx"
)
SUMMARY_FILE = os.path.join(BASE_FOLDER, "SUMMARY FILE - HASHI PROD.xlsx")


def show_title(title):
    line_width = 100
    line = "=" * line_width
    print(f"\n{line}")
    print(title.center(line_width))
    print(f"{line}\n")


def ensure_directories():
    os.makedirs(BASE_FOLDER, exist_ok=True)
    os.makedirs(MAIN_FILES_FOLDER, exist_ok=True)
    os.makedirs(UNIMPORTANT_FOLDER, exist_ok=True)
    os.makedirs(SUCCESS_ERROR_FOLDER, exist_ok=True)
    os.makedirs(DUPLICATE_FOLDER, exist_ok=True)


def prompt_yes_no(message):
    return input(message).strip().lower() == "y"


def get_latest_bulkquery_csv():
    matching_files = [
        f for f in os.listdir(DOWNLOAD_FOLDER)
        if f.lower().endswith(".csv") and "bulkquery_result_" in f.lower()
    ]
    if not matching_files:
        raise FileNotFoundError("No matching bulkQuery_result_ CSV file found in Downloads.")
    latest_file = max(
        matching_files,
        key=lambda f: os.path.getmtime(os.path.join(DOWNLOAD_FOLDER, f))
    )
    return os.path.join(DOWNLOAD_FOLDER, latest_file)


def rename_latest_bulkquery_csv(new_file_path):
    latest_file_path = get_latest_bulkquery_csv()
    if os.path.abspath(latest_file_path) == os.path.abspath(new_file_path):
        return new_file_path
    if os.path.exists(new_file_path):
        os.remove(new_file_path)
    os.rename(latest_file_path, new_file_path)
    return new_file_path


def safe_move(src, dst):
    if os.path.exists(src):
        os.makedirs(os.path.dirname(dst), exist_ok=True)
        if os.path.exists(dst):
            os.remove(dst)
        shutil.move(src, dst)
    else:
        print(f"File not found: {src} (skipping)")


def copy_query_and_wait(query, message):
    pyperclip.copy(query.strip())
    if not prompt_yes_no(message):
        print("Operation cancelled.")
        sys.exit()


def get_day_suffix(day):
    if 11 <= day <= 13:
        return "th"
    return {1: "st", 2: "nd", 3: "rd"}.get(day % 10, "th")


def wait_for_step_confirmation(step_message):
    print(f"\n⏸️ MANUAL ACTION REQUIRED:\n{step_message}\n")
    if not prompt_yes_no("Type 'y' to continue to the next step: "):
        print("Process paused by user.")
        sys.exit()


def step_1_import_account():
    show_title("📂 HASHI CORP LOAD - STEP 1 ACCOUNT 📂")

    df = pd.read_csv(HASHI_OPPTY_FILE)
    unique_ids = df["ACCOUNTID"].dropna().astype(str).unique()
    formatted_ids = ",".join([f"'{i}'" for i in unique_ids])

    query = f"""
    SELECT AccountNumber, Id
    FROM Account
    WHERE AccountNumber IN ({formatted_ids})
    """
    copy_query_and_wait(
        query,
        "\n✅ Account query copied. Paste in Workbench, download CSV, then type 'y': "
    )

    rename_latest_bulkquery_csv(ACCOUNT_EXPORT_FILE)

    hashi_df = pd.read_csv(HASHI_OPPTY_FILE)
    account_df = pd.read_csv(ACCOUNT_EXPORT_FILE)

    hashi_ids = (
        hashi_df["ACCOUNTID"]
        .dropna()
        .astype(str)
        .str.strip()
        .str.upper()
        .unique()
    )
    account_numbers = (
        account_df["AccountNumber"]
        .dropna()
        .astype(str)
        .str.strip()
        .str.upper()
        .unique()
    )

    missing_accounts = sorted(set(hashi_ids) - set(account_numbers))
    output_df = pd.DataFrame(missing_accounts, columns=["ACCOUNTID"])
    output_df.to_excel(ACCOUNTS_TO_IMPORT_FILE, index=False)

    print(f"\n📄 Missing accounts saved to: {ACCOUNTS_TO_IMPORT_FILE}")
    print(f"🔢 Total missing accounts: {len(missing_accounts)}")

    safe_move(HASHI_OPPTY_FILE, os.path.join(MAIN_FILES_FOLDER, "hashi oppty.csv"))
    safe_move(ACCOUNT_EXPORT_FILE, os.path.join(UNIMPORTANT_FOLDER, "Account export.csv"))
    safe_move(ACCOUNTS_TO_IMPORT_FILE, os.path.join(UNIMPORTANT_FOLDER, "Accounts to import.xlsx"))

    print("\n🚨 Check if you need to import accounts, then load the opportunity file.")


def step_2_product_file():
    show_title("📂 HASHI CORP LOAD - STEP 2 PRODUCT FILE 📂")

    query = "select Source_Id__c ,Id from Opportunity where Acquisition_Name__c='Hashicorp'"
    copy_query_and_wait(
        query,
        "\n✅ Source id query copied. Paste in Workbench, download CSV, then type 'y': "
    )

    if not os.path.exists(SOURCE_ID_FILE):
        rename_latest_bulkquery_csv(SOURCE_ID_FILE)

    source_df = pd.read_csv(SOURCE_ID_FILE)
    hashi_df = pd.read_csv(HASHI_LINEITEM_FILE)

    source_df["Source_ID__c_std"] = (
        source_df["Source_ID__c"].astype(str).str.strip().str.upper()
    )
    hashi_df["SOURCE_ID__C_std"] = (
        hashi_df["SOURCE_ID__C"].astype(str).str.strip().str.upper()
    )

    merged_df = hashi_df.merge(
        source_df[["Source_ID__c_std", "Id"]],
        left_on="SOURCE_ID__C_std",
        right_on="Source_ID__c_std",
        how="left"
    )
    merged_df.rename(columns={"Id": "Opportunityid"}, inplace=True)

    merged_df["EXPIRATION_DATE__C"] = pd.to_datetime(
        merged_df["EXPIRATION_DATE__C"],
        errors="coerce"
    ).dt.strftime("%Y-%m-%d")

    merged_df.drop(columns=["SOURCE_ID__C_std", "Source_ID__c_std"], inplace=True)

    output_file = os.path.join(MAIN_FILES_FOLDER, "Hashi lineitem.csv")
    merged_df.to_csv(output_file, index=False)

    safe_move(HASHI_LINEITEM_FILE, os.path.join(UNIMPORTANT_FOLDER, "hashi lineitem.csv"))
    safe_move(SOURCE_ID_FILE, os.path.join(UNIMPORTANT_FOLDER, "Source id file.csv"))

    print(f"\n📄 Hashi lineitem file created: {output_file}")
    print("\n🚨 Load the opportunity product now and download success and error files.")


def find_csv_by_keywords(keywords, exclude=None):
    for file in os.listdir(DOWNLOAD_FOLDER):
        name = file.lower()
        if (
            name.endswith(".csv")
            and all(k in name for k in keywords)
            and (not exclude or not any(e in name for e in exclude))
        ):
            return os.path.join(DOWNLOAD_FOLDER, file)
    return None


def step_3_summary_file():
    show_title("📂 CREATE SUMMARY FILE 📂")

    shutil.copy(SUMMARY_TEMPLATE_FILE, SUMMARY_FILE)

    files = {
        "opp_success": {
            "sheet": "Opportunity success",
            "keywords": ["opportunity", "upsert", "success"],
            "exclude": ["product"]
        },
        "opp_error": {
            "sheet": "Opportunity failures",
            "keywords": ["opportunity", "upsert", "error"],
            "exclude": ["product"]
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

    summary_cells = {
        "opp_success": "E5",
        "opp_error": "F5",
        "opp_total": "G5",
        "prod_success": "E6",
        "prod_error": "F6",
        "prod_total": "G6"
    }

    row_counts = {}

    with pd.ExcelWriter(
        SUMMARY_FILE,
        engine="openpyxl",
        mode="a",
        if_sheet_exists="replace"
    ) as writer:
        for key, info in files.items():
            csv_path = find_csv_by_keywords(info["keywords"], info.get("exclude"))

            if csv_path:
                df = pd.read_csv(csv_path)
                df.to_excel(writer, sheet_name=info["sheet"], index=False)
                row_counts[key] = len(df)
                print(f"✅ {info['sheet']} → {len(df)} rows")
                safe_move(
                    csv_path,
                    os.path.join(SUCCESS_ERROR_FOLDER, os.path.basename(csv_path))
                )
            else:
                pd.DataFrame().to_excel(writer, sheet_name=info["sheet"], index=False)
                row_counts[key] = 0
                print(f"⚠️ {info['sheet']} missing → count 0")

    workbook = load_workbook(SUMMARY_FILE)
    summary = workbook["Summary"]

    summary[summary_cells["opp_success"]] = row_counts["opp_success"]
    summary[summary_cells["opp_error"]] = row_counts["opp_error"]
    summary[summary_cells["prod_success"]] = row_counts["prod_success"]
    summary[summary_cells["prod_error"]] = row_counts["prod_error"]
    summary[summary_cells["opp_total"]] = row_counts["opp_success"] + row_counts["opp_error"]
    summary[summary_cells["prod_total"]] = row_counts["prod_success"] + row_counts["prod_error"]

    workbook.save(SUMMARY_FILE)

    print("\n🎯 Summary file generated successfully")


def compare_mismatch_files(
    isc_csv_name,
    isced_csv_name,
    source_column_name,
    object_name,
    run_stat_name
):
    input_path1 = os.path.join(DOWNLOAD_FOLDER, isc_csv_name)
    input_path2 = os.path.join(DOWNLOAD_FOLDER, isced_csv_name)
    output_record_mismatch = os.path.join(DOWNLOAD_FOLDER, f"{object_name}_Record_Mismatch.csv")
    output_lmd_mismatch = os.path.join(UNIMPORTANT_FOLDER, f"{object_name}_LastModifiedDate_Mismatch.csv")

    df = pd.read_csv(input_path1)
    df.rename(columns={source_column_name: "Id"}, inplace=True)
    df.to_csv(input_path1, index=False)

    isc_data = pd.read_csv(input_path1, usecols=["Id"], skipinitialspace=True)
    isced_data = pd.read_csv(input_path2, usecols=["Id"], skipinitialspace=True)

    isc_record_count = isc_data.shape[0]
    isced_record_count = isced_data.shape[0]

    outer_join = pd.merge(isc_data, isced_data, on="Id", how="outer", indicator=True)
    record_mismatch = outer_join[outer_join["_merge"].isin(["left_only", "right_only"])].copy()
    record_mismatch["_merge"] = record_mismatch["_merge"].astype(str)
    record_mismatch.loc[record_mismatch["_merge"] == "left_only", "_merge"] = "Not_in_ISCED"
    record_mismatch.loc[record_mismatch["_merge"] == "right_only", "_merge"] = "Not_in_ISC"
    record_mismatch.to_csv(output_record_mismatch, index=False)
    record_mismatch_count = record_mismatch.shape[0]

    isc_data2 = pd.read_csv(input_path1, skipinitialspace=True)
    isced_data2 = pd.read_csv(input_path2, skipinitialspace=True)

    date_format_1 = "%Y-%m-%dT%H:%M:%S.%fZ"
    date_format_2 = "%Y-%m-%d-%H.%M.%S.%f"

    isc_data2["LastModifiedDate"] = pd.to_datetime(isc_data2["LastModifiedDate"], format=date_format_1)
    isced_data2["LastModifiedDate"] = pd.to_datetime(isced_data2["LastModifiedDate"], format=date_format_2)

    merged_df = pd.merge(isc_data2, isced_data2, on="Id", suffixes=("_isc_data", "_isced_data"))
    lmd_mismatched_records = merged_df[
        merged_df["LastModifiedDate_isc_data"] != merged_df["LastModifiedDate_isced_data"]
    ]
    lmd_mismatch_count = lmd_mismatched_records.shape[0]
    lmd_mismatched_records.to_csv(output_lmd_mismatch, index=False)

    csv_file_path = os.path.join(UNIMPORTANT_FOLDER, run_stat_name)
    df_new = pd.DataFrame({
        "Object_Name": [object_name],
        "ISC_Record_Count": [isc_record_count],
        "ISCED_Record_Count": [isced_record_count],
        "Record_Mismatch_Count": [record_mismatch_count],
        "LastModifiedDate_Mismatch_Count": [lmd_mismatch_count]
    })
    print("\nStat:")
    print(df_new.to_markdown(index=False))
    df_new.to_csv(csv_file_path, mode="a", header=not os.path.exists(csv_file_path), index=False)


def step_4_oppty_mismatch():
    show_title("📂 Oppty Duplicate Files Process 📂")

    query = "select Source_Id__c,LastModifiedDate from Opportunity where Acquisition_Name__c ='HashiCorp' and RecordTypeId='012Ka0000015C9sIAE'"
    copy_query_and_wait(
        query,
        "\n✅ Oppty_isc query copied. Paste in Workbench, download CSV, then type 'y': "
    )

    rename_latest_bulkquery_csv(os.path.join(DOWNLOAD_FOLDER, "oppty_isc.csv"))

    compare_mismatch_files(
        isc_csv_name="oppty_isc.csv",
        isced_csv_name="oppty_isced.csv",
        source_column_name="Source_ID__c",
        object_name="oppty",
        run_stat_name="oppty_Run_Stat.csv"
    )

    safe_move(
        os.path.join(DOWNLOAD_FOLDER, "oppty_isc.csv"),
        os.path.join(DUPLICATE_FOLDER, "oppty_isc.csv")
    )
    safe_move(
        os.path.join(DOWNLOAD_FOLDER, "oppty_isced.csv"),
        os.path.join(DUPLICATE_FOLDER, "oppty_isced.csv")
    )
    safe_move(
        os.path.join(DOWNLOAD_FOLDER, "oppty_Record_Mismatch.csv"),
        os.path.join(DUPLICATE_FOLDER, "oppty_Record_Mismatch.csv")
    )


def step_5_product_mismatch():
    show_title("📂 Product Duplicate Files Process 📂")

    query = "select Lineitem_Legacy_id__c,LastModifiedDate from OpportunityLineitem where Opportunity.Acquisition_Name__c ='HashiCorp' and Opportunity.RecordTypeId='012Ka0000015C9sIAE'"
    copy_query_and_wait(
        query,
        "\n✅ product_isc query copied. Paste in Workbench, download CSV, then type 'y': "
    )

    rename_latest_bulkquery_csv(os.path.join(DOWNLOAD_FOLDER, "product_isc.csv"))

    compare_mismatch_files(
        isc_csv_name="product_isc.csv",
        isced_csv_name="product_isced.csv",
        source_column_name="Lineitem_Legacy_id__c",
        object_name="product",
        run_stat_name="Run_Stat.csv"
    )

    safe_move(
        os.path.join(DOWNLOAD_FOLDER, "product_isc.csv"),
        os.path.join(DUPLICATE_FOLDER, "product_isc.csv")
    )
    safe_move(
        os.path.join(DOWNLOAD_FOLDER, "product_isced.csv"),
        os.path.join(DUPLICATE_FOLDER, "product_isced.csv")
    )
    safe_move(
        os.path.join(DOWNLOAD_FOLDER, "product_Record_Mismatch.csv"),
        os.path.join(DUPLICATE_FOLDER, "product_Record_Mismatch.csv")
    )


def step_6_delete_files():
    show_title("📂 Create Delete Files 📂")

    product_mismatch_file = os.path.join(DUPLICATE_FOLDER, "product_Record_Mismatch.csv")
    oppty_mismatch_file = os.path.join(DUPLICATE_FOLDER, "oppty_Record_Mismatch.csv")

    product_delete_file = os.path.join(DUPLICATE_FOLDER, "PreDelete_Product.csv")
    oppty_delete_file = os.path.join(DUPLICATE_FOLDER, "PreDelete_Oppty.csv")

    product_df = pd.read_csv(product_mismatch_file)
    product_delete = product_df.loc[
        product_df["_merge"] == "Not_in_ISCED", ["Id"]
    ].copy()
    product_delete.columns = ["Delete Product"]
    product_delete.to_csv(product_delete_file, index=False)

    oppty_df = pd.read_csv(oppty_mismatch_file)
    oppty_delete = oppty_df.loc[
        oppty_df["_merge"] == "Not_in_ISCED", ["Id"]
    ].copy()
    oppty_delete.columns = ["Delete Oppty"]
    oppty_delete.to_csv(oppty_delete_file, index=False)

    df = pd.read_csv(oppty_delete_file)
    values = df["Delete Oppty"].dropna().astype(str)
    formatted_values = ",".join(f"'{v}'" for v in values)

    oppty_query = f"""
    SELECT Id,Source_ID__c
    FROM Opportunity
    WHERE Source_ID__c IN ({formatted_values})
    """
    copy_query_and_wait(
        oppty_query,
        "\n✅ Oppty delete query copied. Paste in Workbench, download CSV, then type 'y': "
    )
    rename_latest_bulkquery_csv(os.path.join(DOWNLOAD_FOLDER, "DELETE OPPTY.csv"))

    df = pd.read_csv(product_delete_file)
    values = df["Delete Product"].dropna().astype(str)
    formatted_values = ",".join(f"'{v}'" for v in values)

    product_query = (
        "select Id,Lineitem_Legacy_Id__c from OpportunityLineitem "
        f"where Lineitem_Legacy_Id__c in ({formatted_values})"
    )
    copy_query_and_wait(
        product_query,
        "\n✅ Product delete query copied. Paste in Workbench, download CSV, then type 'y': "
    )
    rename_latest_bulkquery_csv(os.path.join(DOWNLOAD_FOLDER, "DELETE PRODUCT.csv"))

    delete_oppty_main = os.path.join(MAIN_FILES_FOLDER, "DELETE OPPTY.csv")
    delete_product_main = os.path.join(MAIN_FILES_FOLDER, "DELETE PRODUCT.csv")

    safe_move(os.path.join(DOWNLOAD_FOLDER, "DELETE OPPTY.csv"), delete_oppty_main)
    safe_move(os.path.join(DOWNLOAD_FOLDER, "DELETE PRODUCT.csv"), delete_product_main)

    if not os.path.exists(SUMMARY_FILE):
        raise FileNotFoundError("Summary file not found.")

    df_oppty = pd.read_csv(delete_oppty_main)
    df_product = pd.read_csv(delete_product_main)

    with pd.ExcelWriter(SUMMARY_FILE, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        df_oppty.to_excel(writer, sheet_name="DELETE OPPTY", index=False)
        df_product.to_excel(writer, sheet_name="DELETE PRODUCT", index=False)

    today_date = datetime.today().strftime("%Y-%m-%d")
    dated_summary_file = os.path.join(BASE_FOLDER, f"SUMMARY FILE - HASHI PROD ({today_date}).xlsx")
    if os.path.exists(dated_summary_file):
        os.remove(dated_summary_file)
    os.rename(SUMMARY_FILE, dated_summary_file)

    now = datetime.now()
    new_folder_name = f"Hashi Load ({now.day}{get_day_suffix(now.day)} {now.strftime('%B')})"
    new_folder_path = os.path.join(DOWNLOAD_FOLDER, new_folder_name)

    if os.path.exists(new_folder_path):
        shutil.rmtree(new_folder_path)
    os.rename(BASE_FOLDER, new_folder_path)

    print(f"\n\nSUMMARY FOLDER:- {new_folder_path}\n")


def main():
    ensure_directories()

    show_title("📂 HASHI CORP COMBINED LOAD 📂")
    print("This script combines all 6 Hashi Corp steps into one flow.")
    print("Keep required input files ready in Downloads before each step.")

    step_1_import_account()
    wait_for_step_confirmation(
        "Review 'Accounts to import.xlsx'. If account import is needed, complete it before proceeding."
    )

    step_2_product_file()
    wait_for_step_confirmation(
        "Import the generated product file, then download the Opportunity/Product success and error CSV files to Downloads before proceeding to Summary."
    )

    step_3_summary_file()
    wait_for_step_confirmation(
        "Review the summary file. Keep 'oppty_isced.csv' ready in Downloads before running opportunity mismatch."
    )

    step_4_oppty_mismatch()
    wait_for_step_confirmation(
        "Review opportunity mismatch output. Keep 'product_isced.csv' ready in Downloads before running product mismatch."
    )

    step_5_product_mismatch()
    wait_for_step_confirmation(
        "Review duplicate mismatch outputs before generating delete files."
    )

    step_6_delete_files()

    show_title("📂 Hashi Load Done 📂")


if __name__ == "__main__":
    main()

# Made with Bob
