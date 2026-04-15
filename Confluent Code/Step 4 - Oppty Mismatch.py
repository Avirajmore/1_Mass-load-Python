import pandas as pd
import os
import shutil
import pyperclip
from tabulate import tabulate


def show_title(title):

    line_width = 100
    line = "=" * line_width
    print(f"\n{line}")
    print(title.center(line_width))
    print(f"{line}\n")

# Display the title for the folder creation and file movement process
title = "📂 Oppty Dupliate Files process 📂"
show_title(title)
print("\n❗️ Download Oppty_ised file from Database first and rename it.")
query = "select Source_Id__c,LastModifiedDate from Opportunity where Acquisition_Name__c ='Confluent' and RecordTypeId='012Ka0000015C9sIAE'"
pyperclip.copy(query.strip())
DOWNLOAD_FOLDER = os.path.expanduser("~/Downloads")                # change if needed (e.g. Downloads)
NEW_FILE_NAME = 'oppty_isc.csv'
choice = input("\n✅ Oppty_isc query is copied to clipboard\n\n - Paste this Query in the WorkBench and download the csv file. Once done, type 'y' !")

if choice.lower() == 'y':

    # ---------------- FILE RENAME ----------------

    matching_files = [
        f for f in os.listdir(DOWNLOAD_FOLDER)
        if f.lower().endswith(".csv") and "bulkquery_result_" in f.lower()
    ]

    if not matching_files:
        print("\n❌ No matching bulkQuery_result_ CSV file found.\n")
    else:
        # Pick latest modified file
        latest_file = max(
            matching_files,
            key=lambda f: os.path.getmtime(os.path.join(DOWNLOAD_FOLDER, f))
        )

        old_path = os.path.join(DOWNLOAD_FOLDER, latest_file)
        new_path = os.path.join(DOWNLOAD_FOLDER, NEW_FILE_NAME)

        os.rename(old_path, new_path)

dir_path = os.path.expanduser("~/Downloads/Confluent Load/Unimportant")

# Create directory if it doesn't exist
os.makedirs(dir_path, exist_ok=True)


class CompareCsv():
    def csv_compare(self):
        # path variables
        isc_obj = 'oppty'
        isced_obj = 'oppty'
        output_directory = os.path.expanduser(f"~/Downloads")
        input_directory = os.path.expanduser(f"~/Downloads")
        input_csv_file_name1 = isc_obj + '_isc.csv'
        input_csv_file_name2 = isced_obj + '_isced.csv'
        output_csv_file_name1 = isced_obj + '_Record_Mismatch.csv'
        output_csv_file_name2 = isced_obj + '_LastModifiedDate_Mismatch.csv'
        input_path1 = os.path.join(input_directory, input_csv_file_name1)
        input_path2 = os.path.join(input_directory, input_csv_file_name2)
        output_path1 = os.path.join(output_directory, output_csv_file_name1)
        output_path2 = os.path.join(dir_path, output_csv_file_name2)
        # record compare
        
        df = pd.read_csv(input_path1)
        df.rename(columns={"Source_ID__c": "Id"}, inplace=True)

        # Write back to SAME file
        df.to_csv(input_path1, index=False)

        isc_data = pd.read_csv(input_path1, usecols=['Id'], skipinitialspace=True)
        isced_data = pd.read_csv(input_path2, usecols=['Id'], skipinitialspace=True)
        isc_record_count = isc_data.shape[0]
        isced_record_count = isced_data.shape[0]
        outer_join = pd.merge(isc_data, isced_data, on='Id', how='outer', indicator=True)
        outer_join['_merge'] = outer_join['_merge'].astype(str)
        record_mismatch = outer_join[
            outer_join['_merge'].isin(['left_only', 'right_only'])
        ].replace({'left_only': 'Not_in_ISCED', 'right_only': 'Not_in_ISC'})
        
        record_mismatch.to_csv(output_path1, index=False)
        record_mismatch_count = record_mismatch.shape[0]
        #isc_only = outer_join[outer_join['_merge'] == 'left_only']
        #isced_only = outer_join[outer_join['_merge'] == 'right_only']
        # last modified date compare
        isc_data2 = pd.read_csv(input_path1, skipinitialspace=True)
        isced_data2 = pd.read_csv(input_path2, skipinitialspace=True)
        date_format_1 = "%Y-%m-%dT%H:%M:%S.%fZ"
        date_format_2 = "%Y-%m-%d-%H.%M.%S.%f"
        isc_data2['LastModifiedDate'] = pd.to_datetime(isc_data2['LastModifiedDate'], format=date_format_1)
        isced_data2['LastModifiedDate'] = pd.to_datetime(isced_data2['LastModifiedDate'], format=date_format_2)
        merged_df = pd.merge(isc_data2, isced_data2, on='Id', suffixes=('_isc_data', '_isced_data'))
        lmd_mismatched_records = merged_df[merged_df['LastModifiedDate_isc_data'] != merged_df['LastModifiedDate_isced_data']]
        lmd_mismatch_count = lmd_mismatched_records.shape[0]
        lmd_mismatched_records.to_csv(output_path2, index=False)

        # Write run stat
        csv_file_path = os.path.expanduser("~/Downloads/Confluent Load/Unimportant/oppty_Run_Stat.csv")
        data = {'Object_Name': [isced_obj],
                'ISC_Record_Count': [isc_record_count],
                'ISCED_Record_Count': [isced_record_count],
                'Record_Mismatch_Count': [record_mismatch_count],
                'LastModifiedDate_Mismatch_Count': [lmd_mismatch_count]
                }
        df_new = pd.DataFrame(data)
        print("\nStat:")
        print(df_new.to_markdown(index=False))
        df_new.to_csv(csv_file_path, mode='a', header=False, index=False)
        print("\n")

if __name__ == "__main__":
    class_instance = CompareCsv()
    class_instance.csv_compare()

os.mkdir(os.path.expanduser("~/Downloads/Confluent Load/Duplicate Files"))
shutil.move(os.path.expanduser("~/Downloads/oppty_isc.csv"), os.path.expanduser("~/Downloads/Confluent Load/Duplicate Files/oppty_isc.csv"))
shutil.move(os.path.expanduser("~/Downloads/oppty_isced.csv"), os.path.expanduser("~/Downloads/Confluent Load/Duplicate Files/oppty_isced.csv"))
shutil.move(os.path.expanduser("~/Downloads/oppty_Record_Mismatch.csv"), os.path.expanduser("~/Downloads/Confluent Load/Duplicate Files/oppty_Record_Mismatch.csv"))

title = "📂 Oppty Mismatch done 📂"
show_title(title)