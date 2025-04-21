import os
# test
import pandas as pd
from tabulate import tabulate

line_width = 100
line = "=" * line_width

title = "📝  File Data Validation 📝"
print(f"\n{line}")
print(title.center(line_width))
print(f"{line}\n")

folder_path = '/Users/avirajmore/Downloads'

summary_list = []

print(f"\n🔍 Files Available: ")
for file in os.listdir(folder_path):
    if file.endswith(".xlsx"):
    
        file_path = os.path.join(folder_path, file)
        xls = pd.ExcelFile(file_path)
        sheet_names = xls.sheet_names
        if 'Opportunity' in sheet_names and 'Opportunity_product' in sheet_names:
            print(f"\n    ✅ {file}")
        else:
            print(f"\n    ❌ {file}")

for file in os.listdir(folder_path):
    if file.endswith(".xlsx"):
        file_path = os.path.join(folder_path,file)
        xls = pd.ExcelFile(file_path)
        sheet_names = xls.sheet_names
        if 'Opportunity' in sheet_names and 'Opportunity_product' in sheet_names:
            title = f"✅ {file} ✅"
            print(f"\n{line}")
            print(title.center(line_width))
            print(f"{line}\n")


            # Summary status flag
            file_status = "✅ All Good"
            
            print(f"\n🔍 Step 1: Select a file")
            print(f"\n    📝 File Selected: {file}")

            # --- Load Excel file ---
            xls = pd.ExcelFile(file_path)

            # --- Valid API Names ---
            valid_api_names = [
                'Engage', 'Qualify', 'Design', 'Propose', 'Negotiate', 'Closing', 'Prospecting', 'Developing', 'Negotiation', 'Won', 'Lost',
                'RNL_MAINT_ADMIN', 'RNL_MAINT_NOPAY', 'RNL_MAINT_CATCH', 'RNL_MAINT_ENDSUPP', 'RNL_MAINT_OUTPROD',
                'RNL_MAINT_REPLACE', 'RNL_MAINT_SELFSERV', 'RNL_MAINT_THIRD', 'RNL_COMP_PRICE', 'RNL_CNTR_EARLY',
                'RNL_CNTR_MOVE', 'RNL_CUST_NORESP', 'RNL_CUST_NOBDGT', 'RNL_CUST_OUTOFBUS', 'RNL_CUST_PROJEND',
                'RNL_CUST_TEMP', 'RNL_IBM_PRODCHNG', 'RNL_IBM_NOPROD', 'RNL_PROD_NOSUPP', 'RNL_PROD_NOTRELB',
                'RNL_PROD_NOSOLN', 'RNL_REVN_MOVETERM', 'RNL_REVN_MOVESAAS', 'RNL_REVN_MOVEPERP', 'RNL_REVN_MOVESUBS',
                'RNL_REVN_MOVEPERS', 'RNL_REVN_MOVETYPE', 'RNL_REVN_MOVESALE', 'Business Partner', 'Digital Sales',
                'Field Sales', 'Digital Sales Development', 'TYPEESA', 'TYPEOEM', 'TYPEPCR', 'TYPERFS', 'TYPESRV',
                'TYPEASP', 'TYPEEST', 'TYPESSP', 'TYPEMNL', 'TYPERBK', 'TYPEZSW', 'TYPEESP', 'TYPEASE', 'TYPEPLN',
                'TYPEINS', 'TYPEBLD', 'TYPEMIG', 'TYPEPER', 'TYPEESM', 'TYPEEXC', 'TYPELSU', 'TYPELIN', 'TYPECUS',
                'CLASSNEW', 'CLASSEXP', 'CLASSREN', 'CLASSUPG', 'CLASSDEP', 'CLASSREI', 'CLASSWEX', 'RENEW_QUALIFY',
                'RENEW_DESIGN', 'RENEW_PROPOSE', 'RENEW_NEGOTIATE', 'RENEW_CLOSING', 'RENEW_LOST', 'RENEW_WON',
                'AUTOREN_ORIG', 'AUTOREN_12M', 'AUTOREN_24M', 'AUTOREN_36M', 'TERMINATE', 'Bill CONTINUOUS', 'NOTAPPL',
                'AUTOREN_SUBNEW', 'AUTOREN_SUBUPG', 'TERMINATE_SUBNEW', 'TERMINATE_SUBUPG','SW','SWSUBSCR','SWSVC','SAAS','MAINT',
                'WONSKILL',	'WONBRAND',	'WONPRICE',	'WONINC',	'WONTERMS',	'WONRFP',	'WONISV',	'WONEXEC',	'WONNEEDS',	'WONFIN',
                'LCDNP',	'LIBMDNP',	'LLTC',	'LOSTTLS',	'LOSTFIN',	'LOSTREN',	'FINPART','IBMRESC',	'FINBP', 'IBMREXP',	'CUSPRIOR',	
                'FINCASH',	'CUSSTATU',	'COMPEXP',	'COMPINC',	'COMPTRMS',	'COMPSOLN',	'FINREJECT', 'FINCUST',	'COMPSAT',	'FIND',	'IBMDUP',	
                'IBMERR',	'CUSSPONS',	'FINGOE',	'FINIBM',	'COMPREL',	'COMPPART',	'IBMBUDG',	'FINK',	'IBMSUPP',	'FINCOMP',	'CUSTBUDG',	
                'IBMLOW',	'CUSTRISK',	'COMPMOVE',	'CUSTACT',	'FINNH',	'FINNONF',	'COMPPRCE',	'RNL_COMP_PRICE',	'RNL_CNTR_EARLY',	'RNL_CNTR_MOVE',	
                'RNL_CUST_NOBDGT',	'RNL_CUST_NORESP',	'RNL_CUST_OUTOFBUS',	'RNL_CUST_PROJEND',	'RNL_CUST_TEMP',	'RNL_IBM_PRODCHNG',	'RNL_IBM_NOPROD',	
                'RNL_PROD_NOSUPP',	'RNL_PROD_NOTRELB',	'RNL_PROD_NOSOLN',	'RNL_REVN_MOVESAAS',	'RNL_REVN_MOVETERM',	'RNL_REVN_MOVEPERP',	'RNL_REVN_MOVESUBS',	
                'RNL_REVN_MOVEPERS',	'RNL_REVN_MOVETYPE',	'RNL_REVN_MOVESALE',	'IBMUNDES',	'IBMUNRSP',	'TLSTRIBM',	'TLSTRNON',	'TLSCLOUD',	'TLSOTHER',	'TLSSELFP',	
                'TLSSELFE',	'TLSSELFO',	'TLSTPMP',	'TLSTPME',	'TLSTPMSD',	'TLSMVS',	'TLSIBMNA',	'TLSNORES',	'TLSBUDG',	'TLSDUPE',	'TLSNONE',	'BPDREXP',	'BPSUPP',	
                'BPESA',	'BPQUA',	'BPOTHER',	'TLSRENCO'
            ]

            # --- Required Columns ---
            required_columns = {
                "Opportunity": [
                    "opportunity_legacy_id_c", "name", "accountid", "sales_stage",
                    "expected_close_date", "currency_code", "ownerid", "OI_Source"
                ],
                "Opportunity_product": [
                    "OpportunityId", "Product", "product_type", "unitprice", "Term",
                    "Classification Type", "Type"
                ]
            }

            # --- Columns to Validate against API list ---
            columns_to_validate = [
                "sales_stage", "OI_Source", "Classification Type", "Type", "Renewal type", "Renewal Status","product_type",
                "Won Reason","Lost Category","Lost Reason"
            ]

            # --- Columns to check for blanks ---
            blank_values_columns = {
                "Opportunity": [
                    "opportunity_legacy_id_c", "name", "accountid", "sales_stage","ownerid", "expected_close_date", "currency_code", "ownerid", "OI_Source"
                ],
                "Opportunity_product": [
                    "Product", "product_type", "unitprice", "Classification Type"
                ]
            }

            # --- Check Required Columns ---
            print(f"\n🔍 Step 2: Checking Required Columns: ")
            for sheet, columns in required_columns.items():
                if sheet in xls.sheet_names:
                    df = pd.read_excel(xls, sheet_name=sheet)
                    df_columns_lower = [col.lower() for col in df.columns]
                    missing = [col for col in columns if col.lower() not in df_columns_lower]
                    if missing:
                        file_status = "❌ Issues Found"
                        print(f"\n    ❌ Missing columns in '{sheet}':")
                        for col in missing:
                            print(f"\n        🔸 {col}")
                    else:
                        print(f"\n    ✅ All required columns present in '{sheet}'")
                else:
                    file_status = "❌ Issues Found"
                    print(f"\n    ❌ Sheet '{sheet}' not found.")

            # --- Check API Value Validity (Case-Insensitive) ---
            print(f"\n🔍 Step 3: Checking for Invalid API names:")

            # Convert valid API names to lowercase for case-insensitive check
            valid_api_names_lower = [val.strip().lower() for val in valid_api_names]
            # Track invalid values per sheet and column
            invalid_report = {}

            for sheet_name in ['Opportunity', 'Opportunity_product']:
                if sheet_name in xls.sheet_names:
                    df = pd.read_excel(xls, sheet_name=sheet_name)

                    for col in columns_to_validate:
                        if col in df.columns:
                            non_blank_values = df[col].dropna().astype(str)
                            non_blank_values = non_blank_values[non_blank_values.str.strip() != ""]

                            invalid_values = non_blank_values[~non_blank_values.str.strip().str.lower().isin(valid_api_names_lower)]

                            if not invalid_values.empty:
                                if sheet_name not in invalid_report:
                                    invalid_report[sheet_name] = {}
                                invalid_report[sheet_name][col] = invalid_values.unique()

            # Display results
            if invalid_report:

                for sheet, columns in invalid_report.items():
                    file_status = "❌ Issues Found"
                    print(f"\n    ❗️ Sheet: {sheet}")
                    for col, values in columns.items():
                        print(f"\n        🔹 Column: {col}")
                        for val in values:
                            print(f"\n            🔸 {val}")

                print("\n    ✅ All other values are valid.\n")

            else:
                print("\n    🎉 All API values are valid in both sheets.")


            # --- Check Opportunity: Missing 'Won Reason' and 'Lost Reason/Category' ---
            print(f"\n🔍 Step 4: Check if Won Reason, Lost category and Lost reason are present")
            if "Opportunity" in xls.sheet_names:
                df = pd.read_excel(xls, sheet_name='Opportunity')
                df['Excel Row'] = df.index + 2  # Excel row numbers

                identifier_cols = ['Opportunity Name'] if 'Opportunity Name' in df.columns else []
                
                df['sales_stage'] = df['sales_stage'].str.lower()

                # Check for missing reasons based on stage
                missing_won_reason = df[(df['sales_stage'] == 'won') & (df['Won Reason'].isnull() | (df['Won Reason'].astype(str).str.strip() == ''))]
                missing_lost_info = df[(df['sales_stage'] == 'lost') & (
                    (df['Lost Category'].isnull() | (df['Lost Category'].astype(str).str.strip() == '')) |
                    (df['Lost Reason'].isnull() | (df['Lost Reason'].astype(str).str.strip() == ''))
                )]

                # 🔴 Check for invalid values in 'Lost' fields when stage is 'Won'
                invalid_lost_fields_in_won = df[(df['sales_stage'] == 'Won') & (
                    (df['Lost Category'].notnull() & (df['Lost Category'].astype(str).str.strip() != '')) |
                    (df['Lost Reason'].notnull() & (df['Lost Reason'].astype(str).str.strip() != ''))
                )]

                # 🔴 Check for invalid 'Won Reason' when stage is 'Lost'
                invalid_won_reason_in_lost = df[(df['sales_stage'] == 'Lost') & (
                    (df['Won Reason'].notnull() & (df['Won Reason'].astype(str).str.strip() != ''))
                )]

                # Reporting
                if not missing_won_reason.empty:
                    file_status = "❌ Issues Found"
                    print(f"\n    ❗️ Missing 'Won Reason' where sales_stage = 'Won'")
                    # table_str = tabulate(
                    #     missing_won_reason[['Excel Row'] + identifier_cols + ['sales_stage', 'Won Reason']],
                    #     headers='keys', tablefmt='pretty', showindex=False
                    # )
                    # indented_table = "\n".join("        " + line for line in table_str.splitlines())
                    # print(indented_table)
                else:
                    print("\n    ✅ No missing 'Won Reason' for any rows with sales_stage = 'Won'.")

                if not missing_lost_info.empty:
                    file_status = "❌ Issues Found"
                    print(f"\n    ❗️ Missing 'Lost Category' or 'Lost Reason' for these rows where sales_stage = 'Lost'")
                    # table_str = tabulate(
                    #     missing_lost_info[['Excel Row'] + identifier_cols + ['sales_stage', 'Lost Category', 'Lost Reason']],
                    #     headers='keys', tablefmt='pretty', showindex=False
                    # )
                    # indented_table = "\n".join("        " + line for line in table_str.splitlines())
                    # print(indented_table)
                else:
                    print("\n    ✅ No missing 'Lost Category' or 'Lost Reason' for any rows with sales_stage = 'Lost'.")

                # 🔴 Invalid values in Lost fields for 'Won' stage
                if not invalid_lost_fields_in_won.empty:
                    file_status = "❌ Issues Found"
                    print(f"\n    ❗️ Invalid 'Lost Category' or 'Lost Reason' present for these rows where sales_stage = 'Won'")
                    table_str = tabulate(
                        invalid_lost_fields_in_won[['Excel Row'] + identifier_cols + ['sales_stage', 'Lost Category', 'Lost Reason']],
                        headers='keys', tablefmt='pretty', showindex=False
                    )
                    indented_table = "\n".join("        " + line for line in table_str.splitlines())
                    print(indented_table)
                else:
                    print("\n    ✅ No invalid 'Lost Category' or 'Lost Reason' for any rows with sales_stage = 'Won'.")

                # 🔴 Invalid 'Won Reason' in 'Lost' stage
                if not invalid_won_reason_in_lost.empty:
                    file_status = "❌ Issues Found"
                    print(f"\n    ❗️ Invalid 'Won Reason' present for these rows (sales_stage = 'Lost'):\n")
                    table_str = tabulate(
                        invalid_won_reason_in_lost[['Excel Row'] + identifier_cols + ['sales_stage', 'Won Reason']],
                        headers='keys', tablefmt='pretty', showindex=False
                    )
                    indented_table = "\n".join("        " + line for line in table_str.splitlines())
                    print(indented_table)
                else:
                    print("\n    ✅ No invalid 'Won Reason' for any rows with sales_stage = 'Lost'.")

            # --- Function to Check for Blank Values in Given Columns ---
            print(f"\n🔍 Step 5: Check if there are any blank values in Important columns")
            def check_blank_values(df, column_names):
                blank_columns = []
                for col in column_names:
                    if col in df.columns:
                        blanks_found = df[col].isnull().any() or (df[col].apply(lambda x: isinstance(x, str) and x.strip() == '')).any()
                        if blanks_found:
                            blank_columns.append(col)
                return blank_columns

            # --- Check for Blank Values in 'Opportunity' and 'Opportunity_product' ---
            if 'Opportunity' in xls.sheet_names:
                df_opportunity = pd.read_excel(xls, sheet_name='Opportunity')
                blank_cols_opp = check_blank_values(df_opportunity, blank_values_columns["Opportunity"])
                if blank_cols_opp:
                    file_status = "❌ Issues Found"
                    print("\n    ❗️ Blank values found in columns of 'Opportunity':")
                    for col in blank_cols_opp:
                        print(f'\n        🔸 {col}')
                else:
                    print("\n    ✅ No blank values in required columns of 'Opportunity'.")

            if 'Opportunity_product' in xls.sheet_names:
                df_opportunity_product = pd.read_excel(xls, sheet_name='Opportunity_product')
                blank_cols_prod = check_blank_values(df_opportunity_product, blank_values_columns["Opportunity_product"])
                if blank_cols_prod:
                    file_status = "❌ Issues Found"
                    print("\n    ❗️ Blank values found in columns of 'Opportunity_product':")
                    for col in blank_cols_prod:
                        print(f'\n        🔸 {col}')
                else:
                    print("\n    ✅ No blank values in required columns of 'Opportunity_product'.")

            print(f"\n🔍 Step 6: Check if there are any blank values in Important columns")
            if "Opportunity" in xls.sheet_names:
                df_opportunity = pd.read_excel(xls, sheet_name='Opportunity')
                df_opportunity = df_opportunity.drop_duplicates()
                if df_opportunity['opportunity_legacy_id_c'].duplicated().any():
                    file_status = "❌ Issues Found"
                    duplicated_values = df_opportunity[df_opportunity['opportunity_legacy_id_c'].duplicated(keep=False)]
                    print("\n   ❗️ Duplicate values found in 'opportunity_legacy_id_c'")
                else:
                    print("\n   ✅ No duplicate values found in 'opportunity_legacy_id_c'.")

            # Append file result to summary
            summary_list.append({"File Name": file, "Status": file_status})

title = "✅  File Validation Done ✅"
print(f"\n{line}")
print(title.center(line_width))
print(f"{line}\n")

# ✅ Final Summary Report
title = "📊 Final Summary 📊"
print(f"\n{line}")
print(title.center(line_width))
print(f"{line}\n")

print(tabulate(summary_list, headers="keys", tablefmt="fancy_grid"))
