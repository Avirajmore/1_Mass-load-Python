
import os
import pandas as pd
from tabulate import tabulate
print(f"\n{'='*100}\n{' ' * 30} 📝  File Data Validation 📝 {' ' * 30}\n{'='*100}\n")
folder_path = '/Users/avirajmore/Downloads'

print(f"\n🔍 Files Available: ")
for file in os.listdir(folder_path):
    if file.endswith(".xlsx"):
        print(f"\n    ✅ {file}")

for file in os.listdir(folder_path):
    if file.endswith(".xlsx"):
        file_path = os.path.join(folder_path,file)
        
        
        
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
            'AUTOREN_SUBNEW', 'AUTOREN_SUBUPG', 'TERMINATE_SUBNEW', 'TERMINATE_SUBUPG','SW','SWSUBSCR','SWSVC','SAAS','MAINT'
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
            "sales_stage", "OI_Source", "Classification Type", "Type", "Renewal type", "Renewal Status","product_type"
        ]

        # --- Columns to check for blanks ---
        blank_values_columns = {
            "Opportunity": [
                "opportunity_legacy_id_c", "name", "accountid", "sales_stage","ownerid", "expected_close_date", "currency_code", "ownerid", "OI_Source"
            ],
            "Opportunity_product": [
                "Product", "product_type", "unitprice", "Term", "Classification Type"
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
                    print(f"\n    ❌ Missing columns in '{sheet}':")
                    for col in missing:
                        print(f"\n        - {col}")
                else:
                    print(f"\n    ✅ All required columns present in '{sheet}'")
            else:
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
                print(f"\n    ❗️ Sheet: {sheet}")
                for col, values in columns.items():
                    print(f"\n        • Column: {col}")
                    for val in values:
                        print(f"\n            - {val}")

            print("\n    ✅ All other values are valid.\n")

        else:
            print("\n    🎉 All API values are valid in both sheets.")


        # --- Check Opportunity: Missing 'Won Reason' and 'Lost Reason/Category' ---
        print(f"\n🔍 Step 4: Check if Won Reason, Lost category and Lost reason are present")
        if "Opportunity" in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name='Opportunity')
            df['Excel Row'] = df.index + 2  # Excel row numbers

            identifier_cols = ['Opportunity Name'] if 'Opportunity Name' in df.columns else []

            missing_won_reason = df[(df['sales_stage'] == 'Won') & (df['Won Reason'].isnull() | (df['Won Reason'].astype(str).str.strip() == ''))]
            missing_lost_info = df[(df['sales_stage'] == 'Lost') & (
                (df['Lost Category'].isnull() | (df['Lost Category'].astype(str).str.strip() == '')) |
                (df['Lost Reason'].isnull() | (df['Lost Reason'].astype(str).str.strip() == ''))
            )]

            if not missing_won_reason.empty:
                print(f"\n    ❗️ Missing 'Won Reason' for these rows (sales_stage = 'Won'):\n")
                table_str = tabulate(
                missing_won_reason[['Excel Row'] + identifier_cols + ['sales_stage', 'Won Reason']],
                headers='keys',
                tablefmt='pretty',
                showindex=False
            )
            
                # Add indent to each line of the table
                indented_table = "\n".join("        " + line for line in table_str.splitlines())
                print(indented_table)
            else:
                print("\n    ✅ No missing 'Won Reason' for any rows with sales_stage = 'Won'.")

            if not missing_lost_info.empty:
                print(f"\n    ❗️ Missing 'Lost Category' or 'Lost Reason' for these rows (sales_stage = 'Lost'):\n")
                
                try:
                    # Generate the table only if there is data
                    table_str = tabulate(
                        missing_lost_info[['Excel Row'] + identifier_cols + ['sales_stage', 'Lost Category', 'Lost Reason']],
                        headers='keys',
                        tablefmt='pretty',
                        showindex=False
                    )
                except Exception as e:
                        print(f"Error generating table: {e}")

                # Add indent to each line of the table
                indented_table = "\n".join("        " + line for line in table_str.splitlines())
                print(indented_table)
            else:
                    print("\n    ✅ No missing 'Lost Category' or 'Lost Reason' for any rows with sales_stage = 'Lost'.")

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
                print("\n    ❗️ Blank values found in columns of 'Opportunity':")
                for col in blank_cols_opp:
                    print(f'\n        - {col}')
            else:
                print("\n    ✅ No blank values in required columns of 'Opportunity'.")

        if 'Opportunity_product' in xls.sheet_names:
            df_opportunity_product = pd.read_excel(xls, sheet_name='Opportunity_product')
            blank_cols_prod = check_blank_values(df_opportunity_product, blank_values_columns["Opportunity_product"])
            if blank_cols_prod:
                print("\n    ❗️ Blank values found in columns of 'Opportunity_product':")
                for col in blank_cols_prod:
                    print(f'\n        - {col}')
            else:
                print("\n    ✅ No blank values in required columns of 'Opportunity_product'.")

        print(f"\n{'='*100}\n{' ' * 30} ✅  File Validation Done ✅ {' ' * 30}\n{'='*100}\n")