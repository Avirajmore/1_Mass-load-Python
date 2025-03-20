import sys
import pandas as pd

file_path = '/Users/avirajmore/Downloads/ISC reload for Q2 Renewals 03192025 -  Removing Duplicates copy.xlsx'
print("\n\n🔍 Step 6: Creating or Overwriting the 'Currency' column in the 'Opportunity_product' sheet...")

# Read the relevant sheets from the Excel file
opportunity_product_df = pd.read_excel(file_path, sheet_name="Opportunity_product")
opportunity_df = pd.read_excel(file_path, sheet_name="Opportunity")

try:
    # Perform VLOOKUP operation (merge data)
    # print(f"\n    🔄 Merging data to create or overwrite the 'opportunity currency' column...")
    merged_df = pd.merge(opportunity_product_df, opportunity_df,
                        left_on="opportunityid", right_on="opportunity_legacy_id__c",
                        how="left")
    
    merged_df.to_csv('/Users/avirajmore/Downloads/merged_df.csv')


    # Ensure the column is fully cleared before overwriting
    opportunity_product_df["opportunity currency"] = None  # Clear any existing values

    opportunity_product_df.to_csv('/Users/avirajmore/Downloads/opportunity_product_df1.csv')

    # Explicitly update the 'opportunity currency' column with the correct value from the merged DataFrame
    opportunity_product_df = opportunity_product_df.merge(merged_df[['opportunityid', 'CurrencyIsoCode']], 
                                                           on='opportunityid', 
                                                           how='left')
    
    opportunity_product_df.to_csv('/Users/avirajmore/Downloads/opportunity_product_df.csv')

    # Overwrite or add the "opportunity currency" column
    opportunity_product_df['opportunity currency'] = opportunity_product_df['CurrencyIsoCode']

    # Save the modified DataFrame back to Excel
    with pd.ExcelWriter(file_path, mode='a', if_sheet_exists='replace') as writer:
        opportunity_product_df.to_excel(writer, sheet_name="Opportunity_product", index=False)

    # Success message
    print(f"\n    ✅ Process completed. The 'opportunity currency' column has been successfully created in the 'Opportunity_product' sheet.")

except FileNotFoundError:
    # Handle file not found error
    print(f"\n    ❌ Error: File '{file_path}' not found. ")
    sys.exit()
except KeyError as e:
    # Handle missing column error
    print(f"\n    ❌ Error: The required column '{e.args[0]}' is missing. ")
    sys.exit()
except Exception as e:
    # Handle any other unexpected errors
    print(f"\n    ❌ Error: An unexpected error occurred. Details: {e} ")
    sys.exit()
