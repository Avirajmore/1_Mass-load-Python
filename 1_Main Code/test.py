import sys
import pandas as pd

file_path = '/Users/avirajmore/Downloads/ISC reload for Q2 Renewals 03192025 -  Removing Duplicates copy.xlsx'
print("\n\n🔍 Step 6: Creating or Overwriting the 'Currency' column in the 'Opportunity_product' sheet...")

# Read the relevant sheets from the Excel file
opportunity_product_df = pd.read_excel(file_path, sheet_name="Opportunity_product")
opportunity_df = pd.read_excel(file_path, sheet_name="Opportunity")

# Debug: Print the first few rows of both DataFrames to verify the structure and data
print("First few rows of Opportunity_product DataFrame:")
print(opportunity_product_df.head())

print("First few rows of Opportunity DataFrame:")
print(opportunity_df.head())

try:
    # Normalize the columns to strip spaces and ensure case consistency
    opportunity_product_df["opportunityid"] = opportunity_product_df["opportunityid"].str.strip()
    opportunity_df["opportunity_legacy_id__c"] = opportunity_df["opportunity_legacy_id__c"].str.strip()
    opportunity_df["CurrencyIsoCode"] = opportunity_df["CurrencyIsoCode"].str.strip().str.upper()

    # Debug: Check for duplicates in the Opportunity DataFrame
    duplicates = opportunity_df[opportunity_df.duplicated(subset=["opportunity_legacy_id__c"], keep=False)]
    print(f"\nRows with duplicate opportunity_legacy_id__c:\n{duplicates[duplicates['opportunity_legacy_id__c'] == '006Sc000009xcNS']}")

    # Perform the merge (VLOOKUP) operation
    print("\n🔄 Merging data...")
    merged_df = pd.merge(opportunity_product_df, opportunity_df,
                         left_on="opportunityid", right_on="opportunity_legacy_id__c",
                         how="left")

    # Debug: Check the columns in the merged DataFrame
    print("\nColumns in the merged DataFrame:")
    print(merged_df.columns)

    # Debug: Check the merged DataFrame for the specific opportunityid
    print(f"\nMerged Data for opportunityid '006Sc000009xcNS':")
    print(merged_df[merged_df["opportunityid"] == "006Sc000009xcNS"])

    # Check if the currency for the given opportunity is correct in Opportunity DataFrame
    specific_opportunity = opportunity_df[opportunity_df["opportunity_legacy_id__c"] == "006Sc000009xcNS"]
    print(f"\nCurrency for opportunity '006Sc000009xcNS' in Opportunity DataFrame:")
    print(specific_opportunity[["opportunity_legacy_id__c", "CurrencyIsoCode"]])

    # Ensure the column is fully cleared before overwriting
    opportunity_product_df["opportunity currency"] = None  # Clear any existing values

    # Explicitly update the 'opportunity currency' column with the correct value from the merged DataFrame
    opportunity_product_df = opportunity_product_df.merge(merged_df[['opportunityid', 'CurrencyIsoCode']], 
                                                           on='opportunityid', 
                                                           how='left')

    # Overwrite or update the 'opportunity currency' column in the Opportunity_product DataFrame
    opportunity_product_df['opportunity currency'] = opportunity_product_df['CurrencyIsoCode']

    # Check if the currency in Opportunity_product DataFrame is correct after merge
    print(f"\nCurrency in Opportunity_product DataFrame for opportunityid '006Sc000009xcNS':")
    print(opportunity_product_df[opportunity_product_df["opportunityid"] == "006Sc000009xcNS"][["opportunityid", "opportunity currency"]])

    # Save the modified DataFrame back to Excel
    with pd.ExcelWriter(file_path, mode='a', if_sheet_exists='replace') as writer:
        opportunity_product_df.to_excel(writer, sheet_name="Opportunity_product", index=False)

    # Success message
    print(f"\n✅ Process completed. The 'opportunity currency' column has been successfully created in the 'Opportunity_product' sheet.")

except FileNotFoundError:
    # Handle file not found error
    print(f"\n❌ Error: File '{file_path}' not found. ")
    sys.exit()
except KeyError as e:
    # Handle missing column error
    print(f"\n❌ Error: The required column '{e.args[0]}' is missing. ")
    sys.exit()
except Exception as e:
    # Handle any other unexpected errors
    print(f"\n❌ Error: An unexpected error occurred. Details: {e} ")
    sys.exit()
