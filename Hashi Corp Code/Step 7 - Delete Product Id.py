import pandas as pd
import os

# Read files
lineitem_df = pd.read_csv(os.path.expanduser("~/Downloads/lineitem export.csv"))
predelete_df = pd.read_csv(os.path.expanduser("~/Downloads/PreDelete_Product.csv"))

# Standardize case and strip spaces
lineitem_df["Lineitem_Legacy_id__c_std"] = (
    lineitem_df["Lineitem_Legacy_id__c"]
    .astype(str)
    .str.strip()
    .str.upper()
)

predelete_df["Delete_Product_std"] = (
    predelete_df["Delete Product"]
    .astype(str)
    .str.strip()
    .str.upper()
)

# Filter matching rows
matched_df = lineitem_df[
    lineitem_df["Lineitem_Legacy_id__c_std"].isin(
        predelete_df["Delete_Product_std"]
    )
]

# Extract Ids only
output_df = matched_df[["Id"]]

# Write output
output_df.to_csv(os.path.expanduser("~/Downloads/Delete_product.csv", index=False))

print("✅ Delete_product.csv created successfully")
