# Mass Load Python Scripts — Documentation

This document explains what each of the four Python scripts does, step by step, in detail. These scripts form a **sequential pipeline** for preparing, processing, and reporting Salesforce (ISC) mass load data from Excel workbooks.

---

## Overview — How the Scripts Fit Together

```
1_Validate Files.py   →   2_Generate Queries.py   →   3_Main File.py   →   4_Success and Error file handling.py
     (Pre-check)              (Extract & Query)          (Transform)              (Post-load Reporting)
```

Each script is meant to be run in order. The output of one feeds into the next.

---

## 1. `1_Validate Files.py` — Pre-load File Validation

**Purpose:** Scan all `.xlsx` files in the `~/Downloads` folder and validate their structure and data quality before anything is processed.

### What it does, step by step

#### Step 0 — Discover Files
- Scans `~/Downloads` for all `.xlsx` files.
- Prints a list of all files found so you know what will be validated.
- Loops through each file and runs all the checks below.

---

#### Step 1 — Sheet Presence Check
- Checks that the workbook contains all **required sheets**:
  - `Opportunity`
  - `Opportunity_product`
  - `Opportunity_Team ` *(note: trailing space is intentional)*
  - `Reporting_codes`
- **Auto-renames known variants** before checking:
  - `Opportunity_products` → `Opportunity_product`
  - `Opportunity_Team` → `Opportunity_Team ` (adds the trailing space)
- If a required sheet is missing, it lists the available extra sheets and **asks you which one to rename** to fill the gap. You can also type `skip` to skip renaming a sheet.
- Saves the workbook after any renames.
- Also reports any **extra sheets** beyond the required set.

---

#### Step 2 — Required Column Check
- For each of the two key sheets, checks that all mandatory columns are present (case-insensitive):
  - **Opportunity sheet** must have: `opportunity_legacy_id_c`, `name`, `accountid`, `sales_stage`, `expected_close_date`, `currency_code`, `ownerid`, `OI_Source`, `created_by`
  - **Opportunity_product sheet** must have: `OpportunityId`, `Product`, `product_type`, `unitprice`, `Term`, `Classification Type`, `Type`
- Prints which columns are missing (if any).

---

#### Step 2 (continued) — API Name Conversion from Reference Files
- Loads two reference Excel files:
  - `Product Fields API Name.xlsx` — for the `Opportunity_product` sheet
  - `Opportunity Fields API.xlsx` — for the `Opportunity` sheet
- Each reference file maps human-readable display values (`Values` column) to their Salesforce API names (`API Name` column).
- Goes through every cell in the relevant columns and **converts non-API values to their correct API names** (case-insensitive lookup).
- Saves the corrected file if any conversions were made.

---

#### Step 2.6 — Fix Common API Name Typos
- Checks every cell in both `Opportunity` and `Opportunity_product` sheets for known typos:
  - `AUTOREN_SUBUPGR` → `AUTOREN_SUBUPG`
  - `TERMINATE_SUBUPGR` → `TERMINATE_SUBUPG`
- Fixes and saves the file if any typos were found.

---

#### Step 2.7 — Auto-Cleanup of `(blank)` Values and Mismatched Won/Lost Fields
- Looks at the `Opportunity` sheet and cleans up three types of bad data:
  1. **`(blank)` text values** — if any of the `Won Reason`, `Lost Category`, or `Lost Reason` columns contain the literal text `(blank)`, it clears those cells.
  2. **Lost fields on Won rows** — if `sales_stage` is `won`, it clears `Lost Category` and `Lost Reason` (they should not be filled for won deals).
  3. **Won Reason on Lost rows** — if `sales_stage` is `lost`, it clears `Won Reason` (it should not be filled for lost deals).
- Saves the file if anything was cleaned.

---

#### Step 3 — Invalid API Name Check
- Validates values in these columns against a hard-coded list of **valid Salesforce API values**:
  - `sales_stage`, `OI_Source`, `Classification Type`, `Type`, `Renewal type`, `Renewal Status`, `product_type`, `Won Reason`, `Lost Category`, `Lost Reason`
- The allowed values list covers sales stages (Engage, Qualify, Won, Lost…), renewal types (AUTOREN_12M, TERMINATE…), product types (TYPEESA, TYPEOEM…), classification types (CLASSNEW, CLASSREN…), and won/lost reason codes (WONSKILL, LCDNP, COMPEXP…).
- Reports any values in those columns that are **not** in the approved list, grouped by sheet and column.

---

#### Step 4 — Won/Lost Reason Completeness Check
- Checks for rows where:
  - `sales_stage = won` but `Won Reason` is blank → **flagged as missing**
  - `sales_stage = lost` but `Lost Category` or `Lost Reason` is blank → **flagged as missing**
  - `sales_stage = won` but `Lost Category` or `Lost Reason` is filled → **flagged as invalid** (shown in a table)
  - `sales_stage = lost` but `Won Reason` is filled → **flagged as invalid** (shown in a table)

---

#### Step 5 — Blank Value Check in Important Columns
- Checks these columns for any blank/null values:
  - **Opportunity**: `opportunity_legacy_id_c`, `name`, `accountid`, `sales_stage`, `expected_close_date`, `currency_code`, `ownerid`, `OI_Source`
  - **Opportunity_product**: `Product`, `product_type`, `unitprice`, `Classification Type`
- Filters out completely blank rows before checking, so trailing empty rows don't trigger false positives.

---

#### Step 6 — Duplicate Legacy ID Check
- Checks the `opportunity_legacy_id_c` column in the `Opportunity` sheet for duplicate values.
- First drops completely blank rows and fully duplicate rows, then checks for any repeated legacy IDs.

---

#### Final Summary
- After all files are processed, prints a summary table showing each file name and its overall validation status: `✅ All Good` or `❌ Issues Found`.

---

## 2. `2_Generate Queries.py` — Extract Data and Generate Salesforce SOQL Queries

**Purpose:** Extract key lookup values from all mass load Excel files and generate ready-to-run Salesforce SOQL queries. Also handles renaming downloaded query result CSVs, merging split CSVs, and checking for missing accounts/tags.

### What it does, step by step

#### Setup — Fresh Workspace
- Deletes and recreates the `Extracted_Files/` folder every time the script runs to start clean.

#### Step 1 — Extract Data from Files
- Scans `~/Downloads` for all `.xlsx` files.
- **Auto-renames sheet name variants** before reading (same logic as Script 1):
  - `Opportunity_products` → `Opportunity_product`
  - `Opportunity_Team` → `Opportunity_Team ` (trailing space)
- For each valid file, extracts data into these output sheets in `Extracted_Files/Extracted_data.xlsx`:

| Source Sheet | Source Columns | Output Sheet |
|---|---|---|
| `Opportunity` | `accountid` | `Accounts` |
| `Opportunity` | `ownerid`, `created_by` | `Email_id` |
| `Opportunity` | `currency_code` | `Currency` |
| `Opportunity` | `opportunity_legacy_id_c` | `Legacy_Ids` |
| `Opportunity_product` | `Product` + `product_type` combined as `product_family` | `Product` |
| `Opportunity_Team ` | `Email` | `Email_id` |
| `Reporting_codes` | `tags` + `reporting_codes` columns | `Strategy` |

- Combines data across multiple files (appends to existing sheets if the output file already exists).

#### Step 2 — Clean the Extracted Data
- **Trims whitespace** from all column names and cell values.
  - For `accountid` specifically: removes ALL internal spaces (not just leading/trailing).
- **Expands comma-separated values** in the `Email_id` and `Strategy` sheets — if a cell contains `a,b,c`, it creates separate rows for `a`, `b`, and `c`.
- **Formats Account IDs**: for values starting with `DC`, strips everything after the first `-` (removes country codes). e.g. `DC123-US` → `DC123`.
- **Removes duplicates** in each sheet based on the first column.

#### Step 3 — Create Concatenated File for SQL
- Creates `Extracted_Files/Concatenated_excel_file.xlsx`.
- Wraps every cell value in single quotes and adds a comma: `'value',` — formatted for pasting into SOQL `IN (...)` clauses.

#### Step 4 — Generate Query Files
- Generates 5 SOQL query `.txt` files in `Extracted_Files/Queries/`:

| File | Query Target | Values From |
|---|---|---|
| `1_Account_query.txt` | `Account` — fetches `AccountNumber, Id` | `Accounts` sheet, `accountid` column |
| `2_Userid_query.txt` | `User` — fetches `email, id, Profile.Name, isactive` (excludes Partner Community users, active only) | `Email_id` sheet |
| `3_PricebookEntry_query.txt` | `PricebookEntry` — fetches `Product2.Product_Code_Family__c, CurrencyIsoCode, id, isactive` | `Product` sheet + `Currency` sheet (cross-filter) |
| `4_Strategy_query.txt` | `Strategy__c` — fetches `id, name, Record_Type_Name__c` | `Strategy` sheet |
| `5_Legacy_query.txt` | `Opportunity` — fetches existing legacy IDs | `Legacy_Ids` sheet |

- If a query's values are too large to fit in one query (>80,000 characters), it automatically **splits into multiple part files** (e.g. `1_Account_query_part1.txt`, `1_Account_query_part2.txt`).

#### Step 5 — Interactive Query Copy & CSV Rename Loop
- Asks if you want to copy queries to clipboard.
- Lists all generated query `.txt` files and lets you select one by number.
- **Copies the selected query to the clipboard** (ready to paste into Salesforce Data Loader or Workbench).
- **Waits for the downloaded CSV result** to appear in `~/Downloads` (looks for a file with `bulkQuery` in the name).
- **Auto-renames the downloaded CSV** to a standardised name:

| Query File | Renamed CSV |
|---|---|
| `1_Account_query.txt` | `accounts.csv` |
| `2_Userid_query.txt` | `userid.csv` |
| `3_PricebookEntry_query.txt` | `productfamily.csv` |
| `4_Strategy_query.txt` | `tags.csv` |
| `5_Legacy_query.txt` | `legacyid.csv` |

- You can press `C` to cancel waiting for a CSV.
- After completing userid query: **auto-copies `userid.csv` to `teammember.csv`** (reuses the same user data for team member lookups).

#### Step 6 — Merge Split CSV Files
- Scans `~/Downloads` for CSV files that follow the pattern `prefix1.csv`, `prefix2.csv`, etc.
- Merges all files with the same prefix into a single combined CSV (e.g. `accounts1.csv` + `accounts2.csv` → `accounts.csv`).

#### Step 7 — Missing Accounts and Tags Check (Optional)
- Asks if you want to check for missing accounts and tags.
- **Accounts check**:
  - Compares the extracted `Accounts` list against `accounts.csv` from Salesforce.
  - Produces `Extracted_Files/Accounts_vlookup.xlsx` — each account with its Salesforce `Id` or `Not found in ISC`.
  - Saves accounts not found in ISC to `~/Downloads/Accounts_to_be_imported.xlsx`.
  - Validates account format: `DB-XX` / `DC*` / `C*` are valid formats; anything else is flagged as invalid and saved to `~/Downloads/Invalid_Accounts.xlsx`.
- **Tags/Strategy check**:
  - Compares the extracted `Strategy` list against `tags.csv` from Salesforce.
  - Produces `Extracted_Files/tags_vlookup.xlsx`.
  - Saves missing tags to `~/Downloads/Tags_to_be_inserted.csv` with the full template needed to insert them into Salesforce (includes `RecordTypeId`, `Record_Type_Name__c`, `Active__c`).
- **Legacy ID overlap check**:
  - If `legacyid.csv` exists, scans all Excel files in Downloads and reports how many of their Opportunity Legacy IDs already exist in Salesforce (i.e. opportunities that would be duplicates).
- Prints a final summary of: tags to insert, accounts to import, invalid accounts, and files with no missing accounts.

---

## 3. `3_Main File.py` — Core Mass Load Data Transformation

**Purpose:** The main processing engine. Takes a mass load Excel file, transforms every sheet into clean, Salesforce-ready CSV files, and produces a summary workbook.

### Configuration (top of file)
- `BASE_DIR` — where the iteration folders are created (e.g. `~/Documents/Office Docs/Massload Files/2026`)
- `REF_SUMMARY_FILE_PATH` — template for the summary Excel file
- `product_ref_file` / `opportunity_ref_file` — reference files for API name mapping
- `DOWNLOADS_DIR` — `~/Downloads`

### What it does, step by step

#### Phase 1 — Folder Setup

**Step 1 — Create Iteration Folder**
- Prompts for a folder name (rejects invalid characters like `\/:*?"<>|`).
- Creates the folder inside `BASE_DIR`.
- Creates two subfolders inside it: `Copy files` and `Final iteration files`.

**Step 2 — Move Excel Files from Downloads**
- Moves all `.xlsx`, `.xls`, `.xlsm`, `.xlsb`, `.xltx`, `.xltm` files from `~/Downloads` into the new iteration folder.

**Step 3 — Create Backup Copies**
- For each Excel file, creates a `_Copy` version in the `Copy files` subfolder (e.g. `file.xlsx` → `file_Copy.xlsx`).
- Skips if a copy already exists.

**Step 4 — Create Output Folder Structure**
- For each Excel file, creates a folder named after that file inside `Final iteration files`.
- Each file's folder gets three subfolders: `Removed Rows`, `Success and error files`, `CSV Files`.

**Step 5 — File Selection Loop**
- Lists all files in the `Copy files` folder.
- Lets you pick a file by number to process next.
- After processing one file, asks if you want to process another.

---

#### Phase 2 — Per-File Processing (runs for each selected file)

**Prep — Lowercase Headers & Preserve Numeric Columns**
- Reads all sheets and converts ALL column headers to lowercase.
- Numeric columns (`unitprice`, `expiring amount`, `term`, `expiring term`) are preserved as numbers. Non-numeric values in those columns are reported as warnings.

---

#### Opportunity Sheet Processing

**Step 1 — File Existence Check**
- Confirms the selected file path is valid.

**MPP Column Extraction**
- If `Mpp_Number__c` column exists in the Opportunity sheet, saves `opportunity_legacy_id_c` + `Mpp_Number__c` to `CSV Files/MPP_Column.xlsx` for reference.

**Step 2 — Remove Duplicates and Blank Rows**
- Drops exact duplicate rows.
- Drops rows where all cells are blank.
- Drops rows where `opportunity_legacy_id_c` is blank.
- Trims whitespace from all string columns.
- After cleaning, **exits if any duplicate legacy IDs remain** (to prevent bad data).

**Step 3 — Count Initial Rows**
- Records the initial row count for the Opportunity sheet (used at the end to compare before/after).

**Step 4 — Flag Already-Existing Opportunities**
- Reads `~/Downloads/legacyid.csv` (from Script 2's query).
- Adds an `Already Exist` column: `Already Exist in ISC` or `Does not Exist in ISC` based on whether the legacy ID appears in Salesforce.

**Step 5 — Standardise Email IDs**
- Converts `ownerid` and `created_by` columns to lowercase.
- Fills blank values with the default data migration email: `iscdmig2@in.ibm.com`.
- Reports how many blanks were filled per column.

**Step 6 — Add Pricebook2Id and RecordTypeId**
- Adds two fixed-value columns to every row:
  - `Pricebook2Id = 01s3h000003KXvoAAG`
  - `RecordTypeId = 0123h000000kppcAAA`

**Step 7 — Format the Close Date**
- Converts `expected_close_date` to `YYYY-MM-DD` format using pandas `to_datetime`.
- Exits if any dates are blank or invalid after conversion.

**Step 8 — Add `legacy_opportunity_split_id_c` Column**
- If this column doesn't already exist, creates it as a copy of `opportunity_legacy_id_c`.
- If it does exist but has blank values, exits with an error.

**Step 9 — Trim Account and Owner IDs**
- Trims whitespace from `accountid` and `ownerid`.
- For `accountid`: removes ALL internal spaces (not just leading/trailing).

**Step 10 — Format DC Account Numbers**
- Creates a new `AccountNumber` column.
- For values starting with `DC` or `C`, removes the country code suffix (everything after the first `-`). DB-format values are kept as-is.

**Step 11 — Check for Required CSV Reference Files**
- Confirms `accounts.csv` and `userid.csv` exist in `~/Downloads`.
- Prompts to retry if missing; exits if you decline.

**Step 12 — Map Account IDs (AccountNumber → Salesforce Id)**
- Merges the Opportunity sheet with `accounts.csv` using the `AccountNumber` column.
- If any `AccountNumber` appears multiple times in `accounts.csv` with different Salesforce IDs, **prompts you to manually select which ID to use**.
- Fills unmatched accounts with `Not in ISC`.
- Reports how many accounts were not found.

**Step 13+ — Map User IDs (ownerid/created_by → Salesforce Id)**
- Merges with `userid.csv` for `ownerid` and `created_by` columns.
- Handles duplicate emails in `userid.csv` by asking you to pick the correct ID.
- Fills unmatched users with a default user ID (`0053h000000sdCVAAY`).

**Column Rename**
- Renames columns to their Salesforce API names using a predefined mapping (e.g. `currency_code` → `CurrencyIsoCode`, `sales_stage` → `StageName`, etc.).

**Column Reorder and Save**
- Reorders columns to a specific predefined order for the Salesforce data loader.
- Saves to `CSV Files/1_Opportunity load.csv`.

**Final Row Count Check**
- Compares initial vs final row count and flags if rows were dropped unexpectedly.

---

#### Opportunity_product Sheet Processing

- Reads the `Opportunity_product` sheet.
- Removes blank/duplicate rows.
- Flags rows where the linked Opportunity does not exist in the `Opportunity` sheet (`existing` column).
- Removes rows whose Opportunity was already removed from the main sheet (keeps the sheets in sync).
- Maps PricebookEntry IDs using `productfamily.csv` — matches on `Product-product_type` + `CurrencyIsoCode`. Rows where no pricebook entry is found are flagged.
- Formats `unitprice` and `expiring amount` as plain numbers.
- Renames and reorders columns to Salesforce API names.
- Saves removed rows to `Removed Rows/Removed_Rows - Product.csv`.
- Saves the final output to `CSV Files/2_Opportunity Product load.csv`.

---

#### Opportunity_Team Sheet Processing

- Reads `Opportunity_Team ` sheet.
- Asks if you want to process the team sheet (can skip if empty or not needed).
- Merges email addresses with `userid.csv` to get Salesforce User IDs.
- Removes rows where the linked Opportunity was removed.
- Saves removed rows to `Removed Rows/Removed_Rows - Team.csv`.
- Saves final output to `CSV Files/3_Opportunity Team member Load.csv`.

---

#### Reporting_codes Sheet Processing

- Reads the `Reporting_codes` sheet.
- Looks up strategy/reporting code IDs from `tags.csv`.
- Rows where a strategy code is not found are removed and saved to `Removed Rows/Removed_Rows - ReportingCodes.csv`.
- Removes rows linked to opportunities that were removed.
- Saves final output to `CSV Files/4_Reporting_codes.csv`.

---

#### Tags Sheet Processing

- Processes the tags portion of the `Reporting_codes` sheet separately.
- Same lookup and removal logic as Reporting Codes.
- Saves removed rows to `Removed Rows/Removed_Rows - Tags.csv`.
- Saves final output to `CSV Files/5_Tags.csv`.

---

#### Contact Roles Sheet Processing

- Asks if you want to process Contact Roles.
- If the sheet is empty or you decline, it is skipped.
- Removes rows linked to removed opportunities.
- Saves removed rows to `Removed Rows/Removed_Rows - contact.csv`.
- Saves final output to `CSV Files/6_Contact_Roles.csv`.

---

#### FeedItem Sheet Processing

- Reads the `FeedItem` sheet.
- Asks if you want to process it.
- Removes rows linked to removed opportunities.
- Saves removed rows to `Removed Rows/Removed_Rows - Feed item.csv`.
- Saves final output to `CSV Files/7_Feeditem.csv`.

---

#### Summary File Generation
- Copies the reference summary file template into the output folder, named after the current file.
- Writes key metadata into the `Summary` sheet (e.g. file name, iteration details).
- Reports total elapsed time for the file.
- Asks if you want to process another file (loops back to file selection).

---

## 4. `4_Success and Error file handling.py` — Post-Load Result Processing

**Purpose:** After Salesforce mass load jobs have run, this script renames the raw result CSVs, copies them into the summary workbook, copies removed row data, updates row counts, and prints a final success/failure summary table.

### What it does, step by step

#### Phase 0 — Auto-Rename Raw Result CSVs
- Scans `FOLDER_PATH` (configured at top of file, e.g. `~/Documents/Office docs/Massload Files/2026`) for any CSV files.
- Applies rename rules based on keywords found in the filename:

| If filename contains... | And contains... | Renamed to |
|---|---|---|
| `opportunity insert` | `success` | `opptysuccess.csv` |
| `opportunity insert` | `error` | `opptyerror.csv` |
| `opportunity product insert` | `success` | `productsuccess.csv` |
| `opportunity product insert` | `error` | `producterror.csv` |
| `feed item insert` | `success` | `feedsuccess.csv` |
| `feed item insert` | `error` | `feederror.csv` |
| `opportunity strategy insert` | `success` | `tagssuccess.csv` |
| `opportunity strategy insert` | `error` | `tagserror.csv` |
| `opportunity code insert` | `success` | `codessuccess.csv` |
| `opportunity code insert` | `error` | `codeserror.csv` |
| `team member insert` | `success` | `teamsuccess.csv` |
| `team member insert` | `error` | `teamerror.csv` |

- Prevents overwriting: if a target name already exists, appends `_1`, `_2`, etc.

---

#### Step 1 — Copy/Move Success & Error CSVs
- Opens a **file picker dialog** — you select one or more CSV files (from any location, but usually `~/Downloads`).
- Opens a **folder picker dialog** — you select the `Success and error files` subfolder inside the relevant iteration folder.
- Moves (not copies) all selected CSVs to that folder.
- Reports which files were moved successfully and which failed.

---

#### Step 2 — Write CSV Data into Summary Workbook
- Opens a **file picker dialog** — you select the `_summary file.xlsx` for the current iteration.
- Derives the `Success and error files` folder path automatically from the location of the summary file.
- Loads each of these CSVs from the `Success and error files` folder and writes them into the corresponding sheet in the workbook:

| CSV File | Target Sheet |
|---|---|
| `opptysuccess.csv` | `Opportunity Success` |
| `opptyerror.csv` | `Opportunity Failures` |
| `productsuccess.csv` | `Opportunity Product Success` |
| `producterror.csv` | `Opportunity Product Failures` |
| `teamsuccess.csv` | `Team Member Success` |
| `teamerror.csv` | `Team Member Failure` |
| `codessuccess.csv` | `Reporting Code Success` |
| `codeserror.csv` | `Reporting Code Failure` |
| `tagssuccess.csv` | `Tag Success` |
| `tagserror.csv` | `Tag Failure` |
| `feedsuccess.csv` | `Feed Item Success` |
| `feederror.csv` | `Feed Item Failure` |
| `contactsuccess.csv` | `Contact Success` |
| `contacterror.csv` | `Contact Failure` |

- **Clears existing sheet content** before writing (overwrites, not appends).
- If any sheets already have data, asks you to confirm before proceeding.
- Skips any CSV files that aren't found (reports them at the end).

---

#### Step 3 — Copy Removed Rows into Summary Workbook
- Reads removed row CSVs from the `Removed Rows` subfolder (derived from the summary file path).
- Appends them to the corresponding Failures sheets in the summary workbook:

| CSV File | Target Sheet |
|---|---|
| `Removed_Rows - Oppty.csv` | `Opportunity Failures` |
| `Removed_Rows - Product.csv` | `Opportunity Product Failures` |
| `Removed_Rows - Product (Duplicate Opportunity).csv` | `Opportunity Product Failures` |
| `Removed_Rows - Team.csv` | `Team Member Failure` |
| `Removed_Rows - Team member (Duplicate Opportunity).csv` | `Team Member Failure` |
| `Removed_Rows - ReportingCodes.csv` | `Reporting Code Failure` |
| `Removed_Rows - Codes (Duplicate Opportunity).csv` | `Reporting Code Failure` |
| `Removed_Rows - Tags.csv` | `Tag Failure` |
| `Removed_Rows - Tags (Duplicate Opportunity).csv` | `Tag Failure` |
| `Removed_Rows - Feed Item.csv` | `Feed Item Failure` |
| `Removed_Rows - Contact.csv` | `Contact Failure` |

- Renames the `Reason` or `ERROR` column to `ERRORS` to match the summary file's header.
- Aligns CSV columns to the sheet's existing column order before appending.

---

#### Step 4 — Update Row Counts in Summary Sheet
- Counts the number of data rows (excluding header) in each success/failure sheet.
- Writes the counts into specific cells of the `Summary` sheet:

| Sheet | Success Cell | Failure Cell | Total Cell |
|---|---|---|---|
| Opportunity | E5 | F5 | G5 |
| Opportunity Product | E6 | F6 | G6 |
| Team Member | E7 | F7 | G7 |
| Reporting Code | E8 | F8 | G8 |
| Tag | E9 | F9 | G9 |
| Contact | E10 | F10 | G10 |
| Feed Item | E11 | F11 | G11 |

- Calculates totals (Success + Failures = Total) and writes them into the `G` column cells.

---

#### Final — Print Aggregated Summary Table
- Groups all counts by category.
- Prints a formatted table in the terminal showing, for each load type:
  - Number of Successes
  - Number of Failures
  - Total

Example output:
```
╒══════════════════════╤═══════════╤══════════╤═══════╕
│       Summary        │  Success  │ Failures │ Total │
╞══════════════════════╪═══════════╪══════════╪═══════╡
│     Opportunity      │      150  │        5 │   155 │
│ Opportunity Product  │      320  │        8 │   328 │
│     Team Member      │       45  │        2 │    47 │
│   Reporting Code     │       90  │        1 │    91 │
│         Tag          │       60  │        0 │    60 │
│      Feed Item       │       30  │        3 │    33 │
╘══════════════════════╧═══════════╧══════════╧═══════╛
```

---

## Key Reference Files Required

These files must exist before running Scripts 2 and 3:

| File | Location | Purpose |
|---|---|---|
| `Product Fields API Name.xlsx` | `~/Documents/Office Docs/Massload Files/Reference File/` | Maps product field display names to Salesforce API names |
| `Opportunity Fields API.xlsx` | `~/Documents/Office Docs/Massload Files/Reference File/` | Maps opportunity field display names to Salesforce API names |
| `Reference_Summary_file.xlsx` | `~/Documents/Office Docs/Massload Files/Reference File/` | Template workbook used to generate the per-iteration summary file |

## Key CSV Files Generated During the Process (by Salesforce queries)

These are downloaded from Salesforce and placed in `~/Downloads`:

| File | Contents | Used By |
|---|---|---|
| `accounts.csv` | `AccountNumber`, `Id` | Script 3 — maps account numbers to Salesforce IDs |
| `userid.csv` | `Email`, `Id`, `Profile.Name`, `IsActive` | Script 3 — maps emails to Salesforce User IDs |
| `productfamily.csv` | `Product2.Product_Code_Family__c`, `CurrencyIsoCode`, `Id`, `IsActive` | Script 3 — maps products to PricebookEntry IDs |
| `tags.csv` | `Name`, `Id`, `Record_Type_Name__c` | Script 3 — maps strategy/tag names to Salesforce IDs |
| `legacyid.csv` | `Opportunity_Legacy_Id__c`, `Id` | Script 3 — identifies already-existing opportunities |
