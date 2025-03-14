import os
import pandas as pd
import re
from openpyxl import load_workbook

# File paths (Update these paths as needed)
EXCEL_FILE = "C:\\Users\\n925072\\Downloads\\MacroFile_Conversion-master\\New folder\\convertor\\Macro_Functional_Excel.xlsx"
UPLOAD_FOLDER = "C:\\1"

# Ensure folder exists
if not os.path.exists(UPLOAD_FOLDER):
    print(f"❌ Error: Folder '{UPLOAD_FOLDER}' does not exist.")
    exit()

# Load workbook
wb = load_workbook(EXCEL_FILE)
ws_main = wb["Main"]
ws_bal = wb["Business Approved List"]

# Define correct config load order
CONFIG_LOAD_ORDER = [
    "ValueList", "AttributeType", "UserDefinedTerm", "LineOfBusiness",
    "Product", "ServiceCategory", "BenefitNetwork", "NetworkDefinitionComponent",
    "BenefitPlanComponent", "WrapAroundBenefitPlan", "BenefitPlanRider",
    "BenefitPlanTemplate", "Account", "BenefitPlan", "AccountPlanSelection"
]

# Function to clean and normalize text
def normalize_text(text):
    """Removes special characters, converts to lowercase, and trims spaces."""
    return re.sub(r'[^a-zA-Z0-9]', '', str(text)).strip().lower()

# Load Business Approved List as DataFrame
df_bal = pd.read_excel(EXCEL_FILE, sheet_name="Business Approved List", dtype=str)
df_bal["Config Type"] = df_bal["Config Type"].astype(str).apply(normalize_text)

# Get config types from "Business Approved List"
approved_config_types = set(df_bal["Config Type"].dropna().unique())

# Get list of available folders in UPLOAD_FOLDER
available_folders = {normalize_text(f): os.path.join(UPLOAD_FOLDER, f) 
                     for f in os.listdir(UPLOAD_FOLDER) if os.path.isdir(os.path.join(UPLOAD_FOLDER, f))}

# Select matching folders
selected_folders = {config: path for config, path in available_folders.items() if config in approved_config_types}

if not selected_folders:
    print("❌ Error: No matching config folders found inside the parent folder.")
    exit()

# Process each folder
for config_type, folder_path in selected_folders.items():
    uploaded_files = [f for f in os.listdir(folder_path) if os.path.isfile(os.path.join(folder_path, f))]
    file_count = len(uploaded_files)

    # Append data to "Main" sheet
    ws_main.append([config_type, file_count, "Pending", "Pending", "Pending"])

# Validate Config Order
df_bal["Order"] = df_bal["Config Type"].apply(lambda x: CONFIG_LOAD_ORDER.index(x) if x in CONFIG_LOAD_ORDER else -1)
valid_orders = df_bal[df_bal["Order"] >= 0]["Order"]

if not valid_orders.is_monotonic_increasing:
    print("❌ Error: Invalid Order! Please arrange the data correctly.")
    exit()

df_bal.drop(columns=["Order"], inplace=True)  # Remove temp column

# Function to match config names with uploaded files
def find_matching_file(config_name, folder_path):
    """Strictly matches filenames against the config name and logs mismatches."""
    normalized_config_name = re.sub(r'[^a-zA-Z0-9]', '', config_name).lower()
    matched_files = []

    for filename in os.listdir(folder_path):
        if os.path.isfile(os.path.join(folder_path, filename)):
            cleaned_filename = re.sub(r'[^a-zA-Z0-9]', '', filename).lower()

            # Exact match check
            if normalized_config_name in cleaned_filename:
                matched_files.append(filename)

    if matched_files:
        print(f"✅ Matched: {config_name} → {matched_files[0]}")
        return matched_files[0]  # Return the first matched file
    else:
        print(f"❌ No match for: {config_name} in {folder_path}")
        return None


# Match HRL files
for index, row in df_bal.iterrows():
    config_type = row["Config Type"]
    config_name = row["Config Name"]

    if pd.isna(config_name) or not str(config_name).strip():
        continue  # Skip empty config names

    if config_type in selected_folders:
        matching_file = find_matching_file(config_name, selected_folders[config_type])

        if matching_file:
            df_bal.at[index, "HRL Available?"] = "HRL Found"
            df_bal.at[index, "File Name is correct in export sheet"] = os.path.join(selected_folders[config_type], matching_file)
        else:
            df_bal.at[index, "HRL Available?"] = "Not Found"

# Save updates to Excel
for row_idx, row in df_bal.iterrows():
    for col_idx, value in enumerate(row):
        ws_bal.cell(row=row_idx+2, column=col_idx+1, value=str(value))

wb.save(EXCEL_FILE)
print("✅ Excel file updated successfully!")
