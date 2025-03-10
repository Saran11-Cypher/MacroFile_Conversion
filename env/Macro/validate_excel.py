import os
import pandas as pd
import re
from openpyxl import load_workbook

# Define paths
EXCEL_FILE = "E:\\PYTHON\\Django\\Workspace\\Macro_Generator\\env\\Macro_Functional_Excel.xlsx"  # Update with your actual file path
UPLOAD_FOLDER = "E:\\PYTHON\\ServiceCategory"  # Change to the folder containing uploaded files

# Ensure the folder exists
if not os.path.exists(UPLOAD_FOLDER):
    print(f"❌ Error: Folder '{UPLOAD_FOLDER}' does not exist.")
    exit()

# Load workbook and sheets
wb = load_workbook(EXCEL_FILE)
ws_main = wb["Main"]
ws_bal = wb["Business Approved List"]

# Define the correct config load order
config_load_order = [
    "ValueList", "AttributeType", "UserDefinedTerm", "LineOfBusiness",
    "Product", "ServiceCategory", "BenefitNetwork", "NetworkDefinitionComponent",
    "BenefitPlanComponent", "WrapAroundBenefitPlan", "BenefitPlanRider",
    "BenefitPlanTemplate", "Account", "BenefitPlan", "AccountPlanSelection"
]

# Function to normalize and clean text
def normalize_text(text):
    """Removes special characters, converts to lowercase, and standardizes spaces/hyphens."""
    return re.sub(r'[^a-zA-Z0-9\s-]', '', str(text)).strip().lower()

# Get list of subfolders inside the parent folder
available_folders = {normalize_text(f): f for f in os.listdir(UPLOAD_FOLDER) if os.path.isdir(os.path.join(UPLOAD_FOLDER, f))}

# Load "Business Approved List" into a DataFrame
df_bal = pd.read_excel(EXCEL_FILE, sheet_name="Business Approved List", dtype=str)  # Ensure all columns are strings
df_bal["Config Type"] = df_bal["Config Type"].astype(str).str.strip()

# Filter config_load_order based on what exists in Business Approved List
normalized_config_types = {normalize_text(cfg): cfg for cfg in config_load_order}
normalized_df_bal_types = df_bal["Config Type"].dropna().apply(normalize_text)
available_config_types = [normalized_config_types[cfg] for cfg in normalized_df_bal_types if cfg in normalized_config_types]

# **🔹 Fix: Only check the sequence of available configurations**
expected_order = [cfg for cfg in config_load_order if cfg in available_config_types]

if available_config_types != expected_order:
    print(f"❌ Error: Invalid Order! Expected sequence: {expected_order}, but found {available_config_types}.")
    exit()

# Process each selected folder based on Business Approved List
selected_folders = {
    config: os.path.join(UPLOAD_FOLDER, available_folders[normalize_text(config)])
    for config in available_config_types if normalize_text(config) in available_folders
}

# Validate if required folders exist
if not selected_folders:
    print("❌ Error: No matching config folders found inside the parent folder.")
    exit()

# Update "Main" sheet with folder names in correct order
for config_type, folder_path in selected_folders.items():
    uploaded_files = [f for f in os.listdir(folder_path) if os.path.isfile(os.path.join(folder_path, f))]
    file_count = len(uploaded_files)
    ws_main.append([config_type, file_count, folder_path, "Pending", "Pending"])

# Function to match config names with uploaded files
def find_matching_file(config_name):
    """Finds files that contain all words from the config_name in any order."""
    config_words = normalize_text(config_name).split()  # Normalize config name and split into words
    
    for folder_path in selected_folders.values():
        for filename in os.listdir(folder_path):
            if os.path.isfile(os.path.join(folder_path, filename)):
                cleaned_filename = normalize_text(filename)  # Normalize filename
                if all(word in cleaned_filename for word in config_words):
                    return os.path.join(folder_path, filename)  # Return the first matched file

    return None  # No match found

# Check for HRL availability and update DataFrame
for index, row in df_bal.iterrows():
    config_name = row["Config Name"]
    
    if pd.isna(config_name) or not str(config_name).strip():
        continue  # Skip empty config names

    matching_file = find_matching_file(config_name)

    if matching_file:
        df_bal.at[index, "HRL Available?"] = "HRL Found"
        df_bal.at[index, "File Name is correct in export sheet"] = str(matching_file)  # Ensure string type
    else:
        df_bal.at[index, "HRL Available?"] = "Not Found"

# Save updates to Excel
for row_idx, row in df_bal.iterrows():
    for col_idx, value in enumerate(row):
        ws_bal.cell(row=row_idx+2, column=col_idx+1, value=str(value))  # Ensure everything is saved as string

wb.save(EXCEL_FILE)
print("✅ Excel file updated successfully!")
