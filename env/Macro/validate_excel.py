import os
import shutil
import pandas as pd
import re
from openpyxl import load_workbook

EXCEL_FILE = "C:\\Users\\n925072\\Downloads\\MacroFile_Conversion-master\\MacroFile_Conversion-master\\New folder\\convertor\\Macro_Functional_Excel.xlsx"  # Update with your actual file path
UPLOAD_FOLDER = "C:\\1"  # Change to the folder containing uploaded files
OUTPUT_FOLDER = "C:\\Filtered_Files"  # Folder to store copied files

# Ensure the folder exists
if not os.path.exists(UPLOAD_FOLDER):
    print(f"❌ Error: Folder '{UPLOAD_FOLDER}' does not exist.")
    exit()
if not os.path.exists(OUTPUT_FOLDER):
    os.makedirs(OUTPUT_FOLDER)

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

# Load "Business Approved List" into a DataFrame
df_bal = pd.read_excel(EXCEL_FILE, sheet_name="Business Approved List", dtype=str)  # Ensure all columns are strings
df_bal["Config Type"] = df_bal["Config Type"].astype(str).apply(normalize_text)

# Get the config types mentioned in "Business Approved List"
approved_config_types = set(df_bal["Config Type"].dropna().unique())

# Get list of subfolders inside the parent folder
available_folders = {normalize_text(f): os.path.join(UPLOAD_FOLDER, f) 
                     for f in os.listdir(UPLOAD_FOLDER) if os.path.isdir(os.path.join(UPLOAD_FOLDER, f))}

# Select only folders present in both the config load order and "Business Approved List"
selected_folders = {config: path for config, path in available_folders.items() if config in approved_config_types}

if not selected_folders:
    print("❌ Error: No matching config folders found inside the parent folder.")
    exit()

# Check if any selected folder contains subfolders
for config_type, folder_path in selected_folders.items():
    subfolders = [f for f in os.listdir(folder_path) if os.path.isdir(os.path.join(folder_path, f))]
    if subfolders:
        print(f"❌ Error: The folder '{folder_path}' contains subfolders, which is not allowed.")
        exit()

# Function to find all matching files
def find_all_matching_files(config_name, folder_path):
    """Finds all files that match the config name (ignoring case, special characters, and spacing)."""
    normalized_config_name = re.sub(r'[^a-zA-Z0-9]', '', config_name).lower()
    matching_files = []

    for filename in os.listdir(folder_path):
        if os.path.isfile(os.path.join(folder_path, filename)):
            cleaned_filename = re.sub(r'[^a-zA-Z0-9]', '', filename).lower()
            if normalized_config_name in cleaned_filename:
                matching_files.append(filename)

    return matching_files  # Return all matched files

# Dictionary to store found files
files_to_move = {}

# Search for matching files first
for index, row in df_bal.iterrows():
    config_type = row["Config Type"]
    config_name = row["Config Name"]

    if pd.isna(config_name) or not str(config_name).strip():
        continue  # Skip empty config names

    if config_type in selected_folders:
        matching_files = find_all_matching_files(config_name, selected_folders[config_type])

        if matching_files:
            df_bal.at[index, "HRL Available?"] = "HRL Found"
            df_bal.at[index, "File Name is correct in export sheet"] = "; ".join(matching_files)
            
            # Store files to move later
            files_to_move[config_type] = files_to_move.get(config_type, []) + matching_files

# Now move the files after all searches are complete
for config_type, file_list in files_to_move.items():
    config_folder = os.path.join(OUTPUT_FOLDER, config_type)
    os.makedirs(config_folder, exist_ok=True)
    
    for matching_file in file_list:
        src_path = os.path.join(selected_folders[config_type], matching_file)
        dest_path = os.path.join(config_folder, matching_file)

        if os.path.exists(src_path):
            try:
                shutil.copy2(src_path, dest_path)
            except Exception as e:
                print(f"❌ Error copying file {matching_file}: {e}")
        else:
            print(f"⚠️ Warning: File '{matching_file}' not found in '{selected_folders[config_type]}'")

# Save updates to Excel
for row_idx, row in df_bal.iterrows():
    for col_idx, value in enumerate(row):
        ws_bal.cell(row=row_idx+2, column=col_idx+1, value=str(value))

wb.save(EXCEL_FILE)
print("✅ Excel file updated successfully! Files copied to respective folders.")
