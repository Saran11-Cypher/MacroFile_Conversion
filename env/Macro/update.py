import os
import pandas as pd
import re
import shutil
from datetime import datetime
from openpyxl import load_workbook

EXCEL_FILE = "C:\\Users\\n925072\\Downloads\\MacroFile_Conversion-master\\MacroFile_Conversion-master\\New folder\\convertor\\Macro_Functional_Excel.xlsx"
UPLOAD_FOLDER = "C:\\1"
DATE_STAMP = datetime.now().strftime("%Y%m%d_%H%M%S")  # New parent folder with date and time stamp
HRL_PARENT_FOLDER = f"C:\\Datas\\HRLS_{DATE_STAMP}"  # Updated parent folder structure

if not os.path.exists(UPLOAD_FOLDER):
    print(f"❌ Error: Folder '{UPLOAD_FOLDER}' does not exist.")
    exit()

os.makedirs(HRL_PARENT_FOLDER, exist_ok=True)  # Ensure the parent folder exists

wb = load_workbook(EXCEL_FILE)
ws_main = wb["Main"]
ws_bal = wb["Business Approved List"]

config_load_order = [
    "ValueList", "AttributeType", "UserDefinedTerm", "LineOfBusiness",
    "Product", "ServiceCategory", "BenefitNetwork", "NetworkDefinitionComponent",
    "BenefitPlanComponent", "WrapAroundBenefitPlan", "BenefitPlanRider",
    "BenefitPlanTemplate", "Account", "BenefitPlan", "AccountPlanSelection"
]

def normalize_text(text):
    """Normalize text by converting to lowercase, replacing spaces with hyphens, and removing special characters."""
    return re.sub(r'[^a-zA-Z0-9\s-]', '', str(text)).strip().lower().replace(" ","-")

df_bal = pd.read_excel(EXCEL_FILE, sheet_name="Business Approved List", dtype=str)
df_bal["Config Type"] = df_bal["Config Type"].astype(str).apply(normalize_text)

approved_config_types = set(df_bal["Config Type"].dropna().unique())

available_folders = {normalize_text(f): os.path.join(UPLOAD_FOLDER, f) 
                     for f in os.listdir(UPLOAD_FOLDER) if os.path.isdir(os.path.join(UPLOAD_FOLDER, f))}

selected_folders = {config: path for config, path in available_folders.items() if config in approved_config_types}

if not selected_folders:
    print("❌ Error: No matching config folders found inside the parent folder.")
    exit()

for config_type, folder_path in selected_folders.items():
    uploaded_files = [f for f in os.listdir(folder_path) if os.path.isfile(os.path.join(folder_path, f))]
    file_count = len(uploaded_files)
    ws_main.append([config_type, file_count, "Pending", "Pending", "Pending"])

df_bal["Order"] = df_bal["Config Type"].apply(lambda x: config_load_order.index(x) if x in config_load_order else -1)

valid_orders = df_bal[df_bal["Order"] >= 0]["Order"]

if not valid_orders.is_monotonic_increasing:
    print("❌ Error: Invalid Order! Please arrange the data correctly.")
    exit()

df_bal.drop(columns=["Order"], inplace=True)

def find_matching_file(config_name, file_list):
    """Find the best match for the given config name within a folder."""
    normalized_config_name = normalize_text(config_name)
    config_words = normalized_config_name.split()  # Convert to a list of words

    exact_matches = []
    prefixed_matches = []

    for file_name in file_list:
        normalized_file_name = normalize_text(file_name)
        file_words = normalized_file_name.split()  # Convert filename to words

        try:
            index = file_words.index(config_words[0])  # Check first word of config in filename
            matched_sequence = file_words[index:index + len(config_words)]

            if matched_sequence == config_words:
                if index == 0:  # No prefix before config name
                    exact_matches.append(file_name)
                else:  # Config name exists but has prefixes
                    prefixed_matches.append(file_name)
        except ValueError:
            continue  # Skip if config name not found

    # Prioritize exact matches (no prefix before config name)
    if exact_matches:
        return exact_matches[0]
    elif prefixed_matches:
        return prefixed_matches[0]  # Select prefixed match if no exact match found

    return None  # No match found

for index, row in df_bal.iterrows():
    config_type = row["Config Type"]
    config_name = row["Config Name"]

    if pd.isna(config_name) or not str(config_name).strip():
        continue

    if config_type in selected_folders:
        folder_path = selected_folders[config_type]
        uploaded_files = [f for f in os.listdir(folder_path) if os.path.isfile(os.path.join(folder_path, f))]

        matching_file = find_matching_file(config_name, uploaded_files)

        if matching_file:
            df_bal.at[index, "HRL Available?"] = "HRL Found"
            source_path = os.path.join(folder_path, matching_file)
            target_folder = os.path.join(HRL_PARENT_FOLDER, config_type)
            os.makedirs(target_folder, exist_ok=True)
            target_path = os.path.join(target_folder, matching_file)
            
            shutil.copy2(source_path, target_path)  # Copy HRL file to the new parent folder
            df_bal.at[index, "File Name is correct in export sheet"] = source_path  # Keep original path
        else:
            df_bal.at[index, "HRL Available?"] = "Not Found"

# Save DataFrame back to Excel
for row_idx, row in df_bal.iterrows():
    for col_idx, value in enumerate(row):
        ws_bal.cell(row=row_idx+2, column=col_idx+1, value=str(value))

wb.save(EXCEL_FILE)
print(f"✅ HRL files copied to '{HRL_PARENT_FOLDER}' and Excel file updated successfully.")
