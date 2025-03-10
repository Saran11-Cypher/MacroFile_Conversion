import os
import pandas as pd
import re
from openpyxl import load_workbook

# ✅ Define Paths
EXCEL_FILE = r"C:\Users\n925072\Downloads\MacroFile_Conversion-master\MacroFile_Conversion-master\New folder\convertor\Macro_Functional_Excel.xlsx"  
UPLOAD_FOLDER = r"C:\1"  # Change to your folder path

# ✅ Check if Upload Folder Exists
if not os.path.exists(UPLOAD_FOLDER):
    print(f"❌ Error: Folder '{UPLOAD_FOLDER}' does not exist.")
    exit()

# ✅ Load Workbook and Sheets
wb = load_workbook(EXCEL_FILE)
ws_main = wb["Main"]
ws_bal = wb["Business Approved List"]

# ✅ Define the correct config load order
config_load_order = [
    "ValueList", "AttributeType", "UserDefinedTerm", "LineOfBusiness",
    "Product", "ServiceCategory", "BenefitNetwork", "NetworkDefinitionComponent",
    "BenefitPlanComponent", "WrapAroundBenefitPlan", "BenefitPlanRider",
    "BenefitPlanTemplate", "Account", "BenefitPlan", "AccountPlanSelection"
]

# ✅ Function to normalize text for consistent comparison
def normalize_text(text):
    """Removes special characters, keeps spaces, and converts to lowercase."""
    return re.sub(r'[^a-zA-Z0-9\s-]', '', str(text)).strip().lower()

# ✅ Load "Business Approved List" into a DataFrame
df_bal = pd.read_excel(EXCEL_FILE, sheet_name="Business Approved List", dtype=str)  
df_bal["Config Type"] = df_bal["Config Type"].astype(str).apply(normalize_text)

# ✅ Get Config Types from Business Approved List
approved_config_types = set(df_bal["Config Type"].dropna().unique())

# ✅ List and Normalize Available Folders
available_folders = {normalize_text(f): os.path.join(UPLOAD_FOLDER, f) 
                     for f in os.listdir(UPLOAD_FOLDER) if os.path.isdir(os.path.join(UPLOAD_FOLDER, f))}

# 🔹 Debugging: Print Folder and Config Type Information
print("\n🔎 Available Folders (Normalized):", available_folders.keys())
print("🔎 Approved Config Types:", approved_config_types)

# ✅ Select only folders present in both the config load order and "Business Approved List"
selected_folders = {config: available_folders[config] for config in config_load_order 
                    if config in available_folders and config in approved_config_types}

# ✅ Error Handling: If No Matching Folders Found
if not selected_folders:
    print("\n❌ Error: No matching config folders found in the parent folder.")
    print("👉 Check if folder names match exactly with Business Approved List.")
    exit()

# ✅ Process Each Selected Folder
for config_type, folder_path in selected_folders.items():
    uploaded_files = [f for f in os.listdir(folder_path) if os.path.isfile(os.path.join(folder_path, f))]
    file_count = len(uploaded_files)

    # ✅ Update the "Main" Sheet Dynamically
    ws_main.append([config_type, file_count, "Pending", "Pending", "Pending"])

# ✅ Assign Order Based on Available Configurations
df_bal["Order"] = df_bal["Config Type"].apply(lambda x: config_load_order.index(x) if x in config_load_order else -1)

# ✅ Validate Order: If Not in Increasing Sequence, Show Error
valid_orders = df_bal[df_bal["Order"] >= 0]["Order"]

if not valid_orders.is_monotonic_increasing:
    print("\n❌ Error: Invalid Order! Please arrange the data correctly.")
    exit()

# ✅ Remove Temporary "Order" Column
df_bal.drop(columns=["Order"], inplace=True)

# ✅ Function to Match Config Names with Uploaded Files
def find_matching_file(config_name, folder_path):
    """Finds files that contain all words from config_name in any order."""
    config_words = normalize_text(config_name).split()

    for filename in os.listdir(folder_path):
        if os.path.isfile(os.path.join(folder_path, filename)):
            cleaned_filename = normalize_text(filename)

            # ✅ Ensure all words in config_name exist in the filename
            if all(word in cleaned_filename for word in config_words):
                return filename  # ✅ Return the first matched file

    return None  # ❌ No match found

# ✅ Check for HRL Availability and Update DataFrame
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

# ✅ Save Updates to Excel
for row_idx, row in df_bal.iterrows():
    for col_idx, value in enumerate(row):
        ws_bal.cell(row=row_idx+2, column=col_idx+1, value=str(value))  # Ensure everything is saved as string

wb.save(EXCEL_FILE)
print("\n✅ Excel file updated successfully!")



🔎 Available Folders (Normalized): dict_keys(['benefitplancomponent', 'benefitplantemplate', 'servicecategory', 'valuelist'])
🔎 Approved Config Types: {'benefitplantemplate', 'servicecategory', 'valuelist', 'benefitplancomponent'}

❌ Error: No matching config folders found in the parent folder.
👉 Check if folder names match exactly with Business Approved List.
