import os
import pandas as pd
import re
from openpyxl import load_workbook

# 🔹 Define Constants
EXCEL_FILE = r"C:\Users\n925072\Downloads\Macro_Functional_Excel.xlsx"  # Update with actual file path
UPLOAD_FOLDER = r"C:\1"  # Update with your upload folder path

# ✅ Ensure the folder exists
if not os.path.exists(UPLOAD_FOLDER):
    print(f"❌ Error: Folder '{UPLOAD_FOLDER}' does not exist.")
    exit()

# ✅ Load workbook and sheets
wb = load_workbook(EXCEL_FILE)
ws_main = wb["Main"]
ws_bal = wb["Business Approved List"]

# ✅ Define the correct config load order (Normalized)
config_load_order = [
    "valuelist", "attributetype", "userdefinedterm", "lineofbusiness",
    "product", "servicecategory", "benefitnetwork", "networkdefinitioncomponent",
    "benefitplancomponent", "wraparoundbenefitplan", "benefitplanrider",
    "benefitplantemplate", "account", "benefitplan", "accountplanselection"
]

# ✅ Function to normalize text for consistent matching
def normalize_text(text):
    """Converts to lowercase, removes special characters, and trims spaces."""
    return re.sub(r'[^a-zA-Z0-9]', '', str(text)).strip().lower()

# ✅ Load "Business Approved List" into a DataFrame
df_bal = pd.read_excel(EXCEL_FILE, sheet_name="Business Approved List", dtype=str)
df_bal["Config Type"] = df_bal["Config Type"].astype(str).apply(normalize_text)

# ✅ Get the config types mentioned in "Business Approved List"
approved_config_types = set(df_bal["Config Type"].dropna().unique())

# ✅ Get list of subfolders inside the parent folder (UPLOAD_FOLDER)
available_folders = {normalize_text(f): os.path.join(UPLOAD_FOLDER, f) 
                     for f in os.listdir(UPLOAD_FOLDER) if os.path.isdir(os.path.join(UPLOAD_FOLDER, f))}

# 🔍 Debugging Outputs
print(f"🔎 Available Folders (Normalized): {list(available_folders.keys())}")
print(f"🔎 Approved Config Types: {approved_config_types}")

# ✅ Select only folders that match both the config load order and Business Approved List
selected_folders = {config: available_folders[config] for config in config_load_order 
                    if config in available_folders and config in approved_config_types}

# 🚨 Error Handling: No matching folders
if not selected_folders:
    print("❌ Error: No matching config folders found in the parent folder.")
    print("👉 Check if folder names match exactly with Business Approved List.")
    exit()

# ✅ Process each selected folder
for config_type, folder_path in selected_folders.items():
    uploaded_files = [f for f in os.listdir(folder_path) if os.path.isfile(os.path.join(folder_path, f))]
    file_count = len(uploaded_files)

    # 📝 Update the "Main" sheet dynamically
    ws_main.append([config_type, file_count, "Pending", "Pending", "Pending"])

# ✅ Assign order dynamically based on available configurations
df_bal["Order"] = df_bal["Config Type"].apply(lambda x: config_load_order.index(x) if x in config_load_order else -1)

# 🚨 Validate Order: Ensure it's in increasing sequence
valid_orders = df_bal[df_bal["Order"] >= 0]["Order"]

if not valid_orders.is_monotonic_increasing:
    print("❌ Error: Invalid Order! Please arrange the data correctly.")
    exit()

# 🛑 Remove the temporary "Order" column (not needed in final output)
df_bal.drop(columns=["Order"], inplace=True)

# ✅ Function to match config names with uploaded files
def find_matching_file(config_name, folder_path):
    """Finds files that contain all words from the config_name in any order."""
    config_words = normalize_text(config_name).split()

    for filename in os.listdir(folder_path):
        if os.path.isfile(os.path.join(folder_path, filename)):
            cleaned_filename = normalize_text(filename)

            # Ensure all words in config_name exist in the filename
            if all(word in cleaned_filename for word in config_words):
                return filename  # Return the first matched file

    return None  # No match found

# ✅ Check for HRL availability and update DataFrame
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

# ✅ Save updates to Excel
for row_idx, row in df_bal.iterrows():
    for col_idx, value in enumerate(row):
        ws_bal.cell(row=row_idx+2, column=col_idx+1, value=str(value))  # Ensure everything is saved as string

wb.save(EXCEL_FILE)
print("✅ Excel file updated successfully!")
