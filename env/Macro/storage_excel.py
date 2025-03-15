import os
import shutil
import pandas as pd
import re
from openpyxl import load_workbook

EXCEL_FILE = "C:\\Users\\n925072\\Downloads\\MacroFile_Conversion-master\\MacroFile_Conversion-master\\New folder\\convertor\\Macro_Functional_Excel.xlsx"
UPLOAD_FOLDER = "C:\\1"  # Folder containing uploaded files
PROCESSED_FOLDER = "C:\\Processed_Files"  # Parent folder where organized files will be stored

# Ensure the parent processed folder exists
os.makedirs(PROCESSED_FOLDER, exist_ok=True)

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

# Function to normalize text
def normalize_text(text):
    return re.sub(r'[^a-zA-Z0-9\s-]', '', str(text)).strip().lower()

# Load "Business Approved List" into a DataFrame
df_bal = pd.read_excel(EXCEL_FILE, sheet_name="Business Approved List", dtype=str)
df_bal["Config Type"] = df_bal["Config Type"].astype(str).apply(normalize_text)

# Get approved config types
approved_config_types = set(df_bal["Config Type"].dropna().unique())

# Get existing folders inside the UPLOAD_FOLDER
available_folders = {normalize_text(f): os.path.join(UPLOAD_FOLDER, f) 
                     for f in os.listdir(UPLOAD_FOLDER) if os.path.isdir(os.path.join(UPLOAD_FOLDER, f))}

# Select only folders that match approved config types
selected_folders = {config: path for config, path in available_folders.items() if config in approved_config_types}

if not selected_folders:
    print("❌ Error: No matching config folders found in upload folder.")
    exit()

# Function to match config names with uploaded files
def find_matching_file(config_name, folder_path):
    """Finds files that match the config name (ignoring special characters)."""
    if "&" in config_name:
        config_name = config_name.replace("&", "and")

    normalized_config_name = re.sub(r'[^a-zA-Z0-9]', '', config_name).lower()

    for filename in os.listdir(folder_path):
        if os.path.isfile(os.path.join(folder_path, filename)):
            cleaned_filename = re.sub(r'[^a-zA-Z0-9]', '', filename).lower()
            if normalized_config_name in cleaned_filename:
                return filename  # Return first matched file

    return None  # No match found

# Move files to respective config folders inside PROCESSED_FOLDER
for config_type, folder_path in selected_folders.items():
    processed_config_folder = os.path.join(PROCESSED_FOLDER, config_type)
    os.makedirs(processed_config_folder, exist_ok=True)  # Create config-type folder if not exists

    uploaded_files = [f for f in os.listdir(folder_path) if os.path.isfile(os.path.join(folder_path, f))]
    file_count = len(uploaded_files)

    for file_name in uploaded_files:
        src_path = os.path.join(folder_path, file_name)
        dest_path = os.path.join(processed_config_folder, file_name)

        shutil.move(src_path, dest_path)  # Move file
        print(f"📂 Moved: {file_name} → {processed_config_folder}")

    # Update the "Main" sheet dynamically
    ws_main.append([config_type, file_count, "Pending", "Pending", "Pending"])

# Assign order dynamically based on available configurations
df_bal["Order"] = df_bal["Config Type"].apply(lambda x: config_load_order.index(x) if x in config_load_order else -1)

# Validate order: If not in increasing sequence, show error and exit
valid_orders = df_bal[df_bal["Order"] >= 0]["Order"]

if not valid_orders.is_monotonic_increasing:
    print("❌ Error: Invalid Order! Please arrange the data correctly.")
    exit()

# Remove the temporary "Order" column
df_bal.drop(columns=["Order"], inplace=True)

# Check for HRL availability and update DataFrame
for index, row in df_bal.iterrows():
    config_type = row["Config Type"]
    config_name = row["Config Name"]

    if pd.isna(config_name) or not str(config_name).strip():
        continue  # Skip empty config names

    if config_type in selected_folders:
        matching_file = find_matching_file(config_name, selected_folders[config_type])

        if matching_file:
            df_bal.at[index, "HRL Available?"] = "HRL Found"
            df_bal.at[index, "File Name is correct in export sheet"] = os.path.join(PROCESSED_FOLDER, config_type, matching_file)
        else:
            df_bal.at[index, "HRL Available?"] = "Not Found"

# Save updates to Excel
for row_idx, row in df_bal.iterrows():
    for col_idx, value in enumerate(row):
        ws_bal.cell(row=row_idx+2, column=col_idx+1, value=str(value))  # Ensure everything is saved as string

wb.save(EXCEL_FILE)
print("✅ Files have been stored and Excel updated successfully!")
