import os
import pandas as pd
import re
from datetime import datetime
import shutil
from openpyxl import load_workbook

# ---------------------------------------------
# Constants
# ---------------------------------------------
EXCEL_FILE = "C:\\Users\\n925072\\Downloads\\MacroFile_Conversion-master\\MacroFile_Conversion-master\\New folder\\convertor\\Macro_Functional_Excel.xlsx"
UPLOAD_FOLDER = "C:\\1"
DATE_STAMP = datetime.now().strftime("%Y%m%d_%H%M%S")
HRL_PARENT_FOLDER = f"C:\\Datas\\HRLS_{DATE_STAMP}"

# ---------------------------------------------
# Ensure Folder Exists
# ---------------------------------------------
if not os.path.exists(UPLOAD_FOLDER):
    print(f"❌ Error: Folder '{UPLOAD_FOLDER}' does not exist.")
    exit()

# ---------------------------------------------
# Load Workbook and Sheets
# ---------------------------------------------
wb = load_workbook(EXCEL_FILE)
ws_main = wb["Main"]
ws_bal = wb["Business Approved List"]

# ---------------------------------------------
# Config Load Order
# ---------------------------------------------
config_load_order = [
    "ValueList", "AttributeType", "UserDefinedTerm", "LineOfBusiness",
    "Product", "ServiceCategory", "BenefitNetwork", "NetworkDefinitionComponent",
    "BenefitPlanComponent", "WrapAroundBenefitPlan", "BenefitPlanRider",
    "BenefitPlanTemplate", "Account", "BenefitPlan", "AccountPlanSelection"
]

# ---------------------------------------------
# Helper Functions
# ---------------------------------------------
def trim_suffix(filename):
    """Remove the date part from the filename."""
    return re.sub(r'\.\d{4}-\d{2}-\d{2}\..*$', '', filename)

def normalize_text(text):
    """Normalize text: remove special characters, lowercase, remove spaces."""
    return re.sub(r'[^a-zA-Z0-9.]', '', str(text)).strip().lower()

def extract_date(filename):
    """Extract date from filename."""
    match = re.search(r'(\d{4}-\d{2}-\d{2})', filename)
    if match:
        return datetime.strptime(match.group(1), "%Y-%m-%d")
    return None

# ---------------------------------------------
# Step 1: Prompt for Multi-Version Choice
# ---------------------------------------------
print("\n⚡ Multi-Version Files Detected!")
print("Choose file selection strategy globally:")
print("1️⃣  Latest Version")
print("2️⃣  Oldest Version")

while True:
    user_choice = input("Enter your choice (1 or 2): ").strip()
    if user_choice in ['1', '2']:
        break
    print("Invalid input. Please enter 1 or 2.")

choose_latest = (user_choice == '1')  # Boolean flag

# ---------------------------------------------
# Step 2: Analyze All Files in UPLOAD_FOLDER
# ---------------------------------------------
print("\n🔍 Analyzing all files...")

# Gather all files from subfolders
all_files = {}
for root, dirs, files in os.walk(UPLOAD_FOLDER):
    for file in files:
        normalized_name = normalize_text(trim_suffix(file))
        full_path = os.path.join(root, file)
        all_files.setdefault(normalized_name, []).append(full_path)

# Separate into single and multi-version dictionaries
single_version_files = {}
multi_version_files = {}

for base_name, file_list in all_files.items():
    if len(file_list) == 1:
        single_version_files[base_name] = file_list[0]
    else:
        multi_version_files[base_name] = file_list

print(f"✅ Single Version Files Detected: {len(single_version_files)}")
print(f"✅ Multi Version Files Detected: {len(multi_version_files)}")

# ---------------------------------------------
# Step 3: Load Business Approved List Sheet
# ---------------------------------------------
df_bal = pd.read_excel(EXCEL_FILE, sheet_name="Business Approved List", dtype=str)
df_bal["Config Type"] = df_bal["Config Type"].astype(str).apply(normalize_text)

# ---------------------------------------------
# Step 4: Validate Config Load Order
# ---------------------------------------------
approved_config_types = set(df_bal["Config Type"].dropna().unique())
available_folders = {normalize_text(f): os.path.join(UPLOAD_FOLDER, f) 
                     for f in os.listdir(UPLOAD_FOLDER) if os.path.isdir(os.path.join(UPLOAD_FOLDER, f))}
selected_folders = {config: path for config, path in available_folders.items() if config in approved_config_types}

if not selected_folders:
    print("❌ Error: No matching config folders found inside the parent folder.")
    exit()

df_bal["Order"] = df_bal["Config Type"].apply(lambda x: config_load_order.index(x) if x in config_load_order else -1)
valid_orders = df_bal[df_bal["Order"] >= 0]["Order"]

if not valid_orders.is_monotonic_increasing:
    print("❌ Error: Invalid Order! Please arrange the data correctly.")
    exit()

df_bal.drop(columns=["Order"], inplace=True)

# ---------------------------------------------
# Step 5: Find and Copy Matching Files
# ---------------------------------------------
def find_matching_file(config_type, config_name):
    normalized_type = normalize_text(config_type)
    normalized_name = normalize_text(config_name)
    expected_pattern = f"{normalized_type}.{normalized_name}"

    # Single Version Match
    if expected_pattern in single_version_files:
        print(f"✅ [Single Version] Found for {expected_pattern}")
        return single_version_files[expected_pattern]

    # Multi Version Match
    elif expected_pattern in multi_version_files:
        files = multi_version_files[expected_pattern]
        files_with_dates = [(f, extract_date(f)) for f in files if extract_date(f) is not None]

        if not files_with_dates:
            print(f"⚠️ Warning: No valid dates found for {expected_pattern}. Picking arbitrarily.")
            return files[0]

        files_with_dates.sort(key=lambda x: x[1], reverse=choose_latest)
        chosen_file = files_with_dates[0][0]
        print(f"✅ [Multi Version] Chose {'Latest' if choose_latest else 'Oldest'} for {expected_pattern}: {os.path.basename(chosen_file)}")
        return chosen_file

    # Not Found
    print(f"❌ No file found for {expected_pattern}")
    return None

# ---------------------------------------------
# Step 6: Update Excel and Copy Files
# ---------------------------------------------
for index, row in df_bal.iterrows():
    config_type = row["Config Type"]
    config_name = row["Config Name"]

    if pd.isna(config_name) or not str(config_name).strip():
        continue

    matching_file_path = find_matching_file(config_type, config_name)

    if matching_file_path:
        df_bal.at[index, "HRL Available?"] = "HRL Found"
        target_folder = os.path.join(HRL_PARENT_FOLDER, config_type)
        os.makedirs(target_folder, exist_ok=True)
        target_path = os.path.join(target_folder, os.path.basename(matching_file_path))
        shutil.copy2(matching_file_path, target_path)
        df_bal.at[index, "File Name is correct in export sheet"] = matching_file_path
    else:
        df_bal.at[index, "HRL Available?"] = "Not Found"

# Update Excel
for row_idx, row in df_bal.iterrows():
    for col_idx, value in enumerate(row):
        ws_bal.cell(row=row_idx+2, column=col_idx+1, value=str(value))

wb.save(EXCEL_FILE)
print(f"\n🎉 HRL files copied to '{HRL_PARENT_FOLDER}' and Excel file updated successfully!")

