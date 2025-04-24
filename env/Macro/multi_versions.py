import os
import pandas as pd
import re
import sys
import shutil
from datetime import datetime
from tqdm import tqdm
from openpyxl import load_workbook

# === Configuration ===
EXCEL_FILE = "C:\\Users\\n925072\\Downloads\\MacroFile_Conversion-master\\MacroFile_Conversion-master\\New folder\\convertor\\Macro_Functional_Excel.xlsx"
UPLOAD_FOLDER = "C:\\1"
DATE_STAMP = datetime.now().strftime("%Y%m%d_%H%M%S")
HRL_PARENT_FOLDER = f"C:\\Datas\\HRLS_{DATE_STAMP}"

# === Helper Functions ===
def normalize_text(text):
    return re.sub(r'[^a-zA-Z0-9.]', '', str(text)).strip().lower()

def trim_suffix(filename):
    return re.sub(r'\.\d{4}-\d{2}-\d{2}\..*$', '', filename)

def extract_date(filename):
    match = re.search(r'\.(\d{4}-\d{2}-\d{2})\.', filename)
    if match:
        return datetime.strptime(match.group(1), "%Y-%m-%d")
    return None

def prompt_file_version_choice():
    print("⚠️ Multiple files with the same base name but different dates were detected.")
    print("How would you like to proceed for such files?")
    print("1. Choose the latest version (newest date)")
    print("2. Choose the oldest version (earliest date)")
    choice = input("Enter your choice (1 or 2): ").strip()
    if choice == "1":
        return "latest"
    elif choice == "2":
        return "oldest"
    else:
        print("❌ Invalid input. Exiting...")
        sys.exit()

# === Version Preference Prompt ===
version_preference = prompt_file_version_choice()

def find_matching_file(config_type, config_name, folder_path):
    if "&" in config_name:
        config_name = config_name.replace("&", "and")
    normalized_config_type = normalize_text(config_type)
    normalized_config_name = normalize_text(config_name)

    candidates = []
    for file in os.listdir(folder_path):
        if not os.path.isfile(os.path.join(folder_path, file)):
            continue
        base = normalize_text(trim_suffix(file))
        if base == f"{normalized_config_type}.{normalized_config_name}":
            candidates.append(file)

    if not candidates:
        print("❌ No exact match found")
        return None

    if len(candidates) == 1:
        print(f"✅ Match Found: {candidates[0]}")
        return candidates[0]

    candidates.sort(key=lambda x: extract_date(x) or datetime.min, reverse=(version_preference == "latest"))
    print(f"⚠️ Multiple versions found. {version_preference.capitalize()} version selected: {candidates[0]}")
    return candidates[0]

# === Initialization ===
if not os.path.exists(UPLOAD_FOLDER):
    print(f"❌ Error: Folder '{UPLOAD_FOLDER}' does not exist.")
    exit()

wb = load_workbook(EXCEL_FILE)
ws_main = wb["Main"]
ws_bal = wb["Business Approved List"]

config_load_order = [
    "ValueList", "AttributeType", "UserDefinedTerm", "LineOfBusiness",
    "Product", "ServiceCategory", "BenefitNetwork", "NetworkDefinitionComponent",
    "BenefitPlanComponent", "WrapAroundBenefitPlan", "BenefitPlanRider",
    "BenefitPlanTemplate", "Account", "BenefitPlan", "AccountPlanSelection"
]

df_bal = pd.read_excel(EXCEL_FILE, sheet_name="Business Approved List", dtype=str)
df_bal["Config Type"] = df_bal["Config Type"].astype(str).apply(normalize_text)

approved_config_types = set(df_bal["Config Type"].dropna().unique())
available_folders = {
    normalize_text(f): os.path.join(UPLOAD_FOLDER, f)
    for f in os.listdir(UPLOAD_FOLDER)
    if os.path.isdir(os.path.join(UPLOAD_FOLDER, f))
}
selected_folders = {
    config: path for config, path in available_folders.items() if config in approved_config_types
}

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

for index, row in tqdm(df_bal.iterrows(), total=len(df_bal), desc="🔍 Matching Files"):
    config_type = row["Config Type"]
    config_name = row["Config Name"]

    if pd.isna(config_name) or not str(config_name).strip():
        continue

    if config_type in selected_folders:
        matching_file = find_matching_file(config_type, config_name, selected_folders[config_type])
        if matching_file:
            df_bal.at[index, "HRL Available?"] = "HRL Found"
            source_path = os.path.join(selected_folders[config_type], matching_file)
            target_folder = os.path.join(HRL_PARENT_FOLDER, config_type)
            os.makedirs(target_folder, exist_ok=True)
            target_path = os.path.join(target_folder, matching_file)
            shutil.copy2(source_path, target_path)
            df_bal.at[index, "File Name is correct in export sheet"] = source_path
        else:
            df_bal.at[index, "HRL Available?"] = "Not Found"

# Write updated DataFrame back to Excel
for row_idx, row in df_bal.iterrows():
    for col_idx, value in enumerate(row):
        ws_bal.cell(row=row_idx + 2, column=col_idx + 1, value=str(value))

wb.save(EXCEL_FILE)
print(f"\n✅ HRL files copied to '{HRL_PARENT_FOLDER}' and Excel file updated successfully!")
