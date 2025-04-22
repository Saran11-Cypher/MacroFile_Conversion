import os
import pandas as pd
import re
from datetime import datetime
import shutil
from openpyxl import load_workbook
from tqdm import tqdm  # Ensure tqdm is installed

EXCEL_FILE = "C:\\Users\\n925072\\Downloads\\MacroFile_Conversion-master\\MacroFile_Conversion-master\\New folder\\convertor\\Macro_Functional_Excel.xlsx"
UPLOAD_FOLDER = "C:\\1"
DATE_STAMP = datetime.now().strftime("%Y%m%d_%H%M%S")
HRL_PARENT_FOLDER = f"C:\\Datas\\HRLS_{DATE_STAMP}"

if not os.path.exists(UPLOAD_FOLDER):
    print(f"❌ Error: Folder '{UPLOAD_FOLDER}' does not exist.")
    exit()

# Load workbook and sheets
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
    return re.sub(r'[^a-zA-Z0-9.]','', str(text)).strip().lower()

def get_base_filename(filename):
    return re.sub(r'\.\d{4}-\d{2}-\d{2}\..*$', '', filename)

def extract_date_from_filename(filename):
    match = re.search(r'\.(\d{4}-\d{2}-\d{2})\.', filename)
    return match.group(1) if match else None

def analyze_files(selected_folders):
    file_map = {}
    print("🔍 Scanning folders for files...")
    for folder_name, folder_path in tqdm(selected_folders.items(), desc="Analyzing folders"):
        for file in os.listdir(folder_path):
            full_path = os.path.join(folder_path, file)
            if os.path.isfile(full_path):
                base_name = get_base_filename(file)
                file_map.setdefault(base_name, []).append((file, full_path))
    duplicates = {base: versions for base, versions in file_map.items() if len(versions) > 1}
    return file_map, duplicates

def prompt_user_for_version_choice():
    version_choice = ""
    while version_choice not in ["latest", "oldest"]:
        version_choice = input(
            "\n⚠️ Found multiple files with same name but different dates.\n"
            "Would you like to proceed with the latest or oldest version? [latest/oldest]: "
        ).strip().lower()
    print(f"\n✅ Your choice: Proceed with the **{version_choice}** version for all duplicates.\n")
    return version_choice

def select_files(file_map, duplicates_with_versions, version_choice):
    selected_files = {}
    for base_name, versions in file_map.items():
        if len(versions) == 1:
            selected_files[base_name] = versions[0][1]
        else:
            sorted_versions = sorted(
                versions,
                key=lambda x: extract_date_from_filename(x[0]) or "",
                reverse=(version_choice == "latest")
            )
            selected_files[base_name] = sorted_versions[0][1]
    return selected_files

# Load Business Approved List
df_bal = pd.read_excel(EXCEL_FILE, sheet_name="Business Approved List", dtype=str)
df_bal["Config Type"] = df_bal["Config Type"].astype(str).apply(normalize_text)

approved_config_types = set(df_bal["Config Type"].dropna().unique())

available_folders = {
    normalize_text(f): os.path.join(UPLOAD_FOLDER, f)
    for f in os.listdir(UPLOAD_FOLDER)
    if os.path.isdir(os.path.join(UPLOAD_FOLDER, f))
}

selected_folders = {
    config: path for config, path in available_folders.items()
    if config in approved_config_types
}

if not selected_folders:
    print("❌ Error: No matching config folders found inside the parent folder.")
    exit()

# Analyze files and handle duplicates
file_map, duplicates_with_versions = analyze_files(selected_folders)
version_choice = prompt_user_for_version_choice() if duplicates_with_versions else "latest"
selected_files_map = select_files(file_map, duplicates_with_versions, version_choice)

# Count files and update Main sheet
for config_type, folder_path in selected_folders.items():
    uploaded_files = [f for f in os.listdir(folder_path) if os.path.isfile(os.path.join(folder_path, f))]
    ws_main.append([config_type, len(uploaded_files), "Pending", "Pending", "Pending"])

# Validate config order
df_bal["Order"] = df_bal["Config Type"].apply(lambda x: config_load_order.index(x) if x in config_load_order else -1)
valid_orders = df_bal[df_bal["Order"] >= 0]["Order"]

if not valid_orders.is_monotonic_increasing:
    print("❌ Error: Invalid Order! Please arrange the data correctly.")
    exit()
df_bal.drop(columns=["Order"], inplace=True)

# Match config files to selected files
def find_matching_file(config_type, config_name):
    if "&" in config_name:
        config_name = config_name.replace("&", "and")

    normalized_config_type = normalize_text(config_type)
    normalized_config_name = normalize_text(config_name)
    expected_pattern = f"{normalized_config_type}.{normalized_config_name}"
    return selected_files_map.get(expected_pattern, None)

for index, row in df_bal.iterrows():
    config_type = row["Config Type"]
    config_name = row["Config Name"]

    if pd.isna(config_name) or not str(config_name).strip():
        continue

    if config_type in selected_folders:
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

# Update Excel sheet
for row_idx, row in df_bal.iterrows():
    for col_idx, value in enumerate(row):
        ws_bal.cell(row=row_idx+2, column=col_idx+1, value=str(value))

wb.save(EXCEL_FILE)
print(f"\n✅ HRL files copied to '{HRL_PARENT_FOLDER}' and Excel file updated successfully!")
