import os
import pandas as pd
import re
from datetime import datetime
import shutil
from openpyxl import load_workbook
from collections import defaultdict

EXCEL_FILE = "C:\\Users\\n925072\\Downloads\\MacroFile_Conversion-master\\MacroFile_Conversion-master\\New folder\\convertor\\Macro_Functional_Excel.xlsx"
UPLOAD_FOLDER = "C:\\1"
DATE_STAMP = datetime.now().strftime("%Y%m%d_%H%M%S")
HRL_PARENT_FOLDER = f"C:\\Datas\\HRLS_{DATE_STAMP}"

# Ensure upload folder exists
if not os.path.exists(UPLOAD_FOLDER):
    print(f"❌ Error: Folder '{UPLOAD_FOLDER}' does not exist.")
    exit()

# Load workbook and sheets
# print("🔄 Loading Excel workbook...")
wb = load_workbook(EXCEL_FILE)
ws_main = wb["Main"]
ws_bal = wb["Business Approved List"]

config_load_order = [
    "ValueList", "AttributeType", "UserDefinedTerm", "LineOfBusiness",
    "Product", "ServiceCategory", "BenefitNetwork", "NetworkDefinitionComponent",
    "BenefitPlanComponent", "WrapAroundBenefitPlan", "BenefitPlanRider",
    "BenefitPlanTemplate", "Account", "BenefitPlan", "AccountPlanSelection"
]

def trim_suffix(filename):
    """Trim the suffix that contains the date."""
    return re.sub(r'\.\d{4}-\d{2}-\d{2}\..*$', '', filename)

def normalize_text(text):
    """Normalize text for matching."""
    return re.sub(r'[^a-zA-Z0-9._-]', '', str(text)).strip().lower()

def extract_date_from_filename(filename):
    """Extract date from filename if available."""
    match = re.search(r'\.(\d{4}-\d{2}-\d{2})\.', filename)
    if match:
        try:
            return datetime.strptime(match.group(1), "%Y-%m-%d")
        except ValueError:
            return None
    return None

# Load BAL sheet
# print("🔄 Loading Business Approved List sheet...")
df_bal = pd.read_excel(EXCEL_FILE, sheet_name="Business Approved List", dtype=str)
df_bal["Config Type"] = df_bal["Config Type"].astype(str).apply(normalize_text)

approved_config_types = set(df_bal["Config Type"].dropna().unique())

print(f"✅ Found {len(approved_config_types)} approved config types.")

available_folders = {normalize_text(f): os.path.join(UPLOAD_FOLDER, f)
                     for f in os.listdir(UPLOAD_FOLDER) if os.path.isdir(os.path.join(UPLOAD_FOLDER, f))}

selected_folders = {config: path for config, path in available_folders.items() if config in approved_config_types}

if not selected_folders:
    print("❌ Error: No matching config folders found inside the parent folder.")
    exit()

print(f"✅ Found {len(selected_folders)} matching folders in the upload directory.")

# STEP 1: Prompt user for version preference at the beginning
while True:
    user_choice = input("\n🔎 Do you want to pick the (L)atest or (O)ldest version for multi-versions? (L/O): ").strip().lower()
    
    if user_choice in ('l', 'o'):
        break
    else:
        print("❗ Invalid input. Please type 'L' for latest or 'O' for oldest.")

# Normalize user choice
selected_version = 'latest' if user_choice == 'l' else 'oldest'

print(f"\n✅ You have selected to pick the **{selected_version.upper()}** version for all files.\n")

# STEP 2: Analyze all files and categorize into single-version and multi-version
multi_version_detected = False
single_version_files = {}
multi_version_files = defaultdict(list)

print("🔄 Analyzing files for version categorization...")

def categorize_files(folder_path):
    single_version_files = {}
    multi_version_files = defaultdict(list)
    
    all_files = 0
    base_name_counter = {}

    for root, dirs, files in os.walk(folder_path):
        for file in files:
            all_files += 1

            parts = file.split('.')
            if len(parts) >= 3:
                config_name = parts[1]  # Pick only the middle part (real config name)
            else:
                continue  # Skip files that don't match the pattern

            # Count how many times each config_name appears
            if config_name in base_name_counter:
                base_name_counter[config_name] += 1
            else:
                base_name_counter[config_name] = 1

            # Categorize into single or multi
            if config_name in single_version_files:
                multi_version_files[config_name].append(file)
                multi_version_files[config_name].append(single_version_files.pop(config_name)[0])
            elif config_name in multi_version_files:
                multi_version_files[config_name].append(file)
            else:
                single_version_files[config_name] = [file]

    return single_version_files, multi_version_files, all_files

single_version_files, multi_version_files, all_files = categorize_files(UPLOAD_FOLDER)

multi_files_count = len(multi_version_files)
print(f"✅ Categorization complete. {len(single_version_files)} single-version files and {multi_files_count} multi-version files found.")

# STEP 3: Update "Main" sheet with file counts
# print("🔄 Updating 'Main' sheet with file counts...")
for config_type, folder_path in selected_folders.items():
    uploaded_files = [f for f in os.listdir(folder_path) if os.path.isfile(os.path.join(folder_path, f))]
    ws_main.append([config_type, len(uploaded_files), "Pending", "Pending", "Pending"])

# Assign order dynamically
# print("🔄 Assigning order to configurations based on the load order...")
df_bal["Order"] = df_bal["Config Type"].apply(lambda x: config_load_order.index(x) if x in config_load_order else -1)
valid_orders = df_bal[df_bal["Order"] >= 0]["Order"]

if not valid_orders.is_monotonic_increasing:
    print("❌ Error: Invalid Order! Please arrange the data correctly.")
    exit()

df_bal.drop(columns=["Order"], inplace=True)

# Function to find matching file
def find_matching_file(config_type, config_name):
    print(f"🔍 Finding matching file for {config_type}.{config_name}...")

    normalized_key = f"{normalize_text(config_type)}.{normalize_text(config_name)}"

    # Step 1: Search in single-version files
    if normalized_key in single_version_files:
        print(f"✅ Found in single-version files: {single_version_files[normalized_key][0]}")
        return single_version_files[normalized_key][0]

    # Step 2: Search in multi-version files
    if normalized_key in multi_version_files:
        candidates = multi_version_files[normalized_key]

        # Sort based on date extracted from filenames
        candidates_with_dates = []
        for file in candidates:
            file_date = extract_date_from_filename(file)
            candidates_with_dates.append((file, file_date))
        
        # Sort candidates by date
        candidates_with_dates.sort(key=lambda x: (x[1] or datetime.min))

        if selected_version == 'latest':
            selected_file = candidates_with_dates[-1][0]
        else:
            selected_file = candidates_with_dates[0][0]

        print(f"✅ Found in multi-version files, selected: {selected_file}")
        return selected_file

    print(f"❌ No matching file found for {config_type}.{config_name}.")
    return None

    # Sort candidates based on date
    candidates = sorted(candidates, key=lambda x: (x[1] or datetime.min))

# Check for HRL availability and copy files
print("🔄 Checking HRL availability and copying files...")
for index, row in df_bal.iterrows():
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

# Write updated DataFrame to BAL sheet
print("🔄 Writing updated DataFrame back to 'Business Approved List' sheet...")
for row_idx, row in df_bal.iterrows():
    for col_idx, value in enumerate(row):
        ws_bal.cell(row=row_idx+2, column=col_idx+1, value=str(value))

wb.save(EXCEL_FILE)
print(f"\n✅ HRL files copied to '{HRL_PARENT_FOLDER}' and Excel file updated successfully!")


python multi_versions.py
✅ Found 1 approved config types.
✅ Found 1 matching folders in the upload directory.

🔎 Do you want to pick the (L)atest or (O)ldest version for multi-versions? (L/O): l

✅ You have selected to pick the **LATEST** version for all files.

🔄 Analyzing files for version categorization...
✅ Categorization complete. 4498 single-version files and 6521 multi-version files found.
🔄 Checking HRL availability and copying files...
🔍 Finding matching file for benefitplan.MAC03ESAPPOMAUnisysCorporation...
Normalized CT: benefitplan and Normalized CN: mac03esappomaunisyscorporation
Expected Pattern: benefitplan.mac03esappomaunisyscorporation
Candidatees:[]
❌ No matching files found for benefitplan - MAC03ESAPPOMAUnisysCorporation.

✅ HRL files copied to 'C:\Datas\HRLS_20250427_183202' and Excel file updated successfully!
