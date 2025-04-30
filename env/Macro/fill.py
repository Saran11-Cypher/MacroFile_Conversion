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
    return re.sub(r'\.\d{4}-\d{2}-\d{2}\..*$', '', filename)

def normalize_text(text):
    return re.sub(r'[^a-zA-Z0-9.]', '', str(text)).strip().lower()

def extract_date_from_filename(filename):
    match = re.search(r'\.(\d{4}-\d{2}-\d{2})\.', filename)
    if match:
        try:
            return datetime.strptime(match.group(1), "%Y-%m-%d")
        except ValueError:
            return None
    return None

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

# Step 1: Prompt user globally for latest/oldest/both
while True:
    user_choice = input(
        "\n🔎 Do you want to pick the (L)atest, (O)ldest, or (B)oth versions for multi-versions? (L/O/B): "
    ).strip().lower()
    if user_choice in ('l', 'o', 'b'):
        break
    else:
        print("❗ Invalid input. Please type 'L' for latest, 'O' for oldest, or 'B' for both.")

if user_choice == 'b':
    selected_version = 'both'
    print("\n✅ You have selected to pick **BOTH** the latest and oldest versions for all files.\n")
elif user_choice == 'l':
    selected_version = 'latest'
    print("\n✅ You have selected to pick the **LATEST** version for all files.\n")
else:
    selected_version = 'oldest'
    print("\n✅ You have selected to pick the **OLDEST** version for all files.\n")


# STEP 2: Analyze and process each config type separately
def categorize_files(folder_path):
    single_version_files = {}
    multi_version_files = defaultdict(list)
    
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            parts = file.split('.')
            if len(parts) >= 3:
                config_name = parts[1]
            else:
                continue
            normalized_config_name = normalize_text(config_name)

            if normalized_config_name in single_version_files:
                multi_version_files[normalized_config_name].append(file)
                multi_version_files[normalized_config_name].append(single_version_files.pop(normalized_config_name)[0])
            elif normalized_config_name in multi_version_files:
                multi_version_files[normalized_config_name].append(file)
            else:
                single_version_files[normalized_config_name] = [file]
    
    return single_version_files, multi_version_files

# Assign order dynamically
df_bal["Order"] = df_bal["Config Type"].apply(lambda x: config_load_order.index(x) if x in config_load_order else -1)
valid_orders = df_bal[df_bal["Order"] >= 0]["Order"]

if not valid_orders.is_monotonic_increasing:
    print("❌ Error: Invalid Order! Please arrange the data correctly.")
    exit()

df_bal.drop(columns=["Order"], inplace=True)

# Function to find matching file
def find_matching_file(config_name, single_version_files, multi_version_files):
    if "&" in config_name:
        config_name = config_name.replace("&", "and")

    normalized_key = normalize_text(config_name)
    print(f"🔍 Finding matching file for {normalized_key}")

    if normalized_key in single_version_files:
        print(f"✅ Found in single-version files: {single_version_files[normalized_key][0]}")
        return [single_version_files[normalized_key][0]]
    elif normalized_key in multi_version_files:
        candidates = multi_version_files[normalized_key]
        candidates_with_dates = []
        for file in candidates:
            file_date = extract_date_from_filename(file)
            candidates_with_dates.append((file, file_date))

        candidates_with_dates.sort(key=lambda x: (x[1] or datetime.min))

        if selected_version == 'latest':
            selected_files = [candidates_with_dates[-1][0]]
        elif selected_version == 'oldest':
            selected_files = [candidates_with_dates[0][0]]
        else:  # both
            selected_files = list({candidates_with_dates[0][0], candidates_with_dates[-1][0]})  # avoid duplicates

        print(f"✅ Found in multi-version files, selected: {selected_files}")
        return selected_files

    print(f"❌ No matching file found for {normalized_key}")
    return []
# Main Loop: Process each config type
print("🔄 Checking HRL availability and copying files...")

for config_type, folder_path in selected_folders.items():
    print(f"\n📂 Processing Config Type: {config_type}")

    # Analyze files for this config type folder
    single_version_files, multi_version_files = categorize_files(folder_path)

    # Filter rows belonging to this config type
    config_type_rows = df_bal[df_bal["Config Type"] == config_type]

    for index, row in config_type_rows.iterrows():
        config_name = row["Config Name"]

        if pd.isna(config_name) or not str(config_name).strip():
            continue

        matching_file = find_matching_file(config_name, single_version_files, multi_version_files)

        if matching_file:
            df_bal.at[index, "HRL Available?"] = "HRL Found"
            source_path = os.path.join(folder_path, matching_file)
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



Traceback (most recent call last):
  File "C:\Users\n925072\Downloads\MacroFile_Conversion-master\MacroFile_Conversion-master\New folder\convertor\fill.py", line 171, in <module>
    source_path = os.path.join(folder_path, matching_file)
  File "C:\Program Files\Python\310\lib\ntpath.py", line 143, in join
    genericpath._check_arg_types('join', path, *paths)
  File "C:\Program Files\Python\310\lib\genericpath.py", line 152, in _check_arg_types
    raise TypeError(f'{funcname}() argument must be str, bytes, or '
TypeError: join() argument must be str, bytes, or os.PathLike object, not 'list'
