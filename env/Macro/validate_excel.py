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
    return re.sub(r'[^a-zA-Z0-9\s-]', '', str(text)).strip().lower()

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


def find_matching_file(config_name, config_type, folder_path):
    if "&" in config_name:
        config_name = config_name.replace("&", "and")

    # Normalize config name (remove special characters, lowercase)
    normalized_config_name = re.sub(r'[^a-zA-Z0-9]', '', config_name).lower()
    search_pattern = f"{config_type.lower()}.{normalized_config_name}"  # Creating exact match pattern

    # ✅ Use a dictionary to map normalized filenames to original filenames
    filename_mapping = {}

    for filename in os.listdir(folder_path):
        if os.path.isfile(os.path.join(folder_path, filename)):
            cleaned_filename = re.sub(r'[^a-zA-Z0-9]', '', filename).lower()
            filename_mapping[cleaned_filename] = filename  # ✅ Store original filename

    # Search for an exact match
    for cleaned_filename in filename_mapping:
        if cleaned_filename.startswith(search_pattern):
            return filename_mapping[cleaned_filename]  # ✅ Return original filename

    return None  # Return if no exact match is found



for index, row in df_bal.iterrows():
    config_type = row["Config Type"]
    config_name = row["Config Name"]

    if pd.isna(config_name) or not str(config_name).strip():
        continue

    if config_type in selected_folders:
        matching_file = find_matching_file(config_name, selected_folders[config_type])

        if matching_file:
            df_bal.at[index, "HRL Available?"] = "HRL Found"
            source_path = os.path.join(selected_folders[config_type], matching_file)
            target_folder = os.path.join(HRL_PARENT_FOLDER, config_type)
            os.makedirs(target_folder, exist_ok=True)
            target_path = os.path.join(target_folder, matching_file)
            
            shutil.copy2(source_path, target_path)  # Copy HRL file to the new parent folder
            df_bal.at[index, "File Name is correct in export sheet"] = source_path  # Keep original path
        else:
            df_bal.at[index, "HRL Available?"] = "Not Found"

for row_idx, row in df_bal.iterrows():
    for col_idx, value in enumerate(row):
        ws_bal.cell(row=row_idx+2, column=col_idx+1, value=str(value))

wb.save(EXCEL_FILE)
print(f"✅ HRL files copied to '{HRL_PARENT_FOLDER}' and Excel file updated successfully!")

Config Names:
Chiropractor Supplier Classification - Pend
Vision - Exam
Laboratory - Diagnostic
Surgery
Diagnostic Testing
X-ray - Diagnostic
BRLOAPEND
Highest Copay - Medicare - INN - Benefit Labels - Practitioner
Global Exclusion (Allows OTC Female Contraceptives) 
Global Exclusion (Allows OTC Male & Female Contraceptives) - MO
Global Exclusion - Medicare
Physical and Occupational Therapy (Excludes Home) - Medicare Outpatient Definition (Alternate Office & Clinic Definition) - INN (Copay Deductible Waived)
Physical and Occupational Therapy (Excludes Home) - Medicare Outpatient Definition - INN (Copay)
Physical and Occupational Therapy (Excludes Home) - Medicare Outpatient Definition - INN (Copay Deductible Waived)
Gold S - HMO (FL) 2024


🔹 Searching for: servicecategory.chiropractor supplier classification - pend.YYYY-MM-DD.a.hrl
🔹 Expected Pattern: ^servicecategory\.chiropractor supplier classification - pend\.\d{4}-\d{2}-\d{2}\.a\.hrl$

❌ No exact match found!

🔹 Searching for: servicecategory.vision - exam.YYYY-MM-DD.a.hrl
🔹 Expected Pattern: ^servicecategory\.vision - exam\.\d{4}-\d{2}-\d{2}\.a\.hrl$

❌ No exact match found!

🔹 Searching for: servicecategory.laboratory - diagnostic.YYYY-MM-DD.a.hrl
🔹 Expected Pattern: ^servicecategory\.laboratory - diagnostic\.\d{4}-\d{2}-\d{2}\.a\.hrl$

❌ No exact match found!

🔹 Searching for: servicecategory.surgery.YYYY-MM-DD.a.hrl
🔹 Expected Pattern: ^servicecategory\.surgery\.\d{4}-\d{2}-\d{2}\.a\.hrl$

❌ No exact match found!

🔹 Searching for: servicecategory.diagnostic testing.YYYY-MM-DD.a.hrl
🔹 Expected Pattern: ^servicecategory\.diagnostic testing\.\d{4}-\d{2}-\d{2}\.a\.hrl$

❌ No exact match found!

🔹 Searching for: servicecategory.x-ray - diagnostic.YYYY-MM-DD.a.hrl
🔹 Expected Pattern: ^servicecategory\.x-ray - diagnostic\.\d{4}-\d{2}-\d{2}\.a\.hrl$

❌ No exact match found!

🔹 Searching for: servicecategory.brloapend.YYYY-MM-DD.a.hrl
🔹 Expected Pattern: ^servicecategory\.brloapend\.\d{4}-\d{2}-\d{2}\.a\.hrl$

❌ No exact match found!

🔹 Searching for: servicecategory.highest copay - medicare - inn - benefit labels - practitioner.YYYY-MM-DD.a.hrl
🔹 Expected Pattern: ^servicecategory\.highest copay - medicare - inn - benefit labels - practitioner\.\d{4}-\d{2}-\d{2}\.a\.hrl$

❌ No exact match found!

🔹 Searching for: benefitplancomponent.global exclusion allows otc female contraceptives.YYYY-MM-DD.a.hrl
🔹 Expected Pattern: ^benefitplancomponent\.global exclusion allows otc female contraceptives\.\d{4}-\d{2}-\d{2}\.a\.hrl$

❌ No exact match found!

🔹 Searching for: benefitplancomponent.global exclusion allows otc male and female contraceptives - mo.YYYY-MM-DD.a.hrl
🔹 Expected Pattern: ^benefitplancomponent\.global exclusion allows otc male and female contraceptives - mo\.\d{4}-\d{2}-\d{2}\.a\.hrl$

❌ No exact match found!

🔹 Searching for: benefitplancomponent.global exclusion - medicare.YYYY-MM-DD.a.hrl
🔹 Expected Pattern: ^benefitplancomponent\.global exclusion - medicare\.\d{4}-\d{2}-\d{2}\.a\.hrl$

❌ No exact match found!

🔹 Searching for: benefitplancomponent.physical and occupational therapy excludes home - medicare outpatient definition alternate office and clinic definition - inn copay deductible waived.YYYY-MM-DD.a.hrl
🔹 Expected Pattern: ^benefitplancomponent\.physical and occupational therapy excludes home - medicare outpatient definition alternate office and clinic definition - inn copay deductible waived\.\d{4}-\d{2}-\d{2}\.a\.hrl$

❌ No exact match found!

🔹 Searching for: benefitplancomponent.physical and occupational therapy excludes home - medicare outpatient definition - inn copay.YYYY-MM-DD.a.hrl
🔹 Expected Pattern: ^benefitplancomponent\.physical and occupational therapy excludes home - medicare outpatient definition - inn copay\.\d{4}-\d{2}-\d{2}\.a\.hrl$

❌ No exact match found!

🔹 Searching for: benefitplancomponent.physical and occupational therapy excludes home - medicare outpatient definition - inn copay deductible waived.YYYY-MM-DD.a.hrl
🔹 Expected Pattern: ^benefitplancomponent\.physical and occupational therapy excludes home - medicare outpatient definition - inn copay deductible waived\.\d{4}-\d{2}-\d{2}\.a\.hrl$

❌ No exact match found!

🔹 Searching for: benefitplantemplate.gold s - hmo fl 2024.YYYY-MM-DD.a.hrl
🔹 Expected Pattern: ^benefitplantemplate\.gold s - hmo fl 2024\.\d{4}-\d{2}-\d{2}\.a\.hrl$

❌ No exact match found!
