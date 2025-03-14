import os
import pandas as pd
import re
from openpyxl import load_workbook

EXCEL_FILE = "C:\\Users\\n925072\\Downloads\\MacroFile_Conversion-master\\MacroFile_Conversion-master\\New folder\\convertor\\Macro_Functional_Excel.xlsx"
UPLOAD_FOLDER = "C:\\1"

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

def normalize_text(text):
    return re.sub(r'[^a-zA-Z0-9\s-]', '', str(text)).strip().lower()

def standardize_name(name):
    return (name.replace(" ", "-")
               .replace("(", "")
               .replace(")", "")
               .replace("&", "and")
               .replace(",", "")
               .replace(".", "-"))

def find_matching_file(config_name, folder_path):
    if not config_name or not folder_path:
        return None
        
    std_config_name = standardize_name(config_name).lower()
    best_match = None
    
    for filename in os.listdir(folder_path):
        if os.path.isfile(os.path.join(folder_path, filename)):
            base_name = os.path.splitext(filename)[0]
            std_filename = standardize_name(base_name).lower()
            
            if std_config_name in std_filename:
                if best_match is None or len(std_filename) < len(standardize_name(best_match).lower()):
                    best_match = filename
            
    return best_match

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

for index, row in df_bal.iterrows():
    config_type = row["Config Type"]
    config_name = row["Config Name"]

    if pd.isna(config_name) or not str(config_name).strip():
        continue

    if config_type in selected_folders:
        matching_file = find_matching_file(config_name, selected_folders[config_type])
        
        if matching_file:
            print(f"✅ Matched: {config_name} → {matching_file}")
            df_bal.at[index, "File Available?"] = "Found"
            df_bal.at[index, "File Name is correct in export sheet"] = os.path.join(selected_folders[config_type], matching_file)
        else:
            print(f"❌ No match for: {config_name} in {selected_folders[config_type]}")
            df_bal.at[index, "File Available?"] = "Not Found"

for row_idx, row in df_bal.iterrows():
    for col_idx, value in enumerate(row):
        ws_bal.cell(row=row_idx+2, column=col_idx+1, value=str(value))

wb.save(EXCEL_FILE)
print("✅ Excel file updated successfully!...")




✅ Matched: Medical Supplies - Compression Garments - Medicare Covered - INN (Benefit Specific Coinsurance) → BenefitPlanComponent.MedicalSupplies-CompressionGarments-MedicareCovered-INN_CoinsuranceDeductibleWaived.1800-01-01.a.hrl
nts-MedicareCovered-INN_BenefitSpecificCoinsurance.1800-01-01.a.hrl
s-CompressionGarments-MedicareCovered-INN_BenefitSpecificCoinsuranceDeductibleWaived.1800-01-01.a.hrl
✅ Matched: Medical Supplies - Compression Garments - Medicare Covered - INN (Benefit Specific Coinsurance) → BenefitPlanComponent.MedicalSupplies-CompressionGarms-CompressionGarments-MedicareCovered-INN_BenefitSpecificCoinsuranceDeductibleWaived.1800-01-01.a.hrl
✅ Matched: Medical Supplies - Compression Garments - Medicare Covered - INN (Benefit Specific Coinsurance) → BenefitPlanComponent.MedicalSupplies-CompressionGarme  Matched: Medical Supplies - Compression Garments - Medicare Covered - INN (Coinsurance Deductible Waived) → BenefitPlanComponent.MedicalSupplies-CompressionGarnts-MedicareCovered-INN_BenefitSpecificCoinsurance.1800-01-01.a.hrl
s-CompressionGarments-MedicareCovered-INN_BenefitSpecificCoinsuranceDeductibleWaived.1800-01-01.a.hrl
✅ Matched: Medical Supplies - Compression Garments - Medicare Covered - INN (Benefit Specific Coinsurance) → BenefitPlanComponent.MedicalSupplies-CompressionGarm




BenefitPlanComponent
❌ No match for: Medical Services - Medicare (Alternate Office & Clinic Definition) - INN (Coinsurance) Office (Copay Deductible Waived) in C:\1\BenefitPlanCompone
nt
❌ No match for: Medical Services - Medicare (Alternate Office & Clinic Definition) - OON (Coinsurance) Office (Copay Deductible Waived) in C:\1\BenefitPlanCompone
nt
✅ Matched: Medical Supplies - Compression Garments - Medicare Covered - INN (Benefit Specific Coinsurance Deductible Waived) → BenefitPlanComponent.MedicalSupplie
