import os
import pandas as pd
import re
from openpyxl import load_workbook

EXCEL_FILE = "C:\\Users\\n925072\\Downloads\\MacroFile_Conversion-master\\MacroFile_Conversion-master\\New folder\\convertor\\Macro_Functional_Excel.xlsx"  # Update path
UPLOAD_FOLDER = "C:\\1"  # Folder with uploaded files

# Load workbook and sheets
wb = load_workbook(EXCEL_FILE)
ws_main = wb["Main"]
ws_bal = wb["Business Approved List"]

# Define Config Load Order
config_load_order = [
    "ValueList", "AttributeType", "UserDefinedTerm", "LineOfBusiness",
    "Product", "ServiceCategory", "BenefitNetwork", "NetworkDefinitionComponent",
    "BenefitPlanComponent", "WrapAroundBenefitPlan", "BenefitPlanRider",
    "BenefitPlanTemplate", "Account", "BenefitPlan", "AccountPlanSelection"
]

# Function to normalize text dynamically
def normalize_text(text):
    """Normalize text dynamically while preserving correct patterns."""
    text = str(text).strip().lower()  # Ensure string format and lowercase

    # Keep alphanumeric, spaces, and hyphens, removing all other characters
    text = re.sub(r'[^a-zA-Z0-9\s-]', '', text)  

    # Handle cases like `-_INN` → `-INN` while keeping other structures intact
    text = re.sub(r'_-([a-zA-Z0-9]+)', r'-\1', text)  

    # Normalize multiple spaces and dashes
    text = re.sub(r'\s+', '_', text)  # Convert spaces to underscores
    text = re.sub(r'[-]+', '-', text)  # Normalize multiple dashes
    text = re.sub(r'[_]+', '_', text)  # Normalize multiple underscores

    return text

# Load "Business Approved List" into DataFrame
df_bal = pd.read_excel(EXCEL_FILE, sheet_name="Business Approved List", dtype=str)
df_bal["Config Type"] = df_bal["Config Type"].astype(str).apply(normalize_text)

# Get the config types from the Business Approved List
approved_config_types = set(df_bal["Config Type"].dropna().unique())

# Normalize folder names dynamically
available_folders = {normalize_text(f): os.path.join(UPLOAD_FOLDER, f)
                     for f in os.listdir(UPLOAD_FOLDER) if os.path.isdir(os.path.join(UPLOAD_FOLDER, f))}

# Filter only matching folders
selected_folders = {config: path for config, path in available_folders.items() if config in approved_config_types}

if not selected_folders:
    print("❌ No matching config folders found.")
    exit()

# Process selected folders and count files
for config_type, folder_path in selected_folders.items():
    uploaded_files = [f for f in os.listdir(folder_path) if os.path.isfile(os.path.join(folder_path, f))]
    file_count = len(uploaded_files)
    ws_main.append([config_type, file_count, "Pending", "Pending", "Pending"])

# Assign order dynamically based on available configurations
df_bal["Order"] = df_bal["Config Type"].apply(lambda x: config_load_order.index(x) if x in config_load_order else -1)

# Validate order sequence
valid_orders = df_bal[df_bal["Order"] >= 0]["Order"]
if not valid_orders.is_monotonic_increasing:
    print("❌ Invalid Order! Please arrange the data correctly.")
    exit()

# Drop temporary "Order" column
df_bal.drop(columns=["Order"], inplace=True)

# Function to find matching files dynamically
def find_matching_file(config_name, folder_path):
    """Finds files that strictly match the config name (ignoring case, special characters, and spacing)."""
    # Pre-process the config name once
    normalized_config_name = re.sub(r'[^a-zA-Z0-9]', '', config_name).lower()
    
    # Get all files at once and store in memory
    all_files = [f for f in os.listdir(folder_path) if os.path.isfile(os.path.join(folder_path, f))]
    
    # First try exact matches
    for filename in all_files:
        # Try with modified filename first (replacing '_-' with '-')
        temp_filename = filename.replace('_-', '-')
        cleaned_filename = re.sub(r'[^a-zA-Z0-9]', '', temp_filename).lower()
        
        if normalized_config_name == cleaned_filename:
            return filename
    
    # If no exact match, try partial matches
    for filename in all_files:
        temp_filename = filename.replace('_-', '-')
        cleaned_filename = re.sub(r'[^a-zA-Z0-9]', '', temp_filename).lower()
        
        if normalized_config_name in cleaned_filename:
            return filename

    return None

# ... rest of the code ...

# Add progress tracking for large file processing
total_rows = len(df_bal)
print(f"Processing {total_rows} configurations...")
    if pd.isna(config_name) or not str(config_name).strip():
        continue  # Skip empty config names

    if config_type in selected_folders:
        matching_file = find_matching_file(config_name, selected_folders[config_type])

        if matching_file:
            df_bal.at[index, "HRL Available?"] = "HRL Found"
            df_bal.at[index, "File Name is correct in export sheet"] = os.path.join(selected_folders[config_type], matching_file)
        else:
            df_bal.at[index, "HRL Available?"] = "Not Found"

# Save updates to Excel
for row_idx, row in df_bal.iterrows():
    for col_idx, value in enumerate(row):
        ws_bal.cell(row=row_idx+2, column=col_idx+1, value=str(value))

wb.save(EXCEL_FILE)
print("✅ Excel file updated successfully!")





BenefitPlanComponent.MedicalServices-Medicare_AlternateOffice_and_ClinicDefinition_-INN_CoinsuranceDeductibleWaived_Office_CopayDeductibleWaived.1800-01-01.a.hrl
BenefitPlanComponent.MedicalServices-Medicare_AlternateOffice_and_ClinicDefinition_-INN_Coinsurance_Office_CopayDeductibleWaived.1800-01-01.a.hrl
BenefitPlanComponent.MedicalServices-Medicare_AlternateOffice_and_ClinicDefinition_-OON_Coinsurance_Office_CopayDeductibleWaived.1800-01-01.a.hrl
BenefitPlanComponent.MedicalServices-Medicare-INN_BenefitSpecificCoinsuranceDeductibleWaived_PCP_CoinsuranceDeductibleWaived.1800-01-01.a.hrl

