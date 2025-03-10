import os
import pandas as pd
import re
from openpyxl import load_workbook

# Define paths
EXCEL_FILE = "C:\\Users\\n925072\\Downloads\\MacroFile_Conversion-master\\MacroFile_Conversion-master\\New folder\\convertor\\Macro_Functional_Excel.xlsx"  # Update with your actual file path
UPLOAD_FOLDER = "C:\\1\\BenefitPlanComponent"  # Change to the folder containing uploaded files

# Ensure the folder exists
if not os.path.exists(UPLOAD_FOLDER):
    print(f"❌ Error: Folder '{UPLOAD_FOLDER}' does not exist.")
    exit()

# Load workbook and sheets
wb = load_workbook(EXCEL_FILE)
ws_main = wb["Main"]
ws_bal = wb["Business Approved List"]

# List all uploaded files
uploaded_files = [f for f in os.listdir(UPLOAD_FOLDER) if os.path.isfile(os.path.join(UPLOAD_FOLDER, f))]
file_count = len(uploaded_files)

# Update the counts in main sheet.
ws_main.append(["BenefitPlanComponent", file_count, "Pending", "Pending", "Pending"])

# Define the correct config load order
config_load_order = [
    "ValueList", "AttributeType", "UserDefinedTerm", "LineOfBusiness",
    "Product", "ServiceCategory", "BenefitNetwork", "NetworkDefinitionComponent",
    "BenefitPlanComponent", "WrapAroundBenefitPlan", "BenefitPlanRider",
    "BenefitPlanTemplate", "Account", "BenefitPlan", "AccountPlanSelection"
]

# Load "Business Approved List" into a DataFrame
df_bal = pd.read_excel(EXCEL_FILE, sheet_name="Business Approved List", dtype=str)  # Ensure all columns are strings
df_bal["Config Type"] = df_bal["Config Type"].astype(str)
df_bal["HRL Available?"] = df_bal["HRL Available?"].astype(str)
df_bal["File Name is correct in export sheet"] = df_bal["File Name is correct in export sheet"].astype(str)  # Explicit conversion

# Function to normalize and clean text
def normalize_text(text):
    """Removes special characters, converts to lowercase, and standardizes spaces/hyphens."""
    return re.sub(r'[^a-zA-Z0-9\s-]', '', str(text)).strip().lower()

# Normalize config load order and config type in df_bal
normalized_config_types = {normalize_text(cfg): cfg for cfg in config_load_order}
normalized_df_bal_types = df_bal["Config Type"].dropna().apply(normalize_text)

# Filter config_load_order to only include present config types in df_bal
available_config_types = [normalized_config_types[cfg] for cfg in normalized_df_bal_types if cfg in normalized_config_types]

# Assign order dynamically based on available configurations only
if available_config_types:
    df_bal["Order"] = df_bal["Config Type"].apply(lambda x: available_config_types.index(x) if normalize_text(x) in map(normalize_text, available_config_types) else -1)
else:
    df_bal["Order"] = -1

# Validate order: If not in increasing sequence, show error and exit
valid_orders = df_bal[df_bal["Order"] >= 0]["Order"]

if not valid_orders.is_monotonic_increasing:
    print("❌ Error: Invalid Order! Please arrange the data correctly.")
    exit()

# Remove the temporary "Order" column (not needed in final output)
df_bal.drop(columns=["Order"], inplace=True)

# Function to match config names with uploaded files
def find_matching_file(config_name):
    """Finds files that contain all words from the config_name in any order."""
    config_words = normalize_text(config_name).split()  # Normalize config name and split into words
    
    for filename in uploaded_files:
        cleaned_filename = normalize_text(filename)  # Normalize filename

        # Ensure all words in config_name exist in the filename
        if all(word in cleaned_filename for word in config_words):
            return filename  # Return the first matched file

    return None  # No match found

# Check for HRL availability and update DataFrame
for index, row in df_bal.iterrows():
    config_name = row["Config Name"]
    
    if pd.isna(config_name) or not str(config_name).strip():
        continue  # Skip empty config names

    matching_file = find_matching_file(config_name)

    if matching_file:
        df_bal.at[index, "HRL Available?"] = "HRL Found"
        df_bal.at[index, "File Name is correct in export sheet"] = str(os.path.join(UPLOAD_FOLDER, matching_file))  # Ensure string type
    else:
        df_bal.at[index, "HRL Available?"] = "Not Found"

# Save updates to Excel
for row_idx, row in df_bal.iterrows():
    for col_idx, value in enumerate(row):
        ws_bal.cell(row=row_idx+2, column=col_idx+1, value=str(value))  # Ensure everything is saved as string

wb.save(EXCEL_FILE)
print("✅ Excel file updated successfully!")
