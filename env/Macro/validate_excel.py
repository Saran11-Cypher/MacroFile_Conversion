import os
import shutil
import pandas as pd
import re

# Define file paths
EXCEL_FILE = "C:\\Users\\n925072\\Downloads\\Macro_Functional_Excel.xlsx"  # Update with your actual file path
UPLOAD_FOLDER = "C:\\1"  # Change to the folder containing uploaded files
FILTERED_FOLDER = os.path.join(UPLOAD_FOLDER, "Filtered_Files")  # Folder to store matched files

# Ensure the folder exists
if not os.path.exists(UPLOAD_FOLDER):
    print(f"❌ Error: Folder '{UPLOAD_FOLDER}' does not exist.")
    exit()

# Ensure the filtered files folder exists
if not os.path.exists(FILTERED_FOLDER):
    os.makedirs(FILTERED_FOLDER)

# Load the Business Approved List sheet
df_bal = pd.read_excel(EXCEL_FILE, sheet_name="Business Approved List", dtype=str)

# Normalize text for comparison
def normalize_text(text):
    """Removes special characters, converts to lowercase, and standardizes spaces/hyphens."""
    return re.sub(r'[^a-zA-Z0-9\s-]', '', str(text)).strip().lower()

# Process only rows where HRL is found
for index, row in df_bal.iterrows():
    if row["HRL Available?"] == "HRL Found":
        config_type = row["Config Type"]
        source_file_path = row["File Name is correct in export sheet"]

        # Skip if there's no valid file path
        if not source_file_path or not os.path.exists(source_file_path):
            continue

        # Normalize config type to create a folder
        config_folder_name = normalize_text(config_type)
        config_folder_path = os.path.join(FILTERED_FOLDER, config_folder_name)

        # Create a subfolder for the config type
        if not os.path.exists(config_folder_path):
            os.makedirs(config_folder_path)

        # Copy the file into the respective config type folder
        destination_path = os.path.join(config_folder_path, os.path.basename(source_file_path))
        shutil.copy2(source_file_path, destination_path)

print("✅ All filtered files have been stored in the 'Filtered_Files' folder.")
