import os
import re
from tqdm import tqdm  # pip install tqdm if not already installed

# Example placeholder for selected folders dictionary
# selected_folders = {
#     "ServiceCategory": "/path/to/ServiceCategory",
#     "ValueList": "/path/to/ValueList"
# }

def get_base_filename(filename):
    """Remove date and extension to get the common identifier for duplicate files."""
    return re.sub(r'\.\d{4}-\d{2}-\d{2}\..*$', '', filename)

def extract_date_from_filename(filename):
    """Extracts the date part from the filename if present."""
    match = re.search(r'\.(\d{4}-\d{2}-\d{2})\.', filename)
    return match.group(1) if match else None

def analyze_files(selected_folders):
    """Analyzes all files, detects duplicates, and returns a filtered map."""
    file_map = {}

    print("🔍 Scanning folders for files...")
    for folder_name, folder_path in tqdm(selected_folders.items(), desc="Analyzing folders"):
        for file in os.listdir(folder_path):
            full_path = os.path.join(folder_path, file)
            if os.path.isfile(full_path):
                base_name = get_base_filename(file)
                file_map.setdefault(base_name, []).append((file, full_path))

    # Detect duplicates based on base name
    duplicates_with_versions = {
        base: versions for base, versions in file_map.items() if len(versions) > 1
    }

    return file_map, duplicates_with_versions

def prompt_user_for_version_choice():
    """Prompt user to choose version rule once."""
    version_choice = ""
    while version_choice not in ["latest", "oldest"]:
        version_choice = input(
            "\n⚠️ Found multiple files with same name but different dates.\n"
            "Would you like to proceed with the latest or oldest version? [latest/oldest]: "
        ).strip().lower()
    print(f"\n✅ Your choice: Proceed with the **{version_choice}** version for all duplicates.\n")
    return version_choice

def select_files(file_map, duplicates_with_versions, version_choice):
    """Selects appropriate version of files based on user input."""
    selected_files = {}

    for base_name, versions in file_map.items():
        if len(versions) == 1:
            selected_files[base_name] = versions[0][1]  # Only one version
        else:
            if base_name in duplicates_with_versions:
                # Multiple versions, apply rule
                sorted_versions = sorted(
                    versions,
                    key=lambda x: extract_date_from_filename(x[0]) or "",
                    reverse=(version_choice == "latest")
                )
                selected_files[base_name] = sorted_versions[0][1]  # pick the latest/oldest
            else:
                selected_files[base_name] = versions[0][1]  # fallback, should not occur

    return selected_files

# -------------------------
# 🧠 MAIN EXECUTION
# -------------------------
def run_file_analysis(selected_folders):
    file_map, duplicates_with_versions = analyze_files(selected_folders)

    if duplicates_with_versions:
        version_choice = prompt_user_for_version_choice()
    else:
        version_choice = "latest"  # default silently

    selected_files = select_files(file_map, duplicates_with_versions, version_choice)

    print("📂 Files selected for processing:")
    for key, path in selected_files.items():
        print(f"  - {key}: {os.path.basename(path)}")

    return selected_files  # Use this in your business logic

# Example:
# selected_folders = {
#     "ServiceCategory": "E:/MacroFileUploads/ServiceCategory",
#     "ValueList": "E:/MacroFileUploads/ValueList"
# }
# selected_files = run_file_analysis(selected_folders)
