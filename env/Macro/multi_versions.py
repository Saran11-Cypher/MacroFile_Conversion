import os
import shutil
import pandas as pd

# Constants for folder paths
UPLOAD_FOLDER = "/path/to/your/upload/folder"
HRL_PARENT_FOLDER = "/path/to/your/output/folder"

# Define the user prompt for handling multiple versions
USER_PROMPT = """
Multiple versions found for some files. What would you like to do?
1. Choose latest version
2. Choose oldest version
Enter your choice (1 or 2): """

# Function to normalize text (for comparison purposes)
def normalize_text(text):
    return text.strip().lower()

# Function to get the file version based on the user's choice
def prompt_user_for_multiversion_choice():
    while True:
        print(USER_PROMPT)
        choice = input().strip()
        if choice in ['1', '2']:
            return choice
        else:
            print("Invalid input. Please enter 1 or 2.")

# Function to find matching files
def find_matching_file(config_type, config_name, folder_path, user_choice=None):
    if "&" in config_name:
        config_name = config_name.replace("&", "and")
    
    # Normalize the config type and name
    normalized_config_type = normalize_text(config_type)
    normalized_config_name = normalize_text(config_name)
    
    # Get all filenames without date suffixes
    normalized_files = {
        normalize_text(trim_suffix(f)): f
        for f in os.listdir(folder_path) if os.path.isfile(os.path.join(folder_path, f))
    }

    # Construct the expected pattern (without date & extension)
    expected_pattern = f"{normalized_config_type}.{normalized_config_name}"

    # Find all files that match the pattern (ignoring date suffix)
    matching_files = [
        f for norm_file, f in normalized_files.items() if expected_pattern in norm_file
    ]

    # If there are no matches, return None
    if not matching_files:
        print(f"❌ No match found for {expected_pattern}")
        return None

    # If there is only one match, return it directly
    if len(matching_files) == 1:
        print(f"✅ Match Found: {matching_files[0]}")
        return matching_files[0]

    # If there are multiple versions, use the user's choice
    if user_choice == '1':  # Choose the latest version
        return sorted(matching_files)[-1]  # Return the latest by date
    elif user_choice == '2':  # Choose the oldest version
        return sorted(matching_files)[0]  # Return the oldest by date

    return None

# Function to trim the suffix from the filename (date part)
def trim_suffix(filename):
    return '.'.join(filename.split('.')[:-2])  # Remove the date and extension

# Main function to process the files
def process_files(df_bal, selected_folders):
    user_choice = prompt_user_for_multiversion_choice()  # Ask user about version choice

    for index, row in df_bal.iterrows():
        config_type = row["Config Type"]
        config_name = row["Config Name"]

        if pd.isna(config_name) or not str(config_name).strip():
            continue

        if config_type in selected_folders:
            matching_file = find_matching_file(config_type, config_name, selected_folders[config_type], user_choice)

            if matching_file:
                # Mark the HRL as found
                df_bal.at[index, "HRL Available?"] = "HRL Found"
                source_path = os.path.join(selected_folders[config_type], matching_file)
                
                # Prepare target folder
                target_folder = os.path.join(HRL_PARENT_FOLDER, config_type)
                os.makedirs(target_folder, exist_ok=True)
                target_path = os.path.join(target_folder, matching_file)
                
                # Copy the file to the new folder
                shutil.copy2(source_path, target_path)
                df_bal.at[index, "File Name is correct in export sheet"] = source_path
            else:
                df_bal.at[index, "HRL Available?"] = "Not Found"

    # Save updated data back to Excel
    save_to_excel(df_bal)

# Function to save the updated DataFrame to Excel
def save_to_excel(df):
    output_file = "updated_data.xlsx"
    df.to_excel(output_file, index=False)
    print(f"Data saved to {output_file}")

# Example of how this logic is used
if __name__ == "__main__":
    # Load your Excel data into a pandas DataFrame (replace with your actual data)
    df_bal = pd.read_excel("your_excel_file.xlsx")

    # Define the folder structure where the files are located
    selected_folders = {
        "ServiceCategory": "/path/to/ServiceCategory/folder",
        # Add other config types and folder paths here
    }

    # Call the main processing function
    process_files(df_bal, selected_folders)
