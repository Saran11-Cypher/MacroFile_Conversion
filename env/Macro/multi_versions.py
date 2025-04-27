import os

def categorize_files(folder_path):
    single_version_files = {}
    multi_version_files = {}

    all_files = 0

    # Iterate through files in the given folder and its subdirectories
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            all_files += 1
            
            # Extract the config name (excluding 'configtype.')
            base_name = file.split(".")[1]  # Getting only the part after 'configtype.'
            
            # If the base name already exists in single_version_files or multi_version_files, categorize accordingly
            if base_name in single_version_files:
                multi_version_files[base_name].append(file)
                del single_version_files[base_name]  # Move to multi version list
            elif base_name in multi_version_files:
                multi_version_files[base_name].append(file)
            else:
                single_version_files[base_name] = [file]

    return single_version_files, multi_version_files

# Example usage:
folder_path = 'path_to_your_folder'
single_version_files, multi_version_files = categorize_files(folder_path)

# Output results
print("Single Version Files:")
for base_name, files in single_version_files.items():
    print(f"{base_name}: {files}")

print("\nMulti Version Files:")
for base_name, files in multi_version_files.items():
    print(f"{base_name}: {files}")
