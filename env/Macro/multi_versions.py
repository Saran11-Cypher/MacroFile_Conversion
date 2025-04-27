import os
from collections import defaultdict

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

# Example usage:
folder_path = 'your_folder_path_here'
single_version_files, multi_version_files, all_files = categorize_files(folder_path)

print(f"Total Files: {all_files}")
print(f"Single Versions: {len(single_version_files)}")
print(f"Multi Versions: {len(multi_version_files)}")

print("\nSingle Version Files:")
for config_name, files in single_version_files.items():
    print(f"{config_name}: {files}")

print("\nMulti Version Files:")
for config_name, files in multi_version_files.items():
    print(f"{config_name}: {files}")
