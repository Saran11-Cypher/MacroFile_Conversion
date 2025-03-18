import re

def normalize_text(text):
    """Normalize text by converting to lowercase and replacing separators with spaces."""
    return re.sub(r'[^a-zA-Z0-9\s]', ' ', text).strip().lower()

def preprocess_filenames(file_list):
    """Precompute a dictionary with filenames and their normalized words for quick lookup."""
    file_dict = {}
    for file_name in file_list:
        normalized_file_name = normalize_text(file_name)
        file_dict[file_name] = normalized_file_name.split()  # Store words as a list (preserving order)
    return file_dict

def find_best_match(config_name, file_dict):
    """Find the best match for the given config name with no-prefix priority logic."""
    normalized_config_name = normalize_text(config_name)
    config_words = normalized_config_name.split()  # Convert to a list of words

    exact_matches = []
    prefixed_matches = []

    for file_name, file_words in file_dict.items():
        try:
            index = file_words.index(config_words[0])  # First word of config in filename
            matched_sequence = file_words[index:index + len(config_words)]

            if matched_sequence == config_words:
                if index == 0:  # No prefix before config name
                    exact_matches.append(file_name)
                else:  # Config name exists but has prefixes
                    prefixed_matches.append(file_name)
        except ValueError:
            continue  # Skip if config name not found

    # Prioritize exact matches (no prefix before config name)
    if exact_matches:
        return exact_matches[0]
    elif prefixed_matches:
        return prefixed_matches[0]  # Select prefixed match if no exact match found

    return None  # No match found

# Example Scenarios

# Scenario 1
config_name_1 = "Physical and Occupational Therapy Excludes Home - Medicare Outpatient Definition - INN Copay."
file_names_1 = [
    "ServiceCategory.Physical and Occupational Therapy - Medicare Outpatient Definition - INN (Copay).2025-01-01.a.hrl",
    "ServiceCategory.Physical and Occupational Therapy - Medicare Outpatient Definition - INN (Copay Deductible Waived).2025-01-01.a.hrl"
]

# Scenario 2
config_name_2 = "Surgery"
file_names_2 = [
    "ServiceCategory.Surgery.1800-01-01.a.hrl",
    "ServiceCategory.Heart.1800-01-01.a.hrl",
    "ServiceCategory.Medical_Vision-item.2023-05-12.a.hrl",
    "ServiceCategory.Vision-item.CareFree-Surgery.1800-01-01.a.hrl",
    "ServiceCategory.Surgeryitems.1800-01-01.a.hrl"
]

# Preprocess filenames
file_dict_1 = preprocess_filenames(file_names_1)
file_dict_2 = preprocess_filenames(file_names_2)

# Find the best match for each scenario
matched_file_1 = find_best_match(config_name_1, file_dict_1)
matched_file_2 = find_best_match(config_name_2, file_dict_2)

print("Scenario 1 Matched File:", matched_file_1 if matched_file_1 else "No exact match found")
print("Scenario 2 Matched File:", matched_file_2 if matched_file_2 else "No exact match found")
