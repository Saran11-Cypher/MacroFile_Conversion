import os
import re

def normalize_text(text):
    """
    Normalizes text by:
    - Converting '&' to 'and'
    - Removing special characters except hyphen
    - Removing extra spaces
    - Lowercasing the text
    """
    text = text.lower().strip()
    text = text.replace("&", "and")  # Convert '&' to 'and'
    text = re.sub(r'[^a-zA-Z0-9\s-]', '', text)  # Remove special characters except hyphen
    text = re.sub(r'\s+', '', text)  # Remove all spaces for strict matching
    return text

def extract_config_from_filename(filename):
    """
    Extracts and normalizes the configuration name portion from the filename.
    - Removes 'BenefitPlanComponent.' prefix
    - Removes timestamp and extension (e.g., .1800-01-01.a.hrl)
    """
    if "." in filename:
        filename = filename.split(".", 1)[1]  # Remove prefix
    filename = filename.rsplit(".", 2)[0]  # Remove timestamp & extension
    return normalize_text(filename)

def find_matching_file(config_name, folder_path):
    """
    Matches a given config name to an HRL file in the given folder.
    - Returns the exact matching file name.
    - If no match is found, returns "No HRL found".
    """
    normalized_config_name = normalize_text(config_name)

    for filename in os.listdir(folder_path):
        if os.path.isfile(os.path.join(folder_path, filename)):
            cleaned_filename = extract_config_from_filename(filename)

            # Strict match check
            if cleaned_filename == normalized_config_name:
                return filename  

    return "No HRL found"  # If no exact match, return this message

# Example Usage
config_name = "Medical Services - Medicare Alternate Office and Clinic Definition - INN Coinsurance Deductible Waived Office Copay Deductible Waived"
folder_path = "your_folder_path_here"  # Replace with actual path

matched_file = find_matching_file(config_name, folder_path)
print("Matched File:", matched_file)



print("Config Name:", normalize_text("Medical Services - Medicare (Alternate Office & Clinic Definition) - INN (Coinsurance Deductible Waived) Office (Copay Deductible Waived)"))
print("File Name:", normalize_text("BenefitPlanComponent.MedicalServices-Medicare_AlternateOffice_and_ClinicDefinition_-INN_CoinsuranceDeductibleWaived_Office_CopayDeductibleWaived.1800-01-01.a.hrl"))
Config Name: medical services - medicare alternate office and clinic definition - inn coinsurance deductible waived office copay deductible waived
File Name: benefitplancomponentmedicalservices-medicarealternateofficeandclinicdefinition-inncoinsurancedeductiblewaivedofficecopaydeductiblewaived1800-01-01ahrl

Searching for config: Medical Services - Medicare (Alternate Office & Clinic Definition) - INN (Coinsurance Deductible Waived) Office (Copay Deductible Waived)ng 
Normalized Name: medicalservicesmedicarealternateofficeclinicdefinitioninncoinsurancedeductiblewaivedofficecopaydeductiblewaived

Clr=eaned filename : benefitplancomponentabortionelectivestandardmandaterteonlyooncoinsurance18000101ahrl
Clr=eaned filename : benefitplancomponentmedicalservicesmedicarealternateofficeandclinicdefinitioninncoinsurancedeductiblewaivedofficecopaydeductiblewaived18000101ahrl
Clr=eaned filename : benefitplancomponentmedicalservicesmedicarealternateofficeandclinicdefinitioninncoinsuranceofficecopaydeductiblewaived18000101ahrl
Clr=eaned filename : benefitplancomponentmedicalservicesmedicarealternateofficeandclinicdefinitionooncoinsuranceofficecopaydeductiblewaived18000101ahrl
