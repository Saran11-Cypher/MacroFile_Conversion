import os
import math
from datetime import datetime

def merge_files_in_batches(input_folder, base_output_folder, num_files, batch_size, run_id):
    output_folder = os.path.join(base_output_folder, run_id)
    os.makedirs(output_folder, exist_ok=True)  # Ensure output folder exists
    file_list = sorted([f for f in os.listdir(input_folder) if f.endswith(".html")])[:num_files]  # Get the required files
    
    total_batches = math.ceil(num_files / batch_size)  # Calculate number of batches
    batch_files = []
    
    # if not run_id:
    #     run_id = datetime.now().strftime("%Y%m%d_%H%M%S")
        
    # Step 1: Merge in batches and store them
    for i in range(total_batches):
        batch_start = i * batch_size
        batch_end = min((i + 1) * batch_size, num_files)
        batch_file_list = file_list[batch_start:batch_end]
        
        batch_output_name = f"batch_{i+1}_{run_id}.html"
        batch_output_path = os.path.join(output_folder, batch_output_name)
        batch_files.append(batch_output_path)

        with open(batch_output_path, "w", encoding="utf-8") as batch_file:
            for file_name in batch_file_list:
                file_path = os.path.join(input_folder, file_name)
                with open(file_path, "r", encoding="utf-8") as f:
                    batch_file.write(f.read() + "\n")  # Maintain separation
            
        print(f"Batch {i+1} merged and saved  {len(batch_file_list)} --> {batch_output_path}")

    # Step 2: Merge all batch files into the final merged file
    final_output_name = f"Final_{run_id}.html"
    final_output_path = os.path.join(output_folder, final_output_name)
    
    with open(final_output_path, "w", encoding="utf-8") as final_file:
        for batch_file in batch_files:
            with open(batch_file, "r", encoding="utf-8") as bf:
                final_file.write(bf.read() + "\n")
    
    print(f"All batches merged into {final_output_path}")

# User input
num_files = int(input("Enter the number of HTML files to merge: "))
batch_size = int(input("Enter the number of files per batch: "))
custom_run_id = input("Enter a run ID or leave blank for time-stamp-based Naming: ").strip()

if not custom_run_id:
    print(" Run ID cannot be empty.")
else:
    input_folder = "C:\\Users\\n925072\\Downloads\\BP_MisMatched\\BP_MisMatched"  # Change this to actual folder path
    output_folder = "C:\\Data\\Merged_Files"  # Store batches & final file inside this folder



merge_files_in_batches(input_folder, output_folder, num_files, batch_size, custom_run_id)
