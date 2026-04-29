# In this code, we want to create a file for each teacher based on a predefined template. 
# We need to extract teachers' names from the previous folder and rename the new files.

import os
import shutil

# Set folder paths
teacher_folder = "Project\\sample_data"      
output_folder = "Project\\sample_data"     
template_file = os.path.join("Project\\sample_data", "temp.xlsx")

# Loop through each file in the teacher folder
for filename in os.listdir(teacher_folder):
    if filename.endswith(".xlsx"):
        # Extract teacher name & remove extension
        teacher_name = os.path.splitext(filename)[0]
        
        # Define path for the new file
        new_file_path = os.path.join(output_folder, f"{teacher_name}.xlsx")
        
        # Copy template and rename it
        shutil.copy(template_file, new_file_path)

        print(f"Created: {new_file_path}")
