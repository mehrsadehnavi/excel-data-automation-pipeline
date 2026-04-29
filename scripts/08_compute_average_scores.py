import os
import pandas as pd

folder_path = "Project\sample_data"  
output_file = os.path.join("Project\sample_data", "average.txt")

def calculate_averages(folder_path, output_file):
    results = []

    for filename in os.listdir(folder_path):
        if filename.endswith(".xlsx") or filename.endswith(".xls"):
            file_path = os.path.join(folder_path, filename)

            try:
                # Read the specific "results" sheet
                df = pd.read_excel(file_path, sheet_name="results", header=None)

                # Extract numeric values from columns E (index 4) & F (index 5) for rows 5-14
                values = df.iloc[4:14, [4, 5]].values.flatten()

                # Convert to numeric, ignoring errors (non-numeric values turn to NaN)
                numeric_values = pd.to_numeric(values, errors='coerce')

                # Remove NaN values (empty or non-numeric cells)
                numeric_values = numeric_values[~pd.isna(numeric_values)]

                # Compute average if there are valid numbers
                average = numeric_values.mean() if len(numeric_values) > 0 else 0

                # Store result in "filename    average" format
                results.append(f"{filename}\t{average:.1f}")

            except Exception as e:
                print(f"Error processing {filename}: {e}")

    # Write all results to a single text file
    with open(output_file, "w") as f:
        f.write("\n".join(results))

    print(f"Averages saved to: {output_file}")

calculate_averages(folder_path, output_file)