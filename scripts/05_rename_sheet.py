import os
import win32com.client

# Paths
folder_path = 'Project\\sample_data'
# Password for the protected files
password = '3070'

# Initialize Excel application
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False  # Run Excel in the background

# Loop through all files in the folder
for filename in os.listdir(folder_path):
    if filename.endswith('.xlsx'):  # Process only .xlsx files
        file_path = os.path.join(folder_path, filename)

        try:
            # Open the Excel file with the password
            wb = excel.Workbooks.Open(file_path, Password=password)

            # Access the sheet named 'results'
            if 'results' in [sheet.Name for sheet in wb.Sheets]:
                sheet = wb.Sheets['results']
                sheet.Name = 'result'  # Rename the sheet

                # Save the workbook with the new sheet name and without password
                # Re-save the file with an empty password (removes protection)
                wb.SaveAs(file_path, Password="")
                print(f"Sheet renamed and password removed in {filename}")

            else:
                print(f"Sheet 'results' not found in {filename}")

            # Close the workbook
            wb.Close(False)  # False to not save changes if no renaming occurred

        except Exception as e:
            print(f"Error processing {filename}: {e}")
        finally:
            # Continue to next file
            pass

excel.Quit()
