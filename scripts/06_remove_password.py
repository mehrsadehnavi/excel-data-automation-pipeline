import os
import win32com.client

# Folder containing the Excel files
folder_path = "Project\\sample_data\\"

# Common password for the files
password = "3070"

# Application
excel = win32com.client.Dispatch("Excel.Application")
excel.DisplayAlerts = False

# Loop through all files in the folder
for filename in os.listdir(folder_path):
    if filename.endswith(".xls") or filename.endswith(".xlsx"):
        full_path = os.path.join(folder_path, filename)

        try:
            # Open the workbook with the password
            workbook = excel.Workbooks.Open(full_path, Password=password)

            # Save the workbook without a password
            workbook.SaveAs(full_path, Password="", WritePassword="", ReadOnlyRecommended=False, CreateBackup=False)
            workbook.Close()

            print(f"Password removed for: {filename}")

        except Exception as e:
            print(f"Error processing {filename}: {e}")

excel.Quit()
