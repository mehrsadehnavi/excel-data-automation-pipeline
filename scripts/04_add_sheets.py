# In this program, we need to add 3 new sheets from a template file to 500+ already-existing excel files.
# We also get a backup from the files before modifying them, in case the code doesn't work and messes the whole thing up. :))

import os
import shutil
import traceback
import win32com.client as win32

# Paths
template_path = "Project\\sample_data\\temp.xlsx"  
teacher_folder = "Project\\sample_data\\" 
backup_folder = os.path.join(teacher_folder, "backups") 
sheet_names = ["Bonus", "Personal Info", "Professional Info"] 

# Safety checks
if not os.path.exists(template_path):
    raise FileNotFoundError(f"Template not found: {template_path}")
if not os.path.isdir(teacher_folder):
    raise NotADirectoryError(f"Teacher folder not found: {teacher_folder}")
os.makedirs(backup_folder, exist_ok=True)

excel = None
template_wb = None
try:
    excel = win32.gencache.EnsureDispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    # Open template (read-only is fine & no need to choose the "edit" button)
    template_wb = excel.Workbooks.Open(template_path, ReadOnly=True)

    for fname in os.listdir(teacher_folder):
        # skip non-Excel files and temp lock files
        if not fname.lower().endswith((".xlsx", ".xlsm", ".xls")):
            continue
        if fname.startswith("~$"):
            continue

        full_path = os.path.join(teacher_folder, fname)
        # avoid modifying the template if someone placed it in the same folder
        if os.path.abspath(full_path).lower() == os.path.abspath(template_path).lower():
            print(f"Skipping template file itself: {fname}")
            continue

        try:
            print(f"\n--- Processing: {fname} ---")
            # backup first
            backup_path = os.path.join(backup_folder, fname)
            shutil.copy2(full_path, backup_path)
            print(f"Backup created: {backup_path}")

            # open teacher workbook
            wb = excel.Workbooks.Open(full_path)

            # 1) Delete old sheets if they exist (some sheets were added manually)
            for sname in sheet_names:
                try:
                    ws = wb.Worksheets(sname)
                    ws.Delete()  # DisplayAlerts is False so no prompt
                    print(f"Deleted existing sheet: {sname}")
                except Exception:
                    # sheet not present: ignore
                    pass

            # Save after deletion to avoid name collisions
            wb.Save()

            # 2) Copy each sheet from template into the teacher workbook (append at end)
            for sname in sheet_names:
                try:
                    t_ws = template_wb.Worksheets(sname)
                except Exception:
                    raise RuntimeError(f"Sheet '{sname}' not found in template.")
                # Copy -> this inserts a new sheet as the last sheet
                t_ws.Copy(After=wb.Sheets(wb.Sheets.Count))
                # newly copied sheet will be the last sheet
                new_sheet = wb.Sheets(wb.Sheets.Count)
                # Rename (should be safe because we deleted same-name sheets earlier and saved)
                try:
                    new_sheet.Name = sname
                    print(f"Copied and renamed sheet: {sname}")
                except Exception as e:
                    print(f"Warning: couldn't rename copied sheet to '{sname}': {e}")

            # final save & close
            wb.Save()
            wb.Close(SaveChanges=True)
            print(f"Saved and closed: {fname}")

        except Exception as e:
            print(f"ERROR while processing {fname}: {e}")
            traceback.print_exc()
            try:
                # try to close without saving if something went wrong
                wb.Close(SaveChanges=False)
            except Exception:
                pass

    # close template
    template_wb.Close(SaveChanges=False)
    template_wb = None

    print("\n✅ All done. Check the backups folder in case something unexpected occurred.")

except Exception as e:
    print("Fatal error:", e)
    traceback.print_exc()

finally:
    if template_wb is not None:
        try:
            template_wb.Close(SaveChanges=False)
        except Exception:
            pass
    if excel is not None:
        try:
            excel.DisplayAlerts = True
            excel.Quit()
        except Exception:
            pass
