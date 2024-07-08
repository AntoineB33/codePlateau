import xlwings as xw
import os

# Function to add or update the sheet if it doesn't exist
def add_sheet_if_not_exists(wb, sheet_name):
    if sheet_name not in [sheet.name for sheet in wb.sheets]:
        wb.sheets.add(sheet_name)

# Function to update VBA module with code from .bas file
def update_vba_module_from_bas(wb, bas_file_path, module_name="Module1"):
    # Read the .bas file contents
    with open(bas_file_path, 'r') as file:
        vba_code = file.read()
    
    # Check if the module exists
    module_exists = False
    for component in wb.api.VBProject.VBComponents:
        if component.Name == module_name:
            module_exists = True
            module = component
            break

    # If the module doesn't exist, create it
    if not module_exists:
        module = wb.api.VBProject.VBComponents.Add(1)
        module.Name = module_name

    # Replace the existing code with the new code from the .bas file
    module.CodeModule.DeleteLines(1, module.CodeModule.CountOfLines)
    module.CodeModule.AddFromString(vba_code)

# File name and sheet name
open_all = False
file_name = 'sample2.xlsm'
full_file_path = os.path.abspath(file_name)  # Get the absolute path to the file
recap_folder = r"C:\Users\abarb\Documents\travail\stage et4\travail\codePlateau\uzeir" + "\\"
sheet_name = 'MySheet'
bas_file_path = os.path.abspath('uzeir\\recap.bas')

full_file_path = recap_folder + file_name
the_file_exists = True
if not os.path.exists(full_file_path):
    app = xw.App(visible=open_all)
    the_file_exists = False
    wb = app.books.add()

is_open = False
if the_file_exists:
    # Check if the specific workbook is already open by its full path
    app = xw.apps.active
    if app:
        for book in xw.books:
            if book.fullname == full_file_path:
                wb = book
                is_open = True
                break
    if not is_open:
        app = xw.App(visible=open_all)
        wb = app.books.open(full_file_path)


# Add the sheet if it doesn't exist
add_sheet_if_not_exists(wb, sheet_name)

# Update the VBA module with the code from the .bas file
update_vba_module_from_bas(wb, bas_file_path)

wb.macro("Module1.MyFunctions")("hey", "yi")

# Save the workbook and close it if it was not already open
if not the_file_exists:
    wb.save(full_file_path)
else:
    wb.save()
if not open_all and not is_open:
    wb.close()
    app.quit()

print(f"Workbook '{file_name}' with sheet '{sheet_name}' and updated VBA module from '{bas_file_path}' added successfully.")
