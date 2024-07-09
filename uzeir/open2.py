import os
import win32com.client as win32

def ensure_sheets_exist(workbook, *sheet_names):
    for sheet_name in sheet_names:
        sheet_exists = False
        for sheet in workbook.Sheets:
            if sheet.Name == sheet_name:
                sheet_exists = True
                break
        if not sheet_exists:
            workbook.Sheets.Add().Name = sheet_name

    
def update_vba_module_from_bas(workbook, module_name, bas_file_path):
    with open(bas_file_path, 'r') as file:
        bas_code = file.read()

    vb_component = None
    for vb_comp in workbook.VBProject.VBComponents:
        if vb_comp.Name == module_name:
            vb_component = vb_comp
            break
    
    if vb_component is None:
        vb_component = workbook.VBProject.VBComponents.Add(1)
        vb_component.Name = module_name
    
    vb_component.CodeModule.DeleteLines(1, vb_component.CodeModule.CountOfLines)
    vb_component.CodeModule.AddFromString(bas_code)

def execute_vba_function(filename, vba_function_name, bas_file_path, module_name, open_all=False):
    # Initialize Excel application
    excel = win32.Dispatch('Excel.Application')

    # Check if the file exists, if not, create it
    if not os.path.exists(filename):
        workbook = excel.Workbooks.Add()
        workbook.SaveAs(filename, FileFormat=52)
        file_existed = False
    else:
        file_existed = True

    # Check if the file is already open
    workbook_open = False
    if file_existed:
        for wb in excel.Workbooks:
            if wb.FullName == os.path.abspath(filename):
                workbook_open = True
                workbook = wb
                break

    # If the workbook is not already open, open it
    if not workbook_open:
        workbook = excel.Workbooks.Open(os.path.abspath(filename))

    ensure_sheets_exist(workbook, "hey", "yo")
        
    # Update or add the VBA module from the .bas file
    update_vba_module_from_bas(workbook, module_name, bas_file_path)


    # Run the VBA function
    data = [[1,2], [3]]
    data = excel.Application.Run(f'{workbook.Name}!{vba_function_name}', data)
    excel.Application.Run(f'{workbook.Name}!{vba_function_name}', data)

    # Save the workbook
    workbook.Save()

    # Close the workbook if it wasn't open before, unless open_all is True
    if not workbook_open and not open_all:
        workbook.Close()

    # Make Excel visible if open_all is True
    if open_all:
        excel.Visible = True

# Example usage
execute_vba_function(r"C:\Users\abarb\Documents\travail\stage et4\travail\codePlateau\uzeir\sample2.xlsm", 'MyFunctions', r"C:\Users\abarb\Documents\travail\stage et4\travail\codePlateau\uzeir\recap.bas", "addData", open_all=True)
