import win32com.client

def open_and_run_vba(xlsm_file_path, vba_function_name, open_all=True):
    # Create an instance of Excel
    excel = win32com.client.Dispatch("Excel.Application")
    
    # Try to find the workbook if it's already open
    workbook = None
    for wb in excel.Workbooks:
        if wb.FullName == xlsm_file_path:
            workbook = wb
            break
    
    # If the workbook is not already open, open it
    if not workbook:
        workbook = excel.Workbooks.Open(xlsm_file_path)
    
    # Run the VBA function
    excel.Application.Run(vba_function_name)
    
    # Keep the file open if open_all is True
    if open_all:
        excel.Visible = True
        excel.WindowState = win32com.client.constants.xlNormal
    else:
        workbook.Close(SaveChanges=False)
    
    # Optionally, quit the Excel application if open_all is False and there are no other workbooks open
    if not open_all and len(excel.Workbooks) == 0:
        excel.Quit()

# Example usage
open_and_run_vba(r"C:\Users\abarb\Documents\travail\stage et4\travail\codePlateau\uzeir\sample2.xlsm", 'MyFunction', open_all=True)