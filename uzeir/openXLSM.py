import win32com.client

def run_vba_macro(xlsm_filename, macro_name):
    # Create an instance of the Excel application
    excel_app = win32com.client.Dispatch("Excel.Application")
    
    # Check if the workbook is already open
    workbook_open = False
    for workbook in excel_app.Workbooks:
        if workbook.FullName == xlsm_filename:
            workbook_open = True
            wb = workbook
            break
    
    # If the workbook is not open, open it
    if not workbook_open:
        wb = excel_app.Workbooks.Open(xlsm_filename)
    
    # Run the VBA macro
    excel_app.Application.Run(f"{wb.Name}!{macro_name}")
    
    # If the workbook was opened by this function, close it without saving
    # if not workbook_open:
    #     wb.Close(SaveChanges=False)
    
    # Quit the Excel application if it was opened by this function
    # excel_app.Quit()

# Example usage
xlsm_filename = r'C:\Users\abarb\Documents\travail\stage et4\travail\codePlateau\uzeir\sample2.xlsm'
macro_name = 'MyFunction'
run_vba_macro(xlsm_filename, macro_name)
