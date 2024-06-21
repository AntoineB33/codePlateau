import os
import openpyxl
import xlwings as xw

def save_vba_and_execute(xlsx_file_path, xlsm_file_path):
    """
    Transform an .xlsx file into an .xlsm file if the .xlsx file exists.
    Do nothing if the .xlsm file exists.
    Create a new .xlsm file if none of them exist.
    Save VBA code in the .xlsm file and execute it.

    Parameters:
    xlsx_file_path (str): The path to the source .xlsx file.
    xlsm_file_path (str): The path to the destination .xlsm file.
    """
    if os.path.exists(xlsm_file_path):
        print(f"{xlsm_file_path} already exists. No action taken.")
    elif os.path.exists(xlsx_file_path):
        # Load the .xlsx file
        workbook = openpyxl.load_workbook(xlsx_file_path)
        # Save the workbook as an .xlsm file
        workbook.save(xlsm_file_path)
        print(f"Transformed {xlsx_file_path} to {xlsm_file_path}.")
    else:
        # Create a new .xlsm file
        workbook = openpyxl.Workbook()
        workbook.save(xlsm_file_path)
        print(f"Created new file {xlsm_file_path}.")

    # Add VBA code to the .xlsm file and execute it
    with xw.App(visible=True) as app:
        wb = app.books.open(xlsm_file_path)
        vba_module = wb.api.VBProject.VBComponents.Add(1)  # 1 = Module
        vba_code = """
        Sub WriteData()
            Dim ws As Worksheet
            Set ws = ThisWorkbook.Sheets(1)
            ws.Cells(1, 1).Value = "Hello"
            ws.Cells(2, 1).Value = "World"
        End Sub
        """
        vba_module.CodeModule.AddFromString(vba_code)
        wb.save()
        wb.macro('WriteData')()  # Run the VBA macro
        wb.save()
        wb.close()

    print(f"VBA code added and executed in {xlsm_file_path}.")

# Example usage
xlsx_file_path = 'example.xlsx'
xlsm_file_path = 'example.xlsm'

save_vba_and_execute(xlsx_file_path, xlsm_file_path)
