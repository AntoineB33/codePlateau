import win32com.client

def create_xlsm_with_named_sheet(file_path, sheet_name):
    # Create an Excel application instance
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False  # Set to True if you want Excel to be visible

    try:
        # Add a new workbook
        workbook = excel.Workbooks.Add()

        # Excel starts with 1 sheet by default, rename it
        workbook.Sheets(1).Name = sheet_name

        # Save the workbook as an .xlsm file
        workbook.SaveAs(file_path, FileFormat=52)  # 52 corresponds to xlOpenXMLWorkbookMacroEnabled

    except Exception as e:
        print(f"Error: {e}")
    finally:
        # Close the workbook and quit Excel
        workbook.Close(SaveChanges=False)  # Change to True if you've made changes you want to save
        excel.Quit()

# Example usage
file_path = r"C:\Users\abarb\Documents\travail\stage et4\travail\codePlateau\data\A envoyer_antoine(non corrompue)\A envoyer\recap\recap.xlsm"
sheet_name = "MyCustomSheet"
create_xlsm_with_named_sheet(file_path, sheet_name)