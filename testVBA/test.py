import win32com.client

def call_vba_function(excel_path, function_name, *args):
    # Open the Excel application
    excel = win32com.client.Dispatch("Excel.Application")
    
    # Open the workbook
    workbook = excel.Workbooks.Open(excel_path)
    
    # Call the VBA function
    result = excel.Application.Run(f'addData.{function_name}', *args)
    
    # Close the workbook without saving
    # workbook.Close(SaveChanges=False)
    
    # Quit the Excel application
    # excel.Quit()
    
    return result

if __name__ == "__main__":
    excel = win32com.client.Dispatch("Excel.Application")
    workbook = excel.Workbooks.Open(r"C:\Users\abarb\Documents\travail\stage et4\travail\codePlateau\data\A envoyer_antoine(non corrompue)\A envoyer\recap\recap.xlsm")
    # excel2 = win32com.client.Dispatch("Excel.Application")
    # workbook2 = excel2.Workbooks.Open(r"C:\Users\abarb\Documents\travail\stage et4\travail\codePlateau\data\A envoyer_antoine(non corrompue)\A envoyer\recap\duree_segments.xlsm")
    result = excel.Application.Run(f'addData.test2')
    result2 = excel.Application.Run(f'addData.test')

    # # excel_path = r"C:\Users\abarb\Documents\travail\stage et4\travail\codePlateau\testVBA\example.xlsm"
    # excel_path = r"C:\Users\abarb\Documents\travail\stage et4\travail\codePlateau\data\A envoyer_antoine(non corrompue)\A envoyer\recap\recap.xlsm"
    # # function_name = "MyVBAFunction"
    # function_name = "test"
    # arg1 = 5
    # arg2 = 10
    
    # # result = call_vba_function(excel_path, function_name, arg1, arg2)
    # result = call_vba_function(excel_path, function_name)
    # print(f"The result of the VBA function is: {result}")
