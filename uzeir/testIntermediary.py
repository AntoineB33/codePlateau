import os


file_path = r"C:\\Users\\abarb\\Documents\\travail\\stage et4\\travail\\codePlateau\\data\\A envoyer(pate a modeler)\\A envoyer\\recap\\recap.xlsm"

if os.path.exists(file_path):
    if is_workbook_open(excelLocal, workbook_name):
        workbookLocal = excelLocal.Workbooks.Open(file_path)
        check_and_add_sheet(workbookLocal, sheet_name)
        workbook_opened = False
    else:
        workbookLocal = excelLocal.Workbooks.Open(file_path)
        check_and_add_sheet(workbookLocal, sheet_name)
else:
    workbookLocal = excelLocal.Workbooks.Add()
    workbookLocal.Sheets.Add().Name = sheet_name
    workbookLocal.SaveAs(file_path, FileFormat=52)