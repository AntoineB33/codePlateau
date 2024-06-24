import win32com.client

# Function to check if the workbook is already open
def is_workbook_open(excel, workbook_name):
    for workbook in excel.Workbooks:
        if workbook.Name == workbook_name:
            return True
    return False

# Create a new instance of Excel application
try:
    excel = win32com.client.GetActiveObject("Excel.Application")
    excel_visible = False
except:
    excel = win32com.client.Dispatch("Excel.Application")
    excel_visible = True

# Name of the workbook to check/open
workbook_path = r"testVBA.xlsm"
workbook_name = "testVBA.xlsm"

# Open the workbook if it is not already open
if is_workbook_open(excel, workbook_name):
    workbook = excel.Workbooks(workbook_name)
    workbook_opened = False
else:
    workbook = excel.Workbooks.Open(workbook_path)
    workbook_opened = True

# Define the data to be passed as a string
data = [["A1", "Hello"], ["B2", "World"]]
data_str = ";".join([f"{item[0]}:{item[1]}" for item in data])

# Run the VBA function with the data string
excel.Application.Run("ThisWorkbook.ImportData", data_str)

# Save and close the workbook if it was opened by this script
if workbook_opened:
    workbook.Save()
    workbook.Close()

# Quit the Excel application if it was started by this script
if excel_visible:
    excel.Application.Quit()
