import win32com.client

def edit_open_excel(file_path, sheet_name, cell, new_value, background_color, border_style):
    # Connect to Excel
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True

    # Open the workbook (or attach to it if it's already open)
    try:
        workbook = excel.Workbooks.Open(file_path)
    except:
        workbook = excel.Workbooks(file_path.split("\\")[-1])

    # Access the specified worksheet
    sheet = workbook.Sheets(sheet_name)

    # Edit the specified cell
    cell_range = sheet.Range(cell)
    cell_range.Value = new_value

    # Change the background color
    cell_range.Interior.Color = background_color

    # Set the border style
    # Constants for border styles: https://docs.microsoft.com/en-us/office/vba/api/excel.xllinestyle
    xlEdgeLeft = 7
    xlEdgeTop = 8
    xlEdgeBottom = 9
    xlEdgeRight = 10

    borders = [xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight]
    for border in borders:
        cell_range.Borders(border).LineStyle = border_style

    # Save changes
    workbook.Save()

    # Optionally close the workbook and Excel application
    # workbook.Close()
    # excel.Quit()

# Example usage
file_path = r"C:\Users\abarb\Documents\travail\stage et4\travail\codePlateau\Resultats exp bag_couverts\Resultats exp bag_couverts\test.xlsx"
sheet_name = 'Feuil1'
cell = 'A1'
new_value = 'Hello, World!'
background_color = 0x00FF00  # Green color in RGB
border_style = 1  # Continuous line style

edit_open_excel(file_path, sheet_name, cell, new_value, background_color, border_style)