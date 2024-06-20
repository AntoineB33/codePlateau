import win32com.client
import os

def open_excel(file_path, sheet_name):
    excel = win32com.client.Dispatch("Excel.Application")
    # excel.Visible = True

    file_path = os.path.abspath(file_path)
    try:
        workbook = excel.Workbooks.Open(file_path)
    except:
        workbook = excel.Workbooks(file_path.split("\\")[-1])

    sheet = workbook.Sheets(sheet_name)
    return excel, workbook, sheet

def add_or_replace_vba_module(workbook, module_name, vba_code):
    vb_component = None
    for vb_component in workbook.VBProject.VBComponents:
        if vb_component.Name == module_name:
            workbook.VBProject.VBComponents.Remove(vb_component)
            break
    
    module = workbook.VBProject.VBComponents.Add(1)  # 1 = vbext_ct_StdModule
    module.Name = module_name
    module.CodeModule.AddFromString(vba_code)

def run_vba_macro(excel, workbook, macro_name):
    excel.Application.Run(f"{workbook.Name}!{macro_name}")

def save_as_macro_enabled(workbook, file_path):
    # Change extension to .xlsm
    new_file_path = file_path.replace('.xlsx', '.xlsm')
    workbook.SaveAs(new_file_path, FileFormat=52)  # 52 = xlOpenXMLWorkbookMacroEnabled
    return new_file_path

def main():
    excel_all_path = "path_to_all_workbook.xlsx"
    sheet_name = "Sheet1"  # Adjust as needed
    excel_segments_path = "path_to_segments_workbook.xlsx"
    sheet_name_segment = "Sheet1"  # Adjust as needed
    
    vba_code_all = """
    Sub UpdateMainWorkbook()
        Dim ws As Worksheet
        Set ws = ThisWorkbook.Sheets("Sheet1") ' Adjust sheet name as needed
        
        Dim usedRange As Range
        Set usedRange = ws.UsedRange
        
        Dim rowNum As Long
        Dim searchString As String
        Dim cell As Range
        Dim segmentNum As Integer
        
        For Each cell In usedRange.Columns(20).Cells
            rowNum = cell.Row
            searchString = "some_condition" ' Replace with actual search string logic
            
            If cell.Value = searchString Then
                ' Edit specified cells
                ws.Cells(rowNum, 11).Value = "new_value_1" ' Replace with actual logic
                ws.Cells(rowNum, 12).Value = "new_value_2" ' Replace with actual logic
                ' Add more cells as needed
                
                ' Update segments
                segmentNum = 1
                For i = 2 To 10 ' Replace with actual segment count
                    ws.Cells(rowNum, i).Value = segmentNum ' Replace with actual segment value
                    ws.Cells(rowNum, i).Interior.Color = RGB(255, 255, 255) ' Replace with actual color logic
                    
                    ' Set borders
                    With ws.Cells(rowNum, i).Borders
                        .LineStyle = xlContinuous
                    End With
                    
                    segmentNum = segmentNum + 1
                Next i
            End If
        Next cell
        
        ' Save workbook
        ThisWorkbook.Save
    End Sub
    """
    
    vba_code_segments = """
    Sub UpdateSegmentsWorkbook()
        Dim ws As Worksheet
        Set ws = ThisWorkbook.Sheets("Sheet1") ' Adjust sheet name as needed
        
        ' Your logic to update the segments workbook
        
        ' Save workbook
        ThisWorkbook.Save
    End Sub
    """

    # Open and update the main workbook
    excel_all, workbook_all, sheet_all = open_excel(excel_all_path, sheet_name)
    if excel_all_path.endswith('.xlsx'):
        excel_all_path = save_as_macro_enabled(workbook_all, excel_all_path)
    add_or_replace_vba_module(workbook_all, "Module1", vba_code_all)
    run_vba_macro(excel_all, workbook_all, "UpdateMainWorkbook")

    # Open and update the segments workbook
    excel_segments, workbook_segments, sheet_segments = open_excel(excel_segments_path, sheet_name_segment)
    if excel_segments_path.endswith('.xlsx'):
        excel_segments_path = save_as_macro_enabled(workbook_segments, excel_segments_path)
    add_or_replace_vba_module(workbook_segments, "Module1", vba_code_segments)
    run_vba_macro(excel_segments, workbook_segments, "UpdateSegmentsWorkbook")

    # Clean up
    workbook_all.Close(SaveChanges=True)
    workbook_segments.Close(SaveChanges=True)
    excel_all.Quit()
    excel_segments.Quit()

if __name__ == "__main__":
    main()
