Public Function SearchAndImportData(sheetName As String, columnName As String, data As String) As Variant
    ' MsgBox "yu"
    Dim ws As Worksheet
    Dim dataArr() As String
    Dim i As Integer
    Dim cellData() As String
    Dim searchRange As Range
    Dim found As Range

    Set ws = ThisWorkbook.Sheets(sheetName)
    Set searchRange = ws.Columns(columnName)
    
    Dim startCol As Integer
    If columnName = "A" Then
        startCol = 7
    Else
        startCol = 10
    End If
    
    Dim result As String
    dataArr = Split(data, ";")
    ' MsgBox "yi"
    For i = LBound(dataArr) To UBound(dataArr)
        cellData = Split(dataArr(i), ":")

        ' Search for the name in the specified column
        Set found = searchRange.Find(What:=cellData(LBound(cellData)), LookIn:=xlValues, LookAt:=xlWhole)
        
        If Not found Is Nothing Then
            ' Name found, process the data
            
            For j = LBound(cellData) + 1 To UBound(cellData)
                ' Write the value in the corresponding cell in the same row
                ws.Cells(found.row, startCol + j).Value = cellData(j)
            Next j
            result = result & cellData(LBound(cellData)) & ":" & found.row & ";"
        Else
            result = result & cellData(LBound(cellData)) & ":-1;"
        End If
    Next i
    ' MsgBox SearchAndImportData
    SearchAndImportData = Left(result, Len(result) - 1)
End Function

Public Function SearchAndImportData2(sheetName As String, columnName As String, data As String) As String
    SearchAndImportData2 = "hey there"
End Function

Public Function test4() As String
    test4 = "hey there"
End Function

Sub allFileName(sheetName As String, folderPath As String)
    ' MsgBox "yo"
    Dim fileName As String
    Dim ws As Worksheet
    Dim row As Integer
    
    folderPath = folderPath & "\"
    
    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets(sheetName) ' Change "Sheet1" to your sheet name
    
    ' Initialize the starting row
    row = 2
    
    ' Get the first file name
    fileName = Dir(folderPath)
    
    ' Loop through all files in the folder
    Do While fileName <> ""
        ' Write the file name in the first column of the current row
        ws.Cells(row, 1).Value = Left(fileName, InStrRev(fileName, ".") - 1)
        
        ' Move to the next row
        row = row + 1
        
        ' Get the next file name
        fileName = Dir
    Loop
End Sub

Sub test0(sheetName As String, columnName As String, data As String)
    MsgBox "hu"
End Sub


Sub test()
    Dim result As String
    result = SearchAndImportData("Feuil1", "T", "hey:1:2:3;yo:4:5:6")
    MsgBox result
End Sub

Sub test2()
    Dim result As String
    result = SearchAndImportData("Resultats_merged", "A", "18_06_24_Benjamin_Roxane_P1:539.259:84:105:221.98799999999963:3.469:43.58800000000008:0.20100000000002183:41.2:64")
    MsgBox result
End Sub





Function MyFunction()
    MsgBox "Hello, World!"
End Function
Sub MyFunctions(sheetName As String, sheetName2 As String)
    MsgBox "Hello, World!" & sheetName & sheetName2
End Sub

Function ProcessDictionary(data As Variant) As String
    Dim i As Long
    Dim key As String
    Dim value As String
    Dim result As String
    result = ""

    ' Assuming data is a 2D array with key-value pairs
    For i = LBound(data, 1) To UBound(data, 1)
        key = data(i, 1)
        value = data(i, 2)
        result = result & key & ": " & value & vbNewLine
    Next i

    ProcessDictionary = result
End Function

