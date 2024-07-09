Public Function SearchAndImportData(sheetName As String, columnName As String, tabs As Variant, data As Variant) As Variant
    Dim ws As Worksheet
    Dim i As Integer
    Dim searchRange As Range
    Dim found As Range

    Set ws = ThisWorkbook.Sheets(sheetName)
    Set searchRange = ws.Columns(columnName)
    
    Dim startCol As Integer
    If columnName = "A" Then
        startCol = 1
    Else
        startCol = 10
    End If
    
    For i = LBound(tabs) To UBound(tabs)
        ws.Cells(1, i + startCol + 1).Value = tabs(i)
    Next i
    
    Dim result() As Variant
    ReDim result(LBound(data) To UBound(data), 1 To 2)

    For i = LBound(data) To UBound(data)

        ' Search for the name in the specified column
        result(i, 1) = data(i, LBound(tabs))
        Set found = searchRange.Find(What:=result(i, 1), LookIn:=xlValues, LookAt:=xlWhole)
        
        If Not found Is Nothing Then
            ' Name found, process the data
            
            For j = LBound(data, 2) + 1 To UBound(data, 2)
                ' Write the value in the corresponding cell in the same row
                ws.Cells(found.row, startCol + j).value = data(i, j)
            Next j
            result(i, 2) = found.row
        Else
            result(i, 2) = -1
        End If
    Next i
    ' MsgBox SearchAndImportData
    SearchAndImportData = result
End Function

Public Function SearchAndImportData2(sheetName As String, columnName As String, data As String) As String
    SearchAndImportData2 = "hey there"
End Function

Public Function test4() As String
    test4 = "hey there"
End Function

Sub allFileName(sheetName As String, file_names As Variant)
    Dim ws As Worksheet
    Dim i As Long
    
    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)
    
    ' Loop through the file_names array and write to the first column starting from the second row
    For i = LBound(file_names) To UBound(file_names)
        ws.Cells(i + 2, 1).Value = file_names(i)
    Next i
End Sub


Sub allFileName2(sheetName As String, file_names As Variant)
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
        ws.Cells(row, 1).value = Left(fileName, InStrRev(fileName, ".") - 1)
        
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
Sub MyFunctions(data As Variant)
    MsgBox "Hello, World!" & LBound(data(0))
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

Public Sub func2(dataList As Variant)
    Dim ws As Worksheet
    Dim i As Integer
    Dim sublist As Variant
    
    Set ws = ThisWorkbook.Sheets("MySheet")
    
    ' Loop through each sublist in the array
    For i = LBound(dataList) To UBound(dataList)
        MsgBox "Size of sublist " & i + 1 & ": " & UBound(dataList(i)) - LBound(dataList(i)) + 1
    Next i
End Sub

