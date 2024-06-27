Public Function SearchAndImportData(sheetName As String, columnName As String, data As String) As String
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
    For i = LBound(dataArr) To UBound(dataArr)
        cellData = Split(dataArr(i), ":")

        ' Search for the name in the specified column
        Set found = searchRange.Find(What:=cellData(LBound(cellData)), LookIn:=xlValues, LookAt:=xlWhole)
        
        If Not found Is Nothing Then
            ' Name found, process the data
            
            For j = LBound(cellData) + 1 To UBound(cellData)
                ' Write the value in the corresponding cell in the same row
                ws.Cells(found.Row, startCol + j).Value = cellData(j)
            Next j
            result = result & cellData(LBound(cellData)) & ":" & found.Row & ";"
        Else
            result = result & cellData(LBound(cellData)) & ":-1;"
        End If
    Next i
    SearchAndImportData = Left(result, Len(result) - 1)
End Function

Sub test()
    Dim result As String
    result = SearchAndImportData("Feuil1", "T", "hey:1:2:3;yo:4:5:6")
    MsgBox result
End Sub

