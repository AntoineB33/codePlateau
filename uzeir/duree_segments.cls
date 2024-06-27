Public Sub ImportSegments(sheetName As String, row_found As String, data_str_segments As String)

    Dim xlEdgeLeft As Long
    Dim xlEdgeTop As Long
    Dim xlEdgeBottom As Long
    Dim xlEdgeRight As Long
    Dim borders As Variant
    Dim border As Variant

    Dim ws As Worksheet
    Dim dataRow() As String
    Dim dataSegments() As String
    Dim i As Integer, j As Integer, n As Integer
    Dim dataRowData() As String
    Dim dataSegmentsData() As String
    Dim cell As Range
    Dim segmentMaxI As Integer
    Dim segmentMax As Integer
    segmentMax = 0
    
    xlEdgeLeft = 7
    xlEdgeTop = 8
    xlEdgeBottom = 9
    xlEdgeRight = 10
    
    borders = Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight)

    Set ws = ThisWorkbook.Sheets(sheetName)
    
    ws.Cells(1, 1).Value = "Repas"
    
    dataRow = Split(row_found, ";")
    dataSegments = Split(data_str_segments, ";")
    For i = LBound(dataRow) To UBound(dataRow)
        dataRowData = Split(dataRow(i), ":")
        n = CInt(dataRowData(1))
        Set cell = ws.Cells(n, 1)
        cell.Value = dataRowData(0)
        dataSegmentsData = Split(dataSegments(i), ":")
        segmentMaxI = 1
        For j = LBound(dataSegmentsData) To (UBound(dataSegmentsData) - 1) / 2
            Set cell = ws.Cells(n, j + 2)
            cell.Value = dataSegmentsData(j * 2)
            cell.Interior.Color = dataSegmentsData(j * 2 + 1)
            If segmentMaxI > segmentMax Then
                segmentMax = segmentMaxI
            End If
            segmentMaxI = segmentMaxI + 1
        Next j
        For j = UBound(dataSegmentsData) + 1 To LBound(dataSegmentsData) + 1 Step -1
            ' Define the range you want to apply the borders to
            Set cell = ws.Cells(n, j)
            If cell.borders(xlEdgeRight).LineStyle = xlContinuous Then
                Exit For
            Else
                For Each border In borders
                    cell.borders(border).LineStyle = xlContinuous
                Next border
            End If
        Next j
    Next i
    For i = segmentMax To 1 Step -1
        Set cell = ws.Cells(1, i + 1)
        If cell.Value = "Segment " & i Then
            Exit For
        Else
            cell.Value = "Segment " & i
        End If
    Next i
End Sub

Sub test()
    ImportSegments "Feuil1", "hey:2;yo2:6", "4.5:16711935:5.5:16711935;5.5:16711935"
End Sub






