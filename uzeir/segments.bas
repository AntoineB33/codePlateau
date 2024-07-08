Public Sub ImportSegments(row_found As Variant, tabs As Variant, data As Variant)

    Dim xlEdgeLeft As Long
    Dim xlEdgeTop As Long
    Dim xlEdgeBottom As Long
    Dim xlEdgeRight As Long
    Dim borders As Variant
    Dim border As Variant

    Dim dataRow() As String
    Dim dataSegments() As String
    Dim i As Integer, j As Integer, n As Integer
    Dim dataRowData() As String
    Dim dataSegmentsData() As String
    Dim cell As Range
    
    xlEdgeLeft = 7
    xlEdgeTop = 8
    xlEdgeBottom = 9
    xlEdgeRight = 10
    
    borders = Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight)

    ' create a list of all the Sheets
    Dim sheetList As Variant


    Dim wbs As Object
    Set wbs = CreateObject("Scripting.Dictionary")
    
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        wbs(ws.Name) = ws
    Next ws
    
    For i = LBound(dataRow) To UBound(dataRow)
        n = CInt(row_found(i, 1))
        For k = LBound(data(i)) To UBound(data(i))
            If tabs(k) = "colors" Then
                For m = LBound(data(i)) To UBound(data(i))
                    If tabs(m) <> "colors" Then
                        For j = LBound(data(i, k)) To UBound(data(i, k))
                            wbs(tabs(k)).Cells(n, j + 1).Interior.Color = data(i, k, j)
                        Next j
                    End If
                Next m
            Else
                wbs(tabs(k)).Cells(1, 1).Value = "Repas"
                wbs(tabs(k)).Cells(n, 1).Value = row_found(i, 0)
                For j = LBound(data(i, k)) To UBound(data(i, k))
                    wbs(tabs(k)).Cells(n, j + 1).Value = data(i, k, j)
                Next j
            End If
            
            ' For j = UBound(dataSegmentsData) To LBound(dataSegmentsData) Step -1
            '     ' Define the range you want to apply the borders to
            '     Set cell = ws.Cells(n, j + 2)
            '     exitFor = False
            '     If cell.borders(xlEdgeLeft).LineStyle = xlContinuous Then
            '         exitFor = True
            '     End If
            '     For Each border In borders
            '         cell.borders(border).LineStyle = xlContinuous
            '     Next border
            '     If exitFor Then
            '         Exit For
            '     End If
            ' Next j
        Next k
    Next i
    If segmentMax > -1 Then
        For i = segmentMax + 1 To 1 Step -1
            Set cell = ws.Cells(1, i + 1)
            If cell.Value = "Segment " & i Then
                Exit For
            Else
                cell.Value = "Segment " & i
            End If
        Next i
    End If
End Sub

Sub test()
    ImportSegments "Feuil1", "hey:3;yo2:6", "4.5:16711935:5.5:16711935;5.5:16711935"
End Sub








