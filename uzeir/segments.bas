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
    segmentMax = -1
    
    xlEdgeLeft = 7
    xlEdgeTop = 8
    xlEdgeBottom = 9
    xlEdgeRight = 10
    
    borders = Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight)

    ' create a list of all the Sheets
    Dim sheetList As Variant


    Set ws = ThisWorkbook.Sheets(sheetName)
    
    ws.Cells(1, 1).Value = "Repas"
    
    dataRow = Split(row_found, ";")
    dataSegmentsLines = Split(data_str_segments, ";")
    For i = LBound(dataRow) To UBound(dataRow)
        dataSegmentsTabs = Split(dataSegmentsLines(i), "_")
        For k = LBound(dataSegmentsTabs) To UBound(dataSegmentsTabs)
            If dataSegmentsTabs(k)(0) = "colors" Then
                For k = LBound(dataSegmentsTabs) To UBound(dataSegmentsTabs)
                    If dataSegmentsTabs(k)(0) <> "colors" Then
                        dataSegmentsData = Split(dataSegmentsTabs(i), ":")
                        If UBound(dataSegmentsData) > segmentMax Then
                            segmentMax = UBound(dataSegmentsData)
                        End If
                    End If
                Next k
            Else
                dataRowData = Split(dataRow(i), ":")
                n = CInt(dataRowData(1))
                Set cell = ws.Cells(n, 1)
                cell.Value = dataRowData(0)
                dataSegmentsData = Split(dataSegmentsTabs(i), ":")
                For j = LBound(dataSegmentsData) To UBound(dataSegmentsData)
                    Set cell = ws.Cells(n, j + 1)
                    cell.Value = dataSegmentsData(j)
                    cell.Interior.Color = dataSegmentsData(j * 2 + 1)
                Next j
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
            End If
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








