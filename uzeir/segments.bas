Sub AddScriptingRuntimeReference()
    Dim ref As Object
    Dim refName As String
    refName = "Scripting"

    ' Check if the reference is already added
    For Each ref In ThisWorkbook.VBProject.References
        If ref.Name = refName Then
            Exit Sub
        End If
    Next ref

    ' Add the reference
    ThisWorkbook.VBProject.References.AddFromGuid "{420B2830-E718-11CF-893D-00A0C9054228}", 1, 0
End Sub

Public Sub ImportSegments(row_found As Variant, tabs As Variant, data As Variant)

    Dim xlEdgeLeft As Long
    Dim xlEdgeTop As Long
    Dim xlEdgeBottom As Long
    Dim xlEdgeRight As Long
    Dim borders As Variant
    Dim border As Variant
    Dim segmentMax As Integer

    Dim i As Integer, j As Integer, n As Integer
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
    
    AddScriptingRuntimeReference
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        wbs.Add ws.Name, ws
    Next ws

    segmentMax = -1
    
    For i = LBound(row_found) To UBound(row_found)
        n = CInt(row_found(i, 1))
        For k = LBound(data, 2) To UBound(data, 2)
            If tabs(k) = "colors" Then
                For m = LBound(data, 2) To UBound(data, 2)
                    If tabs(m) <> "colors" Then
                        For j = LBound(data(i, k)) To UBound(data(i, k))
                            wbs(tabs(m)).Cells(n, j + 2).Interior.Color = data(i, k)(j)
                        Next j
                    End If
                Next m
            Else
                wbs(tabs(k)).Cells(1, 1).Value = "Repas"
                wbs(tabs(k)).Cells(n, 1).Value = row_found(i, 0)
                For j = LBound(data(i, k)) To UBound(data(i, k))
                    wbs(tabs(k)).Cells(n, j + 2).Value = data(i, k)(j)
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
        If UBound(data(i, 0)) > segmentMax Then
            segmentMax = UBound(data(i, 0))
        End If
    Next i
    If segmentMax > -1 Then
        ' For i = segmentMax + 1 To 1 Step -1
        '     Set cell = wbs(tabs(0)).Cells(1, i + 1)
        '     If cell.Value = "Segment " & i Then
        '         Exit For
        '     Else
        '         cell.Value = "Segment " & i
        '     End If
        ' Next i

        Dim headers() As String
        ' Initialize the headers array
        ReDim headers(1 To segmentMax + 1)
        ' Generate headers dynamically
        For j = 1 To segmentMax + 1
            headers(j) = "Segment " & j
        Next j
        For k = LBound(data, 2) To UBound(data, 2)
            wbs(tabs(k)).Range(wbs(tabs(k)).Cells(1, 2), wbs(tabs(k)).Cells(1, segmentMax + 1)).Value = headers
        Next k
    End If
End Sub

Sub test()
    ImportSegments "Feuil1", "hey:3;yo2:6", "4.5:16711935:5.5:16711935;5.5:16711935"
End Sub











