Sub HighlightHighExpenses()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(1)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim i As Long
    For i = 2 To lastRow
        If IsNumeric(ws.Cells(i, 2).Value) Then
            If CDbl(ws.Cells(i, 2).Value) > 1000 Then
                ws.Rows(i).Interior.Color = RGB(255, 150, 150)
            End If
        End If
    Next i
End Sub
