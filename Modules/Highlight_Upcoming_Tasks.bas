Sub HighlightUpcomingTasks()

Dim RowSearch As Long

For RowSearch = 2 To Sheet1.Range("F100000").End(xlUp).Row  'Finding non-empty row in date column
'If date is overdue, then highlight red. If data is between today and the next 3 days, highlight yellow
If Sheet1.Cells(RowSearch, "F") < Date Then
    Cells(RowSearch, "F").Interior.Color = vbRed
ElseIf Sheet1.Cells(RowSearch, "F") >= Date And Sheet1.Cells(RowSearch, "F") <= Date + 3 Then
    Cells(RowSearch, "F").Interior.Color = vbYellow
Else
    Cells(RowSearch, "F").Select
    
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
End If
Next

End Sub
