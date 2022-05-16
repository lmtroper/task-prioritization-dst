Private Sub lstEditTask_Click()

End Sub

Private Sub UserForm_Initialize()

Dim msgValue As VbMsgBoxResult
Dim iRow As Long
iRow = [Counta(TaskSheet!A:A)] 'Counts the amount of non-empty entries in column A

    With EditTask_UserForm1

        .lstEditTask.ColumnCount = 6
        .lstEditTask.ColumnHeads = True 'Includes the column headings
        .lstEditTask.ColumnWidths = "65, 80, 90, 80, 75, 75"

        'Intialize the list-box
        If iRow > 1 Then

            .lstEditTask.RowSource = "TaskSheet!A2:F" & iRow ' iRow > 1 so add the information from the last row (assessment) inputted

        Else
            .lstEditTask.RowSource = "TaskSheet!A2:F2" 'Add the row after the column headers

        End If

    End With

End Sub

Private Sub UpdateTaskButton_Click()

Dim i As Integer, check As Boolean

check = False

For i = 0 To lstEditTask.ListCount - 1
    If lstEditTask.Selected(i) = True Then
    
    check = True

    EditTask_UserForm1.Hide
    EditTask_UserForm2.Show
    
    Call HighlightTasks
    Exit For
End If
Next


If check = False Then
    MsgBox ("Nothing has been selected!")

End If


End Sub
