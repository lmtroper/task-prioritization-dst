Private Sub cmdDeleteTask_Click()

Dim i As Integer, check As Boolean

check = False

For i = 0 To lstDeleteTask.ListCount - 1
    If lstDeleteTask.Selected(i) = True Then

Dim msgValue As VbMsgBoxResult
    msgValue = MsgBox("Do you want to delete this task?", vbYesNo + vbInformation, "Confirmation")
    
    If msgValue = vbNo Then Exit Sub
    
    Application.ScreenUpdating = False
    Call Delete_Task
    Call prioritizationSorting
    Unload Me
    
    Application.ScreenUpdating = True
    
    check = True
    
    Exit For
   End If
Next

If check = False Then
    MsgBox ("Nothing has been selected!")

End If

End Sub

Private Sub lstDeleteTask_Click()

End Sub

Private Sub UserForm_Initialize()

Dim iRow As Long
iRow = [Counta(TaskSheet!A:A)] 'Counts the amount of non-empty entries in column A

    With DeleteTask_UserForm
    
        .lstDeleteTask.ColumnCount = 6
        .lstDeleteTask.ColumnHeads = True 'Includes the column headings
        .lstDeleteTask.ColumnWidths = "65, 80, 90, 80, 75, 75"
            
        'Intialize the list-box
        If iRow > 1 Then
             
            .lstDeleteTask.RowSource = "TaskSheet!A2:F" & iRow ' iRow > 1 so add the information from the last row (assessment) inputted
                  
        Else
            .lstDeleteTask.RowSource = "TaskSheet!A2:F2" 'Add the row after the column headers
                  
        End If
    
    End With
    
End Sub
