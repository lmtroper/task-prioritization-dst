Sub Delete_Task()

Dim total_rows As Integer
Dim i As Integer, k As Integer 'Integers used for looping
Dim startCell As Range

Dim overview As Worksheet, dataSheet As Worksheet
Set overview = ThisWorkbook.Sheets("TaskSheet")
Set dataSheet = ThisWorkbook.Sheets("Data Sheet")

For i = 0 To DeleteTask_UserForm.lstDeleteTask.ListCount - 1  'Loop through each assessment in Complete UserForm 1 list-box
    If DeleteTask_UserForm.lstDeleteTask.Selected(i) = True Then
        dataSheet.Rows(i + 2).Delete

                Set startCell = overview.Range(Cells(i + 2, 1), Cells(i + 2, 6))
                startCell.Select
                Selection.Delete Shift:=xlUp 'Delete task from the TaskSheet
                
        End If
        
    Next i

End Sub


Sub Show_Delete_Task_Form()
    
    DeleteTask_UserForm.Show 'Opens the Delete Task User Form when user clicks on the Delete Task button
    
End Sub

