Sub Completed_Task()

Dim total_rows As Integer, completeTotal_rows As Integer
Dim i As Integer, k As Integer 'Intergers used for looping
Dim startCellDelete As Range

Dim overview As Worksheet, dataSheet As Worksheet, completedTasks As Worksheet
Set overview = ThisWorkbook.Sheets("TaskSheet")
Set dataSheet = ThisWorkbook.Sheets("Data Sheet")
Set completedTasks = ThisWorkbook.Sheets("Completed Tasks")

total_rows = Sheet1.Cells(Rows.count, 1).End(xlUp).Row 'Identifies last row entry in column 1 on TaskSheet
completeTotal_rows = Sheets("Completed Tasks").Cells(Rows.count, 1).End(xlUp).Row 'Identifies last row entry in column 1 on the completed task overview sheet

    For i = 0 To CompletedTask_UserForm.lstCompleteTask.ListCount - 1  'Loop through each assessment in Complete UserForm 1 list-box
        If CompletedTask_UserForm.lstCompleteTask.Selected(i) = True Then
        
            With completedTasks: 'Copies the cell values from the selected task in TaskSheet to the Completed Task Sheet
                .Cells(completeTotal_rows + 1, 1) = overview.Cells(i + 2, 2).Value
                .Cells(completeTotal_rows + 1, 2) = overview.Cells(i + 2, 3).Value
                .Cells(completeTotal_rows + 1, 3) = overview.Cells(i + 2, 4).Value
                .Cells(completeTotal_rows + 1, 4) = overview.Cells(i + 2, 6).Value
                .Cells(completeTotal_rows + 1, 5) = Format(Now(), "short date") 'Adds the date the task was marked as completed
                
            End With

            With dataSheet
                dataSheet.Rows(i + 2).Delete
            
           End With
            
            'Deletes the data associated with the selected task in the TaskSheet
            'No need to loop again as the Data Sheet and TaskSheet have the same task list sorted by priority ranks
            Set startCellDelete = overview.Range(Cells(i + 2, 1), Cells(i + 2, 6))
            startCellDelete.Select
            Selection.Delete Shift:=xlUp
                    
        End If
        
    Next i

End Sub
Sub sort_CompletedTasks()

    'Sorts the tasks the based on date of completion
    'Record Macro Method
    
    Range("E2").Activate
    ActiveWorkbook.Worksheets("Completed Tasks").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Completed Tasks").Sort.SortFields.Add2 Key:=Range( _
        "E2:E17"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Completed Tasks").Sort
        .SetRange Range("E1:E17")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
End Sub

Sub Show_Complete_Task_Form()
    
    CompletedTask_UserForm.Show 'Opens the Completed Task User Form when user clicks on the Complete Task button
    

End Sub

