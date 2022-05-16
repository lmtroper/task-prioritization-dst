Sub prioritization_NewTask() 'Assigns new tasks with a prioritization value

Dim length As Integer
Dim difficulty As Integer
Dim importance As Integer
Dim differenceDays As Integer

Dim urgency As Variant
Dim effort As Variant
Dim prioritization As Variant

Dim taskLength As String
Dim taskDifficulty As String
Dim taskImportance As String
        
        'Determining the value for taskLength and length
        'taskLength is the value inputted in the Data Sheet
        'length is the value used in the prioritization algorithm
        If AddTask_UserForm.optTime1.Value = True Then 'Checks which option button is filled
            length = 1
            taskLength = AddTask_UserForm.optTime1.Caption
        ElseIf AddTask_UserForm.optTime2.Value = True Then
            length = 3
            taskLength = AddTask_UserForm.optTime2.Caption
        ElseIf AddTask_UserForm.optTime3.Value = True Then
            length = 6
            taskLength = AddTask_UserForm.optTime3.Caption
        ElseIf AddTask_UserForm.optTime4.Value = True Then
            length = 10
            taskLength = AddTask_UserForm.optTime4.Caption
        ElseIf AddTask_UserForm.optTime5.Value = True Then
            length = 13
            taskLength = AddTask_UserForm.optTime5.Caption
            
        End If
        
        'Determining the value for taskDifficulty and difficulty
        'taskDifficulty is the value inputted in the Data Sheet
        'difficulty is the value used in the prioritization algorithm
        If AddTask_UserForm.optDifficulty1.Value = True Then 'Checks which option button is filled
            difficulty = 1
            taskDifficulty = AddTask_UserForm.optDifficulty1.Caption
        ElseIf AddTask_UserForm.optDifficulty2.Value = True Then
            difficulty = 2
            taskDifficulty = AddTask_UserForm.optDifficulty2.Caption
        ElseIf AddTask_UserForm.optDifficulty3.Value = True Then
            difficulty = 3
            taskDifficulty = AddTask_UserForm.optDifficulty3.Caption
        ElseIf AddTask_UserForm.optDifficulty4.Value = True Then
            difficulty = 4
            taskDifficulty = AddTask_UserForm.optDifficulty4.Caption
        ElseIf AddTask_UserForm.optDifficulty5.Value = True Then
            difficulty = 5
            taskDifficulty = AddTask_UserForm.optDifficulty5.Caption
    
        End If
        
        'Determining the value for taskImportance and importance
        'taskImportance is the value inputted in the Data Sheet
        'importance is the value used in the prioritization algorithm
        If AddTask_UserForm.optImportance1.Value = True Then 'Checks which option button is filled
            importance = 1
            taskImportance = AddTask_UserForm.optImportance1.Caption
        ElseIf AddTask_UserForm.optImportance2.Value = True Then
            importance = 2
            taskImportance = AddTask_UserForm.optImportance2.Caption
        ElseIf AddTask_UserForm.optImportance3.Value = True Then
            importance = 3
            taskImportance = AddTask_UserForm.optImportance3.Caption
        ElseIf AddTask_UserForm.optImportance4.Value = True Then
            importance = 4
            taskImportance = AddTask_UserForm.optImportance4.Caption
        ElseIf AddTask_UserForm.optImportance5.Value = True Then
            importance = 5
            taskImportance = AddTask_UserForm.optImportance5.Caption
    
        End If
        
  
    effort = (0.5 * difficulty) + (0.5 * length) 'Assigning Formula
    
    'Finding the difference between today's date and the deadline date inputted by the user
    'Dependent on whether the user inputs an additional urgent deadline or just inputs a regular deadline
    'The value for urgency is increased by 1 for tasks with urgent deadlines
    
    If AddTask_UserForm.txtUrgentDeadline.Value <> "" Then 'Only if user submits a date for the urgent deadline
        differenceDays = (DateDiff("d", Date, CDate(AddTask_UserForm.txtUrgentDeadline.Value)))
        
        If differenceDays > 0 Then
            urgency = 2 / differenceDays
        ElseIf differenceDays = 0 Then
            urgency = 2
        ElseIf differenceDays < 0 Then
            urgency = 2 + -(differenceDays)
        End If
        
    Else
        differenceDays = (DateDiff("d", Date, CDate(AddTask_UserForm.txtDeadline.Value)))
        
        
        'If the difference between days is greater than zero, urgency = 1/t where t represents the difference in days
        'If t = 0, urgency = 1
        'For every day that passes the due date, urgency = 1 + t
        
        If differenceDays > 0 Then
            urgency = 1 / differenceDays
        ElseIf differenceDays = 0 Then
            urgency = 1
        ElseIf differenceDays < 0 Then
            urgency = 1 + -(differenceDays)
        End If
        
    End If
    
   
   
   'If importance is greater than or equal to 3, the greater the value of effort, the greater the prioritization
   'If importance is less than 3, tasks that require more effort will be less prioritized than easy, less important tasks
   
    If importance >= 3 Then
        prioritization = (40 / 100) * importance + (36 / 100) * urgency + (24 / 100) * effort
        
    ElseIf importance < 3 Then
        prioritization = (40 / 100) * importance + (36 / 100) * urgency + (24 / 100) * 1 / effort
    End If
    
'Assigning the priority rankings beside each task
Dim overview As Worksheet, dataSheet As Worksheet
Set overview = ThisWorkbook.Sheets("TaskSheet")
Set dataSheet = ThisWorkbook.Sheets("Data Sheet")

Dim iRow As Long
iRow = [Counta(TaskSheet!A:A)] + 1  'Counts how many entries are in the column A of TaskSheet, iRow assigned to the row after the last entry(empty row)



    With dataSheet: 'Sets the all of the cells associated with the task with the values from the Add Task Userform onto the Data Sheet
        .Cells(iRow, 1) = prioritization
        .Cells(iRow, 2) = AddTask_UserForm.txtName.Value
        .Cells(iRow, 3) = taskLength
        .Cells(iRow, 4) = taskDifficulty
        .Cells(iRow, 5) = taskImportance
        .Cells(iRow, 6).Value = iRow - 1
    
    End With
    
    With overview:
        .Cells(iRow, 1) = prioritization 'Sets the priotization cell with the value obtained onto TaskSheet
        .Cells(iRow, 1).Select
        Selection.Font.Color = vbWhite
        
    End With


End Sub

Sub prioritization_ExistingTasks()

Dim length As Integer
Dim difficulty As Integer
Dim importance As Integer
Dim differenceDays As Integer

Dim iRow As Long
iRow = [Counta(TaskSheet!A:A)]

Dim urgency As Variant
Dim effort As Variant
Dim prioritization As Variant

Dim overview As Worksheet, dataSheet As Worksheet
Set overview = ThisWorkbook.Sheets("TaskSheet")
Set dataSheet = ThisWorkbook.Sheets("Data Sheet")


For i = 2 To iRow
    If dataSheet.Cells(i, 3).Text = "0 - 1 hour" Then 'Checks which option button is filled
        length = 1
    ElseIf dataSheet.Cells(i, 3).Text = "1-3 hours" Then
        length = 3
    ElseIf dataSheet.Cells(i, 3).Text = "3-6 hours" Then
        length = 6
    ElseIf dataSheet.Cells(i, 3).Text = "6-10 hours" Then
        length = 10
    ElseIf dataSheet.Cells(i, 3).Text = "10+ hours" Then
            length = 13
    End If
    
    difficulty = dataSheet.Cells(i, 4)
    importance = dataSheet.Cells(i, 5)
    
    effort = (0.5 * difficulty) + (0.5 * length) 'Assigning Formula
    
    If overview.Cells(i, 5).Text <> "" Then 'Only if user submits a date for the urgent deadline
        differenceDays = (DateDiff("d", Date, CDate(overview.Cells(i, 5).Text)))
        
        If differenceDays > 0 Then
            urgency = 2 / differenceDays
        ElseIf differenceDays = 0 Then
            urgency = 2
        ElseIf differenceDays < 0 Then
            urgency = 2 + -(differenceDays)
        End If
        
    Else
        differenceDays = (DateDiff("d", Date, CDate(overview.Cells(i, 6).Text)))
        
        
        'If the difference between days is greater than zero, urgency = 1/t where t represents the difference in days
        'If t = 0, urgency = 1
        'For every day that passes the due date, urgency = 1 + t
        
        If differenceDays > 0 Then
            urgency = 1 / differenceDays
        ElseIf differenceDays = 0 Then
            urgency = 1
        ElseIf differenceDays < 0 Then
            urgency = 1 + -(differenceDays)
        End If
        
    End If

    If importance >= 3 Then
        prioritization = (40 / 100) * importance + (36 / 100) * urgency + (24 / 100) * effort
        
    ElseIf importance < 3 Then
        prioritization = (40 / 100) * importance + (36 / 100) * urgency + (24 / 100) * 1 / effort
    End If

    With dataSheet: 'Sets the all of the cells associated with the task with the values from the Add Task Userform onto the Data Sheet
        .Cells(i, 1) = prioritization
    End With
    
    With overview: 'Sets the all of the cells associated with the task with the values from the Add Task Userform onto the Data Sheet
        .Cells(i, 1) = prioritization
        .Cells(i, 1).Select
        Selection.Font.Color = vbWhite
    End With

Next


End Sub

Sub prioritization_Following_EditTask() 'Re-prioritizes an edited task and assigns a new priority value

Dim length As Integer
Dim difficulty As Integer
Dim importance As Integer
Dim differenceDays As Integer

Dim urgency As Variant
Dim effort As Variant
Dim prioritization As Variant
Dim taskLength As String
Dim taskDifficulty As String
Dim taskImportance As String
        
        'Determining the value for taskLength and length
        'taskLength is the value inputted in the Data Sheet
        'length is the value used in the prioritization algorithm
        If EditTask_UserForm2.optTime12.Value = True Then 'Checks which option button is filled
            length = 1
            taskLength = EditTask_UserForm2.optTime12.Caption
        ElseIf EditTask_UserForm2.optTime22.Value = True Then 'Checks which option button is filled
            length = 3
            taskLength = EditTask_UserForm2.optTime22.Caption
        ElseIf EditTask_UserForm2.optTime32.Value = True Then 'Checks which option button is filled
            length = 6
            taskLength = EditTask_UserForm2.optTime32.Caption
        ElseIf EditTask_UserForm2.optTime42.Value = True Then
            length = 10
            taskLength = EditTask_UserForm2.optTime42.Caption
        ElseIf EditTask_UserForm2.optTime52.Value = True Then
            length = 13
            taskLength = EditTask_UserForm2.optTime52.Caption
            
        End If
        
        'Determining the value for taskDifficulty and difficulty
        'taskDifficulty is the value inputted in the Data Sheet
        'difficulty is the value used in the prioritization algorithm
        If EditTask_UserForm2.optDifficulty12.Value = True Then 'Checks which option button is filled
            difficulty = 1
            taskDifficulty = EditTask_UserForm2.optDifficulty12.Caption
        ElseIf EditTask_UserForm2.optDifficulty22.Value = True Then
            difficulty = 2
            taskDifficulty = EditTask_UserForm2.optDifficulty22.Caption
        ElseIf EditTask_UserForm2.optDifficulty32.Value = True Then
            difficulty = 3
            taskDifficulty = EditTask_UserForm2.optDifficulty32.Caption
        ElseIf EditTask_UserForm2.optDifficulty42.Value = True Then
            difficulty = 4
            taskDifficulty = EditTask_UserForm2.optDifficulty42.Caption
        ElseIf EditTask_UserForm2.optDifficulty52.Value = True Then
            difficulty = 5
            taskDifficulty = EditTask_UserForm2.optDifficulty52.Caption
    
        End If
        
        'Determining the value for taskImportance and importance
        'taskImportance is the value inputted in the Data Sheet
        'importance is the value used in the prioritization algorithm
        If EditTask_UserForm2.optImportance12.Value = True Then 'Checks which option button is filled
            importance = 1
            taskImportance = EditTask_UserForm2.optImportance12.Caption
        ElseIf EditTask_UserForm2.optImportance22.Value = True Then
            importance = 2
            taskImportance = EditTask_UserForm2.optImportance22.Caption
        ElseIf EditTask_UserForm2.optImportance32.Value = True Then
            importance = 3
            taskImportance = EditTask_UserForm2.optImportance32.Caption
        ElseIf EditTask_UserForm2.optImportance42.Value = True Then
            importance = 4
            taskImportance = EditTask_UserForm2.optImportance42.Caption
        ElseIf EditTask_UserForm2.optImportance52.Value = True Then
            importance = 5
            taskImportance = EditTask_UserForm2.optImportance52.Caption
    
        End If
        
  
    effort = (0.5 * difficulty) + (0.5 * length)
    
    
    'Finding the difference between today's date and the deadline date inputted by the user
    'Dependent on whether the user inputs an additional urgent deadline or just inputs a regular deadline
    'The value for urgency is increased by 1 for tasks with urgent deadlines
    If EditTask_UserForm2.txtUrgentDeadline2.Value <> "" Then
        differenceDays = (DateDiff("d", Date, CDate(EditTask_UserForm2.txtUrgentDeadline2.Value)))
        
        If differenceDays > 0 Then
            urgency = 2 / differenceDays
        ElseIf differenceDays = 0 Then
            urgency = 2
        ElseIf differenceDays < 0 Then
            urgency = 2 + -(differenceDays)
        End If
        
    Else
        differenceDays = (DateDiff("d", Date, CDate(EditTask_UserForm2.txtDeadline2.Value)))
        
        
        'If the difference between days is greater than zero, urgency = 1/t where t represents the difference in days
        'If t = 0, urgency = 1
        'For every day that passes the due date, urgency = 1 + t
        
        If differenceDays > 0 Then
            urgency = 1 / differenceDays
        ElseIf differenceDays = 0 Then
            urgency = 1
        ElseIf differenceDays < 0 Then
            urgency = 1 + -(differenceDays)
        End If
        
    End If
    
   
   
   'If importance is greater than or equal to 3, the greater the value of effort, the greater the prioritization
   'If importance is less than 3, tasks that require more effort will be less prioritized than easy, less important tasks
   
    If importance >= 3 Then
        prioritization = (40 / 100) * importance + (36 / 100) * urgency + (24 / 100) * effort
        
    ElseIf importance < 3 Then
        prioritization = (40 / 100) * importance + (36 / 100) * urgency + (24 / 100) * 1 / effort
    End If
    
    
'Assigning the priority rankings beside each task in both WorkSheets
Dim overview As Worksheet, dataSheet As Worksheet
Set overview = ThisWorkbook.Sheets("TaskSheet")
Set dataSheet = ThisWorkbook.Sheets("Data Sheet")

Dim task_row As Integer, task_row2
Dim i As Integer, k As Integer 'Integers for Looping


task_row = overview.Cells(Rows.count, 2).End(xlUp).Row 'Identifies last row entry in column 2 on TaskSheet
task_row2 = dataSheet.Cells(Rows.count, 2).End(xlUp).Row 'Identifies last row entry in column 2 on Data Sheet


For i = 0 To EditTask_UserForm1.lstEditTask.ListCount - 1  'Loop through each assessment in Complete UserForm 1 list-box
    If EditTask_UserForm1.lstEditTask.Selected(i) = True Then
         
            With dataSheet:
                .Cells(i + 2, 1) = prioritization 'Inputs new priority value
            End With
            
            With overview:
                .Cells(i + 2, 1) = prioritization
                 .Cells(i + 2, "A").Select
                 Selection.Font.Color = vbWhite
            End With
        
        
        Exit For
        End If
    Next


End Sub

Sub prioritizationSorting()

Dim overview As Worksheet, dataSheet As Worksheet
Set overview = ThisWorkbook.Sheets("TaskSheet")
Set dataSheet = ThisWorkbook.Sheets("Data Sheet")

Dim iRow As Long
iRow = dataSheet.Cells(Rows.count, 1).End(xlUp).Row

For k = 2 To iRow
    dataSheet.Cells(k, 6).Value = k - 1
    overview.Cells(k, 1).Value = dataSheet.Cells(k, 6).Value
    overview.Cells(k, 1).Select
    Selection.Font.Color = vbBlack
Next


End Sub
