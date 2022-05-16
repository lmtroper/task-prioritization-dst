
'Initializes and resets the values for the AddTask_UserForm

Sub Reset_AddTask_UserForm()

Dim iRow As Long
    
    iRow = [Counta(TaskSheet!A:A)] ''Counts how many entries are in column A of TaskSheet
    
    
    'Clears the values of each element on AddTask_UserForm each time the user adds a new task
    With AddTask_UserForm:

        .txtName.Value = "" 'Clears the task name input box
        
        'Adds the task categories to the combo-box on the UserForm
        .cmbCategory.Clear
        .cmbCategory.AddItem "Front End"
        .cmbCategory.AddItem "Back End"
        .cmbCategory.AddItem "Modelling"
        .cmbCategory.AddItem "Report Writing"
        .cmbCategory.AddItem "Presentation Design"
        
        'Adds the team members to the combo-box on the UserForm
        .cmbTeamMember.Clear
        .cmbTeamMember.AddItem "Ghazal"
        .cmbTeamMember.AddItem "Oriana"
        .cmbTeamMember.AddItem "Tanushree"
        .cmbTeamMember.AddItem "Toni"
        .cmbTeamMember.AddItem "N/A"
        
        'Resets the value of the options buttons to unfilled
        .optTime1.Value = False
        .optTime2.Value = False
        .optTime3.Value = False
        .optTime4.Value = False
        .optTime5.Value = False
        
        'Resets the value of the options buttons to unfilled
        .optDifficulty1.Value = False
        .optDifficulty2.Value = False
        .optDifficulty3.Value = False
        .optDifficulty4.Value = False
        .optDifficulty5.Value = False
        
        'Resets the value of the options buttons to unfilled
        .optImportance1.Value = False
        .optImportance2.Value = False
        .optImportance3.Value = False
        .optImportance4.Value = False
        .optImportance5.Value = False
        
        .txtDeadline = "" 'Clears the deadline input box
        .txtUrgentDeadline.Value = "" 'Clears the urgent deadline input box
    
        
        
        'List Form within the AddTask_UserForm: shows all the tasks that have been added and updates right after you add a new task
        'Set the properties of the columns for the tasksheet to fit the inputted information
        .lstTaskSheet.ColumnCount = 6
        .lstTaskSheet.ColumnHeads = True 'Includes the column headings
        .lstTaskSheet.ColumnWidths = "70, 100, 90, 90, 90, 90"
        
            If iRow > 1 Then
                  
                .lstTaskSheet.RowSource = "TaskSheet!A2:F" & iRow ' iRow > 1 so add the information from the last row (task) inputted
                  
            Else
                .lstTaskSheet.RowSource = "TaskSheet!A2:F2" 'Add the row after the column headers
                  
        End If
              
          
    End With


End Sub


Sub Submit()

    'overview is the name of the sheet that will correspond with TaskSheet
    Dim overview As Worksheet
    Dim iRow As Long
    
    'Assign the worksheet variable to TaskSheet
    Set overview = ThisWorkbook.Sheets("TaskSheet")
    
    'Counts how many entries are in the column A of TaskSheet, iRow assigned to the row after the last entry(empty row)
    iRow = [Counta(TaskSheet!A:A)] + 1
    
    'Fill in information from AddTask_UserForm into TaskSheet into the empty row
    With overview
        
        .Cells(iRow, 2) = AddTask_UserForm.txtName.Value
        .Cells(iRow, 3) = AddTask_UserForm.cmbCategory.Value
        .Cells(iRow, 4) = AddTask_UserForm.cmbTeamMember.Value
        
        If AddTask_UserForm.txtUrgentDeadline.Value <> "" Then 'If the user chooses to input an urgent deadline
            .Cells(iRow, 5) = CDate(AddTask_UserForm.txtUrgentDeadline.Value) 'Converts user urgent deadline input into date format
        End If
        
        .Cells(iRow, 6) = CDate(AddTask_UserForm.txtDeadline.Value) 'Converts user deadline input into date format
         
    
    End With
    
    

End Sub


Sub sort_TaskSheet()

Sheet1.Range("A1").CurrentRegion.Sort Key1:=Range("A2"), Order1:=xlDescending, Header:=xlYes 'Sort TaskSheet by highest to lowest priority ranking value


End Sub


Sub sort_DataSheet()

'Sorts Data Sheet by highest to lowest prioritization ranking
'Record Macro Method

Range("A2").Activate
ActiveWorkbook.Worksheets("Data Sheet").Sort.SortFields.Clear
ActiveWorkbook.Worksheets("Data Sheet").Sort.SortFields.Add2 Key:=Range("A2"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Data Sheet").Sort 'Sort the
        .SetRange Range("A1:E6")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
        
    End With
    
    
End Sub


Sub Show_AddTask_Form()

    AddTask_UserForm.Show 'Opens the Add Task User Form when user clicks on the Add a Task button
    
End Sub
