
Sub InitializeEditTaskUserForm2()


Dim task_row As Integer
Dim i As Integer 'Used for Looping
Dim index As Integer

'Assign the worksheets being referrenced
Dim dataSheet As Worksheet, overview As Worksheet
Set dataSheet = ThisWorkbook.Sheets("Data Sheet")
Set overview = ThisWorkbook.Sheets("TaskSheet")

For i = 0 To EditTask_UserForm1.lstEditTask.ListCount - 1  'Loop through each assessment in Complete UserForm 1 list-box
    If EditTask_UserForm1.lstEditTask.Selected(i) = True Then
    
                With EditTask_UserForm2
                        .txtName2 = overview.Cells(i + 2, 2) 'Displays the original task name as the default in the input box
                        
                        'Add the categories onto the combo-box list for the Edit UserForm
                        .cmbCategory2.AddItem "Front End"
                        .cmbCategory2.AddItem "Back End"
                        .cmbCategory2.AddItem "Modelling"
                        .cmbCategory2.AddItem "Report Writing"
                        .cmbCategory2.AddItem "Presentation Design"
                        
                        .cmbCategory2.Value = overview.Cells(i + 2, 3).Value 'Displays the original selected task category as the default value in the combo box
                        
                        'Add the members onto the combo-box for the Edit UserForm
                        .cmbTeamMember2.AddItem "Ghazal"
                        .cmbTeamMember2.AddItem "Oriana"
                        .cmbTeamMember2.AddItem "Tanushree"
                        .cmbTeamMember2.AddItem "Toni"
                        .cmbTeamMember2.AddItem "N/A"
                        
                        .cmbTeamMember2.Value = overview.Cells(i + 2, 4).Value 'Displays the original assigned team member as the default value in the combo box
                    
                        .txtUrgentDeadline2.Value = overview.Cells(i + 2, 5).Value 'Displays the original urgent deadline on the UserForm
                        .txtDeadline2.Value = overview.Cells(i + 2, 6).Value 'Displays the original deadline on the UserForm
                
                       .optTime12.Value = False 'Resets all options as unfilled
                       .optTime22.Value = False
                       .optTime32.Value = False
                       .optTime42.Value = False
                       .optTime52.Value = False
                    
                       .optDifficulty12.Value = False 'Resets all options as unfilled
                       .optDifficulty22.Value = False
                       .optDifficulty32.Value = False
                       .optDifficulty42.Value = False
                       .optDifficulty52.Value = False
                    
                    
                       .optImportance12.Value = False 'Resets all options as unfilled
                       .optImportance22.Value = False
                       .optImportance32.Value = False
                       .optImportance42.Value = False
                       .optImportance52.Value = False
                       
                        If dataSheet.Cells(i + 2, 3).Value = .optTime12.Caption Then 'Checks which option button was filled previously
                            .optTime12.Value = True 'Sets the original filled radio button as the default filled in button
                        ElseIf dataSheet.Cells(i + 2, 3).Value = .optTime22.Caption Then
                            .optTime22.Value = True
                        ElseIf dataSheet.Cells(i + 2, 3).Value = .optTime32.Caption Then
                            .optTime32.Value = True
                        ElseIf dataSheet.Cells(i + 2, 3).Value = .optTime42.Caption Then
                            .optTime42.Value = True
                        ElseIf dataSheet.Cells(i + 2, 3).Value = .optTime52.Caption Then
                            .optTime52.Value = True
                        
                        End If
                        
                    
                        If dataSheet.Cells(i + 2, 4).Value = .optDifficulty12.Caption Then 'Checks which option button was filled previously
                            .optDifficulty12.Value = True 'Sets the original filled radio button as the default filled in button
                        ElseIf dataSheet.Cells(i + 2, 4).Value = .optDifficulty22.Caption Then
                            .optDifficulty22.Value = True
                        ElseIf dataSheet.Cells(i + 2, 4).Value = .optDifficulty32.Caption Then
                            .optDifficulty32.Value = True
                        ElseIf dataSheet.Cells(i + 2, 4).Value = .optDifficulty42.Caption Then
                            .optDifficulty42.Value = True
                        ElseIf dataSheet.Cells(i + 2, 4).Value = .optDifficulty52.Caption Then
                            .optDifficulty52.Value = True
                        
                        End If
                        
                    
                        If dataSheet.Cells(i + 2, 5).Value = .optImportance12.Caption Then 'Checks which option button was filled previously
                            .optImportance12.Value = True 'Sets the original filled radio button as the default filled in button
                        ElseIf dataSheet.Cells(i + 2, 5).Value = .optImportance22.Caption Then
                            .optImportance22.Value = True
                        ElseIf dataSheet.Cells(i + 2, 5).Value = .optImportance32.Caption Then
                            .optImportance32.Value = True
                        ElseIf dataSheet.Cells(i + 2, 5).Value = .optImportance42.Caption Then
                            .optImportance42.Value = True
                        ElseIf dataSheet.Cells(i + 2, 5).Value = .optImportance52.Caption Then
                            .optImportance52.Value = True
                        
                        End If
  
                    End With
           
                End If
        Next i
        

End Sub

Sub EditTask()

Dim i As Integer, k As Integer 'Used for Looping
Dim task_row As Integer, task_row2 As Integer

Dim overview As Worksheet, dataSheet As Worksheet
Set overview = ThisWorkbook.Sheets("TaskSheet")
Set dataSheet = ThisWorkbook.Sheets("Data Sheet")

task_row = overview.Cells(Rows.count, 2).End(xlUp).Row 'Identifies last row entry in column 2 on TaskSheet
task_row2 = dataSheet.Cells(Rows.count, 2).End(xlUp).Row 'Identifies last row entry in column 2 on Data Sheet

For i = 0 To EditTask_UserForm1.lstEditTask.ListCount - 1  'Loop through each assessment in Complete UserForm 1 list-box
    If EditTask_UserForm1.lstEditTask.Selected(i) = True Then
        
         With overview
            .Cells(i + 2, 2) = EditTask_UserForm2.txtName2.Value 'Overwrites cell value with new input data
            .Cells(i + 2, 3) = EditTask_UserForm2.cmbCategory2.Value 'Overwrites cell value with new input data
            .Cells(i + 2, 4) = EditTask_UserForm2.cmbTeamMember2.Value 'Overwrites cell value with new input data
            
            If EditTask_UserForm2.txtUrgentDeadline2.Value <> "" Then 'If a date is entered
            .Cells(i + 2, 5) = CDate(EditTask_UserForm2.txtUrgentDeadline2.Value) 'Overwrites cell value with new urgent deadline (converted into date format)
            End If
            .Cells(i + 2, 6) = CDate(EditTask_UserForm2.txtDeadline2.Value) 'Overwrites cell value with new deadline (converted into date format)
        
        
        End With
        
            With dataSheet
            .Cells(i + 2, 2) = EditTask_UserForm2.txtName2.Value
            
            If EditTask_UserForm2.optTime12.Value = True Then 'Checks which option button is filled
                .Cells(i + 2, 3) = EditTask_UserForm2.optTime12.Caption 'Overwrites cell value with new input data
            ElseIf EditTask_UserForm2.optTime22.Value = True Then
                .Cells(i + 2, 3) = EditTask_UserForm2.optTime22.Caption
            ElseIf EditTask_UserForm2.optTime32.Value = True Then
                .Cells(i + 2, 3) = EditTask_UserForm2.optTime32.Caption
            ElseIf EditTask_UserForm2.optTime42.Value = True Then
                .Cells(i + 2, 3) = EditTask_UserForm2.optTime42.Caption
            ElseIf EditTask_UserForm2.optTime52.Value = True Then
                .Cells(i + 2, 3) = EditTask_UserForm2.optTime52.Caption
                
            End If
            
            'Determining the value for task difficulty
            If EditTask_UserForm2.optDifficulty12.Value = True Then 'Checks which option button is filled
                .Cells(i + 2, 4) = EditTask_UserForm2.optDifficulty12.Caption 'Overwrites cell value with new input data
            ElseIf EditTask_UserForm2.optDifficulty22.Value = True Then
                .Cells(i + 2, 4) = EditTask_UserForm2.optDifficulty22.Caption
            ElseIf EditTask_UserForm2.optDifficulty32.Value = True Then
                .Cells(i + 2, 4) = EditTask_UserForm2.optDifficulty32.Caption
            ElseIf EditTask_UserForm2.optDifficulty42.Value = True Then
                .Cells(i + 2, 4) = EditTask_UserForm2.optDifficulty42.Caption
            ElseIf EditTask_UserForm2.optDifficulty52.Value = True Then
                .Cells(i + 2, 4) = EditTask_UserForm2.optDifficulty52.Caption
        
            End If
            
            'Determining the value for importance
            If EditTask_UserForm2.optImportance12.Value = True Then 'Checks which option button is filled
                .Cells(i + 2, 5) = EditTask_UserForm2.optImportance12.Caption 'Overwrites cell value with new input data
            ElseIf EditTask_UserForm2.optImportance22.Value = True Then
                .Cells(i + 2, 5) = EditTask_UserForm2.optImportance22.Caption
            ElseIf EditTask_UserForm2.optImportance32.Value = True Then
                .Cells(i + 2, 5) = EditTask_UserForm2.optImportance32.Caption
            ElseIf EditTask_UserForm2.optImportance42.Value = True Then
                .Cells(i + 2, 5) = EditTask_UserForm2.optImportance42.Caption
            ElseIf EditTask_UserForm2.optImportance52.Value = True Then
                .Cells(i + 2, 5) = EditTask_UserForm2.optImportance52.Caption
            End If
        
        
        End With
        
        Exit For
        
        End If
    
    Next i

End Sub


Sub Show_Edit_Update_TaskForm()

    EditTask_UserForm1.Show 'Opens the Edit Task User Form when user clicks on the Edit Task button

End Sub
