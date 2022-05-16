Sub viewTaskDeadlineForm_Click()

viewTasksByDeadline_UserForm.Show 'Opens the View Tasks By Date User Form when user clicks on the Tasks By Date button

End Sub
Sub viewTaskMemberForm_Click()

viewTasksByMember_UserForm.Show 'Opens the View Tasks by Member User Form when user clicks on the Tasks By Member button

End Sub

Sub ClearViewTasks()


Dim overview As Worksheet, viewTaskSheet As Worksheet
Set overview = ThisWorkbook.Sheets("TaskSheet")
Set viewTaskSheet = ThisWorkbook.Sheets("View Tasks")

Dim iRow As Long
iRow = [Counta(TaskSheet!A:A)]

For i = 2 To iRow
    Sheet2.Cells(i, 1) = overview.Cells(i, 6)
    Sheet2.Cells(i, 2) = overview.Cells(i, 1)
    Sheet2.Cells(i, 3) = overview.Cells(i, 2)
    Sheet2.Cells(i, 4) = overview.Cells(i, 3)
    Sheet2.Cells(i, 5) = overview.Cells(i, 4)
    Sheet2.Cells(i, 6) = overview.Cells(i, 5)
Next
    
    
End Sub

Sub viewTasksFunction()

Dim overview As Worksheet, viewTaskSheet
Set overview = ThisWorkbook.Sheets("TaskSheet")
Set viewTaskSheet = ThisWorkbook.Sheets("View Tasks")

Dim iRow As Long, iRow2 As Long
iRow = [Counta(TaskSheet!A:A)]

With Sheet2
    Rows("2:" & Rows.count).ClearContents
End With

k = 2

With viewTasksByDeadline_UserForm
If .cmb_searchMember <> "" Then
    For i = 2 To iRow
    If .txt_startDate <> "" And .txt_endDate <> "" Then
            If overview.Cells(i, 6) >= CDate(.txt_startDate) And overview.Cells(i, 6) <= CDate(.txt_endDate) And overview.Cells(i, 4) = .cmb_searchMember.Value Then
                
                viewTaskSheet.Cells(k, 1) = overview.Cells(i, 6)
                viewTaskSheet.Cells(k, 2) = overview.Cells(i, 1)
                viewTaskSheet.Cells(k, 3) = overview.Cells(i, 2)
                viewTaskSheet.Cells(k, 4) = overview.Cells(i, 3)
                viewTaskSheet.Cells(k, 5) = overview.Cells(i, 4)
                viewTaskSheet.Cells(k, 6) = overview.Cells(i, 5)
                
                k = k + 1
            End If
           
    ElseIf .opt_3days.Value = True Then
            If overview.Cells(i, 6) <= (Date + 3) And overview.Cells(i, 4).Value = .cmb_searchMember.Value Then
                
                viewTaskSheet.Cells(k, 1) = overview.Cells(i, 6)
                viewTaskSheet.Cells(k, 2) = overview.Cells(i, 1)
                viewTaskSheet.Cells(k, 3) = overview.Cells(i, 2)
                viewTaskSheet.Cells(k, 4) = overview.Cells(i, 3)
                viewTaskSheet.Cells(k, 5) = overview.Cells(i, 4)
                viewTaskSheet.Cells(k, 6) = overview.Cells(i, 5)
                
                k = k + 1
            End If
                    
    ElseIf .opt_1week.Value = True Then
            If overview.Cells(i, 6) <= Date + 7 And overview.Cells(i, 4) = .cmb_searchMember.Value Then
                
                viewTaskSheet.Cells(k, 1) = overview.Cells(i, 6)
                viewTaskSheet.Cells(k, 2) = overview.Cells(i, 1)
                viewTaskSheet.Cells(k, 3) = overview.Cells(i, 2)
                viewTaskSheet.Cells(k, 4) = overview.Cells(i, 3)
                viewTaskSheet.Cells(k, 5) = overview.Cells(i, 4)
                viewTaskSheet.Cells(k, 6) = overview.Cells(i, 5)
                
                k = k + 1
            End If
    
                    
    ElseIf .opt_2weeks.Value = True Then
            If overview.Cells(i, 6) <= Date + 14 And overview.Cells(i, 4).Text = .cmb_searchMember.Value Then
                
                viewTaskSheet.Cells(k, 1) = overview.Cells(i, 6)
                viewTaskSheet.Cells(k, 2) = overview.Cells(i, 1)
                viewTaskSheet.Cells(k, 3) = overview.Cells(i, 2)
                viewTaskSheet.Cells(k, 4) = overview.Cells(i, 3)
                viewTaskSheet.Cells(k, 5) = overview.Cells(i, 4)
                viewTaskSheet.Cells(k, 6) = overview.Cells(i, 5)
                
                k = k + 1
            End If
        
    ElseIf .opt_1month.Value = True Then
            If overview.Cells(i, 6) <= Date + 31 And overview.Cells(i, 4).Text = .cmb_searchMember.Value Then
                viewTaskSheet.Cells(k, 1) = overview.Cells(i, 6)
                viewTaskSheet.Cells(k, 2) = overview.Cells(i, 1)
                viewTaskSheet.Cells(k, 3) = overview.Cells(i, 2)
                viewTaskSheet.Cells(k, 4) = overview.Cells(i, 3)
                viewTaskSheet.Cells(k, 5) = overview.Cells(i, 4)
                viewTaskSheet.Cells(k, 6) = overview.Cells(i, 5)
                
                k = k + 1
            End If
    
    Else
            If overview.Cells(i, 4) = .cmb_searchMember.Value Then

                viewTaskSheet.Cells(k, 1) = overview.Cells(i, 6)
                viewTaskSheet.Cells(k, 2) = overview.Cells(i, 1)
                viewTaskSheet.Cells(k, 3) = overview.Cells(i, 2)
                viewTaskSheet.Cells(k, 4) = overview.Cells(i, 3)
                viewTaskSheet.Cells(k, 5) = overview.Cells(i, 4)
                viewTaskSheet.Cells(k, 6) = overview.Cells(i, 5)
                
                k = k + 1
            End If
    End If
Next
End If


If .cmb_searchMember = "" Then
    For i = 2 To iRow
    If .txt_startDate <> "" And .txt_endDate <> "" Then
            If overview.Cells(i, 6) >= CDate(.txt_startDate) And overview.Cells(i, 6) <= CDate(.txt_endDate) Then
                
                viewTaskSheet.Cells(k, 1) = overview.Cells(i, 6)
                viewTaskSheet.Cells(k, 2) = overview.Cells(i, 1)
                viewTaskSheet.Cells(k, 3) = overview.Cells(i, 2)
                viewTaskSheet.Cells(k, 4) = overview.Cells(i, 3)
                viewTaskSheet.Cells(k, 5) = overview.Cells(i, 4)
                viewTaskSheet.Cells(k, 6) = overview.Cells(i, 5)
                
                k = k + 1
            End If
        
    ElseIf .opt_3days.Value = True Then
            If overview.Cells(i, 6) <= Date + 3 Then
                
                viewTaskSheet.Cells(k, 1) = overview.Cells(i, 6)
                viewTaskSheet.Cells(k, 2) = overview.Cells(i, 1)
                viewTaskSheet.Cells(k, 3) = overview.Cells(i, 2)
                viewTaskSheet.Cells(k, 4) = overview.Cells(i, 3)
                viewTaskSheet.Cells(k, 5) = overview.Cells(i, 4)
                viewTaskSheet.Cells(k, 6) = overview.Cells(i, 5)
                
                k = k + 1
            End If
                    
    ElseIf .opt_1week.Value = True Then
            If overview.Cells(i, 6) <= Date + 7 Then
                viewTaskSheet.Cells(k, 1) = overview.Cells(i, 6)
                viewTaskSheet.Cells(k, 2) = overview.Cells(i, 1)
                viewTaskSheet.Cells(k, 3) = overview.Cells(i, 2)
                viewTaskSheet.Cells(k, 4) = overview.Cells(i, 3)
                viewTaskSheet.Cells(k, 5) = overview.Cells(i, 4)
                viewTaskSheet.Cells(k, 6) = overview.Cells(i, 5)
                
                k = k + 1
            End If
                    
    ElseIf .opt_2weeks.Value = True Then
            If overview.Cells(i, 6) <= Date + 14 Then
                viewTaskSheet.Cells(k, 1) = overview.Cells(i, 6)
                viewTaskSheet.Cells(k, 2) = overview.Cells(i, 1)
                viewTaskSheet.Cells(k, 3) = overview.Cells(i, 2)
                viewTaskSheet.Cells(k, 4) = overview.Cells(i, 3)
                viewTaskSheet.Cells(k, 5) = overview.Cells(i, 4)
                viewTaskSheet.Cells(k, 6) = overview.Cells(i, 5)
                
                k = k + 1
            End If
        
    ElseIf .opt_1month.Value = True Then
            If overview.Cells(i, 6) <= Date + 31 Then
                
                viewTaskSheet.Cells(k, 1) = overview.Cells(i, 6)
                viewTaskSheet.Cells(k, 2) = overview.Cells(i, 1)
                viewTaskSheet.Cells(k, 3) = overview.Cells(i, 2)
                viewTaskSheet.Cells(k, 4) = overview.Cells(i, 3)
                viewTaskSheet.Cells(k, 5) = overview.Cells(i, 4)
                viewTaskSheet.Cells(k, 6) = overview.Cells(i, 5)
                
                k = k + 1
        
            End If
    End If
Next
End If
End With

End Sub
