Private Sub cmdEdit2_Click()
 
 
 Dim msgValue As VbMsgBoxResult
 
 'Accounting for any empty textboxes that are necessary for our prioritization function
    
    If txtName2.Text = "" Then
        errorMsg = MsgBox("Please enter the Task Name")
    ElseIf txtDeadline2.Text = "" Then
        errorMsg = MsgBox("Please enter a Deadline")
    ElseIf Not IsDate(txtDeadline2.Text) Then
        MsgBox "Please enter deadline in the correct format"
     ElseIf Not IsDate(txtUrgentDeadline2.Text) And txtUrgentDeadline2.Text <> "" Then
        MsgBox "Please enter the additional deadline in the correct format"
    ElseIf txtUrgentDeadline2.Value >= txtDeadline2.Value Then
        MsgBox "Please enter an additional deadline that occurs before the actual deadline"

        
    Else
        msgValue = MsgBox("Do you want to edit this task?", vbYesNo + vbInformation, "Confirmation")
        
        Application.ScreenUpdating = False
        Call EditTask
        Call prioritization_Following_EditTask
        Call prioritization_ExistingTasks
        Call sort_TaskSheet
        Call sort_DataSheet
        Call prioritizationSorting
        Call HighlightTasks
        'EditTask_UserForm2.Hide
        Unload EditTask_UserForm1
        Unload EditTask_UserForm2
        
        Application.ScreenUpdating = True
        
    
        If msgValue = vbNo Then Exit Sub

    End If
    
End Sub



Private Sub UserForm_Initialize()

    Call InitializeEditTaskUserForm2


End Sub

Private Sub txtDeadline2_Exit(ByVal Cancel As MSForms.ReturnBoolean)
On Error Resume Next 'Allows code to continue runnning even if an error occurs

Me.txtDeadline2 = CDate(Me.txtDeadline2) 'This converts the value of the entry in the textbox to the "Date" type

End Sub

Private Sub txtUrgentDeadline2_Exit(ByVal Cancel As MSForms.ReturnBoolean)
On Error Resume Next 'Allows code to continue runnning even if an error occurs

Me.txtUrgentDeadline2 = CDate(Me.txtUrgentDeadline2) 'This converts the value of the entry in the textbox to the "Date" type

End Sub
