

Private Sub UserForm_Initialize()
    
    'Whenever the UserForm is opened, all data is reset and values are re-added
    Call Reset_AddTask_UserForm

    
End Sub

Private Sub cmdAdd_Click()
    
    Dim msgValue As VbMsgBoxResult
    
    'Accounting for any empty textboxes that are necessary for our prioritization function
    
    If txtName.Text = "" Then
        errorMsg = MsgBox("Please enter the Task Name")
    ElseIf txtDeadline.Text = "" Then
        errorMsg = MsgBox("Please enter a Deadline")
    ElseIf Not IsDate(txtDeadline.Text) Then
        MsgBox "Please enter deadline in the correct format"
     ElseIf Not IsDate(txtUrgentDeadline.Text) And txtUrgentDeadline.Text <> "" Then
        MsgBox "Please enter the additional deadline in the correct format"
    ElseIf txtUrgentDeadline.Value >= txtDeadline.Value Then
        MsgBox "Please enter an additional deadline that occurs before the actual deadline"
    ElseIf optImportance1.Value <> True And optImportance2.Value <> True And optImportance3.Value <> True And optImportance4.Value <> True And optImportance5.Value <> True Then
        errorMsg = MsgBox("Please select an option describing how critical the task is to your project")
    ElseIf optDifficulty1.Value <> True And optDifficulty2.Value <> True And optDifficulty3.Value <> True And optDifficulty4.Value <> True And optDifficulty5.Value <> True Then
        errorMsg = MsgBox("Please select an option describing the task's difficulty")
    ElseIf optTime1.Value <> True And optTime2.Value <> True And optTime3.Value <> True And optTime4.Value <> True And optTime5.Value <> True Then
        errorMsg = MsgBox("Please select an option estimating how long it will take to complete the task")
    
    'When all the conditions above are met, we will submit the data and call the prioritization function
    Else
         msgValue = MsgBox("Do you want to add this task?", vbYesNo + vbInformation, "Confirmation")
    
        If msgValue = vbNo Then Exit Sub
        
        Application.ScreenUpdating = False
        Call Submit
        Call prioritization_NewTask
        Call prioritization_ExistingTasks
        Call sort_TaskSheet
        Call sort_DataSheet
        Call prioritizationSorting
        Call HighlightTasks
        Call Reset_AddTask_UserForm

         Application.ScreenUpdating = True

        
        
    End If
    
     
End Sub

Private Sub cmdReset_Click()

    Dim msgValue As VbMsgBoxResult
    msgValue = MsgBox("Do you want to reset this form?", vbYesNo + vbInformation, "Confirmation")
    
    If msgValue = vbNo Then Exit Sub
    
    Call Reset_AddTask_UserForm

End Sub

Private Sub txtDeadline_Exit(ByVal Cancel As MSForms.ReturnBoolean)
On Error Resume Next 'Allows code to continue runnning even if an error occurs

Me.txtDeadline = CDate(Me.txtDeadline) 'This converts the value of the entry in the textbox to the "Date" type

End Sub

Private Sub txtUrgentDeadline_Exit(ByVal Cancel As MSForms.ReturnBoolean)
On Error Resume Next 'Allows code to continue runnning even if an error occurs

Me.txtUrgentDeadline = CDate(Me.txtUrgentDeadline) 'This converts the value of the entry in the textbox to the "Date" type

End Sub


Private Sub optTime1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
optTime1.SpecialEffect = fmButtonEffectFlat
optTime2.SpecialEffect = fmButtonEffectSunken
optTime3.SpecialEffect = fmButtonEffectSunken
optTime4.SpecialEffect = fmButtonEffectSunken
optTime5.SpecialEffect = fmButtonEffectSunken

End Sub
Private Sub optTime2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
optTime2.SpecialEffect = fmButtonEffectFlat
optTime1.SpecialEffect = fmButtonEffectSunken
optTime3.SpecialEffect = fmButtonEffectSunken
optTime4.SpecialEffect = fmButtonEffectSunken
optTime5.SpecialEffect = fmButtonEffectSunken

End Sub
Private Sub optTime3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
optTime3.SpecialEffect = fmButtonEffectFlat
optTime1.SpecialEffect = fmButtonEffectSunken
optTime2.SpecialEffect = fmButtonEffectSunken
optTime4.SpecialEffect = fmButtonEffectSunken
optTime5.SpecialEffect = fmButtonEffectSunken

End Sub
Private Sub optTime4_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
optTime4.SpecialEffect = fmButtonEffectFlat
optTime1.SpecialEffect = fmButtonEffectSunken
optTime2.SpecialEffect = fmButtonEffectSunken
optTime3.SpecialEffect = fmButtonEffectSunken
optTime5.SpecialEffect = fmButtonEffectSunken

End Sub
Private Sub optTime5_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
optTime5.SpecialEffect = fmButtonEffectFlat
optTime1.SpecialEffect = fmButtonEffectSunken
optTime2.SpecialEffect = fmButtonEffectSunken
optTime3.SpecialEffect = fmButtonEffectSunken
optTime4.SpecialEffect = fmButtonEffectSunken

End Sub
Private Sub optDifficulty1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
optDifficulty1.SpecialEffect = fmButtonEffectFlat
optDifficulty2.SpecialEffect = fmButtonEffectSunken
optDifficulty3.SpecialEffect = fmButtonEffectSunken
optDifficulty4.SpecialEffect = fmButtonEffectSunken
optDifficulty5.SpecialEffect = fmButtonEffectSunken

End Sub
Private Sub optDifficulty2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
optDifficulty2.SpecialEffect = fmButtonEffectFlat
optDifficulty1.SpecialEffect = fmButtonEffectSunken
optDifficulty3.SpecialEffect = fmButtonEffectSunken
optDifficulty4.SpecialEffect = fmButtonEffectSunken
optDifficulty5.SpecialEffect = fmButtonEffectSunken

End Sub
Private Sub optDifficulty3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
optDifficulty3.SpecialEffect = fmButtonEffectFlat
optDifficulty1.SpecialEffect = fmButtonEffectSunken
optDifficulty2.SpecialEffect = fmButtonEffectSunken
optDifficulty4.SpecialEffect = fmButtonEffectSunken
optDifficulty5.SpecialEffect = fmButtonEffectSunken

End Sub
Private Sub optDifficulty4_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
optDifficulty4.SpecialEffect = fmButtonEffectFlat
optDifficulty1.SpecialEffect = fmButtonEffectSunken
optDifficulty2.SpecialEffect = fmButtonEffectSunken
optDifficulty3.SpecialEffect = fmButtonEffectSunken
optDifficulty5.SpecialEffect = fmButtonEffectSunken

End Sub
Private Sub optDifficulty5_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
optDifficulty5.SpecialEffect = fmButtonEffectFlat
optDifficulty1.SpecialEffect = fmButtonEffectSunken
optDifficulty2.SpecialEffect = fmButtonEffectSunken
optDifficulty3.SpecialEffect = fmButtonEffectSunken
optDifficulty4.SpecialEffect = fmButtonEffectSunken


End Sub
Private Sub optImportance1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
optImportance1.SpecialEffect = fmButtonEffectFlat
optImportance2.SpecialEffect = fmButtonEffectSunken
optImportance3.SpecialEffect = fmButtonEffectSunken
optImportance4.SpecialEffect = fmButtonEffectSunken
optImportance5.SpecialEffect = fmButtonEffectSunken

End Sub
Private Sub optImportance2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
optImportance2.SpecialEffect = fmButtonEffectFlat
optImportance1.SpecialEffect = fmButtonEffectSunken
optImportance3.SpecialEffect = fmButtonEffectSunken
optImportance4.SpecialEffect = fmButtonEffectSunken
optImportance5.SpecialEffect = fmButtonEffectSunken

End Sub
Private Sub optImportance3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
optImportance3.SpecialEffect = fmButtonEffectFlat
optImportance1.SpecialEffect = fmButtonEffectSunken
optImportance2.SpecialEffect = fmButtonEffectSunken
optImportance4.SpecialEffect = fmButtonEffectSunken
optImportance5.SpecialEffect = fmButtonEffectSunken

End Sub
Private Sub optImportance4_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
optImportance4.SpecialEffect = fmButtonEffectFlat
optImportance1.SpecialEffect = fmButtonEffectSunken
optImportance2.SpecialEffect = fmButtonEffectSunken
optImportance3.SpecialEffect = fmButtonEffectSunken
optImportance5.SpecialEffect = fmButtonEffectSunken

End Sub
Private Sub optImportance5_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
optImportance5.SpecialEffect = fmButtonEffectFlat
optImportance1.SpecialEffect = fmButtonEffectSunken
optImportance2.SpecialEffect = fmButtonEffectSunken
optImportance3.SpecialEffect = fmButtonEffectSunken
optImportance4.SpecialEffect = fmButtonEffectSunken

End Sub


Private Sub Frame1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)


optTime1.SpecialEffect = fmButtonEffectSunken
optTime2.SpecialEffect = fmButtonEffectSunken
optTime3.SpecialEffect = fmButtonEffectSunken
optTime4.SpecialEffect = fmButtonEffectSunken
optTime5.SpecialEffect = fmButtonEffectSunken

optDifficulty1.SpecialEffect = fmButtonEffectSunken
optDifficulty2.SpecialEffect = fmButtonEffectSunken
optDifficulty3.SpecialEffect = fmButtonEffectSunken
optDifficulty4.SpecialEffect = fmButtonEffectSunken
optDifficulty5.SpecialEffect = fmButtonEffectSunken

optImportance1.SpecialEffect = fmButtonEffectSunken
optImportance2.SpecialEffect = fmButtonEffectSunken
optImportance3.SpecialEffect = fmButtonEffectSunken
optImportance4.SpecialEffect = fmButtonEffectSunken
optImportance5.SpecialEffect = fmButtonEffectSunken


End Sub
