Private Sub CommandButton1_Click()

opt_3days.Value = False
opt_1week.Value = False
opt_2weeks.Value = False
opt_1month.Value = False
    
txt_startDate = ""
txt_endDate = ""

cmb_searchMember = ""

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

Private Sub txt_startDate_Exit(ByVal Cancel As MSForms.ReturnBoolean)
On Error Resume Next 'Allows code to continue runnning even if an error occurs

Me.txt_startDate = CDate(Me.txt_startDate) 'This converts the value of the entry in the textbox to the "Date" type

End Sub

Private Sub txt_endDate_Exit(ByVal Cancel As MSForms.ReturnBoolean)
On Error Resume Next

Me.txt_endDate = CDate(Me.txt_endDate)

End Sub

Private Sub btn_submitDate_Click()

'Convert the entries in the deadline column to the date format

Dim c As Long

For c = 2 To 1000

Sheet2.Cells(c, "A").NumberFormat = "yyyy-mm-dd;@"

Next c


'When the button is clicked, old values are deleted
Sheet2.Range("A2:H500").ClearContents
Application.ScreenUpdating = False

'If textboxes are left blank then display an error message
If txt_startDate = "" And txt_endDate = "" And opt_3days.Value = False And opt_1week.Value = False And opt_2weeks.Value = False And opt_1month.Value = False And cmb_searchMember.Value = "" Then
errorMsg = MsgBox("Please add a date range or select a member")
'Prevent the user from selecting an option button and entering a date range
ElseIf txt_startDate.Value <> "" And txt_endDate.Value <> "" _
    And opt_3days.Value = True Then
        MsgBox ("Multiple date ranges entered!")
ElseIf txt_startDate.Value <> "" And txt_endDate.Value <> "" _
    And opt_1week.Value = True Then
        MsgBox ("Multiple date ranges entered!")
ElseIf txt_startDate.Value <> "" And txt_endDate.Value <> "" _
    And opt_2weeks.Value = True Then
        MsgBox ("Multiple date ranges entered!")
ElseIf txt_startDate.Value <> "" And txt_endDate.Value <> "" _
    And opt_1month.Value = True Then
        MsgBox ("Multiple date ranges entered!")
ElseIf txt_startDate = "" And txt_endDate <> "" And opt_3days.Value = False And opt_1week.Value = False And opt_2weeks.Value = False And opt_1month.Value = False Then
    MsgBox ("Please add a start date to the range")
ElseIf txt_endDate = "" And txt_startDate <> "" And opt_3days.Value = False And opt_1week.Value = False And opt_2weeks.Value = False And opt_1month.Value = False Then
    MsgBox ("Please add an end date to the range")
'If dates are entered in an incorrect format then display an error message
ElseIf Not IsDate(txt_startDate.Text) And opt_3days.Value = False And opt_1week.Value = False And opt_2weeks.Value = False And opt_1month.Value = False And cmb_searchMember.Value = "" Then
    MsgBox "Please enter in the correct format"
ElseIf Not IsDate(txt_startDate.Text) And opt_3days.Value = False And opt_1week.Value = False And opt_2weeks.Value = False And opt_1month.Value = False And cmb_searchMember.Value = "" Then
    MsgBox "Please enter in the correct format"


'ScreenUpdating was added to make the macro run faster

Unload Me
End If


Call viewTasksFunction

Application.ScreenUpdating = True
End Sub

Private Sub UserForm_Initialize()

With viewTasksByDeadline_UserForm

    .opt_3days.Value = False
    .opt_1week.Value = False
    .opt_2weeks.Value = False
    .opt_1month.Value = False
    
End With

With cmb_searchMember

    .AddItem "Ghazal"
    .AddItem "Oriana"
    .AddItem "Tanushree"
    .AddItem "Toni"
    .AddItem "N/A"

End With


End Sub


Private Sub txt_startDate_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
txt_startDate.SpecialEffect = fmSpecialEffectEtched

End Sub

Private Sub txt_endDate_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
txt_endDate.SpecialEffect = fmSpecialEffectEtched

End Sub
Private Sub opt_3days_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
opt_3days.SpecialEffect = fmSpecialEffectFlat
opt_1week.SpecialEffect = fmButtonEffectSunken
opt_2weeks.SpecialEffect = fmButtonEffectSunken
opt_1month.SpecialEffect = fmButtonEffectSunken

End Sub

Private Sub opt_1week_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
opt_1week.SpecialEffect = fmSpecialEffectFlat
opt_3days.SpecialEffect = fmButtonEffectSunken
opt_2weeks.SpecialEffect = fmButtonEffectSunken
opt_1month.SpecialEffect = fmButtonEffectSunken

End Sub
Private Sub opt_2weeks_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
opt_2weeks.SpecialEffect = fmSpecialEffectFlat
opt_3days.SpecialEffect = fmButtonEffectSunken
opt_1month.SpecialEffect = fmButtonEffectSunken
opt_1week.SpecialEffect = fmButtonEffectSunken

End Sub
Private Sub opt_1month_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
opt_1month.SpecialEffect = fmSpecialEffectFlat
opt_3days.SpecialEffect = fmButtonEffectSunken
opt_1week.SpecialEffect = fmButtonEffectSunken
opt_2weeks.SpecialEffect = fmButtonEffectSunken

End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

txt_startDate.SpecialEffect = fmSpecialEffectFlat
txt_endDate.SpecialEffect = fmSpecialEffectFlat
cmb_searchMember.SpecialEffect = fmSpecialEffectFlat
opt_3days.SpecialEffect = fmButtonEffectSunken
opt_1week.SpecialEffect = fmButtonEffectSunken
opt_2weeks.SpecialEffect = fmButtonEffectSunken
opt_1month.SpecialEffect = fmButtonEffectSunken


End Sub


Private Sub Frame2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

txt_startDate.SpecialEffect = fmSpecialEffectFlat
txt_endDate.SpecialEffect = fmSpecialEffectFlat
opt_3days.SpecialEffect = fmButtonEffectSunken
opt_1week.SpecialEffect = fmButtonEffectSunken
opt_2weeks.SpecialEffect = fmButtonEffectSunken
opt_1month.SpecialEffect = fmButtonEffectSunken


End Sub
