VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_EmployeeRepositoryFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cboFilter_Change()
    Dim strFilter As String
    Dim sSearch As String
    On Error Resume Next

If Me.cboFilter.Text <> "" Then
    sSearch = "'*" & Replace(Me.cboFilter.Text, "'", "''") & "*'"
    strFilter = "[FIRSTNAME] Like " & sSearch & " OR [LASTNAME] Like " & sSearch & " OR [CC-DESC] Like " & sSearch & " OR [JOBTITLE] Like " & sSearch & " OR [POSTYP DESC] Like " & sSearch & " OR [EMPSTATUS] Like " & sSearch
    Me.Filter = strFilter
    Me.FilterOn = True
Else
    Me.Filter = ""
    Me.FilterOn = False
End If

If Me.Recordset.RecordCount = 0 Then 'new line of code
    Me.Filter = "" 'new line of code
    Me.FilterOn = False 'new line of code
    Me.cboFilter.SetFocus 'new line of code
    Me.cboFilter.Text = "" 'new line of code
Exit Sub 'new line of code
End If 'new line of code

With Me.cboFilter
    .SetFocus
    .SelStart = Len(Me.cboFilter.Text)
End With
End Sub





'This code is what I have that will reset the textbox to blank and requery, giving me all the people in the DB
'Private Sub txtSearch_Click()
'    Me.txtSearch.Text = ""
'    Me.Requery
'    With Me.txtSearch
'        .SetFocus
'        .SelStart = Len(Me.txtSearch.Text)
'    End With
'End Sub

'If Nz(Me.cboFilter.Text) = "" Then
'    Me.Form.Filter = ""
'    Me.FilterOn = False
'
'ElseIf Me.cboFilter.ListIndex <> -1 Then
'  Me.Form.Filter = "[FIRSTNAME] ='" & Me.cboFilter.Text & "'"
'  Me.FilterOn = True
'ElseIf Me.cboFilter.ListIndex <> -1 Then
'  Me.Form.Filter = "[LASTNAME] ='" & Me.cboFilter.Text & "'"
'  Me.FilterOn = True
'ElseIf Me.cboFilter.ListIndex <> -1 Then
'   Me.Form.Filter = "[FIRSTNAME] = '" & Replace(Me.cboFilter.Text, "'", """")
'   Me.FilterOn = True
'ElseIf Me.cboFilter.ListIndex <> -1 Then
'   Me.Form.Filter = "[LASTNAME] = '" & Replace(Me.cboFilter.Text, "'", """")
'   Me.FilterOn = True
'ElseIf Me.cboFilter.ListIndex <> -1 Then
'    Me.Form.Filter = "[FIRSTNAME] Like '*" & Me.cboFilter.Text & "*'"
'    Me.FilterOn = True
'Else
'    Me.Form.Filter = "[LASTNAME] Like '*" & Me.cboFilter.Text & "*'"
'    Me.FilterOn = True
'End If
'Me.cboFilter.SetFocus
'Me.cboFilter.SelStart = Len(Me.cboFilter.Text)

Private Sub EMP_ID_Click()
DoCmd.GoToRecord acDataForm, "EmployeeRepository", acGoTo, [EMP ID]
End Sub



Private Sub Command40_Click()
DoCmd.OpenForm "EmployeeRepository", acNormal, "", "", , acNormal
DoCmd.GoToControl "EMPID2"
DoCmd.FindRecord Forms!EmployeeRepositoryList!EmployeeRepositoryFind.Form!EMPID
DoCmd.Close acForm, "EmployeeRepositoryList"
'Forms![EmployeeRepositoryFind]![EMPID]
End Sub

Private Sub Command41_Click()
DoCmd.OutputTo acOutputForm, Me.Name, acFormatXLS, "C:\Users\" & Me!Alias & "\DesktopEmployeeList.xls", True
End Sub

Private Sub Form_Current()
Me.Alias = Environ("USERNAME")
End Sub
