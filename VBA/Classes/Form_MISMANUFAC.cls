VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_MISMANUFAC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub AddManu_Click()
Me.cboSubmit.Visible = True
Me.SubmitManu.Visible = True
Me.cboFilter.Visible = False
Me.cboSubmit.SetFocus
Me.AddManu.Visible = False
End Sub

Private Sub cboFilter_Change()
    Dim strFilter As String
    Dim sSearch As String
    On Error Resume Next

If Me.cboFilter.Text <> "" Then
    sSearch = "'*" & Replace(Me.cboFilter.Text, "'", "''") & "*'"
    strFilter = "[MANUID] Like " & sSearch & " OR [MANUDESC] Like " & sSearch
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


Private Sub SubmitManu_Click()


If IsNull(DLookup("MANUDESC", "MISMANUFAC", "[MANUDESC] = '" & [cboSubmit] & "'")) = False Then

MsgBox "You have already submitted for this record. (If you have submitted incorrect data please contact Brad Wilson at (757)-208-7361)", vbCritical

Else
    
    If IsNull(Me.cboSubmit) Then
     
    MsgBox "You must fill out required feilds.", vbCritical, "Data entry error... All Fields Required"
    Me.cboSubmit.BorderColor = vbRed
    GoTo TheEnd
    End If
    
    CurrentDb.Execute "Insert Into MISMANUFAC (MANUID, MANUDESC) " & "Values ('" & Me.recordid & "', '" & Me.cboSubmit & "')"
    
    Me.cboFilter.Visible = True
    Me.cboFilter.SetFocus
    Me.AddManu.Visible = True
    Me.SubmitManu.Visible = False
    Me.cboSubmit.Visible = False


    MsgBox "Thank you, this manufacturer has been added."
    Me.cboSubmit = Null
    Me.cboSubmit.BorderColor = RGB(192, 80, 77)
    
    
        
    
End If
TheEnd:

End Sub
