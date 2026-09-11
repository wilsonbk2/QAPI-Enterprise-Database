Option Compare Database

'This code closes all possible instances of forms within a VBA instance and keeps only the Main Menu form open, so the user can return to the home page.

Public Sub CloseAllForms()

'DoCmd.Close acForm, "Admin Menu"
Dim frms As AllForms
Dim lngX As Long

Set frms = CurrentProject.AllForms
For lngX = 0 To frms.Count - 1
If frms(lngX).Name <> "Splash Screen" Then
If frms(lngX).Name <> "frmLogin" Then
If frms(lngX).Name <> "frmLogoutTimer" Then
If frms(lngX).Name <> "frmLogoutStatus" Then
If frms(lngX).Name <> "frmDemo" Then
If frms(lngX).IsLoaded Then
DoCmd.Close acForm, frms(lngX).Name, acSaveNo
End If
End If
End If
End If
End If
End If
Next lngX
Set frms = Nothing

End Sub
