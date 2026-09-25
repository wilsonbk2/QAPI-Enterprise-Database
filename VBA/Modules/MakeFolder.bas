Attribute VB_Name = "MakeFolder"
Option Explicit
Option Compare Text


#If Win64 Then
Public Declare PtrSafe Function apiCreatePath Lib "Imagehlp.dll" Alias "MakeSureDirectoryPathExists" (ByVal strPath As String) As Long
#Else
Public Declare Function apiCreatePath Lib "Imagehlp.dll" Alias "MakeSureDirectoryPathExists" (ByVal strPath As String) As Long
#End If


Sub Test()

    MakeFolder "Mike Tompson"
    MakeFolder "Mike Tompson", "12345678-2"
    MakeFolder "Mike Tompson", "12345678"
    MakeFolder "Mike A. Tompson"
    MakeFolder "Mike A. Tompson", "1"
    MakeFolder "Mike A. Tompson", "2"


End Sub


Public Sub MakeFolder(ByVal vntName As Variant, Optional ByVal vntAcc As Variant = Null, Optional ByVal vntDept As Variant = Null)

    Dim strPath As String
    
    Const conRoot As String = "\\esh15\QAPI Committee\QAPI Data\QAPI Backup Documents\"
    
    If Len(vntName) Then
        strPath = conRoot & Trim(vntName) & "\"
        
        If Len(vntAcc) Then
            strPath = strPath & Trim(vntAcc) & "\"
        End If
        
        If Len(vntDept) Then
            strPath = strPath & Trim(vntDept) & "\"
        End If
        
        apiCreatePath strPath
    End If

End Sub

