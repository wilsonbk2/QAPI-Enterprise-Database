VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_ViewDataCODESfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub BUI_AfterUpdate()
Call FilterSubform
MU.Requery
Me.MU = Null
SU.Requery
Me.SU = Null
End Sub

Private Sub CD_AfterUpdate()
Call FilterSubform
End Sub

Private Sub Command16_Click()
Me.ViewDataCODESds.Form.FilterOn = False
Me.CD = Null
Me.BUI = Null
Me.MU = Null
Me.SU = Null
End Sub

Private Sub MU_AfterUpdate()
Call FilterSubform
SU.Requery
End Sub

Private Sub SU_AfterUpdate()
Call FilterSubform
End Sub



Private Sub FilterSubform()
    Dim strWhere As String

    'Make string
    If Nz(Me.CD, "") <> "" Then
        strWhere = strWhere & "[CNAME] = '" & Me.CD & "' AND "
    End If

    If Nz(Me.BUI, "") <> "" Then
        strWhere = strWhere & "[BUILDING#] = '" & Me.BUI & "' AND "
    End If

    If Nz(Me.MU, "") <> "" Then
        strWhere = strWhere & "[UNITMAIN] LIKE '*" & Replace(Me.MU, "'", "''") & "*' AND "
    End If
    
    If Nz(Me.SU, "") <> "" Then
        strWhere = strWhere & "[UNITSUB] LIKE '*" & Replace(Me.SU, "'", "''") & "*' AND "
    End If

    'Apply filter
    If strWhere <> "" Then
        strWhere = Left(strWhere, Len(strWhere) - 5) 'Remove the extra AND
        Me.ViewDataCODESds.Form.Filter = strWhere
        Me.ViewDataCODESds.Form.FilterOn = True
    Else
        Me.ViewDataCODESds.Form.Filter = ""
        Me.ViewDataCODESds.Form.FilterOn = False
    End If
End Sub


