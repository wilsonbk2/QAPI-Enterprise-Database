VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_TEST12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Private Sub Command3_Click()
'Dim dlg As FileDialog
Set dlg = Application.FileDialog(msoFileDialogFilePicker)
With dlg
.Title = "Select the Excel file to import"
.AllowMultiSelect = False
.Filters.Clear
.Filters.Add "Excel Files", "*.xlsx", 1
.Filters.Add "All Files", "*.*", 2
If .Show = -1 Then
strFileName = .SelectedItems(1)
'DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel8, "SECRES", strFileName, True
Else
Exit Sub
End If
End With
End Sub

Private Sub Command4_Click()
DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel8, "SECRES", strFileName, True
End Sub
