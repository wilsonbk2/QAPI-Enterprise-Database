VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_TEST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Command18_Click()
MakeFolder.MakeFolder Me.Year1, Me.Month3, Me.DEPT
End Sub

Private Sub Command5_Click()
Dim dlg As FileDialog
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

Private Sub Command8_Click()

Dim SourceFile, DestinationFile
 
SourceFile = Me![strFileName1]     'e.g. "C:\Test\Test_File.doc"
DestinationFile = "C:\QAPI\"   ' Define target file name.
 
FileCopy SourceFile, DestinationFile


End Sub

Private Sub Command9_Click()
Dim dlg As FileDialog
Set dlg = Application.FileDialog(msoFileDialogFilePicker)
With dlg
.Title = "Select the Excel file to import"
.AllowMultiSelect = False
.Filters.Clear
.Filters.Add "Excel Files", "*.xlsx", 1
.Filters.Add "All Files", "*.*", 2
If .Show = -1 Then
strFileName1 = .SelectedItems(1)
'DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel8, "SECRES", strFileName, True
Else
Exit Sub
End If
End With
End Sub

Private Sub Form_Current()
Year1 = DatePart("yyyy", Now())
Month1 = Month(Date)
End Sub
