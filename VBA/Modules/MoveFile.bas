Attribute VB_Name = "MoveFile"
Option Compare Database

Public Sub moveCurrent(file$, pathfrom$, pathto$)
    Dim fso As Object
    Dim SourceFile As String
    Dim targetFile As String
    Dim answer As Integer

    Set fso = CreateObject("Scripting.FileSystemObject")
    SourceFile = pathfrom$ & "\" & file$
    targetFile = pathto$ & "\" & file$
      
    If fso.FileExists(targetFile) Then
        answer = MsgBox("File already exists in this location. " _
            & "Are you sure you want to continue? If you continue " _
            & "the file at destination will be deleted!", _
            vbInformation + vbYesNo)
        If answer = vbNo Then
            Exit Sub
        End If
        Kill targetFile
    End If
    fso.MoveFile SourceFile, targetFile
    Set fso = Nothing
End Sub

