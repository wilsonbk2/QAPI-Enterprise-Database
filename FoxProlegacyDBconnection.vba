Option Compare Database

'This code was used to connect to an older Legacy FoxPro database (If you know you know lol)

Sub FoxProLink(dbase_path As String, Optional oConn As ADODB.Connection, Optional rst As Recordset, Optional file_name As String)
'connect to a foxpro table based on the dbase_path information
 Dim file_path As String, file_ext As String
 ExtractDbaseInfo dbase_path, file_path, file_name, file_ext
 Set oConn = New ADODB.Connection
' conn.Attributes = 3
 oConn.Open "Driver={Microsoft Visual FoxPro Driver};" & _
            "SourceType=DBF;" & _
            "SourceDB=" & file_path & ";" & _
            "Exclusive=No;"
End Sub

Sub ExtractDbaseInfo(source As String, Optional FilePath As String, Optional FileName As String, Optional FileExt As String)

    Dim intCurrentPos As Integer, intFileNamePos As Integer
    Dim intFileExtPos As Integer
    
    'clear UDT passed in if there is any existing data
    FileName = ""
    FileExt = ""
    FilePath = ""

    If source <> "" Then
        'get file name by finding the last slash in the string
        Do
            intCurrentPos = InStr(intCurrentPos + 1, source, "\", vbTextCompare)
            If intCurrentPos <> 0 Then
                intFileNamePos = intCurrentPos
            Else
                intFileNamePos = intFileNamePos + 1
                Exit Do
            End If
        Loop
        'get file extension position if any
        intFileExtPos = InStr(intFileNamePos, source, ".", vbTextCompare) + 1
        If intFileExtPos = 1 Then
            Exit Sub
        End If
        'put data in UDT passed in
        If intFileNamePos <> 1 Then
            FileName = Mid(source, intFileNamePos, Len(source) - intFileNamePos - 3)
            FileExt = Mid(source, intFileExtPos, 3)
            FilePath = Mid(source, 1, intFileNamePos - 2)
        End If

    End If

End Sub



Sub foxprotry()

Set oConn = Server.CreateObject("ADODB.Connection") & _
ConnStr = "Driver=Microsoft Visual Foxpro Driver;" & _
"SourceType=DBC;" & _
"SourceDB=p:\dbs\mother\epicurr.dbf" & _
Conn.Open ConnStr


End Sub