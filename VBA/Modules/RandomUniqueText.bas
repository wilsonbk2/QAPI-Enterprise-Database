Attribute VB_Name = "RandomUniqueText"
Option Compare Database

Public Function UniqueTimeBasedString() As String
    Dim alphanumericCharacters As String
    alphanumericCharacters = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    Dim timeString As String
    timeString = Format(Now, "yymmddhhmmss") & Getmtimestring(3)
    Dim c As Integer
    Dim intTimepart As Integer
    c = 1
    Do While c < Len(timeString)
        intTimepart = CInt(Mid(timeString, c, 3))
        c = c + 3
        UniqueTimeBasedString = UniqueTimeBasedString & Mid(alphanumericCharacters, Int(intTimepart / 61) + 1, 1) & Mid(alphanumericCharacters, intTimepart Mod 61 + 1, 1)
    Loop
End Function

Public Function Getmtimestring(length As Integer) As String
    Dim mtime As Double
    mtime = Timer()
    Dim mtimeLng As Long
    mtimeLng = Int((mtime - Int(mtime)) * (10 ^ length))
    Getmtimestring = CStr(mtimeLng)
    Do While Len(Getmtimestring) < length
        Getmtimestring = "0" & Getmtimestring
    Loop
End Function
