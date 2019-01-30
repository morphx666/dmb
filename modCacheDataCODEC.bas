Attribute VB_Name = "modCacheDataCODEC"
Option Explicit

Public Function Str2HEX(ByVal s As String) As String

    Dim i As Long
    Dim r As String
    Dim d As String

    For i = 1 To Len(s)
        d = Hex(AscW(Mid(s, i, 1)))
        d = String(2 - Len(d), "0") + d
        r = r + d
    Next i

    Str2HEX = r

End Function

Public Function HEX2Str(ByVal s As String) As String

    Dim i As Long
    Dim r As String

    For i = 1 To Len(s) Step 2
        r = r + ChrW("&h" & Mid(s, i, 2))
    Next i

    HEX2Str = r

End Function
