Attribute VB_Name = "Binary"
Option Explicit

Public Function Dec2Bin(d As Long) As String

    Dim i As Long
    Dim r As String
    
    i = 1
    Do While i <= d
        If Int(d / i) Mod 2 = 0 Then
            r = "0" + r
        Else
            r = "1" + r
        End If
        i = i * 2
    Loop
    
    Dec2Bin = r

End Function

Public Function Bin2Dec(b As String) As Double

    Dim i As Integer
    Dim r As Double
    Dim s As Integer
    
    s = Len(b)
    
    For i = 1 To s
        r = r + 2 ^ (i - 1) * Val(Mid$(b, s - i + 1, 1))
    Next i
    
    Bin2Dec = r

End Function

Public Function BinOR(b1 As String, b2 As String) As String

    BinOR = Dec2Bin(Bin2Dec(b1) Or Bin2Dec(b2))

End Function

Public Function BinAND(b1 As String, b2 As String) As String

    BinAND = Dec2Bin(Bin2Dec(b1) And Bin2Dec(b2))

End Function

Public Function BinXOR(b1 As String, b2 As String) As String

    BinXOR = Dec2Bin(Bin2Dec(b1) Xor Bin2Dec(b2))

End Function

Public Function BinNOT(b As String) As String

    Dim i As Integer
    Dim r As String
    Dim s As Integer
    
    s = Len(b)
    For i = 1 To s
        r = r + CStr(Abs(Not Val(Mid$(b, s - i + 1, 1)) = 1))
    Next i
    
    BinNOT = r

End Function
