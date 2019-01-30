Attribute VB_Name = "modAlphaWin"
Option Explicit

Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Const GWL_EXSTYLE = (-20)
Private Const LWA_ALPHA = &H2
Private Const WS_EX_LAYERED = &H80000

Public Sub MakeTransparent(ByVal hWnd As Long, ByVal Alpha As Integer)
    
    Dim Msg As Long
    
    On Error Resume Next
    
    If (Alpha < 0) Or (Alpha > 255) Then Alpha = 255

    Msg = GetWindowLong(hWnd, GWL_EXSTYLE)
    Msg = Msg Or WS_EX_LAYERED
    SetWindowLong hWnd, GWL_EXSTYLE, Msg
    SetLayeredWindowAttributes hWnd, 0, Alpha, LWA_ALPHA
    
End Sub

Public Sub RevealMe(frm As Form)

    Dim i As Integer
    
    If Not IsWinXP Then Exit Sub
    
    MakeTransparent frm.hWnd, 0
    frm.Show
    DoEvents
    For i = 0 To 255 Step 2
        MakeTransparent frm.hWnd, i
    Next i
    frm.Hide

End Sub

Public Sub VanishMe(frm As Form)

    Dim i As Integer
    
    If Not IsWinXP Then Exit Sub

    For i = 255 To 0 Step -4
        DoEvents
        MakeTransparent frm.hWnd, i
    Next i

End Sub

