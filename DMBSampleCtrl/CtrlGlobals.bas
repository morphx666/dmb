Attribute VB_Name = "CtrlGlobals"
Option Explicit

Global frmMain As Form

Public Const VK_LBUTTON = &H1
Public Const VK_RBUTTON = &H2

Public Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Public Function CreateUrlFromPath(p As String) As String

End Function

Public Function EncodeUrl(u As String) As String

End Function
