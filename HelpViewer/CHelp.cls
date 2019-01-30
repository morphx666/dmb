VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CHelp"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const HWND_TOP = 0
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2

Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const GWL_EXSTYLE = &HFFEC
Private Const WS_EX_APPWINDOW = &H40000

Private Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function SetActiveWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long

Private Sub Class_Initialize()

    StartMode = ActiveEXE
    Load frmHelpViewer
    
    Dim lCurStyle As Long
    Dim gwIndex As Long
    
    lCurStyle = GetWindowLong(frmHelpViewer.hwnd, GWL_EXSTYLE)
    SetWindowLong frmHelpViewer.hwnd, GWL_EXSTYLE, lCurStyle Or WS_EX_APPWINDOW

End Sub

Public Sub ShowHelp(CmdStr As String)

    On Error Resume Next
    
    frmHelpViewer.Visible = True
    frmHelpViewer.SetFocus

    Select Case CmdStr
        Case "%%SEARCH%%"
            frmHelpViewer.tsContents.Tabs("tsSearch").Selected = True
            frmHelpViewer.tsContents_Click
            frmHelpViewer.txtKeywords.SetFocus
        Case "%%CLOSE%%"
            Unload Me
        Case Else
            frmHelpViewer.ForceSelection CmdStr
    End Select
    
    ForceTopMost frmHelpViewer.hwnd
    
End Sub

Private Sub ForceTopMost(hwnd As Long)

    Dim i As Integer
    Dim h As Long
    
    For i = 1 To 10
        DoEvents
    
        SetActiveWindow hwnd
        SetForegroundWindow hwnd
        BringWindowToTop hwnd
        SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
        
        h = GetForegroundWindow
        If h = hwnd Then Exit For
    Next i
    
    SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE

End Sub

Private Sub Class_Terminate()

    ClassIsClosing = True
    
    Unload frmHelpViewer

End Sub
