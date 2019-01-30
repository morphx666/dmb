Attribute VB_Name = "modGlobals"
Option Explicit

Private Type CREATESTRUCT
    lpCreateParams As Long
    hInstance As Long
    hMenu As Long
    hWndParent As Long
    cy As Long
    cx As Long
    y As Long
    x As Long
    style As Long
    lpszName As Long    ' String
    lpszClass As Long   ' String
    ExStyle As Long
End Type

Public Const WM_CREATE = &H1
Public Const GWL_EXSTYLE = (-20)
Public Const GWL_WNDPROC = (-4)
Public Const WS_EX_APPWINDOW = &H40000
Public Const WS_EX_TOOLWINDOW = &H80&

Public m_lHookWndProc As Long

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Function Form_WndProc(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim lSetStyleEX As Long
    ' SPM - specific wnd proc for a form.  Only called once for the WM_CREATE message.
    Select Case Msg
        Case WM_CREATE
            Dim tCS As CREATESTRUCT
            
            CopyMemory tCS, ByVal lParam, Len(tCS)
            lSetStyleEX = GetWindowLong(hwnd, GWL_EXSTYLE)
            lSetStyleEX = lSetStyleEX Or WS_EX_APPWINDOW
            lSetStyleEX = lSetStyleEX And (Not WS_EX_TOOLWINDOW)
            tCS.ExStyle = lSetStyleEX
            CopyMemory ByVal lParam, tCS, Len(tCS)
            SetWindowLong hwnd, GWL_WNDPROC, m_lHookWndProc
            SetWindowLong hwnd, GWL_EXSTYLE, tCS.ExStyle
    End Select
    
    Form_WndProc = CallWindowProc(m_lHookWndProc, hwnd, Msg, wParam, lParam)
End Function
