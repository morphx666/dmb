Attribute VB_Name = "HVGlobals"
Option Explicit

Public Enum StartModeConstants
    Standalone = 0
    ActiveEXE = 1
    RegEXE = 2
End Enum

Global StartMode As StartModeConstants
Global ClassIsClosing As Boolean

Private Declare Sub InitCommonControls Lib "Comctl32.dll" ()
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Const LVM_FIRST = &H1000
Public Const LVM_SETEXTENDEDLISTVIEWSTYLE = LVM_FIRST + 54
Public Const LVM_GETEXTENDEDLISTVIEWSTYLE = LVM_FIRST + 55

Public Const LVS_EX_FULLROWSELECT = &H20
Public Const LVSCW_AUTOSIZE = -1
Public Const LVSCW_AUTOSIZE_USEHEADER = -2
Public Const LVM_SETCOLUMNWIDTH As Long = &H101E
Public Const HDS_BUTTONS = &H2
Public Const LVM_GETHEADER = (LVM_FIRST + 31)
Public Const GWL_STYLE = (-16)

Public Sub Main()

    If FileExists(App.Path + "\dmbhelp.exe.manifest") Then InitCommonControls

    If LCase(Command$) = "/exeselfregister" Then
        StartMode = RegEXE
        End
    End If
    
    StartMode = Standalone
    frmHelpViewer.Show vbModal

End Sub

Public Sub CoolListView(lv As ListView)

    Dim rStyle As Long
    Dim i As Long
    
    With lv
        If .View = lvwReport Then
            For i = 0 To .ColumnHeaders.Count - 1
                SendMessageLong .hWnd, LVM_SETCOLUMNWIDTH, i, ByVal LVSCW_AUTOSIZE_USEHEADER
            Next i
        End If
        
        rStyle = SendMessageLong(.hWnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0&, 0&)
        rStyle = rStyle Or LVS_EX_FULLROWSELECT
        SendMessageLong .hWnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0&, rStyle
    End With

End Sub

Public Sub CenterForm(Frm As Form)

    With Screen
        Frm.Left = (.Width - Frm.Width) / 2
        Frm.Top = (.Height - Frm.Height) / 2
    End With

End Sub
