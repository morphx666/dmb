Attribute VB_Name = "modSubclass"
Option Explicit

Private Const GWL_WNDPROC As Long = (-4)
Public Const WM_GETMINMAXINFO As Long = &H24
Public Const WM_PAINT = &HF
Public Const WM_ACTIVATE = &H6
Public Const WM_ERASEBKGND = &H14
Public Const WM_SETFOCUS = &H7
Public Const WM_DRAWITEM = &H2B
Public Const WM_CTLCOLORSTATIC = &H138
Public Const WM_CTLCOLOREDIT = &H133
Public Const WM_MOVE = &H3

Private Type POINTAPI
    x As Long
    y As Long
End Type

Type MINMAXINFO
    ptReserved As POINTAPI
    ptMaxSize As POINTAPI
    ptMaxPosition As POINTAPI
    ptMinTrackSize As POINTAPI
    ptMaxTrackSize As POINTAPI
End Type

Public Const FLOODFILLBORDER = 0
Public Const FLOODFILLSURFACE = 1

Public Declare Function GetPixel Lib "gdi32" (ByVal hDc As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function SetPixel Lib "gdi32" (ByVal hDc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Public Declare Function ExtFloodFill Lib "gdi32" (ByVal hDc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long, ByVal wFillType As Long) As Long

Public Enum SelViewContants
    svcNormal = 0
    svcMap = 1
End Enum

Public SelView As SelViewContants

Global frmMainHWND As Long
Global frmLCManHWND As Long
Global frmLCManMinWidth As Long
Global frmLCManMinHeight As Long
Global frmBLReportHWND As Long
Global frmIHWHWND As Long
Global frmPOAHWND As Long

Public Function HandleSubclassMsg(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Boolean

    On Error Resume Next
    
    Dim MMI As MINMAXINFO
  
    Select Case hwnd
        Case frmMainHWND
            Select Case uMsg
                Case WM_GETMINMAXINFO
                    CopyMemory MMI, ByVal lParam, LenB(MMI)
                    With MMI
                        .ptMinTrackSize.x = 425
                        .ptMinTrackSize.y = 528
                        .ptMaxTrackSize.x = Screen.Width / Screen.TwipsPerPixelX
                        .ptMaxTrackSize.y = Screen.Height / Screen.TwipsPerPixelY
                    End With
                    CopyMemory ByVal lParam, MMI, LenB(MMI)
                    HandleSubclassMsg = True
                Case WM_PAINT
                    DrawMainSplitters
                    HandleSubclassMsg = False
            End Select
        Case frmLCManHWND
            Select Case uMsg
                Case WM_GETMINMAXINFO
                    CopyMemory MMI, ByVal lParam, LenB(MMI)
                    With MMI
                        .ptMinTrackSize.x = frmLCManMinWidth
                        .ptMinTrackSize.y = frmLCManMinHeight
                        .ptMaxTrackSize.x = Screen.Width / Screen.TwipsPerPixelX
                        .ptMaxTrackSize.y = Screen.Height / Screen.TwipsPerPixelY
                    End With
                    CopyMemory ByVal lParam, MMI, LenB(MMI)
                    HandleSubclassMsg = True
                Case WM_PAINT
                    DrawLCManSplitters
                    HandleSubclassMsg = False
            End Select
        Case frmBLReportHWND
            Select Case uMsg
                Case WM_GETMINMAXINFO
                    CopyMemory MMI, ByVal lParam, LenB(MMI)
                    With MMI
                        .ptMinTrackSize.x = 500
                        .ptMinTrackSize.y = 400
                        .ptMaxTrackSize.x = Screen.Width / Screen.TwipsPerPixelX
                        .ptMaxTrackSize.y = Screen.Height / Screen.TwipsPerPixelY
                    End With
                    CopyMemory ByVal lParam, MMI, LenB(MMI)
                    HandleSubclassMsg = True
            End Select
        Case frmPOAHWND
            Select Case uMsg
                Case WM_GETMINMAXINFO
                    CopyMemory MMI, ByVal lParam, LenB(MMI)
                    With MMI
                        .ptMinTrackSize.x = 480
                        .ptMinTrackSize.y = 494
                        .ptMaxTrackSize.x = Screen.Width / Screen.TwipsPerPixelX
                        .ptMaxTrackSize.y = Screen.Height / Screen.TwipsPerPixelY
                    End With
                    CopyMemory ByVal lParam, MMI, LenB(MMI)
                    HandleSubclassMsg = True
            End Select
        Case frmIHWHWND
            Select Case uMsg
                Case WM_GETMINMAXINFO
                    CopyMemory MMI, ByVal lParam, LenB(MMI)
                    With MMI
                        .ptMinTrackSize.x = 400
                        .ptMinTrackSize.y = 400
                        .ptMaxTrackSize.x = Screen.Width / Screen.TwipsPerPixelX
                        .ptMaxTrackSize.y = Screen.Height / Screen.TwipsPerPixelY
                    End With
                    CopyMemory ByVal lParam, MMI, LenB(MMI)
                    HandleSubclassMsg = True
                Case WM_PAINT
                    DrawIHWSplitters
            End Select
    End Select

End Function

Private Sub DrawIHWSplitters()

    Dim r As RECT
    Dim tppx As Single
    Dim tppy As Single
    
    tppx = Screen.TwipsPerPixelX
    tppy = Screen.TwipsPerPixelY
    
    With frmItemHighlightWizard
        .Cls
        
        ' Draw Files Pane
        r.Left = .tvBrowser.Left / tppx - 1
        r.Right = (.tvBrowser.Left + .tvBrowser.Width) / tppx + 2
        r.Top = .tvBrowser.Top / tppy - 1
        r.bottom = (.tvBrowser.Top + .tvBrowser.Height) / tppy + 2
        DrawEdge .hDc, r, BDR_SUNKENOUTER, BF_RECT
        
        ' Draw Menus Pane
        r.Left = .tvMenus.Left / tppx - 1
        r.Right = (.tvMenus.Left + .tvMenus.Width) / tppx + 2
        r.Top = .tvMenus.Top / tppy - 1
        r.bottom = (.tvMenus.Top + .tvMenus.Height) / tppy + 2
        DrawEdge .hDc, r, BDR_SUNKENOUTER, BF_RECT
        
        ' Draw div
        r.Left = .tvBrowser.Left / tppx - 1
        r.Right = (.tvMenus.Left + .tvMenus.Width) / tppx + 2
        r.Top = (.chkHideDocsNoMenus.Top + .chkHideDocsNoMenus.Height) / tppy + 8
        r.bottom = r.Top + 2
        DrawEdge .hDc, r, BDR_SUNKENOUTER, BF_RECT
    End With

End Sub

Private Sub DrawLCManSplitters()

    Dim r As RECT
    Dim tppx As Single
    Dim tppy As Single
    
    tppx = Screen.TwipsPerPixelX
    tppy = Screen.TwipsPerPixelY
    
    With frmLCMan
        .Cls
        
        ' Draw Folders Pane
        r.Left = .tvFolders.Left / tppx - 1
        r.Right = .picSplit.Left / tppx - 1
        r.Top = .lblFolders.Top / tppy - 3
        r.bottom = (.tvFolders.Top + .tvFolders.Height) / tppy + r.Top - 4
        DrawEdge .hDc, r, BDR_SUNKENOUTER, BF_RECT
        
        ' Draw Files Pane
        r.Top = .lblFiles.Top / tppy - 3
        r.bottom = (.lvFiles.Top + .lvFiles.Height) / tppy + 2
        DrawEdge .hDc, r, BDR_SUNKENOUTER, BF_RECT
        
        ' Draw Label Folders Pane
        r.Left = r.Left + 1
        r.Right = r.Right - 1
        r.Top = .lblFolders.Top / tppy - 2
        r.bottom = (.lblFolders.Top + .lblFolders.Height) / tppy + 4
        DrawEdge .hDc, r, BDR_RAISEDINNER, BF_RECT
        
        ' Draw Label Files Pane
        r.Top = .lblFiles.Top / tppy - 2
        r.bottom = (.lblFiles.Top + .lblFiles.Height) / tppy + 4
        DrawEdge .hDc, r, BDR_RAISEDINNER, BF_RECT
        
        ' Draw Loader Code Pane
        r.Left = .tvLC.Left / tppx - 1
        r.Right = .tvLC.Width / tppx + r.Left + 3
        r.Top = .lblLC.Top / tppy - 3
        r.bottom = (.tvLC.Top + .tvLC.Height) / tppy + r.Top - 4
        DrawEdge .hDc, r, BDR_SUNKENOUTER, BF_RECT
        
        ' Draw Frames Loader Code Pane
        If .lblFLC.Visible Then
            r.Top = .lblFLC.Top / tppy - 3
            r.bottom = (.tvFLC.Top + .tvFLC.Height) / tppy + 2
            DrawEdge .hDc, r, BDR_SUNKENOUTER, BF_RECT
        End If
        
        ' Draw Label LC Pane
        r.Left = r.Left + 1
        r.Right = r.Right - 1
        r.Top = .lblLC.Top / tppy - 2
        r.bottom = (.lblLC.Top + .lblLC.Height) / tppy + 4
        DrawEdge .hDc, r, BDR_RAISEDINNER, BF_RECT
        
        ' Draw Label FLC Pane
        If .lblFLC.Visible Then
            r.Top = .lblFLC.Top / tppy - 2
            r.bottom = (.lblFLC.Top + .lblFLC.Height) / tppy + 4
            DrawEdge .hDc, r, BDR_RAISEDINNER, BF_RECT
        End If
    End With

End Sub

Private Sub DrawMainSplitters()

    Dim r As RECT
    Dim TopPos As Long
    Dim tppx As Single
    Dim tppy As Single
    Dim BotPos As Long
    
    tppx = Screen.TwipsPerPixelX
    tppy = Screen.TwipsPerPixelY
    
    With frmMain
        TopPos = Abs(.tbMenu.Visible + .tbMenu2.Visible + .tbCmd.Visible) * .tbMenu.Height + 15
        BotPos = .Height - GetClientTop(.hwnd) - .sbDummy.Height - 90
        
        .Cls
        
        ' Draw Tabs Background a-la-ACID
        r.Left = 1
        r.Right = (.picSplit.Left / tppx) - 1
        r.Top = (.tvMenus.Top + .tvMenus.Height) / tppy + 2
        r.bottom = r.Top + 22
        frmMain.Line (r.Left * tppx, r.Top * tppy)-(r.Right * tppx, r.bottom * tppy), .FillColor, BF
        
        ' Draw First Tab
        If SelView = svcNormal Then
            DrawTabAt 65, (.tvMenus.Top + .tvMenus.Height) / tppy + 2, 60, False
        Else
            DrawTabAt 1, (.tvMenus.Top + .tvMenus.Height) / tppy + 2, 60, False
        End If

        r.Left = 0
        r.Right = (.picSplit.Left / tppx)
        r.Top = (.tvMenus.Top + .tvMenus.Height) / tppy + 2
        r.bottom = r.Top + 2
        DrawEdge .hDc, r, BDR_SUNKENOUTER, BF_RECT

        If SelView = svcNormal Then
            DrawTabAt 1, (.tvMenus.Top + .tvMenus.Height) / tppy + 2, 60, True
        Else
            DrawTabAt 65, (.tvMenus.Top + .tvMenus.Height) / tppy + 2, 60, True
        End If
        
        ' Draw Treeview Pane
        r.Left = 0
        r.Right = (.picSplit.Left / tppx)
        r.Top = (TopPos / tppy)
        r.bottom = (BotPos / tppy) + 5
        DrawEdge .hDc, r, BDR_SUNKENOUTER, BF_RECT
        
        ' Draw Data Pane
        r.Left = (.picSplit.Left / tppx) + 8
        r.Right = r.Left + (.Width - .picSplit.Left) / tppx - 16
        r.Top = (TopPos / tppy)
        r.bottom = (BotPos / tppy) + 5
        DrawEdge .hDc, r, BDR_SUNKENOUTER, BF_RECT
        
        ' Draw data Title Pane
        r.Left = (.picSplit.Left / tppx) + 9
        r.Right = r.Left + (.Width - .picSplit.Left) / tppx - 18
        r.Top = (TopPos / tppy) + 1
        r.bottom = r.Top + 28
        DrawEdge .hDc, r, BDR_RAISEDINNER, BF_RECT
        
        If .wbMainPreview.Visible Then
            ' Draw LivePreview Pane
            r.Left = 1
            r.Right = (.picSplit.Left / tppx) - 1
            r.Top = ((TopPos + .tvMenus.Height + 405) / tppy) - 1
            r.bottom = (BotPos / tppy) + 4
            DrawEdge .hDc, r, BDR_RAISEDINNER, BF_RECT
        End If
    End With
    
End Sub

Private Sub DrawTabAt(x As Long, y As Long, w As Long, Selected As Boolean)

    Dim r As RECT
    Dim i As Long
    Dim h As Long
    
    If Not Selected Then
        h = 19
        y = y + 1
    Else
        h = 20
    End If
    
    With frmMain
        ' \
        r.Left = x
        r.Right = r.Left + 8
        r.Top = y
        r.bottom = r.Top + h
        DrawEdge .hDc, r, BDR_SUNKENOUTER Or BDR_RAISEDINNER, BF_DIAGONAL_ENDBOTTOMRIGHT
        
        ' -
        r.Left = r.Right - 2
        r.Right = r.Left + w
        r.Top = r.bottom - 3
        r.bottom = r.Top
        r.Top = r.bottom
        DrawEdge .hDc, r, BDR_RAISEDINNER, BF_RECT
        
        ' /
        r.Left = r.Right - 1
        r.Right = r.Left + 8
        r.Top = y - 2
        r.bottom = r.Top + h
        DrawEdge .hDc, r, BDR_SUNKENOUTER Or BDR_RAISEDINNER, BF_DIAGONAL_ENDBOTTOMLEFT
        
        r.Top = y - 1
        r.Left = x
        r.Right = r.Left + w + 13
        For i = w To w - 15 + IIf(Selected, 0, 1) Step -1
            If Int(i / 2) = i / 2 Then
                r.Left = r.Left + 1
                r.Right = r.Right - 1
            End If
            r.Top = r.Top + 1
            r.bottom = r.Top + 1
            DrawEdge .hDc, r, BDR_SUNKENINNER, BF_RECT Or BF_MIDDLE
            If i < w - 1 And Not Selected Then SetPixel .hDc, r.Right, r.Top - 1, GetPixel(.hDc, x, y + 2)
        Next i
        SetPixel .hDc, x - 1, y + 1, GetPixel(.hDc, x, y + 2)
        SetPixel .hDc, x + w + 12, y + 1, GetPixel(.hDc, x, y + 2)
    End With
    
End Sub
