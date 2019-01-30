VERSION 5.00
Begin VB.UserControl ucDMBSampleCtrl 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   ClientHeight    =   2280
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4845
   LockControls    =   -1  'True
   ScaleHeight     =   152
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   323
   Begin VB.PictureBox picCnt 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2025
      Left            =   60
      ScaleHeight     =   135
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   300
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   60
      Width           =   4500
      Begin VB.PictureBox picCTopLeft 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   945
         ScaleHeight     =   14
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   12
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   405
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.PictureBox picCTopCenter 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   1335
         ScaleHeight     =   14
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   105
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   420
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.PictureBox picCTopRight 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   3240
         ScaleHeight     =   14
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   12
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   255
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.PictureBox picCLeft 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   960
         ScaleHeight     =   27
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   12
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   750
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.PictureBox picCRight 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   390
         Left            =   3450
         ScaleHeight     =   26
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   12
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   750
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.PictureBox picCBottomLeft 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   1050
         ScaleHeight     =   14
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   12
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1275
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.PictureBox picCBottomCenter 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   1305
         ScaleHeight     =   14
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   119
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1530
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.PictureBox picCBottomRight 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   3450
         ScaleHeight     =   14
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   12
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   1440
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.PictureBox picGrp 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   690
         Left            =   1245
         ScaleHeight     =   46
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   138
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   660
         Width           =   2070
         Begin VB.PictureBox picCmd 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00800000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   420
            Left            =   90
            MouseIcon       =   "ucDMBSampleCtrl.ctx":0000
            ScaleHeight     =   28
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   123
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   105
            Width           =   1845
            Begin VB.Label lblSample 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Sample"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   195
               Left            =   480
               TabIndex        =   11
               Top             =   150
               UseMnemonic     =   0   'False
               Visible         =   0   'False
               Width           =   630
            End
            Begin VB.Line lnSep 
               BorderColor     =   &H80000005&
               X1              =   37
               X2              =   91
               Y1              =   6
               Y2              =   5
            End
            Begin VB.Image imgR 
               Appearance      =   0  'Flat
               Height          =   210
               Left            =   1515
               MouseIcon       =   "ucDMBSampleCtrl.ctx":0152
               Stretch         =   -1  'True
               Top             =   105
               Width           =   210
            End
            Begin VB.Image imgL 
               Appearance      =   0  'Flat
               Height          =   210
               Left            =   75
               MouseIcon       =   "ucDMBSampleCtrl.ctx":02A4
               Stretch         =   -1  'True
               Top             =   105
               Width           =   210
            End
         End
      End
   End
End
Attribute VB_Name = "ucDMBSampleCtrl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Event SampleUpdated()

Private Enum LastModeConstants
    [Normal]
    [Hover]
    [None]
End Enum

Dim cmd As MenuCmd
Dim grp As MenuGrp
Dim LastMode As LastModeConstants
Dim UsesToolbar As Boolean
Dim SelIsGroup As Boolean

Dim leftCorner As Integer
Dim topCorner As Integer
Dim rightCorner As Integer
Dim bottomCorner As Integer

Public Sub ctrlVars(vars As Variant)

    Project = vars(0)
    AppPath = vars(1)
    Preferences = vars(3)
    StatesPath = vars(4)
    Set frmMain = vars(5)
    
    Set FloodPanel.PictureControl = vars(2)
    Dim UIObjects(1 To 3) As Object
    Set UIObjects(1) = frmMain
    Set UIObjects(2) = FloodPanel
    Set UIObjects(3) = frmMain
    SetUI UIObjects
    
    Dim VarObjects(1 To 7) As Variant
    VarObjects(1) = GetSetting("DMB", "RegInfo", "InstallPath")
    VarObjects(2) = ""
    VarObjects(3) = ""
    VarObjects(4) = GetTEMPPath
    VarObjects(5) = cSep
    VarObjects(6) = nwdPar
    VarObjects(7) = StatesPath
    SetVars VarObjects
    
    grp.CornersImages.gcBottomLeft = ""
    grp.CornersImages.gcBottomCenter = ""
    grp.CornersImages.gcBottomRight = ""
    
    grp.CornersImages.gcTopLeft = ""
    grp.CornersImages.gcTopCenter = ""
    grp.CornersImages.gcTopRight = ""
    
    grp.CornersImages.gcLeft = ""
    grp.CornersImages.gcRight = ""
    RenderCorners True

End Sub

Private Sub imgL_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    picCmd_MouseMove Button, Shift, x / Screen.TwipsPerPixelX + imgL.Left, y / Screen.TwipsPerPixelY + imgL.Top

End Sub

Private Sub imgR_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    picCmd_MouseMove Button, Shift, x / Screen.TwipsPerPixelX + imgR.Left, y / Screen.TwipsPerPixelY + imgR.Top

End Sub

Private Sub lblSample_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    picCmd_MouseMove Button, Shift, x / Screen.TwipsPerPixelX + lblSample.Left, y / Screen.TwipsPerPixelY + lblSample.Top

End Sub

Private Sub picCmd_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    If (x >= 0 And x <= picCmd.Width) And (y >= 0 And y <= picCmd.Height) Then
        SetHoverState
        SetCapture picCmd.hwnd
    Else
        If (GetAsyncKeyState(VK_LBUTTON) = 0) And (GetAsyncKeyState(VK_RBUTTON) = 0) Then
            ReleaseCapture
            SetNormalState False
        End If
    End If

End Sub

Private Sub UserControl_Initialize()
    
    SetTemplateDefaults
    
    ReDim MenuGrps(0)
    ReDim MenuCmds(0)
    
    DoUNICODE = (GetSetting("DMB", "Preferences", "DoUNICODE", 1) = 1)
    cSep = Chr(255) + Chr(255)
    
    cmd = TemplateCommand
    grp = TemplateGroup
    
    picCnt.Move 0, 0, Width, Height
    picCnt.BackColor = &H8000000F
    
End Sub

Public Sub DefItems(idx As Integer, cmds As Variant, grps As Variant, Optional IsGroup As Boolean)
    
    Dim g As Integer

    MenuCmds = cmds
    MenuGrps = grps
    
    SelIsGroup = IsGroup
    
    On Error Resume Next
    UsesToolbar = CreateToolbar
    On Error GoTo DisablePreview

    If SelIsGroup Then
        g = idx
        grp = MenuGrps(idx)
        With cmd
            .Name = grp.Name
            .Alignment = grp.CaptionAlignment
            .Caption = grp.Caption
            .hBackColor = grp.hBackColor
            .HoverFont = grp.DefHoverFont
            .hTextColor = grp.hTextColor
            .iCursor = grp.iCursor
            .LeftImage = grp.tbiLeftImage
            .nBackColor = grp.nBackColor
            .NormalFont = grp.DefNormalFont
            .nTextColor = grp.nTextColor
            .Parent = idx
            .RightImage = grp.tbiRightImage
            .BackImage = grp.tbiBackImage
            
            .CmdsFXhColor = grp.CmdsFXhColor
            .CmdsFXnColor = grp.CmdsFXnColor
            .CmdsFXNormal = grp.CmdsFXNormal
            .CmdsFXOver = grp.CmdsFXOver
            .CmdsFXSize = grp.CmdsFXSize
            .CmdsMarginX = grp.CmdsMarginX
            .CmdsMarginY = grp.CmdsMarginY
        End With
    Else
        g = MenuCmds(idx).Parent
        cmd = MenuCmds(idx)
        grp = MenuGrps(cmd.Parent)
        
        With grp
            .CmdsFXhColor = cmd.CmdsFXhColor
            .CmdsFXnColor = cmd.CmdsFXnColor
            .CmdsFXNormal = cmd.CmdsFXNormal
            .CmdsFXOver = cmd.CmdsFXOver
            .CmdsFXSize = cmd.CmdsFXSize
            .CmdsMarginX = cmd.CmdsMarginX
            .CmdsMarginY = cmd.CmdsMarginY
        End With
    End If
    cmd.Caption = xUNI2Unicode(cmd.Caption)
    leftCorner = CornerSize(g, "left")
    topCorner = CornerSize(g, "top")
    rightCorner = CornerSize(g, "right")
    bottomCorner = CornerSize(g, "bottom")
    
    picCnt.Visible = True
    
    UpdateSample

    Exit Sub

DisablePreview:
    picCnt.Visible = False
    
    Debug.Print "DefItems: " + Err.Description
    
End Sub

Private Sub UpdateSample()

    Dim ItemHeight As Integer
    Dim ItemWidth As Integer
    Dim InvalidData As Boolean
    Dim g As Integer
    Dim SepWidth As Integer
    Dim c As Integer
    
    If UBound(MenuGrps) = 0 Then
        cmd = TemplateCommand
        grp = TemplateGroup
        ItemHeight = 20
        ItemWidth = 130
        InvalidData = True
    Else
        RenderCorners
        If SelIsGroup Then
            g = GetIDByName(grp.Name)
            If UsesToolbar Then
                ItemHeight = GetHotSpotHeight(MemberOf(g), g, True) - 2 * grp.CmdsFXSize
                ItemWidth = GetHotSpotWidth(MemberOf(g), g, True) '- grp.CmdsFXSize
            Else
                ItemHeight = 25
                ItemWidth = GetDivWidth(g)
            End If
        Else
            c = GetIDByName(cmd.Name)
            ItemHeight = GetSingleCommandHeight(c)
            Select Case grp.AlignmentStyle
                Case ascVertical
                    ItemWidth = GetDivWidth(cmd.Parent) - 2 * grp.FrameBorder
                Case ascHorizontal
                    ItemWidth = GetSingleCommandWidth(c) + 2 * grp.FrameBorder
            End Select
        End If
        InvalidData = False
    End If
    
    With grp
        picGrp.Move .FrameBorder + leftCorner, _
                    .FrameBorder + topCorner, _
                    ItemWidth, _
                    ItemHeight + IIf(SelIsGroup, 0, 2 * .Leading) + 2 * (.CmdsFXSize + .ContentsMarginV)
                     
        picCmd.Move .ContentsMarginH + .CmdsFXSize, _
                    IIf(SelIsGroup, 0, .Leading) + .ContentsMarginV + .CmdsFXSize, _
                    Abs(ItemWidth - 2 * .ContentsMarginH - 2 * .CmdsFXSize), _
                    ItemHeight
    End With
    
    picCmd.Visible = True
    If cmd.Name = "[SEP]" Then
        lnSep.Visible = True
        imgL.Visible = False
        imgR.Visible = False
        lblSample.Visible = False
        
        SepWidth = (ItemWidth - 2 * grp.ContentsMarginH) * cmd.SeparatorPercent / 100
        lnSep.X1 = ((ItemWidth - 2 * grp.ContentsMarginH) - SepWidth) \ 2
        lnSep.Y1 = ItemHeight / 2
        lnSep.X2 = lnSep.X1 + SepWidth
        lnSep.Y2 = lnSep.Y1
        lnSep.BorderColor = cmd.nTextColor
    Else
        lnSep.Visible = False
        imgL.Visible = True
        imgR.Visible = True
        lblSample.Visible = True
        
        With imgL
            '.Picture = LoadPictureRes(cmd.LeftImage.NormalImage)
            .Width = cmd.LeftImage.w ' IIf(.Picture, cmd.LeftImage.w, 0)
            .Height = cmd.LeftImage.h ' IIf(.Picture, cmd.LeftImage.h, 0)
            .Move cmd.CmdsMarginX, (picCmd.Height - .Height) / 2
        End With
        
        With imgR
            '.Picture = LoadPictureRes(cmd.RightImage.NormalImage)
            .Width = cmd.RightImage.w ' IIf(.Picture, cmd.RightImage.w, 0)
            .Height = cmd.RightImage.h ' IIf(.Picture, cmd.RightImage.h, 0)
            .Move picCmd.Width - .Width - cmd.CmdsMarginX, (picCmd.Height - .Height) / 2
        End With
        
        lblSample.Caption = IIf(InvalidData, "Preview not Available", ParseCaptionHTML(cmd.Caption))
        'If DoUNICODE Then
        '    lblSample.Caption = Replace(StrConv(lblSample.Caption, vbUnicode), vbNullChar, "")
        'End If
        Select Case cmd.Alignment
            Case tacLeft
                lblSample.Alignment = vbLeftJustify
            Case tacCenter
                lblSample.Alignment = vbCenter
            Case tacRight
                lblSample.Alignment = vbRightJustify
        End Select
        lblSample.Alignment = lblSample.Alignment
    End If
    
    LastMode = None

    SetNormalState True

End Sub

Private Sub PaintGroupBackImage()

    If SelIsGroup And UsesToolbar Then
        With Project.Toolbars(MemberOfByName(grp.Name))
            SetColor .BackColor, picGrp
            If LenB(.Image) <> 0 Then TileImage .Image, picGrp
        End With
    Else
        SetColor grp.bColor, picGrp
        If LenB(grp.Image) <> 0 Then
            TileImage grp.Image, picGrp
        Else
            If grp.bColor = -2 Then
                picGrp.BackColor = Parent.BackColor
            End If
        End If
    End If

End Sub

Private Sub UserControl_Resize()

    Static IsResizing As Boolean
    
    If IsResizing Then Exit Sub
    
    IsResizing = True

    Width = (picGrp.Width + 2 * grp.FrameBorder + leftCorner + rightCorner) * 15
    Height = (picGrp.Height + 2 * grp.FrameBorder + topCorner + bottomCorner) * 15
    
    picCnt.Move 0, 0, Width, Height
    
    IsResizing = False

End Sub

Private Sub SetHoverState()

    Dim SelCrs As Integer
    
    On Error GoTo ExitSub
    
    If LastMode = Hover Or cmd.Disabled Or cmd.Name = "[SEP]" Then Exit Sub
    
    If SelIsGroup And Not UsesToolbar Then
        Exit Sub
    End If

    SetColor cmd.hBackColor, picCmd
    If LenB(cmd.BackImage.HoverImage) <> 0 Then TileImage cmd.BackImage.HoverImage, picCmd, Not cmd.BackImage.Tile
    
    If cmd.hBackColor = -2 Then
        TileImage grp.Image, picGrp
        BitBlt picCmd.hDc, 0, 0, picCmd.Width, picCmd.Height, picGrp.hDc, picCmd.Left, picCmd.Top, vbSrcCopy
    Else
        PaintGroupBackImage
    End If
    
    With cmd
        If .CmdsFXhColor <> -2 Then
            Select Case .CmdsFXOver
                Case cfxcNone
                    DoBevel picGrp, picCmd, .CmdsFXSize, .CmdsFXhColor, .CmdsFXhColor, .CmdsFXhColor, .CmdsFXhColor
                Case cfxcRaised
                    DoBevel picGrp, picCmd, .CmdsFXSize, LightenColor(.CmdsFXhColor), LightenColor(.CmdsFXhColor), DarkenColor(.CmdsFXhColor), DarkenColor(.CmdsFXhColor)
                Case cfxcSunken
                    DoBevel picGrp, picCmd, .CmdsFXSize, DarkenColor(.CmdsFXhColor), DarkenColor(.CmdsFXhColor), LightenColor(.CmdsFXhColor), LightenColor(.CmdsFXhColor)
                Case cfxcDouble
                    DoBevel picGrp, picCmd, .CmdsFXSize, .CmdsFXhColor, .CmdsFXhColor, .CmdsFXhColor, .CmdsFXhColor, 1
                Case cfxcBevel
                    DoBevel picGrp, picCmd, .CmdsFXSize, DarkenColor(.CmdsFXhColor, 2), DarkenColor(.CmdsFXhColor, 2), DarkenColor(.CmdsFXhColor, 1.3), DarkenColor(.CmdsFXhColor, 1.3), 2
            End Select
        End If
    End With
    
    imgL.Picture = LoadPictureRes(cmd.LeftImage.HoverImage)
    imgR.Picture = LoadPictureRes(cmd.RightImage.HoverImage)
    If LenB(cmd.BackImage.HoverImage) <> 0 Then TileImage cmd.BackImage.HoverImage, picCmd, Not cmd.BackImage.Tile
    
    On Error Resume Next
    With lblSample
        With .Font
            .Name = cmd.HoverFont.FontName
            .Size = px2pt(cmd.HoverFont.FontSize)
            .Bold = cmd.HoverFont.FontBold
            .Italic = cmd.HoverFont.FontItalic
            .Underline = cmd.HoverFont.FontUnderline
        End With
        .ForeColor = cmd.hTextColor
    End With
    
    Select Case cmd.iCursor.cType
        Case iccDefault
            SelCrs = vbArrow
        Case iccHand
            SelCrs = vbCustom
        Case iccCrosshair
            SelCrs = vbCrosshair
        Case iccHelp
            SelCrs = vbArrowQuestion
        Case iccText
            SelCrs = vbIbeam
        Case iccResizeE, iccResizeW
            SelCrs = vbSizeWE
        Case iccResizeNE, iccResizeSW
            SelCrs = vbSizeNESW
        Case iccResizeNW, iccResizeSE
            SelCrs = vbSizeNWSE
        Case iccResizeN, iccResizeS
            SelCrs = vbSizeNS
        Case iccResizeAll
            SelCrs = vbSizeAll
        Case iccWait
            SelCrs = vbHourglass
        Case iccCustom
            SelCrs = vbCustom
            'If IsANI(cmd.iCursor.CFile) Then
            '    SelCrs = vbArrow
            'Else
                picCmd.MouseIcon = LoadPictureRes(cmd.iCursor.cFile)
            'End If
    End Select
    
    picCmd.MousePointer = SelCrs
    imgL.MousePointer = SelCrs
    imgR.MousePointer = SelCrs
    lblSample.MousePointer = SelCrs
    
    LastMode = Hover
    
ExitSub:
    
End Sub

Public Sub SetNormalState(Optional ForceUpdate As Boolean)

    Dim imgLWidth As Integer
    Dim imgRWidth As Integer

    If LastMode = Normal And Not ForceUpdate Then Exit Sub
    
    If SelIsGroup And Not UsesToolbar Then
        lblSample.Visible = False
        picCmd.Visible = False
        GoTo ExitSub
    End If
    
    SetColor cmd.nBackColor, picCmd
    If LenB(cmd.BackImage.NormalImage) <> 0 Then
        TileImage cmd.BackImage.NormalImage, picCmd, Not cmd.BackImage.Tile
    End If
    
    If cmd.nBackColor = -2 Then
        If SelIsGroup And UsesToolbar Then
            With Project.Toolbars(MemberOfByName(grp.Name))
                If LenB(.Image) <> 0 Then TileImage .Image, picGrp
            End With
        Else
            If LenB(grp.Image) <> 0 Then TileImage grp.Image, picGrp
        End If
        BitBlt picCmd.hDc, 0, 0, picCmd.Width, picCmd.Height, picGrp.hDc, picCmd.Left, picCmd.Top, vbSrcCopy
    Else
        PaintGroupBackImage
    End If
    
    With cmd
        If .CmdsFXnColor <> -2 Then
            Select Case .CmdsFXNormal
                Case cfxcNone
                    DoBevel picGrp, picCmd, .CmdsFXSize, .CmdsFXnColor, .CmdsFXnColor, .CmdsFXnColor, .CmdsFXnColor
                Case cfxcRaised
                    DoBevel picGrp, picCmd, .CmdsFXSize, LightenColor(.CmdsFXnColor), LightenColor(.CmdsFXnColor), DarkenColor(.CmdsFXnColor), DarkenColor(.CmdsFXnColor)
                Case cfxcSunken
                    DoBevel picGrp, picCmd, .CmdsFXSize, DarkenColor(.CmdsFXnColor), DarkenColor(.CmdsFXnColor), LightenColor(.CmdsFXnColor), LightenColor(.CmdsFXnColor)
                Case cfxcDouble
                    DoBevel picGrp, picCmd, .CmdsFXSize, .CmdsFXnColor, .CmdsFXnColor, .CmdsFXnColor, .CmdsFXnColor, 1
                Case cfxcBevel
                    DoBevel picGrp, picCmd, .CmdsFXSize, DarkenColor(.CmdsFXnColor, 2), DarkenColor(.CmdsFXnColor, 2), DarkenColor(.CmdsFXnColor, 1.3), DarkenColor(.CmdsFXnColor, 1.3), 2
            End Select
        End If
    End With
    
    If cmd.Name <> "[SEP]" Then
        imgL.Picture = LoadPictureRes(cmd.LeftImage.NormalImage)
        imgR.Picture = LoadPictureRes(cmd.RightImage.NormalImage)
        If LenB(cmd.BackImage.NormalImage) <> 0 Then TileImage cmd.BackImage.NormalImage, picCmd, Not cmd.BackImage.Tile
        
        If LenB(cmd.LeftImage.NormalImage) <> 0 Then
            imgLWidth = imgL.Left + imgL.Width + Preferences.ImgSpace
        Else
            imgLWidth = imgL.Left
        End If
        If LenB(cmd.RightImage.NormalImage) <> 0 Then
            imgRWidth = imgL.Left + imgR.Width + Preferences.ImgSpace
        Else
            imgRWidth = imgL.Left
        End If
        
        On Error Resume Next
        With lblSample
            With .Font
                .Name = cmd.NormalFont.FontName
                .Size = px2pt(cmd.NormalFont.FontSize)
                .Bold = cmd.NormalFont.FontBold
                .Italic = cmd.NormalFont.FontItalic
                .Underline = cmd.NormalFont.FontUnderline
                
                lblSample.Font.Name = .Name
                lblSample.Font.Size = .Size
                lblSample.Font.Bold = .Bold
                lblSample.Font.Italic = .Italic
                lblSample.Font.Underline = .Underline
            End With
            .ForeColor = cmd.nTextColor
            
            lblSample.Height = lblSample.Height
            Select Case cmd.Alignment
                Case tacLeft
                    .Move imgLWidth, (picCmd.Height - .Height) / 2
                Case tacCenter
                    .Move imgLWidth + ((picCmd.Width - (imgLWidth + imgRWidth)) - .Width) / 2, (picCmd.Height - .Height) / 2
                Case tacRight
                    .Move picCmd.Width - imgRWidth - .Width, (picCmd.Height - .Height) / 2
            End Select
        End With
    End If
    
ExitSub:
    
    UserControl_Resize
    
    RaiseEvent SampleUpdated
    
    LastMode = Normal

End Sub

Private Sub RenderCorners(Optional Force As Boolean = False)

    'Top Left
    Dim tll As Integer
    Dim tlt As Integer
    Dim tlw As Integer
    Dim tlh As Integer
    
    'Top Center
    Dim tcl As Integer
    Dim tct As Integer
    Dim tcw As Integer
    Dim tch As Integer
    
    'Top Right
    Dim trl As Integer
    Dim trt As Integer
    Dim trw As Integer
    Dim trh As Integer
    
    'Center Left
    Dim cll As Integer
    Dim clt As Integer
    Dim clw As Integer
    Dim clh As Integer
    
    'Center Right
    Dim crl As Integer
    Dim crt As Integer
    Dim crw As Integer
    Dim crh As Integer
    
    'Bottom Left
    Dim bll As Integer
    Dim blt As Integer
    Dim blw As Integer
    Dim blh As Integer
    
    'Bottom Center
    Dim bcl As Integer
    Dim bct As Integer
    Dim bcw As Integer
    Dim bch As Integer
    
    'Bottom Right
    Dim brl As Integer
    Dim brt As Integer
    Dim brw As Integer
    Dim brh As Integer
    
    Dim DivWidth As Integer
    Dim DivHeight As Integer
    
    With grp.CornersImages
        If LenB(.gcBottomCenter) = 0 And LenB(.gcBottomLeft) = 0 And LenB(.gcBottomRight) = 0 And _
            LenB(.gcLeft) = 0 And LenB(.gcRight) = 0 And _
            LenB(.gcTopCenter) = 0 And LenB(.gcTopLeft) = 0 And LenB(.gcTopRight) = 0 And _
            Not Force Then Exit Sub
            
        DivWidth = picGrp.Width + 2 * grp.FrameBorder
        DivHeight = picGrp.Height + 2 * grp.FrameBorder
    
        'Center Width/Height
        clw = GetImageSize(.gcLeft)(0)
        clh = DivHeight
        crw = GetImageSize(.gcRight)(0)
        crh = DivHeight
    
        'Top Width/Height
        tlw = GetImageSize(.gcTopLeft)(0)
        tlh = GetImageSize(.gcTopLeft)(1)
        trw = GetImageSize(.gcTopRight)(0)
        trh = GetImageSize(.gcTopRight)(1)
        tcw = DivWidth + IIf(tlw = 0, leftCorner, 0) + IIf(trw = 0, rightCorner, 0)
        tch = GetImageSize(.gcTopCenter)(1)
        
        'Bottom Width/Height
        blw = GetImageSize(.gcBottomLeft)(0)
        blh = GetImageSize(.gcBottomLeft)(1)
        brw = GetImageSize(.gcBottomRight)(0)
        brh = GetImageSize(.gcBottomRight)(1)
        bcw = DivWidth + IIf(blw = 0, leftCorner, 0) + IIf(brw = 0, rightCorner, 0)
        bch = GetImageSize(.gcBottomCenter)(1)
        
        'Center Adjustment
        clh = clh + IIf(tlh = 0, topCorner, 0) + IIf(blh = 0, bottomCorner, 0)
        crh = crh + IIf(trh = 0, topCorner, 0) + IIf(brh = 0, bottomCorner, 0)
        
        'Top Left/Top
        tll = leftCorner - tlw
        tlt = topCorner - tlh
        tcl = tll + tlw
        tct = topCorner - tch
        trl = leftCorner + DivWidth
        trt = topCorner - trh
        
        'Center Left/Top
        cll = leftCorner - clw
        clt = tlt + tlh
        crl = leftCorner + DivWidth
        crt = trt + trh
        
        'Bottom Left/Top
        bll = leftCorner - blw
        blt = topCorner + DivHeight
        bcl = bll + blw
        bct = topCorner + DivHeight
        brl = leftCorner + DivWidth
        brt = topCorner + DivHeight
        
        'TOP CORNERs
        If LenB(.gcTopLeft) <> 0 Then
            picCTopLeft.Visible = True
            picCTopLeft.Picture = LoadPictureRes(.gcTopLeft)
            picCTopLeft.Move tll, tlt, tlw, tlh
            MkTransparent picCTopLeft
        Else
            picCTopLeft.Visible = False
        End If
        
        If LenB(.gcTopCenter) <> 0 Then
            picCTopCenter.Visible = True
            picCTopCenter.Move tcl, tct, tcw, tch
            TileImage .gcTopCenter, picCTopCenter
            MkTransparent picCTopCenter
        Else
            picCTopCenter.Visible = False
        End If
        
        If LenB(.gcTopRight) <> 0 Then
            picCTopRight.Visible = True
            picCTopRight.Picture = LoadPictureRes(.gcTopRight)
            picCTopRight.Move trl, trt, trw, trh
            MkTransparent picCTopRight
        Else
            picCTopRight.Visible = False
        End If
        
        'CENTER CORNERs
        If LenB(.gcLeft) <> 0 Then
            picCLeft.Visible = True
            picCLeft.Move cll, clt, clw, clh
            TileImage .gcLeft, picCLeft
            MkTransparent picCLeft
        Else
            picCLeft.Visible = False
        End If
        
        If LenB(.gcRight) <> 0 Then
            picCRight.Visible = True
            picCRight.Move crl, crt, crw, crh
            TileImage .gcRight, picCRight
            MkTransparent picCRight
        Else
            picCRight.Visible = False
        End If
        
        'BOTTOM CORNERs
        If LenB(.gcBottomLeft) <> 0 Then
            picCBottomLeft.Visible = True
            picCBottomLeft.Picture = LoadPictureRes(.gcBottomLeft)
            picCBottomLeft.Move bll, blt, blw, blh
            MkTransparent picCBottomLeft
        Else
            picCBottomLeft.Visible = False
        End If
    
        If LenB(.gcBottomCenter) <> 0 Then
            picCBottomCenter.Visible = True
            picCBottomCenter.Move bcl, bct, bcw, bch
            TileImage .gcBottomCenter, picCBottomCenter
            MkTransparent picCBottomCenter
        Else
            picCBottomCenter.Visible = False
        End If
        
        If LenB(.gcBottomRight) <> 0 Then
            picCBottomRight.Visible = True
            picCBottomRight.Picture = LoadPictureRes(.gcBottomRight)
            picCBottomRight.Move brl, brt, brw, brh
            MkTransparent picCBottomRight
        Else
            picCBottomRight.Visible = False
        End If
    End With

End Sub

Private Sub MkTransparent(pic As PictureBox)

    Dim x As Integer
    Dim y As Integer
    Dim tc As Long
    
    Exit Sub
    
    tc = pic.Point(0, 0)
    
    For x = 0 To pic.Width
        For y = 0 To pic.Height
            If pic.Point(x, y) = tc Then pic.PSet (x, y), &H8000000F
            'If pic.Point(x, Y) = tc Then pic.PSet (x, Y), picCnt.Point(x + pic.Left, Y + pic.Top)
        Next y
    Next x
    
End Sub

Public Sub SetupCharset(cs As Long)

    Dim ctrl As Control
    Dim FontObj As Object
    
    On Error Resume Next
    
    For Each ctrl In Controls
        Err.Clear
        Set FontObj = ctrl.Font
        If Err.number = 0 Then
            FontObj.Charset = cs
        End If
    Next ctrl

End Sub
