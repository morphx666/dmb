VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSelPreset 
   Caption         =   "Select Style from Preset"
   ClientHeight    =   6300
   ClientLeft      =   4740
   ClientTop       =   3480
   ClientWidth     =   7965
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSelPreset.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6300
   ScaleWidth      =   7965
   Begin VB.FileListBox flbPresets 
      Height          =   1065
      Left            =   4005
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   75
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   5910
      TabIndex        =   8
      Top             =   4695
      Width           =   900
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6930
      TabIndex        =   7
      Top             =   4695
      Width           =   900
   End
   Begin VB.Frame frameInfo 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2460
      Left            =   60
      TabIndex        =   1
      Top             =   2130
      Width           =   3105
      Begin VB.TextBox txtComments 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   1380
         Left            =   90
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Text            =   "frmSelPreset.frx":038A
         Top             =   1005
         Width           =   2925
      End
      Begin VB.Label lblLabelAuthor 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Author:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   90
         TabIndex        =   6
         Top             =   285
         Width           =   630
      End
      Begin VB.Label lblLabelComments 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Comments:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   90
         TabIndex        =   5
         Top             =   780
         Width           =   960
      End
      Begin VB.Label lblAuthor 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TheAUTHOR"
         Height          =   195
         Left            =   90
         TabIndex        =   4
         Top             =   495
         Width           =   900
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         BackColor       =   &H80000002&
         Caption         =   "TheTITLE"
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
         Height          =   210
         Left            =   45
         TabIndex        =   3
         Top             =   0
         Width           =   3000
      End
   End
   Begin VB.PictureBox picFlood 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   420
      ScaleHeight     =   255
      ScaleWidth      =   3270
      TabIndex        =   0
      Top             =   5760
      Visible         =   0   'False
      Width           =   3270
   End
   Begin VB.Timer tmrInit 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   2910
      Top             =   5040
   End
   Begin VB.Timer tmrCloseDlg 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   7080
      Top             =   5550
   End
   Begin MSComctlLib.ImageList ilIcons 
      Left            =   1140
      Top             =   4650
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelPreset.frx":0398
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelPreset.frx":0A6A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelPreset.frx":0E04
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar sbInfo 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   9
      Top             =   5985
      Width           =   7965
      _ExtentX        =   14049
      _ExtentY        =   556
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin SHDocVwCtl.WebBrowser wbPreview 
      Height          =   4350
      Left            =   3255
      TabIndex        =   11
      Top             =   240
      Width           =   4575
      ExtentX         =   8070
      ExtentY         =   7673
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   1
      RegisterAsDropTarget=   0
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin MSComctlLib.TreeView tvPresets 
      Height          =   1845
      Left            =   60
      TabIndex        =   12
      Top             =   240
      Width           =   3105
      _ExtentX        =   5477
      _ExtentY        =   3254
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   441
      LabelEdit       =   1
      Style           =   7
      ImageList       =   "ilIcons"
      Appearance      =   1
   End
   Begin VB.Label lblPresets 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Presets"
      Height          =   195
      Left            =   60
      TabIndex        =   14
      Top             =   0
      Width           =   540
   End
   Begin VB.Label lblPreview 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Preview"
      Height          =   195
      Left            =   2955
      TabIndex        =   13
      Top             =   0
      Width           =   570
   End
End
Attribute VB_Name = "frmSelPreset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Prj As ProjectDef
Dim mg() As String
Dim sc() As String
Dim mc() As String

Private Type SelPropertiesDef
    tbToolbarStyle As Boolean
    tbBorder As Boolean
    tbMargins As Boolean
    tbSeparation As Boolean
    tbJustification As Boolean
    tbBackColor As Boolean
    tbBackImage As Boolean
    
    tbAlignment As Boolean
    tbSpanning As Boolean
    tbOffset As Boolean
    
    tbFollowScrolling As Boolean
    tbToolbarSize As Boolean
    
    '----------
    
    gColor As Boolean
    gFont As Boolean
    gCursor As Boolean
    gImage As Boolean
    gLeading As Boolean
    gMargins As Boolean
    gSFX As Boolean
    gBorders As Boolean
    
    '----------
    
    cColor As Boolean
    cFont As Boolean
    cCursor As Boolean
    cImage As Boolean
End Type

Private Type SelItemsDef
    SelCmd As Integer
    SelCmdProp As SelPropertiesDef
    SelGrp As Integer
    SelGrpProp As SelPropertiesDef
    SelSubCmd As Integer
    SelSubCmdProp As SelPropertiesDef
    SelTB As Integer
    SelTBProp As SelPropertiesDef
End Type

Private SelItemProp As SelItemsDef
Private IsUpdating As Boolean

Private FloodPanel As New clsFlood

Private Sub cmdCancel_Click()

    tmrCloseDlg.Enabled = False

    CleanPresetsDirs
    Unload Me

End Sub

Private Sub cmdApply_Click()

    ApplyStyle
    Unload Me

End Sub

Private Sub Form_Load()

    With Screen
        Move (.Width - Width) \ 2, (.Height - Height) \ 2, Width, Height
    End With
    
    lblTitle.Caption = ""
    lblAuthor.Caption = ""
    txtComments.Text = ""
    
    wbPreview.Navigate "about:blank"
    Set FloodPanel.PictureControl = picFlood
    tmrInit.Enabled = True

End Sub

Private Sub tmrCloseDlg_Timer()

    cmdCancel_Click

End Sub

Private Sub tmrInit_Timer()

    tmrInit.Enabled = False
    
    LoadPresets

End Sub

Private Sub tvPresets_NodeClick(ByVal Node As MSComctlLib.Node)

    If Node.key = "WebOCP" Then
        If MsgBox("Are you sure you want to download all available presets from the Online Catalog?", vbQuestion + vbYesNo, "Online Catalog") = vbYes Then
            tvPresets.Nodes.Remove Node.Index
            
            If tvPresets.Nodes.Count > 0 Then
                With tvPresets.Nodes(1)
                    .EnsureVisible
                    .Selected = True
                End With
            End If
        End If
    Else
        If Node.Parent Is Nothing Then
            CleanPresetsDirs
            wbPreview.Navigate "about:blank"
            SetControlsState False
            UpdateInfo Nothing
        Else
            ShowPresetPreview Node
            SetControlsState True
            UpdateInfo Node
                
            Prj = GetProjectProperties(tvPresets.SelectedItem.Tag)
            LoadMenuItems
            Close #ff
        End If
    End If
    
End Sub

Private Sub LoadPresets(Optional ByVal ForceSelection As String)

    Dim i As Integer
    Dim dFile As String
    Dim dPath As String
    Dim nNode As Node
    Dim pNode As Node
    Dim sCat As String
    
    On Error Resume Next
    
    Form_Resize
    
    Screen.MousePointer = vbArrowHourglass
    Me.Enabled = False
    
    lblAuthor.Caption = ""
    txtComments.Text = ""
    
    wbPreview.Navigate "about:blank"
    tvPresets.Nodes.Clear
    
    dPath = AppPath + "Presets\"
    With flbPresets
        .Path = dPath
        .Pattern = "*.dpp"
        .Refresh
    End With
    
    DoEvents
    
    FloodPanel.Caption = "Please wait... Loading Available Presets"
    
    For i = 0 To flbPresets.ListCount - 1
        FloodPanel.Value = i / (flbPresets.ListCount - 1) * 100
        dFile = dPath + flbPresets.List(i)
        
        Set nNode = Nothing
        
        sCat = GetPresetProperty(dFile, piCategory)
        If sCat = "" Then sCat = "(uncategorized)"
        
        Set nNode = tvPresets.Nodes("K" + sCat)
        If nNode Is Nothing Then
            Set nNode = tvPresets.Nodes.Add(, , "K" + sCat, sCat, 1)
            nNode.Expanded = True
        End If
        
        Set pNode = tvPresets.Nodes.Add(nNode, tvwChild, , GetPresetProperty(dFile, piTitle), 2)
        If PresetWorkingMode = pwmSubmit Then
            If GetPresetProperty(dFile, piAuthor) <> USER Then
                'pNode.ForeColor = &H80000011
            End If
        End If
        
        If ForceSelection = pNode.Text Then
            pNode.EnsureVisible
            pNode.Selected = True
        End If
    Next i
    
    CleanPresetsDirs
    
    If ForceSelection <> "" Then
        tvPresets_NodeClick pNode
    Else
        If tvPresets.Nodes.Count > 0 Then
        With tvPresets.Nodes(1)
                .Selected = True
                .EnsureVisible
        End With
            tvPresets_NodeClick tvPresets.Nodes(1)
        End If
    End If
    
    FloodPanel.Value = 0
    Me.Enabled = True
    Screen.MousePointer = vbDefault

End Sub

Private Sub LoadMenuItems()

    Dim sStr As String
    Dim i As Integer
    
    ReDim mg(0)
    ReDim mc(0)
    ReDim sc(0)

    If (LOF(ff) = Loc(ff)) Then Exit Sub
    Line Input #ff, sStr
    Do Until (LOF(ff) = Loc(ff)) Or sStr = "[RSC]"
        If sStr <> "" Then Add2Array mg, Mid$(sStr, 4)
        Do Until (LOF(ff) = Loc(ff)) Or sStr = "[RSC]"
            Line Input #ff, sStr
            If sStr <> "" Then
                If Left$(sStr, 3) = "[C]" Then
                    If InStr(Mid$(sStr, 6), "[SEP]") = 0 Then Add2Array mc, Mid$(sStr, 6)
                Else
                    Exit Do
                End If
            End If
        Loop
        If (LOF(ff) = Loc(ff)) Then Exit Do
    Loop
    
ReStart:
    For i = 1 To UBound(mc)
        If Val(GetParam(mc(i), 39)) = atcCascade Or Val(GetParam(mc(i), 43)) = atcCascade Or Val(GetParam(mc(i), 47)) = atcCascade Then
            sStr = mc(i)
            Add2Array sc, sStr
            DelFromArray mc, sStr
            GoTo ReStart
        End If
    Next i
    
    With SelItemProp
        .SelTB = 0
        .SelGrp = 0
        .SelCmd = 0
        
        .SelTB = IIf(UBound(Prj.Toolbars) > 0, 1, 0)
        
        If .SelTB > 0 Then
            For i = 1 To UBound(mg)
                If GetParam(mg(i), 43) <> "" Then
                    .SelGrp = i
                    Exit For
                End If
            Next i
        End If
        If .SelGrp = 0 Then .SelGrp = IIf(UBound(mg) > 0, 1, 0)
        
        For i = 1 To UBound(mc)
            If GetParam(mc(i), 2) <> "" Then
                .SelCmd = i
                Exit For
            End If
        Next i
        If .SelCmd = 0 Then .SelCmd = IIf(UBound(mc) > 0, 1, 0)
        
        If .SelCmd > 0 Then .SelGrp = GetParam(mc(.SelCmd), 19)
        
        .SelSubCmd = IIf(UBound(sc) > 0, 1, 0)
        
        With .SelTBProp
            .cColor = True
            .cCursor = True
            .cFont = True
            .cImage = True
            .gColor = True
            .gCursor = True
            .gFont = True
            .gImage = True
            .gLeading = True
            .gMargins = True
            .gBorders = True
            .gSFX = True
            .tbAlignment = True
            .tbBackColor = True
            .tbBackImage = True
            .tbBorder = True
            .tbFollowScrolling = True
            .tbJustification = True
            .tbMargins = True
            .tbOffset = True
            .tbSeparation = True
            .tbSpanning = True
            .tbToolbarSize = True
            .tbToolbarStyle = True
        End With
        .SelCmdProp = .SelTBProp
        .SelGrpProp = .SelTBProp
        .SelSubCmdProp = .SelTBProp
    End With

End Sub

Private Sub Add2Array(a() As String, s As String)

    ReDim Preserve a(UBound(a) + 1)
    a(UBound(a)) = s

End Sub

Private Sub DelFromArray(a() As String, s As String)

    Dim i As Integer
    Dim j As Integer

    For i = 0 To UBound(a)
        If a(i) = s Then
            For j = i To UBound(a) - 1
                a(j) = a(j + 1)
            Next j
            ReDim Preserve a(UBound(a) - 1)
            Exit Sub
        End If
    Next i

End Sub

Private Sub SetControlsState(State As Boolean)

    cmdApply.Enabled = State
    
End Sub

Private Sub Form_Resize()

    Dim rHeight As Long

    On Error Resume Next
    
    rHeight = (Height - GetClientTop(Me.hWnd))

    tvPresets.Height = rHeight - (tvPresets.Top + frameInfo.Height + 900)
    frameInfo.Top = tvPresets.Top + tvPresets.Height + 45
    
    wbPreview.Move tvPresets.Width + 120, 300, Width - (tvPresets.Left + tvPresets.Width + 270), rHeight - 1155
                    
    cmdApply.Move Width - 2250, wbPreview.Top + wbPreview.Height + 120
    cmdCancel.Move Width - 1110, wbPreview.Top + wbPreview.Height + 120
    
    lblPreview.Left = wbPreview.Left
    
    picFlood.Move 30, sbInfo.Top + 60, sbInfo.Width - 470, sbInfo.Height - 75

End Sub

Private Sub UpdateInfo(sNode As Node)

    Dim dFile As String

    If sNode Is Nothing Then
        lblAuthor.Caption = ""
        txtComments.Text = ""
        lblTitle.Caption = "No Preset Selected"
    Else
        dFile = AppPath + "Presets\" + sNode.Text + ".dpp"
        lblAuthor.Caption = GetPresetProperty(dFile, piAuthor)
        txtComments.Text = GetPresetProperty(dFile, piComments)

        lblTitle.Caption = sNode.Text
    End If
    
End Sub

Private Sub ShowPresetPreview(nNode As Node)

    Dim ff As Integer
    Dim dFile As String
    Dim pName As String
    Dim pkey As String
    Dim pid As String
    
    Screen.MousePointer = vbArrowHourglass
    
    wbPreview.Navigate "about:blank"
    Do
        DoEvents
    Loop While wbPreview.Busy And (wbPreview.ReadyState <> READYSTATE_COMPLETE)
        
    pName = nNode.Text
    UncompressPreset pName
    
    dFile = TempPath + "Presets\tmp\index.html"
    
    ff = FreeFile
    Open dFile For Output As #ff
        Print #ff, "<html><body>"
        Print #ff, "<script language=JavaScript src=menu.js></script>"
        Print #ff, "</body></html>"
    Close #ff
    
    tvPresets.SelectedItem.Tag = GetFilePath(dFile) + Dir(GetFilePath(dFile) + "*.dmb")
    
    wbPreview.Navigate dFile
    
    Screen.MousePointer = vbDefault

End Sub

Private Sub ApplyStyle()

    Dim i As Integer
    Dim sTB As ToolbarDef
    Dim sgI As MenuGrp
    Dim scI As MenuCmd
    Dim sPath As String
    
    On Error Resume Next
    
    Me.Enabled = False
    IsReplacing = True
    
    sPath = TempPath + "Presets\tmp\"
    
    With Project
        With .MenusOffset
            .RootMenusX = Prj.MenusOffset.RootMenusX
            .RootMenusY = Prj.MenusOffset.RootMenusY
            .SubMenusX = Prj.MenusOffset.SubMenusX
            .SubMenusY = Prj.MenusOffset.SubMenusY
        End With
        
        .AnimSpeed = Prj.AnimSpeed
        .DXFilter = Prj.DXFilter
        .FontSubstitutions = Prj.FontSubstitutions
        .FX = Prj.FX
        .HideDelay = Prj.HideDelay
        .RootMenusDelay = Prj.RootMenusDelay
        .SubMenusDelay = Prj.SubMenusDelay
    End With

    With SelItemProp
        If .SelTB > 0 Then
            sTB = Prj.Toolbars(.SelTB)
            With .SelTBProp
                For i = 1 To UBound(Project.Toolbars)
                    If .tbAlignment Then
                        Project.Toolbars(i).Alignment = sTB.Alignment
                        Project.Toolbars(i).CustX = sTB.CustX
                        Project.Toolbars(i).CustY = sTB.CustY
                        Project.Toolbars(i).AttachTo = sTB.AttachTo
                        Project.Toolbars(i).AttachToAlignment = sTB.AttachToAlignment
                    End If
                    If .tbBackColor Then Project.Toolbars(i).BackColor = sTB.BackColor
                    If .tbBackImage Then Project.Toolbars(i).Image = FixImagePath(sTB.Image)
                    If .tbBorder Then
                        Project.Toolbars(i).border = sTB.border
                        Project.Toolbars(i).BorderColor = sTB.BorderColor
                        Project.Toolbars(i).BorderStyle = sTB.BorderStyle
                    End If
                    If .tbFollowScrolling Then
                        Project.Toolbars(i).FollowHScroll = sTB.FollowHScroll
                        Project.Toolbars(i).FollowVScroll = sTB.FollowVScroll
                    End If
                    If .tbJustification Then Project.Toolbars(i).JustifyHotSpots = sTB.JustifyHotSpots
                    If .tbMargins Then
                        Project.Toolbars(i).ContentsMarginH = sTB.ContentsMarginH
                        Project.Toolbars(i).ContentsMarginV = sTB.ContentsMarginV
                    End If
                    If .tbOffset Then
                        Project.Toolbars(i).OffsetH = sTB.OffsetH
                        Project.Toolbars(i).OffsetV = sTB.OffsetV
                    End If
                    If .tbSeparation Then Project.Toolbars(i).Separation = sTB.Separation
                    If .tbSpanning Then Project.Toolbars(i).Spanning = sTB.Spanning
                    If .tbToolbarSize Then
                         Project.Toolbars(i).Width = sTB.Width
                         Project.Toolbars(i).Height = sTB.Height
                    End If
                    If .tbToolbarStyle Then Project.Toolbars(i).Style = sTB.Style
                Next i
            End With
            
            If .SelGrp > 0 Then
                mg(.SelGrp) = "templateGroup" + Left(mg(.SelGrp), InStr(mg(.SelGrp), cSep) - 1) + Mid(mg(.SelGrp), InStr(mg(.SelGrp), cSep))
                AddMenuGroup mg(.SelGrp), True
                sgI = MenuGrps(UBound(MenuGrps))
                With .SelGrpProp
                    For i = 1 To UBound(MenuGrps) - 1
                        MenuGrps(i).Alignment = sgI.Alignment
                        If .gColor Then
                            MenuGrps(i).nTextColor = sgI.nTextColor
                            MenuGrps(i).nBackColor = sgI.nBackColor
                            MenuGrps(i).hTextColor = sgI.hTextColor
                            MenuGrps(i).hBackColor = sgI.hBackColor
                            MenuGrps(i).bColor = sgI.bColor
                        End If
                        If .gFont Then
                            MenuGrps(i).DefNormalFont = sgI.DefNormalFont
                            MenuGrps(i).DefHoverFont = sgI.DefHoverFont
                            MenuGrps(i).CaptionAlignment = sgI.CaptionAlignment
                        End If
                        If .gCursor Then MenuGrps(i).iCursor = sgI.iCursor
                        If .gImage Then
                            MenuGrps(i).tbiLeftImage = FixImage(sgI.tbiLeftImage)
                            MenuGrps(i).tbiRightImage = FixImage(sgI.tbiRightImage)
                            MenuGrps(i).tbiBackImage = FixImage(sgI.tbiBackImage)
                            MenuGrps(i).BackImage = FixImage(sgI.BackImage)
                            MenuGrps(i).Image = FixImagePath(sgI.Image)
                        End If
                        If .gBorders Then
                            MenuGrps(i).frameBorder = sgI.frameBorder
                            MenuGrps(i).BorderStyle = sgI.BorderStyle
                            MenuGrps(i).Corners = sgI.Corners
                            MenuGrps(i).CornersImages.gcBottomCenter = FixImagePath(sgI.CornersImages.gcBottomCenter)
                            MenuGrps(i).CornersImages.gcBottomLeft = FixImagePath(sgI.CornersImages.gcBottomLeft)
                            MenuGrps(i).CornersImages.gcBottomRight = FixImagePath(sgI.CornersImages.gcBottomRight)
                            MenuGrps(i).CornersImages.gcLeft = FixImagePath(sgI.CornersImages.gcLeft)
                            MenuGrps(i).CornersImages.gcRight = FixImagePath(sgI.CornersImages.gcRight)
                            MenuGrps(i).CornersImages.gcTopCenter = FixImagePath(sgI.CornersImages.gcTopCenter)
                            MenuGrps(i).CornersImages.gcTopLeft = FixImagePath(sgI.CornersImages.gcTopLeft)
                            MenuGrps(i).CornersImages.gcTopRight = FixImagePath(sgI.CornersImages.gcTopRight)
                        End If
                        If .gMargins Then
                            MenuGrps(i).ContentsMarginH = sgI.ContentsMarginH
                            MenuGrps(i).ContentsMarginV = sgI.ContentsMarginV
                        End If
                        If .gLeading Then MenuGrps(i).Leading = sgI.Leading
                        If .gSFX Then
                            MenuGrps(i).CmdsFXnColor = sgI.CmdsFXnColor
                            MenuGrps(i).CmdsFXhColor = sgI.CmdsFXhColor
                            MenuGrps(i).CmdsFXNormal = sgI.CmdsFXNormal
                            MenuGrps(i).CmdsFXOver = sgI.CmdsFXOver
                            MenuGrps(i).CmdsFXSize = sgI.CmdsFXSize
                            MenuGrps(i).CmdsMarginX = sgI.CmdsMarginX
                            MenuGrps(i).CmdsMarginY = sgI.CmdsMarginY
                            MenuGrps(i).fWidth = sgI.fWidth
                            MenuGrps(i).fHeight = sgI.fHeight
                        End If
                    Next i
                End With
                ReDim Preserve MenuGrps(UBound(MenuGrps) - 1)
            End If
            
            If .SelCmd > 0 Then
                mc(.SelCmd) = "templateCommand" + Left(mc(.SelCmd), InStr(mc(.SelCmd), cSep) - 1) + Mid(mc(.SelCmd), InStr(mc(.SelCmd), cSep))
                set_scI_Params scI, mc(.SelCmd)
                With .SelCmdProp
                    For i = 1 To UBound(MenuCmds)
                        If MenuCmds(i).Name <> "[SEP]" Then
                            If .cColor Then
                                MenuCmds(i).nTextColor = scI.nTextColor
                                MenuCmds(i).nBackColor = scI.nBackColor
                                MenuCmds(i).hTextColor = scI.hTextColor
                                MenuCmds(i).hBackColor = scI.hBackColor
                            End If
                            If .cFont Then
                                MenuCmds(i).NormalFont = scI.NormalFont
                                MenuCmds(i).HoverFont = scI.HoverFont
                                MenuCmds(i).Alignment = scI.Alignment
                            End If
                            If .cCursor Then MenuCmds(i).iCursor = scI.iCursor
                            If .cImage Then
                                MenuCmds(i).LeftImage = FixImage(scI.LeftImage)
                                MenuCmds(i).BackImage = FixImage(scI.BackImage)
                                MenuCmds(i).RightImage = FixImage(scI.RightImage)
                            End If
                            MenuCmds(i).CmdsFXnColor = scI.CmdsFXnColor
                            MenuCmds(i).CmdsFXhColor = scI.CmdsFXhColor
                            MenuCmds(i).CmdsFXNormal = scI.CmdsFXNormal
                            MenuCmds(i).CmdsFXOver = scI.CmdsFXOver
                            MenuCmds(i).CmdsFXSize = scI.CmdsFXSize
                            MenuCmds(i).CmdsMarginX = scI.CmdsMarginX
                            MenuCmds(i).CmdsMarginY = scI.CmdsMarginY
                        End If
                    Next i
                End With
            End If
            
            If .SelSubCmd > 0 Then
                sc(.SelSubCmd) = "templateSubCommand" + Left(sc(.SelSubCmd), InStr(sc(.SelSubCmd), cSep) - 1) + Mid(sc(.SelSubCmd), InStr(sc(.SelSubCmd), cSep))
                set_scI_Params scI, sc(.SelSubCmd)
                With .SelSubCmdProp
                    For i = 1 To UBound(MenuCmds)
                        If MenuCmds(i).Actions.onclick.Type = atcCascade Or MenuCmds(i).Actions.onmouseover.Type = atcCascade Or MenuCmds(i).Actions.OnDoubleClick.Type = atcCascade Then
                            MenuCmds(i).Actions.onclick.TargetMenuAlignment = scI.Actions.onclick.TargetMenuAlignment
                            MenuCmds(i).Actions.onmouseover.TargetMenuAlignment = scI.Actions.onmouseover.TargetMenuAlignment
                            MenuCmds(i).Actions.OnDoubleClick.TargetMenuAlignment = scI.Actions.OnDoubleClick.TargetMenuAlignment
                            If .cColor Then
                                MenuCmds(i).nTextColor = scI.nTextColor
                                MenuCmds(i).nBackColor = scI.nBackColor
                                MenuCmds(i).hTextColor = scI.hTextColor
                                MenuCmds(i).hBackColor = scI.hBackColor
                            End If
                            If .cFont Then
                                MenuCmds(i).NormalFont = scI.NormalFont
                                MenuCmds(i).HoverFont = scI.HoverFont
                                MenuCmds(i).Alignment = scI.Alignment
                            End If
                            If .cCursor Then MenuCmds(i).iCursor = scI.iCursor
                            If .cImage Then
                                MenuCmds(i).LeftImage = FixImage(scI.LeftImage)
                                MenuCmds(i).BackImage = FixImage(scI.BackImage)
                                MenuCmds(i).RightImage = FixImage(scI.RightImage)
                                
                                If scI.LeftImage.NormalImage = "" Then
                                    If Right(scI.Caption, 2) <> " »" Then
                                        scI.Caption = scI.Caption + " »"
                                    End If
                                Else
                                    If Right(scI.Caption, 2) = " »" Then
                                        scI.Caption = Left(scI.Caption, Len(scI.Caption) - 2)
                                    End If
                                End If
                            End If
                            MenuCmds(i).CmdsFXnColor = scI.CmdsFXnColor
                            MenuCmds(i).CmdsFXhColor = scI.CmdsFXhColor
                            MenuCmds(i).CmdsFXNormal = scI.CmdsFXNormal
                            MenuCmds(i).CmdsFXOver = scI.CmdsFXOver
                            MenuCmds(i).CmdsFXSize = scI.CmdsFXSize
                            MenuCmds(i).CmdsMarginX = scI.CmdsMarginX
                            MenuCmds(i).CmdsMarginY = scI.CmdsMarginY
                        End If
                    Next i
                End With
            End If
        End If
    End With
    
    IsReplacing = False
    Me.Enabled = True
    
End Sub

Private Sub set_scI_Params(c As MenuCmd, params As String)

    Dim cVer As Long
    cVer = CLng(Prj.version)

    With c
        .Caption = GetParam(params, 2)
        .WinStatus = GetParam(params, 21)
        .iCursor.cType = GetParam(params, 18)
        .iCursor.cFile = GetParam(params, 20)
        
        'FOR COMPATIBILITY
        If GetParam(params, 17) <> "" Or Val(GetParam(params, 23)) > 0 Then 'Is old Format
            Select Case Val(GetParam(params, 20))
                Case ByClicking
                    If -Val(GetParam(params, 22)) Then
                        .Actions.onclick.TargetMenu = Val(GetParam(params, 23))
                        .Actions.onclick.Type = atcCascade
                    Else
                        .Actions.onclick.url = GetParam(params, 17)
                        .Actions.onclick.Type = atcURL
                    End If
                Case ByHovering
                    If -Val(GetParam(params, 22)) Then
                        .Actions.onmouseover.TargetMenu = Val(GetParam(params, 23))
                        .Actions.onmouseover.Type = atcCascade
                    Else
                        .Actions.onmouseover.url = GetParam(params, 17)
                        .Actions.onmouseover.Type = atcURL
                    End If
            End Select
        End If
            
        .Alignment = Val(GetParam(params, 29))
        '.TargetFrame = GetParam(Params, 31)
        .disabled = -Val(GetParam(params, 34))
        .hBackColor = Val(GetParam(params, 3))
        .hTextColor = Val(GetParam(params, 4))
        .nBackColor = Val(GetParam(params, 5))
        .nTextColor = Val(GetParam(params, 6))
        .HoverFont.FontName = GetParam(params, 7)
        .HoverFont.FontSize = Val(GetParam(params, 8))
        .HoverFont.FontBold = -Val(GetParam(params, 9))
        .HoverFont.FontItalic = -Val(GetParam(params, 10))
        .HoverFont.FontUnderline = -Val(GetParam(params, 11))
        .NormalFont.FontName = GetParam(params, 12)
        .NormalFont.FontSize = Val(GetParam(params, 13))
        .NormalFont.FontBold = -Val(GetParam(params, 14))
        .NormalFont.FontItalic = -Val(GetParam(params, 15))
        .NormalFont.FontUnderline = -Val(GetParam(params, 16))
        .LeftImage.NormalImage = GetParam(params, 24)
        .LeftImage.HoverImage = GetParam(params, 25)
        .LeftImage.w = Val(GetParam(params, 27))
        .LeftImage.h = Val(GetParam(params, 28))
        .RightImage.NormalImage = GetParam(params, 35)
        .RightImage.HoverImage = GetParam(params, 36)
        .RightImage.w = Val(GetParam(params, 37))
        .RightImage.h = Val(GetParam(params, 38))
        .BackImage.NormalImage = GetParam(params, 51)
        .BackImage.HoverImage = GetParam(params, 52)
        .BackImage.Tile = Val(GetParam(params, 67))
        .BackImage.AllowCrop = Val(GetParam(params, 68))
        If Val(GetParam(params, 26)) = 1 Then
            'Upgrade from old versions
            .RightImage.NormalImage = .LeftImage.NormalImage
            .RightImage.HoverImage = .LeftImage.HoverImage
            .RightImage.w = .LeftImage.w
            .RightImage.h = .LeftImage.h
            .LeftImage.NormalImage = ""
            .LeftImage.HoverImage = ""
            .LeftImage.w = 0
            .LeftImage.h = 0
            .BackImage.AllowCrop = True
            .BackImage.Tile = True
        End If
        If .RightImage.NormalImage = "" Then .RightImage.w = 0: .RightImage.h = 0
        If .LeftImage.NormalImage = "" Then .LeftImage.w = 0: .LeftImage.h = 0
        
        If cVer < 420000 Then
            .BackImage.AllowCrop = True
            .BackImage.Tile = True
        End If
        
        'Check if this its an old project
        'to avoid overwriting stuff...
        If GetParam(params, 17) = "" And Val(GetParam(params, 23)) = 0 Then 'Is not old Format
        With .Actions.onclick
            .Type = Val(GetParam(params, 39))
            .url = GetParam(params, 40)
            .TargetFrame = GetParam(params, 41)
            .TargetMenu = Val(GetParam(params, 42))
            .WindowOpenParams = Replace(GetParam(params, 53), "|", cSep)
            .TargetMenuAlignment = Val(GetParam(params, 31))
        End With
        With .Actions.onmouseover
            .Type = Val(GetParam(params, 43))
            .url = GetParam(params, 44)
            .TargetFrame = GetParam(params, 45)
            .TargetMenu = Val(GetParam(params, 46))
            .WindowOpenParams = Replace(GetParam(params, 54), "|", cSep)
            .TargetMenuAlignment = Val(GetParam(params, 32))
        End With
        With .Actions.OnDoubleClick
            .Type = Val(GetParam(params, 47))
            .url = GetParam(params, 48)
            .TargetFrame = GetParam(params, 49)
            .TargetMenu = Val(GetParam(params, 50))
            .WindowOpenParams = Replace(GetParam(params, 55), "|", cSep)
            .TargetMenuAlignment = Val(GetParam(params, 33))
        End With
        End If
        
        .Sound.onmouseover = GetParam(params, 56)
        .Sound.onclick = GetParam(params, 57)
        
        .CmdsFXhColor = Val(GetParam(params, 59))
        .CmdsFXnColor = Val(GetParam(params, 60))
        .CmdsFXNormal = Val(GetParam(params, 61))
        .CmdsFXOver = Val(GetParam(params, 62))
        .CmdsFXSize = Val(GetParam(params, 63))
        .CmdsMarginX = Val(GetParam(params, 64))
        .CmdsMarginY = Val(GetParam(params, 65))
    End With

End Sub

Private Function FixImagePath(imgPath As String) As String

    If imgPath <> "" Then
        If FileExists(imgPath) Then
            FixImagePath = imgPath
        Else
            FixImagePath = TempPath + "Presets\tmp\" + GetFileName(imgPath)
        End If
    End If
    
End Function

Private Function FixImage(img As tImage) As tImage
    
    img.NormalImage = FixImagePath(img.NormalImage)
    img.HoverImage = FixImagePath(img.HoverImage)
    
    FixImage = img

End Function
