VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{DBF30C82-CAF3-11D5-84FF-0050BA3D926D}#8.5#0"; "VLMnuPlus.ocx"
Begin VB.Form frmPresetsManager 
   Caption         =   "Presets Manager"
   ClientHeight    =   7680
   ClientLeft      =   825
   ClientTop       =   3960
   ClientWidth     =   13845
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPresets.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7680
   ScaleWidth      =   13845
   Begin VB.Timer tmrCloseDlg 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   7080
      Top             =   5610
   End
   Begin InetCtlsObjects.Inet ftpCtrl 
      Left            =   9375
      Top             =   5130
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      Protocol        =   2
      RemotePort      =   21
      URL             =   "ftp://"
   End
   Begin VB.Timer tmrInit 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   2910
      Top             =   5100
   End
   Begin VB.PictureBox picFlood 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   810
      ScaleHeight     =   255
      ScaleWidth      =   3270
      TabIndex        =   17
      Top             =   7245
      Visible         =   0   'False
      Width           =   3270
   End
   Begin MSComctlLib.ImageList ilIcons 
      Left            =   1155
      Top             =   5355
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
            Picture         =   "frmPresets.frx":038A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresets.frx":0A5C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresets.frx":0DF6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picImportOptions 
      Height          =   3315
      Left            =   8310
      ScaleHeight     =   3255
      ScaleWidth      =   5265
      TabIndex        =   6
      Top             =   1020
      Visible         =   0   'False
      Width           =   5325
      Begin VB.ComboBox cmbItems 
         Enabled         =   0   'False
         Height          =   315
         Left            =   45
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   60
         Width           =   2175
      End
      Begin MSComctlLib.TreeView tvProperties 
         Height          =   2775
         Left            =   45
         TabIndex        =   8
         Top             =   450
         Width           =   5190
         _ExtentX        =   9155
         _ExtentY        =   4895
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   353
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         Checkboxes      =   -1  'True
         Appearance      =   1
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin MSComctlLib.TabStrip tsImportTabs 
      Height          =   3750
      Left            =   8235
      TabIndex        =   5
      Top             =   660
      Visible         =   0   'False
      Width           =   5460
      _ExtentX        =   9631
      _ExtentY        =   6615
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Toolbar"
            Key             =   "tbToolbar"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Group"
            Key             =   "tbGroup"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Cascading"
            Key             =   "tbSubCommand"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Command"
            Key             =   "tbCommand"
            ImageVarType    =   2
         EndProperty
      EndProperty
      Enabled         =   0   'False
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
      TabIndex        =   9
      Top             =   2190
      Width           =   3105
      Begin VB.TextBox txtComments 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   1380
         Left            =   90
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   14
         Top             =   1005
         Width           =   2925
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         BackColor       =   &H80000002&
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
         TabIndex        =   10
         Top             =   0
         Width           =   3000
      End
      Begin VB.Label lblAuthor 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   90
         TabIndex        =   12
         Top             =   495
         Width           =   45
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
         TabIndex        =   13
         Top             =   780
         Width           =   960
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
         TabIndex        =   11
         Top             =   285
         Width           =   630
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6930
      TabIndex        =   16
      Top             =   4755
      Width           =   900
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   5910
      TabIndex        =   15
      Top             =   4755
      Width           =   900
   End
   Begin MSComctlLib.StatusBar sbInfo 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   18
      Top             =   7365
      Width           =   13845
      _ExtentX        =   24421
      _ExtentY        =   556
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.FileListBox flbPresets 
      Height          =   1065
      Left            =   4005
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   135
      Visible         =   0   'False
      Width           =   1320
   End
   Begin SHDocVwCtl.WebBrowser wbPreview 
      Height          =   4350
      Left            =   3255
      TabIndex        =   4
      Top             =   300
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
      TabIndex        =   3
      Top             =   300
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
   Begin VLMnuPlus.VLMenuPlus vlmCtrl 
      Left            =   4305
      Top             =   5250
      _ExtentX        =   847
      _ExtentY        =   847
      _CXY            =   4
      _CGUID          =   43495.5104050926
      Language        =   0
   End
   Begin VB.Label lblPreview 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Preview"
      Height          =   195
      Left            =   2955
      TabIndex        =   1
      Top             =   60
      Width           =   570
   End
   Begin VB.Label lblPresets 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Presets"
      Height          =   195
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   540
   End
   Begin VB.Menu mnuPresets 
      Caption         =   "Preset"
      Begin VB.Menu ctxMenuLoad 
         Caption         =   "Load"
      End
      Begin VB.Menu ctxMenuSep01 
         Caption         =   "-"
      End
      Begin VB.Menu ctxMenuEdit 
         Caption         =   "Edit..."
      End
      Begin VB.Menu ctxMenuSep02 
         Caption         =   "-"
      End
      Begin VB.Menu ctxMenuDelete 
         Caption         =   "Delete"
      End
      Begin VB.Menu ctxMenuSep023 
         Caption         =   "-"
      End
      Begin VB.Menu ctxMenuRepack 
         Caption         =   "Repack"
      End
      Begin VB.Menu ctxMenuRepackAll 
         Caption         =   "Repack All"
      End
   End
End
Attribute VB_Name = "frmPresetsManager"
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
Private xMenu As CMenu
Private SubmitRes As String

Private FloodPanel As New clsFlood
Private UseOnlineCatalog As Boolean

Private ForceCancel As Boolean
Private FTPIsBusy As Boolean

Private Sub cmbItems_Click()

    If IsUpdating Then Exit Sub
    UpdatePropertiesSelection
    
    tvProperties.Enabled = cmbItems.ListIndex > 0

End Sub

Private Sub cmdCancel_Click()

    ForceCancel = True
    If FTPIsBusy Then
        tmrCloseDlg.Enabled = True
        Exit Sub
    End If
    tmrCloseDlg.Enabled = False

    CleanPresetsDirs
    Unload Me

End Sub

Private Sub cmdLoad_Click()

    ForceCancel = False

    Select Case PresetWorkingMode
        Case pwmApplyStyle
            ApplyStyle
            Unload Me
        Case pwmNormal
            If LoadSelectedPreset Then
                frmMain.UpdateTitleBar
                Me.Visible = False
                If Err.number = 0 Then frmMain.FileSaveAs
                Unload Me
            End If
        Case pwmSubmit
            If tvPresets.SelectedItem.ForeColor = &H80000011 Then
                MsgBox GetLocalizedStr(912), vbInformation + vbOKOnly, GetLocalizedStr(900)
                Exit Sub
            Else
                SubmitPreset
                If Err.number Then MsgBox "Unexpected Error " & Err.number & ": " & Err.Description
            End If
            Unload Me
    End Select

End Sub

Private Sub SubmitPreset()

    Dim sPath As String
    Dim i As Integer
    Dim fName As String
    
    On Error GoTo ReportError
    
    If Not IsOnline Then
        MsgBox GetLocalizedStr(901), vbInformation + vbOKOnly, GetLocalizedStr(900)
        Do
            DoEvents
        Loop While wbPreview.Busy And (wbPreview.ReadyState <> READYSTATE_COMPLETE)
    End If
    
    If MsgBox(GetLocalizedStr(902) + " '" + tvPresets.SelectedItem.Text + "' " + GetLocalizedStr(903) + vbCrLf + vbCrLf + GetLocalizedStr(548), vbQuestion + vbYesNo, GetLocalizedStr(904)) = vbNo Then
        Exit Sub
    End If
    
    tvPresets.Enabled = False
    mnuPresets.Enabled = False
    
    SetControlsState False
    FTPIsBusy = True
    
    sPath = AppPath + "Presets\"
    
    SaveFile TempPath + "Presets\tmp\rsc.txt", tvPresets.SelectedItem.Text + Chr(1) + _
                                    lblAuthor.caption + Chr(1) + _
                                    tvPresets.SelectedItem.parent.Text + Chr(1) + _
                                    txtComments.Text
    With ftpCtrl
        .AccessType = icUseDefault
        .protocol = icFTP
        
        ' CREATE directory based on user's unlock code
        FloodPanel.caption = GetLocalizedStr(905)
        FloodPanel.Value = 1
'        .Execute "ftp://pmanup:pmanup@software.xfx.net/", _
'                    "MKDIR /web/" + USERSN
        Do While .StillExecuting And Not ForceCancel
            DoEvents
        Loop
        If ForceCancel Then GoTo AbortFTP
        
        FloodPanel.caption = GetLocalizedStr(906)
        .Execute "ftp://xfx.net/", _
                    "SEND " + Long2Short(sPath + tvPresets.SelectedItem.Text + ".dpp") + " /incoming/" + USERSN + "_preset.dpp"
        
        ' UPLOAD all files used by the Preset
        With flbPresets
            .Path = TempPath + "Presets\tmp\"
            .Pattern = "*.*"
            .Refresh
        End With
        
        For i = 0 To flbPresets.ListCount - 1
            Do While .StillExecuting And Not ForceCancel
                DoEvents
            Loop
            If ForceCancel Then GoTo AbortFTP
            
            fName = flbPresets.List(i)
            If fName <> "hRef.txt" And GetFileExtension(fName) <> "dmb" Then
                If fName = "menu.js" Then
                    Compress TempPath + "Presets\tmp\" + fName, LoadFile(sPath + "tmp/" + fName)
                End If
                FloodPanel.Value = (i / (flbPresets.ListCount - 1)) * 100
                .Execute "ftp://xfx.net/", _
                    "SEND " + SetSlashDir(Long2Short(TempPath + "Presets\tmp\" + fName), sdFwd) + " /incoming/" + USERSN + "_" + fName
            End If
        Next i
        
        ' PROCESS the submission and add it to the database
        SubmitRes = ""
        FloodPanel.caption = GetLocalizedStr(907)
        FloodPanel.Value = 1
        wbPreview.Navigate EncodeUrl("https://xfx.net/utilities/dmbuilder/presets/submit.php?us=" & USERSN)
        Do While LenB(SubmitRes) = 0 And Not ForceCancel
            DoEvents
        Loop
        If ForceCancel Then GoTo AbortFTP
        
        wbPreview.Navigate "about:blank"
        
        ' DELETE all uploaded files
        Do While .StillExecuting And Not ForceCancel
            DoEvents
        Loop
        If ForceCancel Then GoTo AbortFTP
        
'        .Execute "ftp://pmanup:pmanup@software.xfx.net/", _
'                "DELETE /web/" + USERSN + "/preset.dpp"
'
'        For i = 0 To flbPresets.ListCount - 1
'            Do While .StillExecuting And Not ForceCancel
'                DoEvents
'            Loop
'            If ForceCancel Then GoTo AbortFTP
'
'            FloodPanel.Value = (i / (flbPresets.ListCount - 1)) * 100
'            .Execute "ftp://pmanup:pmanup@software.xfx.net/", _
'                "DELETE /web/" + USERSN + "/" + flbPresets.List(i)
'        Next i
        
        ' PROVIDE feedback to the user
        If SubmitRes = "DUPLICATE" Then
            MsgBox GetLocalizedStr(908), vbInformation + vbOKOnly, GetLocalizedStr(900)
        End If
        If Left(SubmitRes, 5) = "ERROR" Then
            MsgBox GetLocalizedStr(909) + vbCrLf + vbCrLf + SubmitRes, vbInformation + vbOKOnly, GetLocalizedStr(900)
        End If
        If SubmitRes = "UNDEFINED ERROR" Then
            MsgBox GetLocalizedStr(910) + vbCrLf + vbCrLf + "Error " & Err.number & ": " + Err.Description, vbInformation + vbOKOnly, GetLocalizedStr(900)
        End If
        If SubmitRes = "DONE" Then
            MsgBox GetLocalizedStr(911), vbInformation + vbOKOnly, GetLocalizedStr(913)
        End If
        'RunShellExecute "Open", EncodeUrl("https://xfx.net/utilities/dmbuilder/presets/submit.php?us=" & USERSN), 0, 0, 0
    End With
    
ExitSub:
    tvPresets.Enabled = True
    mnuPresets.Enabled = True
    FTPIsBusy = False
    SetControlsState True
    
    Exit Sub
    
AbortFTP:
    tvPresets.Enabled = True
    mnuPresets.Enabled = True
    ftpCtrl.Cancel
    DoEvents
    GoTo ExitSub
    
ReportError:
    MsgBox "Error " & Err.number & ": " & Err.Description, vbCritical, "Error Submiting Preset"
    GoTo AbortFTP
    
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
    
    frmMain.SaveState "Apply Style from " + tvPresets.SelectedItem.Text
    
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
        .SelChangeDelay = Prj.SelChangeDelay
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
                        Project.Toolbars(i).bOrder = sTB.bOrder
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
                frmMain.SelectItem frmMain.tvMenus.Nodes("G" & GetIDByName(sgI.Name)), True
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
                frmMain.RemoveItem True
            End If
            
            If .SelCmd > 0 Then
                mc(.SelCmd) = "templateCommand" + Left(mc(.SelCmd), InStr(mc(.SelCmd), cSep) - 1) + Mid(mc(.SelCmd), InStr(mc(.SelCmd), cSep))
                AddMenuCommand mc(.SelCmd), True, True
                scI = MenuCmds(UBound(MenuCmds))
                frmMain.SelectItem frmMain.tvMenus.Nodes("C" & GetIDByName(scI.Name)), True
                With .SelCmdProp
                    For i = 1 To UBound(MenuCmds) - 1
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
                frmMain.RemoveItem True
            End If
            
            If .SelSubCmd > 0 Then
                sc(.SelSubCmd) = "templateSubCommand" + Left(sc(.SelSubCmd), InStr(sc(.SelSubCmd), cSep) - 1) + Mid(sc(.SelSubCmd), InStr(sc(.SelSubCmd), cSep))
                AddMenuCommand sc(.SelSubCmd)
                scI = MenuCmds(UBound(MenuCmds))
                frmMain.SelectItem frmMain.tvMenus.Nodes("C" & GetIDByName(scI.Name)), True
                With .SelSubCmdProp
                    For i = 1 To UBound(MenuCmds) - 1
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
                frmMain.RemoveItem True
            End If
        End If
    End With
    
    If UBound(Project.Toolbars) > 0 Then AdjustMenusAlignment 1

    IsReplacing = False
    Me.Enabled = True
    
    frmMain.UpdateControls
    frmMain.UpdateLivePreview

End Sub

Private Function FixImagePath(imgPath As String) As String

    If LenB(imgPath) <> 0 Then
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

Private Function LoadSelectedPreset() As Boolean

    Dim sFile As String
    Dim tFile As String
    Dim b() As Byte
    
    On Error GoTo ExitSub
    
    If IsFromWeb(tvPresets.SelectedItem) Then
    
        FloodPanel.caption = "Downloading Preset from the Online Catalog..."
        FloodPanel.Value = 25
    
        sFile = tvPresets.SelectedItem.Text + ".dpp"
        tFile = AppPath + "Presets\" + sFile
        
        With ftpCtrl
            .AccessType = icUseDefault
            .protocol = icHTTP
            
            b = .OpenURL("https://xfx.net/~pmanup/" + Split(tvPresets.SelectedItem.tag, "|")(0) + "/" + sFile, icByteArray)
            
            FloodPanel.Value = 50
            
            Do While .StillExecuting Or ForceCancel
                DoEvents
            Loop
        End With
        
        Open tFile For Binary As #1
            Put #1, , b
        Close #1
        
        FloodPanel.Value = 0
        
        LoadPresets tvPresets.SelectedItem.Text
        LoadSelectedPreset = False
        Exit Function
    End If
    
    sFile = tvPresets.SelectedItem.tag
    tFile = TempPath + GetFileName(sFile)
    FileCopy sFile, tFile
    
    frmMain.LoadMenu tFile, , , True
    
    tFile = GetFileName(tFile)
    tFile = Left(tFile, Len(tFile) - Len(GetFileExtension(tFile)) - 1)
    
    With Project
        .Name = tFile
        ReDim .UserConfigs(0)
        With .UserConfigs(0)
            .Name = GetLocalizedStr(400)
            .Type = ctcCDROM
        End With
        .DefaultConfig = 0
        .CompilehRefFile = False
        .CompileIECode = True
        .CompileNSCode = True
        .DoFormsTweak = True
        .HasChanged = True
        .CodeOptimization = cocAggressive
    End With
    
    
ExitSub:
    CleanPresetsDirs
    Err.Clear
    LoadSelectedPreset = True

End Function

Private Sub FinalizeRepacking()

    Dim oPref As PrgPrefs

    CleanPresetsDirs
    
    Project.HasChanged = False
    oPref = Preferences
    Preferences.ShowPPOnNewProject = False
    frmMain.NewMenu
    Preferences = oPref

End Sub

Private Sub Repack(sNode As Node)

    Dim sFile As String
    
    If IsFromWeb(sNode) Then Exit Sub
    
    Me.Enabled = False

    sNode.Selected = True
    tvPresets_NodeClick sNode

    sFile = AppPath + "Presets\" + sNode.Text + ".dpp"
    UncompressPreset sNode.Text
    
    With flbPresets
        .Path = AppPath + "Presets\tmp\"
        .Pattern = "*.*"
        .Refresh
    End With
    
    Project.HasChanged = False
    LoadSelectedPreset
    
    CompressPreset GetPresetProperty(sFile, piTitle), _
                    GetPresetProperty(sFile, piAuthor), _
                    GetPresetProperty(sFile, piComments), _
                    GetPresetProperty(sFile, piCategory), _
                    flbPresets
                    
    Me.Enabled = True

End Sub

Private Sub ctxMenuDelete_Click()

    If IsFromWeb(tvPresets.SelectedItem) Then
        MsgBox "Online Presets cannot be deleted", vbInformation + vbOKOnly, "Error Deleting Preset"
    Else
        If MsgBox("Are you sure you want to completely remove the selected Preset from your system?", vbQuestion + vbYesNo, "Delete Confirmation") = vbYes Then
            Kill AppPath + "Presets\" + tvPresets.SelectedItem.Text + ".dpp"
            LoadPresets
        End If
    End If

End Sub

Private Sub ctxMenuEdit_Click()

    Dim sNode As Node
    Dim sFile As String
    
    Set sNode = tvPresets.SelectedItem

    sNode.Selected = True
    tvPresets_NodeClick sNode

    sFile = AppPath + "Presets\" + sNode.Text + ".dpp"
    
    If GetPresetProperty(sFile, piAuthor) <> USER Then
        MsgBox "Only Presets that you have created can be edited." + vbCrLf + vbCrLf + "If you would like to make changes to the selected Preset click the 'Load' button to open the Preset in DHTML Menu Builder," + vbCrLf + "make the desired changes and click 'File->Save As Preset' to save it as a new Preset", vbInformation + vbOKOnly, "Error Editing Preset"
        Exit Sub
    End If
    
    UncompressPreset sNode.Text
    
    With flbPresets
        .Path = AppPath + "Presets\tmp\"
        .Pattern = "*.*"
        .Refresh
    End With
    
    Project.HasChanged = False
    LoadSelectedPreset
    
    With frmPresetCreate
        .txtAuthor.Text = GetPresetProperty(sFile, piAuthor)
        .txtComments.Text = GetPresetProperty(sFile, piComments)
        .txtTitle.Text = GetPresetProperty(sFile, piTitle)
        .cmbCategory.Text = GetPresetProperty(sFile, piCategory)
        .Show vbModal
    End With
    
    FinalizeRepacking
    LoadPresets

End Sub

Private Sub ctxMenuLoad_Click()

    cmdLoad_Click

End Sub

Private Sub ctxMenuRepack_Click()

    Repack tvPresets.SelectedItem
    FinalizeRepacking

End Sub

Private Sub ctxMenuRepackAll_Click()

    Dim sNode As Node
    
    For Each sNode In tvPresets.Nodes
        If Not sNode.parent Is Nothing Then
            Repack sNode
        End If
    Next sNode
    
    FinalizeRepacking
    
    LoadPresets

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    Dim hFile As String

    If KeyCode = vbKeyF1 Then
        Select Case PresetWorkingMode
            Case pwmApplyStyle
                hFile = "dialogs/presets_applying.htm"
            Case pwmNormal
                hFile = "dialogs/presets_main.htm"
            Case pwmSubmit
                hFile = "dialogs/presets_sharing.htm"
        End Select
        
        showHelp hFile
    End If
    
End Sub

Private Sub Form_Load()

    LocalizeUI
    
    wbPreview.Navigate "about:blank"
    
    Select Case PresetWorkingMode
        Case pwmApplyStyle
            cmdLoad.caption = GetLocalizedStr(802)
            ctxMenuLoad.caption = GetLocalizedStr(802)
            tsImportTabs.Visible = True
            picImportOptions.Visible = True
            picImportOptions.BorderStyle = 0
            tvProperties.ImageList = frmMain.ilIcons
        Case pwmSubmit
            cmdLoad.caption = GetLocalizedStr(894)
            ctxMenuLoad.caption = GetLocalizedStr(894)
    End Select
    
    If Val(GetSetting(App.EXEName, "PresetManWinPos", "X")) = 0 Then
        Width = 10815
        Height = 8370
        CenterForm Me
    Else
        Top = GetSetting(App.EXEName, "PresetManWinPos", "X")
        Left = GetSetting(App.EXEName, "PresetManWinPos", "Y")
        Width = GetSetting(App.EXEName, "PresetManWinPos", "W")
        Height = GetSetting(App.EXEName, "PresetManWinPos", "H")
        WindowState = Val(GetSetting(App.EXEName, "PresetManWinPos", "State"))
    End If
    
    If Not IsDebug Then
        Set xMenu = New CMenu
        xMenu.Initialize Me
    Else
        vlmCtrl.Enabled = False
    End If
    
    Set FloodPanel.PictureControl = picFlood
    
    tmrInit.Enabled = True

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
    
    UseOnlineCatalog = False
    lblAuthor.caption = ""
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
    
    FloodPanel.caption = "Please wait... Loading Available Presets"
    
    For i = 0 To flbPresets.ListCount - 1
        FloodPanel.Value = i / (flbPresets.ListCount - 1) * 100
        dFile = dPath + flbPresets.List(i)
        
        Set nNode = Nothing
        
        sCat = GetPresetProperty(dFile, piCategory)
        If LenB(sCat) = 0 Then sCat = "(uncategorized)"
        
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
    
    If PresetWorkingMode = pwmNormal Then LoadPresetsFromWeb
    
    CleanPresetsDirs
    
    If LenB(ForceSelection) <> 0 Then
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

Private Sub LoadPresetsFromWeb()

    Dim p() As String
    Dim info() As String
    Dim i As Integer
    Dim nNode As Node
    Dim k As Integer
    
    On Error Resume Next
    
    If Not UseOnlineCatalog Then
        tvPresets.Nodes.Add , , "WebOCP", "Online Catalog", 3
        Exit Sub
    End If
    
    wbPreview.Visible = False

    If IsOnline Then
        wbPreview.Navigate "https://xfx.net/utilities/dmbuilder/presets/getlist.php"
        
        SubmitRes = ""
        Do
            DoEvents
        Loop Until LenB(SubmitRes) <> 0
        wbPreview.Navigate "about:blank"
        
        p = Split(SubmitRes, "&")
        For i = 0 To UBound(p) - 1
            info = Split(p(i), "|")
            
            If Not FileExists(AppPath + "Presets\" + info(3) + ".dpp") Then
                Set nNode = tvPresets.Nodes("K" + info(2))
                If nNode Is Nothing Then
                    Set nNode = tvPresets.Nodes.Add(, , "K" + info(2), info(2), 1)
                    nNode.Expanded = True
                End If
                
                tvPresets.Nodes.Add(nNode, tvwChild, , info(3), 3).tag = info(4) & "|" & info(5) & "|" & info(1) & "|" & info(0)
                k = k + 1
            End If
        Next i
    End If
    
    wbPreview.Visible = True
    
    MsgBox k & " Presets have been downloaded from the Online Catalog", vbInformation + vbOKOnly, "Online Catalog"

End Sub

Private Sub ShowPresetPreview(nNode As Node)

          Dim ff As Integer
          Dim dFile As String
          Dim pName As String
          Dim pkey As String
          Dim pid As String
          
10        On Error GoTo DisplayError
          
20        Screen.MousePointer = vbArrowHourglass
          
30        wbPreview.Navigate "about:blank"
40        Do While wbPreview.Busy And (wbPreview.ReadyState <> READYSTATE_COMPLETE)
50            DoEvents
60        Loop
          
70        If IsFromWeb(nNode) Then
80            cmdLoad.caption = GetLocalizedStr(469)
90            pkey = Split(nNode.tag, "|")(0)
100           pid = Split(nNode.tag, "|")(1)
110           dFile = "https://xfx.net/utilities/dmbuilder/presets/preview.php?key=" + pkey + "&pname=" + nNode.Text + "&id=" + pid + "&app=1"
120       Else
130           cmdLoad.caption = "Download"
140           Select Case PresetWorkingMode
                  Case pwmNormal
150                   cmdLoad.caption = GetLocalizedStr(796)
160               Case pwmApplyStyle
170                   cmdLoad.caption = GetLocalizedStr(802)
180               Case pwmSubmit
190                   cmdLoad.caption = GetLocalizedStr(894)
200           End Select
              
210           pName = nNode.Text
220           UncompressPreset pName
              
230           dFile = TempPath + "Presets\tmp\index.html"
              
240           ff = FreeFile
250           Open dFile For Output As #ff
260               Print #ff, "<html><body>"
270               Print #ff, "<script language=JavaScript src=menu.js></script>"
280               Print #ff, "</body></html>"
290           Close #ff
              
300           tvPresets.SelectedItem.tag = GetFilePath(dFile) + Dir(GetFilePath(dFile) + "*.dmb")
310       End If
          
320       wbPreview.Navigate dFile
          
330       Screen.MousePointer = vbDefault
          
340       Exit Sub
          
DisplayError:
350       MsgBox "Error " & Err.number & ": " & Err.Description & " at " & Erl

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If WindowState = vbNormal Then
        SaveSetting App.EXEName, "PresetManWinPos", "X", Top
        SaveSetting App.EXEName, "PresetManWinPos", "Y", Left
        SaveSetting App.EXEName, "PresetManWinPos", "W", Width
        SaveSetting App.EXEName, "PresetManWinPos", "H", Height
    End If
    SaveSetting App.EXEName, "PresetManWinPos", "State", WindowState

End Sub

Private Sub Form_Resize()

    Dim rHeight As Long

    On Error Resume Next
    
    rHeight = (Height - GetClientTop(Me.hwnd))

    tvPresets.Height = rHeight - (tvPresets.Top + frameInfo.Height + 900)
    frameInfo.Top = tvPresets.Top + tvPresets.Height + 45
    
    Select Case PresetWorkingMode
        Case pwmNormal, pwmSubmit
            wbPreview.Move tvPresets.Width + 120, 300, Width - (tvPresets.Left + tvPresets.Width + 270), rHeight - 1155
                            
            cmdLoad.Move Width - 2250, wbPreview.Top + wbPreview.Height + 120
            cmdCancel.Move Width - 1110, wbPreview.Top + wbPreview.Height + 120
        Case pwmApplyStyle
            wbPreview.Move tvPresets.Width + 120, 300, Width - (tvPresets.Left + tvPresets.Width + 270), rHeight - (wbPreview.Top + tsImportTabs.Height + 915)
            tsImportTabs.Move wbPreview.Left, wbPreview.Top + wbPreview.Height + 45, wbPreview.Width
            picImportOptions.Move tsImportTabs.Left + 75, tsImportTabs.Top + 360, tsImportTabs.Width - 135
            tvProperties.Width = picImportOptions.Width - 135
            
            cmdLoad.Move Width - 2250, tsImportTabs.Top + tsImportTabs.Height + 120
            cmdCancel.Move Width - 1110, tsImportTabs.Top + tsImportTabs.Height + 120
    End Select
    lblPreview.Left = wbPreview.Left
    
    picFlood.Move 30, sbInfo.Top + 60, sbInfo.Width - 470, sbInfo.Height - 75

End Sub

Private Sub UpdateInfo(sNode As Node)

    Dim dFile As String

    If sNode Is Nothing Then
        lblAuthor.caption = ""
        txtComments.Text = ""
        lblTitle.caption = "No Preset Selected"
    Else
        dFile = AppPath + "Presets\" + sNode.Text + ".dpp"
        If IsFromWeb Then
            lblAuthor.caption = Split(sNode.tag, "|")(3)
            txtComments.Text = Split(sNode.tag, "|")(2)
        Else
            lblAuthor.caption = GetPresetProperty(dFile, piAuthor)
            txtComments.Text = GetPresetProperty(dFile, piComments)
        End If
        lblTitle.caption = sNode.Text
    End If
    
End Sub

Private Sub tmrCloseDlg_Timer()

    cmdCancel_Click

End Sub

Private Sub tmrInit_Timer()

    tmrInit.Enabled = False
    
    LoadPresets

End Sub

Private Sub tsImportTabs_Click()

    PopulateImportOptions

End Sub

Private Sub tvPresets_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim sNode As Node

    If Button = vbRightButton Then
        Set sNode = tvPresets.HitTest(x, y)
        If Not sNode Is Nothing Then
            If Not sNode.parent Is Nothing Then
                sNode.EnsureVisible
                sNode.Selected = True
                SetControlsState True
                tvPresets_NodeClick sNode
                DoEvents
                PopupMenu mnuPresets, vbRightButton, tvPresets.Left + x, tvPresets.Top + y, ctxMenuLoad
            End If
        End If
    End If
    
End Sub

Private Sub tvPresets_NodeClick(ByVal Node As MSComctlLib.Node)

    Me.Enabled = False

    If Node.key = "WebOCP" Then
        If MsgBox("Are you sure you want to download all available presets from the Online Catalog?", vbQuestion + vbYesNo, "Online Catalog") = vbYes Then
            tvPresets.Nodes.Remove Node.Index
            UseOnlineCatalog = True
            LoadPresetsFromWeb
            
            If tvPresets.Nodes.Count > 0 Then
                With tvPresets.Nodes(1)
                    .EnsureVisible
                    .Selected = True
                End With
            End If
        End If
    Else
        If Node.parent Is Nothing Then
            CleanPresetsDirs
            wbPreview.Navigate "about:blank"
            SetControlsState False
            UpdateInfo Nothing
        Else
            ShowPresetPreview Node
            SetControlsState True
            UpdateInfo Node
            If PresetWorkingMode = pwmApplyStyle Then
                Prj = GetProjectProperties(tvPresets.SelectedItem.tag)
                LoadMenuItems
                PopulateImportOptions
                Close #ff
            End If
        End If
    End If
    
    Me.Enabled = True
    
End Sub

Private Function IsFromWeb(Optional sNode As Node) As Boolean

    If sNode Is Nothing Then Set sNode = tvPresets.SelectedItem
    
    IsFromWeb = InStr(sNode.tag, "|")

End Function

Private Sub PopulateImportOptions()

    Dim i As Integer
    
    IsUpdating = True

    cmbItems.Clear
    cmbItems.AddItem "(do not import)"
    tvProperties.Nodes.Clear
    
    Select Case tsImportTabs.SelectedItem.key
        Case "tbToolbar"
            For i = 1 To UBound(Prj.Toolbars)
                cmbItems.AddItem Prj.Toolbars(i).Name
            Next i
            
            With tvProperties.Nodes
                .Add , , "ALL", GetLocalizedStr(507), IconIndex("All Properties")
                    .Add "ALL", tvwChild, "Appearance", GetLocalizedStr(352), IconIndex("Properties")
                        .Add "Appearance", tvwChild, "AppearanceTS", GetLocalizedStr(354), IconIndex("Commands Layout")
                        .Add "Appearance", tvwChild, "AppearanceB", GetLocalizedStr(355), IconIndex("Size")
                        .Add "Appearance", tvwChild, "AppearanceM", GetLocalizedStr(216), IconIndex("Margins")
                        .Add "Appearance", tvwChild, "AppearanceS", GetLocalizedStr(356), IconIndex("Leading")
                        .Add "Appearance", tvwChild, "AppearanceJ", GetLocalizedStr(357), IconIndex("Justify")
                        .Add "Appearance", tvwChild, "AppearanceBC", GetLocalizedStr(182), IconIndex("Back Color")
                        .Add "Appearance", tvwChild, "AppearanceBI", GetLocalizedStr(205), IconIndex("Image")
                    .Add "ALL", tvwChild, "Positioning", GetLocalizedStr(353), IconIndex("Properties")
                        .Add "Positioning", tvwChild, "PositioningA", "Alignment", IconIndex("Toolbar Alignment")
                        .Add "Positioning", tvwChild, "PositioningS", "Spanning", IconIndex("Spanning")
                        .Add "Positioning", tvwChild, "PositioningO", "Offset", IconIndex("Toolbar Offset")
                    .Add "ALL", tvwChild, "Advanced", GetLocalizedStr(325), IconIndex("Properties")
                        .Add "Advanced", tvwChild, "AdvancedFS", GetLocalizedStr(115), IconIndex("Follow Scrolling")
                        .Add "Advanced", tvwChild, "AdvancedTS", GetLocalizedStr(203), IconIndex("Size")
            End With
        Case "tbGroup"
            For i = 1 To UBound(mg)
                cmbItems.AddItem IIf(LenB(GetParam(mg(i), 43)) = 0, "[" + GetParam(mg(i), 1) + "]", GetParam(mg(i), 43))
            Next i
            
            With tvProperties.Nodes
                .Add , , "ALL", GetLocalizedStr(507), IconIndex("All Properties")
                    .Add "ALL", tvwChild, "Color", GetLocalizedStr(212), IconIndex("Color")
                    .Add "ALL", tvwChild, "Font", GetLocalizedStr(213), IconIndex("Font")
                    .Add "ALL", tvwChild, "Cursor", GetLocalizedStr(215), IconIndex("Cursor")
                    .Add "ALL", tvwChild, "Image", GetLocalizedStr(214), IconIndex("Image")
                    .Add "ALL", tvwChild, "Leading", GetLocalizedStr(295), IconIndex("Leading")
                    .Add "ALL", tvwChild, "Margins", GetLocalizedStr(216), IconIndex("Margins")
                    .Add "ALL", tvwChild, "Border", GetLocalizedStr(355), IconIndex("Frame")
                    .Add "ALL", tvwChild, "SFX", GetLocalizedStr(231), IconIndex("Special Effects")
            End With
        Case "tbSubCommand"
            For i = 1 To UBound(sc)
                cmbItems.AddItem IIf(LenB(GetParam(sc(i), 2)) = 0, "[" + GetParam(sc(i), 1) + "]", GetParam(sc(i), 2))
            Next i
            
            With tvProperties.Nodes
                .Add , , "ALL", GetLocalizedStr(507), IconIndex("All Properties")
                    .Add "ALL", tvwChild, "Color", GetLocalizedStr(212), IconIndex("Color")
                    .Add "ALL", tvwChild, "Font", GetLocalizedStr(213), IconIndex("Font")
                    .Add "ALL", tvwChild, "Cursor", GetLocalizedStr(215), IconIndex("Cursor")
                    .Add "ALL", tvwChild, "Image", GetLocalizedStr(214), IconIndex("Image")
            End With
        Case "tbCommand"
            For i = 1 To UBound(mc)
                cmbItems.AddItem IIf(LenB(GetParam(mc(i), 2)) = 0, "[" + GetParam(mc(i), 1) + "]", GetParam(mc(i), 2))
            Next i
            
            With tvProperties.Nodes
                .Add , , "ALL", GetLocalizedStr(507), IconIndex("All Properties")
                    .Add "ALL", tvwChild, "Color", GetLocalizedStr(212), IconIndex("Color")
                    .Add "ALL", tvwChild, "Font", GetLocalizedStr(213), IconIndex("Font")
                    .Add "ALL", tvwChild, "Cursor", GetLocalizedStr(215), IconIndex("Cursor")
                    .Add "ALL", tvwChild, "Image", GetLocalizedStr(214), IconIndex("Image")
            End With
    End Select
    
    tvProperties.Nodes("ALL").Expanded = True
    
    UpdateInterfaceFromSelItemProp
    
    IsUpdating = False

End Sub

Private Sub UpdateInterfaceFromSelItemProp()

    Dim sNode As Node
    
    On Error GoTo ExitSub

    With SelItemProp
        Select Case tsImportTabs.SelectedItem.key
            Case "tbToolbar"
                cmbItems.ListIndex = .SelTB
                tvProperties.Nodes("AppearanceTS").Checked = .SelTBProp.tbToolbarStyle
                tvProperties.Nodes("AppearanceB").Checked = .SelTBProp.tbBorder
                tvProperties.Nodes("AppearanceM").Checked = .SelTBProp.tbMargins
                tvProperties.Nodes("AppearanceS").Checked = .SelTBProp.tbSeparation
                tvProperties.Nodes("AppearanceJ").Checked = .SelTBProp.tbJustification
                tvProperties.Nodes("AppearanceBC").Checked = .SelTBProp.tbBackColor
                tvProperties.Nodes("AppearanceBI").Checked = .SelTBProp.tbBackImage
                tvProperties.Nodes("PositioningA").Checked = .SelTBProp.tbAlignment
                tvProperties.Nodes("PositioningS").Checked = .SelTBProp.tbSpanning
                tvProperties.Nodes("PositioningO").Checked = .SelTBProp.tbOffset
                tvProperties.Nodes("AdvancedFS").Checked = .SelTBProp.tbFollowScrolling
                tvProperties.Nodes("AdvancedTS").Checked = .SelTBProp.tbToolbarSize
            Case "tbSubCommand"
                cmbItems.ListIndex = .SelSubCmd
                tvProperties.Nodes("Color").Checked = .SelSubCmdProp.cColor
                tvProperties.Nodes("Font").Checked = .SelSubCmdProp.cFont
                tvProperties.Nodes("Cursor").Checked = .SelSubCmdProp.cCursor
                tvProperties.Nodes("Image").Checked = .SelSubCmdProp.cImage
            Case "tbGroup"
                cmbItems.ListIndex = .SelGrp
                tvProperties.Nodes("Color").Checked = .SelGrpProp.gColor
                tvProperties.Nodes("Font").Checked = .SelGrpProp.gFont
                tvProperties.Nodes("Cursor").Checked = .SelGrpProp.gCursor
                tvProperties.Nodes("Image").Checked = .SelGrpProp.gImage
                tvProperties.Nodes("Leading").Checked = .SelGrpProp.gLeading
                tvProperties.Nodes("Margins").Checked = .SelGrpProp.gMargins
                tvProperties.Nodes("Border").Checked = .SelGrpProp.gBorders
                tvProperties.Nodes("SFX").Checked = .SelGrpProp.gSFX
            Case "tbCommand"
                cmbItems.ListIndex = .SelCmd
                tvProperties.Nodes("Color").Checked = .SelCmdProp.cColor
                tvProperties.Nodes("Font").Checked = .SelCmdProp.cFont
                tvProperties.Nodes("Cursor").Checked = .SelCmdProp.cCursor
                tvProperties.Nodes("Image").Checked = .SelCmdProp.cImage
        End Select
    End With
    
    Select Case tsImportTabs.SelectedItem.key
        Case "tbToolbar"
            For Each sNode In tvProperties.Nodes
                If sNode.children = 0 Then
                    If sNode.Checked Then
                        sNode.parent.Checked = True
                        sNode.parent.parent.Checked = True
                    End If
                End If
            Next sNode
        Case Else
            For Each sNode In tvProperties.Nodes
                If sNode.Checked Then
                    tvProperties.Nodes("ALL").Checked = True
                    Exit For
                End If
            Next sNode
    End Select
    
ExitSub:

End Sub

Private Sub LoadMenuItems()

    Dim sStr As String
    Dim i As Integer
    
    ReDim mg(0)
    ReDim mc(0)
    ReDim sc(0)
    
    On Error GoTo ExitSub

    If (LOF(ff) = Loc(ff)) Then Exit Sub
    Line Input #ff, sStr
    Do Until (LOF(ff) = Loc(ff)) Or sStr = "[RSC]"
        If LenB(sStr) <> 0 Then Add2Array mg, Mid$(sStr, 4)
        Do Until (LOF(ff) = Loc(ff)) Or sStr = "[RSC]"
            Line Input #ff, sStr
            If LenB(sStr) <> 0 Then
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
        .SelTB = IIf(UBound(Prj.Toolbars) > 0, 1, 0)
        .SelGrp = IIf(UBound(mg) > 0, 1, 0)
        .SelCmd = IIf(UBound(mc) > 0, 1, 0)
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
    
ExitSub:

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

    ctxMenuDelete.Enabled = State
    ctxMenuEdit.Enabled = State
    ctxMenuLoad.Enabled = State
    ctxMenuRepack.Enabled = State
    cmdLoad.Enabled = State
    
    If PresetWorkingMode = pwmApplyStyle Then
        tsImportTabs.Enabled = State
        cmbItems.Enabled = State
        tvProperties.Enabled = State
    End If

End Sub

Private Sub tvProperties_NodeCheck(ByVal Node As MSComctlLib.Node)

    Dim cNode As Node
    Static Recursive As Integer
    
    Set cNode = Node.Child
    Do Until cNode Is Nothing
        cNode.Checked = Node.Checked
        Recursive = Recursive + 1
        tvProperties_NodeCheck cNode
        Recursive = Recursive - 1
        Set cNode = cNode.Next
    Loop
    
    If Recursive = 0 Then
        Set Node = Node.parent
        Do Until Node Is Nothing
            Set cNode = Node.Child
            Node.Checked = False
            Do Until cNode Is Nothing
                If cNode.Checked Then
                    cNode.parent.Checked = True
                    Exit Do
                End If
                Set cNode = cNode.Next
            Loop
            Set Node = Node.parent
        Loop
    End If
    
    UpdatePropertiesSelection

End Sub

Private Sub UpdatePropertiesSelection()

    With SelItemProp
        Select Case tsImportTabs.SelectedItem.key
            Case "tbToolbar"
                .SelTB = cmbItems.ListIndex
                .SelTBProp.tbToolbarStyle = tvProperties.Nodes("AppearanceTS").Checked
                .SelTBProp.tbBorder = tvProperties.Nodes("AppearanceB").Checked
                .SelTBProp.tbMargins = tvProperties.Nodes("AppearanceM").Checked
                .SelTBProp.tbSeparation = tvProperties.Nodes("AppearanceS").Checked
                .SelTBProp.tbJustification = tvProperties.Nodes("AppearanceJ").Checked
                .SelTBProp.tbBackColor = tvProperties.Nodes("AppearanceBC").Checked
                .SelTBProp.tbBackImage = tvProperties.Nodes("AppearanceBI").Checked
                .SelTBProp.tbAlignment = tvProperties.Nodes("PositioningA").Checked
                .SelTBProp.tbSpanning = tvProperties.Nodes("PositioningS").Checked
                .SelTBProp.tbOffset = tvProperties.Nodes("PositioningO").Checked
                .SelTBProp.tbFollowScrolling = tvProperties.Nodes("AdvancedFS").Checked
                .SelTBProp.tbToolbarSize = tvProperties.Nodes("AdvancedTS").Checked
            Case "tbSubCommand"
                .SelSubCmd = cmbItems.ListIndex
                .SelSubCmdProp.cColor = tvProperties.Nodes("Color").Checked
                .SelSubCmdProp.cFont = tvProperties.Nodes("Font").Checked
                .SelSubCmdProp.cCursor = tvProperties.Nodes("Cursor").Checked
                .SelSubCmdProp.cImage = tvProperties.Nodes("Image").Checked
            Case "tbGroup"
                .SelGrp = cmbItems.ListIndex
                .SelGrpProp.gColor = tvProperties.Nodes("Color").Checked
                .SelGrpProp.gFont = tvProperties.Nodes("Font").Checked
                .SelGrpProp.gCursor = tvProperties.Nodes("Cursor").Checked
                .SelGrpProp.gImage = tvProperties.Nodes("Image").Checked
                .SelGrpProp.gLeading = tvProperties.Nodes("Leading").Checked
                .SelGrpProp.gBorders = tvProperties.Nodes("Border").Checked
                .SelGrpProp.gMargins = tvProperties.Nodes("Margins").Checked
                .SelGrpProp.gSFX = tvProperties.Nodes("SFX").Checked
            Case "tbCommand"
                .SelCmd = cmbItems.ListIndex
                .SelCmdProp.cColor = tvProperties.Nodes("Color").Checked
                .SelCmdProp.cFont = tvProperties.Nodes("Font").Checked
                .SelCmdProp.cCursor = tvProperties.Nodes("Cursor").Checked
                .SelCmdProp.cImage = tvProperties.Nodes("Image").Checked
        End Select
    End With
    
End Sub

Private Sub LocalizeUI()

    lblPresets.caption = GetLocalizedStr(797)
    lblPreview.caption = GetLocalizedStr(801)
    
    lblLabelAuthor.caption = GetLocalizedStr(798)
    lblLabelComments.caption = GetLocalizedStr(799)
    
    tsImportTabs.Tabs("tbToolbar").caption = GetLocalizedStr(323)
    tsImportTabs.Tabs("tbSubCommand").caption = GetLocalizedStr(800)
    tsImportTabs.Tabs("tbGroup").caption = GetLocalizedStr(270)
    tsImportTabs.Tabs("tbCommand").caption = GetLocalizedStr(271)
    
    mnuPresets.caption = GetLocalizedStr(807)
    ctxMenuDelete.caption = GetLocalizedStr(808)
    ctxMenuEdit.caption = GetLocalizedStr(339)
    ctxMenuLoad.caption = GetLocalizedStr(796)

    cmdCancel.caption = GetLocalizedStr(187)

End Sub

Private Sub wbPreview_NavigateComplete2(ByVal pDisp As Object, url As Variant)

    Dim doc As IHTMLDocument2
    
    On Error GoTo ExitWithError
    
    If url = "http:///" Or url = "about:blank" Then Exit Sub
    
    If PresetWorkingMode = pwmSubmit Then
        If InStr(url, USERSN) = 0 Then Exit Sub
    
        Do While (wbPreview.Document Is Nothing)
            DoEvents
        Loop
        Set doc = wbPreview.Document
        
        SubmitRes = Trim(doc.body.innerText)
    Else
        If InStr(url, "getlist.php") Then
            Do While (wbPreview.Document Is Nothing)
                DoEvents
            Loop
            Set doc = wbPreview.Document
            SubmitRes = Trim(doc.body.innerText)
        End If
    End If
    
    Exit Sub
    
ExitWithError:
    SubmitRes = "UNDEFINED ERROR"

End Sub

