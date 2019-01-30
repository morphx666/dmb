VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{DBF30C82-CAF3-11D5-84FF-0050BA3D926D}#8.5#0"; "VLMnuPlus.ocx"
Begin VB.Form frmLCMan 
   Caption         =   "Install Menus"
   ClientHeight    =   7275
   ClientLeft      =   5760
   ClientTop       =   5250
   ClientWidth     =   8610
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLCMan.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7275
   ScaleWidth      =   8610
   Begin VB.CheckBox chkIncludeFPSubWebs 
      Caption         =   "Include FrontPage SubWebs"
      Height          =   195
      Left            =   720
      TabIndex        =   17
      Top             =   5445
      Width           =   3045
   End
   Begin VB.CheckBox chkFilter 
      Caption         =   "Hide Unsupported Documents"
      Height          =   195
      Left            =   795
      TabIndex        =   15
      Top             =   4680
      Value           =   1  'Checked
      Width           =   3045
   End
   Begin VB.CommandButton cmdAuto 
      Caption         =   "Auto"
      Height          =   360
      Left            =   930
      TabIndex        =   16
      Top             =   4845
      Width           =   1155
   End
   Begin VB.PictureBox picLCSplit 
      BorderStyle     =   0  'None
      Height          =   120
      Left            =   420
      MouseIcon       =   "frmLCMan.frx":014A
      MousePointer    =   99  'Custom
      ScaleHeight     =   120
      ScaleWidth      =   2475
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1560
      Width           =   2475
   End
   Begin VB.PictureBox picSplit 
      BorderStyle     =   0  'None
      Height          =   2880
      Left            =   2415
      MouseIcon       =   "frmLCMan.frx":029C
      MousePointer    =   99  'Custom
      ScaleHeight     =   2880
      ScaleWidth      =   120
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   -135
      Width           =   120
   End
   Begin VB.PictureBox picFFSplit 
      BorderStyle     =   0  'None
      Height          =   120
      Left            =   390
      MouseIcon       =   "frmLCMan.frx":03EE
      MousePointer    =   99  'Custom
      ScaleHeight     =   120
      ScaleWidth      =   2475
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1350
      Width           =   2475
   End
   Begin MSComctlLib.ImageList ilColHeadIcons 
      Left            =   1410
      Top             =   1980
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   8
      ImageHeight     =   8
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLCMan.frx":0540
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLCMan.frx":069C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLCMan.frx":07F8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvFiles 
      Height          =   645
      Left            =   600
      TabIndex        =   9
      Top             =   3030
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   1138
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      OLEDragMode     =   1
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ilIcons"
      ColHdrIcons     =   "ilColHeadIcons"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      OLEDragMode     =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "chName"
         Text            =   "Name"
         Object.Width           =   2540
         ImageIndex      =   2
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Key             =   "chSize"
         Text            =   "Size"
         Object.Width           =   2540
         ImageIndex      =   3
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "chType"
         Text            =   "Type"
         Object.Width           =   2540
         ImageIndex      =   3
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Key             =   "chDate"
         Text            =   "Date Modified"
         Object.Width           =   2540
         ImageIndex      =   3
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Key             =   "chLC"
         Text            =   "Loader Code"
         Object.Width           =   2540
         ImageIndex      =   3
      EndProperty
   End
   Begin MSComctlLib.TreeView tvFLC 
      Height          =   1290
      Left            =   3450
      TabIndex        =   10
      Top             =   2580
      Width           =   2205
      _ExtentX        =   3889
      _ExtentY        =   2275
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   529
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      ImageList       =   "ilIcons"
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDropMode     =   1
   End
   Begin MSComctlLib.TreeView tvLC 
      Height          =   1035
      Left            =   3570
      TabIndex        =   5
      Top             =   630
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1826
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   529
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      ImageList       =   "ilIcons"
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDropMode     =   1
   End
   Begin VB.Timer tmrInit 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   6255
      Top             =   2295
   End
   Begin VB.PictureBox picIcon 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   105
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   4515
      Visible         =   0   'False
      Width           =   240
   End
   Begin MSComctlLib.StatusBar sbInfo 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   18
      Top             =   6960
      Width           =   8610
      _ExtentX        =   15187
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14658
            Key             =   "pInfo"
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtDummy 
      Height          =   315
      Left            =   135
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   4065
      Visible         =   0   'False
      Width           =   2460
   End
   Begin VB.CommandButton cmdInstall 
      Caption         =   "Install"
      Height          =   360
      Left            =   4485
      TabIndex        =   13
      Top             =   4410
      Width           =   1155
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   5835
      TabIndex        =   14
      Top             =   4410
      Width           =   1155
   End
   Begin MSComctlLib.ImageList ilIcons 
      Left            =   735
      Top             =   1980
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLCMan.frx":084B
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLCMan.frx":0F1D
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvFolders 
      Height          =   1515
      Left            =   60
      TabIndex        =   3
      Top             =   270
      Width           =   2685
      _ExtentX        =   4736
      _ExtentY        =   2672
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   529
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      ImageList       =   "ilIcons"
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDragMode     =   1
   End
   Begin VLMnuPlus.VLMenuPlus vlmCtrl 
      Left            =   6990
      Top             =   3180
      _ExtentX        =   847
      _ExtentY        =   847
      _CXY            =   4
      _CGUID          =   43165.2824652778
      Language        =   0
   End
   Begin VB.Label lblFiles 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Files on"
      Height          =   195
      Left            =   495
      TabIndex        =   8
      Top             =   2730
      Width           =   540
   End
   Begin VB.Label lblFLC 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Files that will receive the Menus"
      Height          =   195
      Left            =   3075
      TabIndex        =   7
      Top             =   2220
      Width           =   2280
   End
   Begin VB.Label lblLC 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Files that will receive the Menus"
      Height          =   195
      Left            =   3060
      TabIndex        =   2
      Top             =   45
      Width           =   2280
   End
   Begin VB.Label lblFolders 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Folders on web site"
      Height          =   195
      Left            =   60
      TabIndex        =   0
      Top             =   45
      Width           =   1395
   End
   Begin VB.Menu mnuContext 
      Caption         =   "mnuContext"
      Begin VB.Menu mnuContextILC 
         Caption         =   "Install Menus"
      End
      Begin VB.Menu mnuContextIFLC 
         Caption         =   "Install Menus"
      End
      Begin VB.Menu mnuContextSep01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuContextRemove 
         Caption         =   "Remove Menus"
      End
   End
End
Attribute VB_Name = "frmLCMan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private RootWeb As String
Private WithEvents msgSubClass As xfxSC
Attribute msgSubClass.VB_VarHelpID = -1
Private IsResizing As Boolean
Private DragData() As String
Private Const DimmedColor = &H80000011
Private IsSplittingFF As Boolean
Private IsSplitting As Boolean
Private IsSplittingLC As Boolean

Private Enum LCType
    lctNone = 0
    lctStandard = 1
    lctFrames = 2
    lctInvalid = 4
    lctAll = 7
    lctReadOnly = 8
End Enum

Private Type CFile
    fName As String
    fSource As String
    fType As LCType
    fIndex As Long
    fXHTMLCompliance As Boolean
    fYahooSiteBuilderSupport As Boolean
End Type
Private Files() As CFile

Private UsesFrames As Boolean
Private ValidFiles As String
Private FolderIconIdx As Integer
Private ShowAllFiles As Boolean
Private IncludeFPSubWebs As Boolean
Private IsLoading As Boolean

Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private xMenu As CMenu

Private Sub chkFilter_Click()

    If IsLoading Then Exit Sub

    ShowAllFiles = Not ShowAllFiles
    LoadFiles
    
    DoEvents
    
    tvFolders.SetFocus

End Sub

Private Sub chkIncludeFPSubWebs_Click()

    If IsLoading Then Exit Sub

    IncludeFPSubWebs = Not IncludeFPSubWebs
    LoadFiles
    
    DoEvents
    
    tvFolders.SetFocus

End Sub

Private Sub cmdAuto_Click()

    Dim nFiles As Integer
    
    nFiles = tvLC.Nodes.Count + tvFLC.Nodes.Count
    AutoDetect
    
    If nFiles = tvLC.Nodes.Count + tvFLC.Nodes.Count Then
        UpdateStatusbar ""
        MsgBox GetLocalizedStr(770), vbInformation + vbOKOnly, GetLocalizedStr(761)
    Else
        UpdateStatusbar GetLocalizedStr(815)
    End If

End Sub

Private Sub cmdClose_Click()

    Unload Me

End Sub

Private Sub cmdInstall_Click()

    Dim fNode As Node
    Dim fn As Long
    Dim i As Integer
    Dim oSource As String
    
    If tvLC.Nodes.Count < 2 And tvFLC.Nodes.Count < 2 Then
        MsgBox GetLocalizedStr(771), vbInformation + vbOKOnly, GetLocalizedStr(761)
        Exit Sub
    End If
    
    'DisplayTip "Please make sure that none of the files selected to receive the menus are open on another application as this will prevent them from updating.", "Recompile your project", True
    'DisplayTip "Please remember to recompile your project after installing the menus. To recompile click Tools->Compile from the main menu.", "Recompile your project", True
    
    For i = 1 To UBound(Files)
        With Files(i)
            If .fType = lctInvalid And LenB(.fSource) <> 0 Then
                oSource = .fSource
                .fSource = RemoveLoaderCode(.fSource, .fName, .fYahooSiteBuilderSupport)
                If oSource = .fSource And InStr(oSource, LoaderCodeSTART) > 0 Then GoTo ExitSub
                
                If oSource <> .fSource Then SaveFile .fName, .fSource
            End If
        End With
    Next i
    
    For Each fNode In tvLC.Nodes
        fn = Val(fNode.tag)
        If fn > 0 Then
            With Files(fn)
                UpdateStatusbar GetLocalizedStr(772) + " " + .fName
                
                oSource = .fSource
                .fSource = RemoveLoaderCode(.fSource, .fName, .fYahooSiteBuilderSupport)
                If oSource = .fSource And InStr(oSource, LoaderCodeSTART) > 0 Then GoTo ExitSub
                
                .fSource = AttachLoaderCode(.fSource, GenLoaderCode(False, False, .fName, Join(SelSecProjects, "|"), .fXHTMLCompliance, .fYahooSiteBuilderSupport), .fYahooSiteBuilderSupport)
                
                SaveFile .fName, .fSource
            End With
        End If
    Next fNode
    
    For Each fNode In tvFLC.Nodes
        fn = Val(fNode.tag)
        If fn > 0 Then
            With Files(fn)
                UpdateStatusbar GetLocalizedStr(773) + ": " + .fName
                
                oSource = .fSource
                .fSource = RemoveLoaderCode(.fSource, .fName, .fYahooSiteBuilderSupport)
                If oSource = .fSource And InStr(oSource, LoaderCodeSTART) > 0 Then GoTo ExitSub
                
                .fSource = AttachLoaderCode(.fSource, GenLoaderCode(True, False, .fName, , .fXHTMLCompliance, .fYahooSiteBuilderSupport), .fYahooSiteBuilderSupport)
                
                SaveFile .fName, .fSource
            End With
        End If
    Next fNode
    
ExitSub:
    Dim Path As String
    Dim selNode() As String
    selNode = Split(tvFolders.SelectedItem.FullPath, "\")
    
    Set fNode = Nothing
    
    LoadFiles
    
    For i = 1 To UBound(selNode)
        Path = Path + selNode(i) + "\"
        Set fNode = tvFolders.Nodes("K" + Path)
        LoadFolderContents fNode
        fNode.Expanded = True
        fNode.EnsureVisible
    Next i
    
    If Not fNode Is Nothing Then
        fNode.Selected = True
        LoadFilesForFolder fNode
    End If
    
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
        Case vbKeyDelete
            If ActiveControl.Name = "tvLC" Or ActiveControl.Name = "tvFLC" Then mnuContextRemove_Click
        Case 112
            If UsesFrames Then
                showHelp "dialogs/iflcman.htm"
            Else
                showHelp "dialogs/ilcman.htm"
            End If
    End Select

End Sub

Private Sub Form_Load()

    IsLoading = True

    FolderIconIdx = GetImage("")

    mnuContext.Visible = False
    
    DoEvents

    If Val(GetSetting(App.EXEName, "LCMWinPos", "X")) = 0 Then
        picSplit.Left = Width / 2
        picFFSplit.Top = Height / 2
        picLCSplit.Top = Height / 2
        CenterForm Me
    Else
        Left = GetSetting(App.EXEName, "LCMWinPos", "X")
        Top = GetSetting(App.EXEName, "LCMWinPos", "Y")
        Width = GetSetting(App.EXEName, "LCMWinPos", "W")
        Height = GetSetting(App.EXEName, "LCMWinPos", "H")
        picFFSplit.Top = GetSetting(App.EXEName, "LCMWinPos", "S0", Height / 2)
        picSplit.Left = GetSetting(App.EXEName, "LCMWinPos", "S1", Width / 2)
        picLCSplit.Top = GetSetting(App.EXEName, "LCMWinPos", "S2", Height / 2)
        
        If Left + Width / 2 > Screen.Width Or Top + Height / 2 > Screen.Height Then
            Left = Screen.Width / 2 - Width / 2
            Top = Screen.Height / 2 - Height / 2
        End If
    End If
    chkFilter.Value = GetSetting(App.EXEName, "LCMWinPos", "Filter", vbChecked)
    ShowAllFiles = (chkFilter.Value = vbUnchecked)
    chkIncludeFPSubWebs.Value = GetSetting(App.EXEName, "LCMWinPos", "IncludeFPSubWebs", vbUnchecked)
    IncludeFPSubWebs = (chkIncludeFPSubWebs.Value = vbChecked)
    
    LocalizeUI
    SetupCharset Me
    
    UsesFrames = Project.UserConfigs(Project.DefaultConfig).Frames.UseFrames
    
    If UsesFrames Then
        mnuContextILC.caption = GetLocalizedStr(768)
        mnuContextIFLC.caption = GetLocalizedStr(761)
        lblLC.caption = GetLocalizedStr(IIf(UBound(Project.Toolbars) > 1, 767, 766))
    Else
        If SecProjMode = spmcFromInstallMenus Then
            If UBound(SelSecProjects) = 0 And UBound(Project.SecondaryProjects) > 0 Then
                frmSecProjDef.Show vbModal
            End If
        End If
        
        If CreateToolbar Then
            lblLC.caption = GetLocalizedStr(IIf(UBound(Project.Toolbars) > 1, 813, 812))
        Else
            lblLC.caption = GetLocalizedStr(814)
        End If
        mnuContextILC.caption = GetLocalizedStr(761)
        mnuContextIFLC.Visible = False
        picLCSplit.Visible = False
        lblFLC.Visible = False
        tvFLC.Visible = False
    End If
    
    lblLC.caption = lblLC.caption
    caption = GetDlgCaption
    
    MenusFrame = 0
    
    tmrInit.Enabled = True
    
    IsLoading = False
    
    If Not IsDebug Then
        Set xMenu = New CMenu
        xMenu.Initialize Me
    End If
    
    If Not IsDebug Then SetupSubclassing True
    
End Sub

Private Function GetDlgCaption(Optional SimpleCaption As Boolean = False) As String

    Dim i As Integer
    Dim sStr As String

    If SimpleCaption Then
        sStr = Project.Name
    Else
        sStr = GetLocalizedStr(761) + ": " + Project.Name
    End If
    
    For i = 1 To UBound(SelSecProjects)
        sStr = sStr + " + " + SelSecProjectsTitles(i)
    Next i
    
    GetDlgCaption = sStr

End Function

Private Sub SetDlgState(State As Boolean)

    DoEvents

    Me.Enabled = State
    MousePointer = IIf(State, vbDefault, vbArrowHourglass)
    
    If State = False Then
        IsResizing = True
        Me.AutoRedraw = True
        BitBlt hDc, 0, 0, Width / 15, Height / 15, GetDC(hwnd), 0, 0, vbSrcCopy
    Else
        IsResizing = False
        Cls
        Me.AutoRedraw = False
    End If
    
    tvFolders.Visible = State
    lvFiles.Visible = State
    tvFLC.Visible = State
    tvLC.Visible = State
    tvFLC.Visible = State And UsesFrames

End Sub

Private Sub LoadFiles()

    Dim i As Integer
    Dim c As Integer
    Dim t As Integer
    
    SetDlgState False
    
    On Error Resume Next
    
    If MenusFrame = 0 Then
        If UsesFrames Then
            t = UBound(FramesInfo.Frames)
            For c = 1 To UBound(MenuCmds)
                For i = 1 To t
                    If MenuCmds(c).Actions.onclick.TargetFrame = FramesInfo.Frames(i).Name Then
                        MenusFrame = i
                        Exit For
                    End If
                Next i
            Next c
            If MenusFrame = 0 Then frmSelMenusFrame.Show vbModal
            lblFLC.caption = GetLocalizedStr(765) + " (" + FramesInfo.Frames(MenusFrame).Name + ")"
            For i = 1 To UBound(FramesInfo.Frames)
                If FramesInfo.Frames(i).srcFile = Project.UserConfigs(Project.DefaultConfig).HotSpotEditor.HotSpotsFile Then
                    lblLC.caption = IIf(UBound(Project.Toolbars) > 1, GetLocalizedStr(767), GetLocalizedStr(766)) + " (" + FramesInfo.Frames(i).Name + ")"
                    Exit For
                End If
            Next i
        End If
    End If
    
    ValidFiles = Split(Mid(SupportedHTMLDocs, InStr(SupportedHTMLDocs, "(") + 1), "|")(0)
    ValidFiles = Left(ValidFiles, Len(ValidFiles) - 1)
    
    CompileProject MenuGrps, MenuCmds, Project, Preferences, params, False, True
      
    lvFiles.ListItems.Clear

    tvFolders.Nodes.Clear
    tvFolders.Nodes.Add , , "[ROOT]", "Root Web", 2
    
    tvFLC.Nodes.Clear
    tvFLC.Nodes.Add , , "[ROOT]", "Root Web", 2
    
    tvLC.Nodes.Clear
    tvLC.Nodes.Add , , "[ROOT]", "Root Web", 2
    
    ReDim Files(0)

    RootWeb = GetRealLocal.RootWeb
    AddFolder RootWeb
    
    For i = 1 To UBound(Files)
        UpdateStatusbar GetLocalizedStr(774) + ": " + Files(i).fName
        AddFile Files(i), False, tvFolders, lctAll
    Next i
    
    OptimizeFoldersDisplay tvFolders, False
    OptimizeFoldersDisplay tvLC
    OptimizeFoldersDisplay tvFLC
    
    LoadFilesForFolder tvFolders.Nodes(1)

    SetDlgState True

End Sub

Private Sub AutoDetect()

    Dim i As Long
    Dim lnk As String
    Dim nf As Long
    Dim a() As String
    Dim j As Integer
    Dim tf As String
        
    SetDlgState False

    For i = 1 To UBound(MenuCmds)
        With MenuCmds(i).Actions
            If .onclick.Type = atcURL Or .onclick.Type = atcNewWindow Then
                lnk = .onclick.url
                lnk = Replace(lnk, Project.UserConfigs(Project.DefaultConfig).RootWeb, RootWeb)
                If InStr(lnk, RootWeb) Then
                    nf = AddFile2Array(lnk)
                    If nf > 0 Then
                        UpdateStatusbar GetLocalizedStr(775) + " " + MenuCmds(i).Name + ": " + Files(nf).fName
                        If NeedsLC(Files(nf)) Then
                            If UsesFrames Then
                                AddFile Files(nf), True, tvFLC, lctAll, True
                            Else
                                AddFile Files(nf), True, tvLC, lctAll, True
                            End If
                        End If
                    End If
                End If
            End If
        End With
    Next i
    
    For i = 1 To UBound(MenuGrps)
        With MenuGrps(i).Actions
            If .onclick.Type = atcURL Or .onclick.Type = atcNewWindow Then
                lnk = .onclick.url
                lnk = Replace(lnk, Project.UserConfigs(Project.DefaultConfig).RootWeb, RootWeb)
                If InStr(lnk, RootWeb) Then
                    nf = AddFile2Array(lnk)
                    If nf > 0 Then
                        If NeedsLC(Files(nf)) Then
                            UpdateStatusbar GetLocalizedStr(776) + " " + MenuGrps(i).Name + ": " + Files(nf).fName
                            If UsesFrames Then
                                AddFile Files(nf), True, tvFLC, lctAll, True
                            Else
                                AddFile Files(nf), True, tvLC, lctAll, True
                            End If
                        End If
                    End If
                End If
            End If
        End With
    Next i
    
    nf = AddFile2Array(GetRealLocal.HotSpotEditor.HotSpotsFile)
    If nf > 0 Then
        If NeedsLC(Files(nf)) Then
            AddFile Files(nf), True, tvLC, lctAll, True
        End If
    End If
    
    If UsesFrames Then
        nf = AddFile2Array(FramesInfo.Frames(MenusFrame).srcFile)
        If nf > 0 Then
            If NeedsLC(Files(nf)) Then
                AddFile Files(nf), True, tvFLC, lctAll, True
            End If
        End If
    End If
    
    For i = 1 To UBound(Files)
        If Files(i).fType <> lctInvalid Then
            If UsesFrames Then
                a = Split(Replace(Files(i).fSource, "<a" + vbCrLf, "<a "), "<a ")
                For j = 1 To UBound(a)
                    lnk = GetParamVal(a(j), "href")
                    If LenB(lnk) <> 0 Then
                        lnk = ObtainAbsPath(Files(i).fName, lnk)
                        If FileExists(lnk) Then
                            If MatchSpec(lnk, ValidFiles) Then
                                If LenB(lnk) <> 0 Then
                                    tf = GetParamVal(a(j), "target")
                                    If LenB(tf) <> 0 Then
                                        If Right(FramesInfo.Frames(MenusFrame).Name, Len(tf)) = tf Then
                                            If NeedsLC(Files(i)) Then
                                                nf = AddFile2Array(lnk)
                                                AddFile Files(nf), True, tvFLC, lctAll
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next j
            End If
            
            UpdateStatusbar GetLocalizedStr(777) + ": " + Files(i).fName
            AddFile Files(i), False, tvFolders, lctAll
            If NeedsLC(Files(i)) Then
                If Not UsesFrames Then
                    'If Files(i).fType = lctNone Then
                        AddFile Files(i), True, tvLC, lctAll
                    'End If
                Else
                    Select Case Files(i).fType
                        Case lctFrames
                            AddFile Files(i), True, tvFLC, lctAll
                        Case lctStandard
                            AddFile Files(i), True, tvLC, lctAll
                    End Select
                End If
            End If
        End If
    Next i
    
    tvLC.Nodes(1).Expanded = True
    tvFLC.Nodes(1).Expanded = True
    
    SetDlgState True

End Sub

Private Sub AddFolder(FolderName As String, Optional doRecursion As Boolean = True)

    Dim dFile As String
    Dim Folders() As String
    Dim i As Integer
    Static recursion As Integer
    
    On Error Resume Next
    
    DoEvents
    
    recursion = recursion + 1
    
    Screen.MousePointer = vbHourglass
    
    ReDim Folders(0)
    
    If InStr(1, FolderName, "_vti_", vbTextCompare) > 0 Then GoTo ExitSub
    If InStr(1, FolderName, "_private", vbTextCompare) > 0 Then GoTo ExitSub
    If InStr(1, FolderName, "_themes", vbTextCompare) > 0 Then GoTo ExitSub
    If InStr(1, FolderName, "_scriptlibrary", vbTextCompare) > 0 Then GoTo ExitSub
    If InStr(1, FolderName, "_derived", vbTextCompare) > 0 Then GoTo ExitSub
    If InStr(1, FolderName, "_overlay", vbTextCompare) > 0 Then GoTo ExitSub
    If InStr(1, FolderName, "_fpclass", vbTextCompare) > 0 Then GoTo ExitSub
    If InStr(1, FolderName, "_colab", vbTextCompare) > 0 Then GoTo ExitSub
    If recursion = 2 And InStr(1, FolderName, "\sitebuilder\", vbTextCompare) > 0 Then GoTo ExitSub
    
    UpdateStatusbar GetLocalizedStr(778) + ": " + EllipseText(txtDummy, FolderName, DT_PATH_ELLIPSIS)
    
    dFile = Dir(FolderName, vbDirectory Or vbHidden Or vbSystem)
    Do While LenB(dFile) <> 0
        Err.Clear
        If dFile <> "." And dFile <> ".." And ((GetAttr(FolderName + dFile) And vbDirectory) = vbDirectory) Then
            If Err.number = 0 Then
                If (dFile = "_vti_pvt") And (FolderName <> RootWeb) And Not IncludeFPSubWebs Then GoTo ExitSub
                
                'If Not ((dFile = "_vti_pvt") And (FolderName <> RootWeb) And Not IncludeFPSubWebs) Then
                    ReDim Preserve Folders(UBound(Folders) + 1)
                    Folders(UBound(Folders)) = FolderName + dFile
                'End If
            Else
                Debug.Print "Error scanning: " + FolderName + dFile
            End If
        End If
        dFile = Dir
    Loop
    
    Dim hasFiles As Boolean
    hasFiles = hasValidFiles(FolderName)
    
    If doRecursion Then
        For i = 1 To UBound(Folders)
            AddFolder AddTrailingSlash(Folders(i), "\"), Not hasFiles
        Next i
    End If
    
    dFile = Dir(FolderName + "*.*")
    Do While LenB(dFile) <> 0
        If InStr(dFile, ".") > 0 Then AddFile2Array FolderName + dFile
        dFile = Dir
    Loop
    
ExitSub:
    
    recursion = recursion - 1
    
    If recursion = 0 Then
        Screen.MousePointer = vbDefault
    End If
    
End Sub

Private Function hasValidFiles(ByVal FolderName As String) As Boolean

    Dim dFile As String
    
    dFile = Dir(FolderName + "*.*")
    Do While LenB(dFile) <> 0
        If InStr(1, ValidFiles, GetFileExtension(dFile) + ";", vbTextCompare) > 0 Then
            hasValidFiles = True
            Exit Function
        End If
        dFile = Dir
    Loop
    
    hasValidFiles = False

End Function

Private Function AddFile2Array(ByVal FileName As String) As Long

    Dim i As Long
    Dim l As Long
    Dim k As Integer

    If Left(GetFileName(FileName), 5) = "_vti_" Then Exit Function
    If Not FileExists(FileName) Then Exit Function
    If Not ShowAllFiles Then
        If Not MatchSpec(FileName, ValidFiles) Then Exit Function
    End If
    
    l = UBound(Files)
    For i = 1 To l
        If Files(i).fName = FileName Then
            AddFile2Array = i
            Exit Function
        End If
    Next i

    l = l + 1
    ReDim Preserve Files(l)
    With Files(l)
        .fName = FileName
        If Not MatchSpec(FileName, ValidFiles) Then
            .fType = lctInvalid
        Else
            .fSource = LoadFile(FileName)
            .fType = GetLCType(.fSource)
            If (GetAttr(FileName) And vbReadOnly) = vbReadOnly Then .fType = lctReadOnly
            .fXHTMLCompliance = (InStr(.fSource, "W3C//DTD XHTML") > 0)
            .fYahooSiteBuilderSupport = (InStr(.fSource, "<!--$sitebuilder version=") > 0)
        End If
        .fIndex = l
    End With
    
    If UsesFrames Then
        If Files(l).fName <> Project.UserConfigs(Project.DefaultConfig).HotSpotEditor.HotSpotsFile Then
            For k = 1 To UBound(FramesInfo.Frames)
                If FramesInfo.Frames(k).srcFile = Files(l).fName And MenusFrame <> k Then
                    Files(l).fType = lctInvalid
                End If
            Next k
        End If
    End If
    
    AddFile2Array = l

End Function

Private Function GetLCType(fSource As String) As LCType
    
    On Error Resume Next
    
    If (InStr(LCase(fSource), "<frameset") > 0) Then
        GetLCType = lctInvalid
        Exit Function
    End If
    
    If InStr(fSource, LoaderCodeSTART) = 0 Then
        GetLCType = lctNone
        Exit Function
    End If
    
    If InStr(fSource, "_frames.js") Then
        GetLCType = lctFrames
        Exit Function
    End If
    
    GetLCType = lctStandard

End Function

Private Sub AddFile(File As CFile, IncludeFiles As Boolean, TargetTV As TreeView, FilterLC As LCType, Optional ExpandFolders As Boolean = False)

    Dim f() As String
    Dim i As Integer
    Dim pNode As Node
    Dim fp As String
    Dim nItem As Object
    Dim FileName As String
    
    On Error Resume Next
    
    FileName = Mid(File.fName, Len(RootWeb) + 1)
    f = Split(FileName, "\")
    
    Set pNode = TargetTV.Nodes("[ROOT]")
    For i = 0 To UBound(f) - 1
        Err.Clear
        fp = fp + f(i) + "\"
        Set pNode = TargetTV.Nodes.Add(pNode.Index, tvwChild, "K" + fp, f(i), FolderIconIdx)
        If Err.number <> 0 Then
            Set pNode = TargetTV.Nodes("K" + fp)
        Else
            Err.Clear
        End If
        If ExpandFolders Then pNode.Expanded = True
    Next i
    
    If IncludeFiles Then
        Set nItem = TargetTV.Nodes.Add(pNode.Index, tvwChild, pNode.FullPath + "\" + "K" + fp + f(i), f(i))
        nItem.tag = File.fIndex
        nItem.Image = GetImage(File.fName)
        SetItemAppearance nItem, File.fType
    End If
    
End Sub

Private Function NeedsLC(File As CFile) As Boolean

    Dim cCode As String
    
    If GetLCType(File.fSource) = lctInvalid Then Exit Function
    
    cCode = GenLoaderCode(File.fType = lctFrames, False, File.fName, Join(SelSecProjects, "|"), File.fXHTMLCompliance, File.fYahooSiteBuilderSupport)
    
    NeedsLC = (InStr(1, File.fSource, cCode, vbTextCompare) = 0)

End Function

Private Function GetItemDescription(fExt As String) As String

    Dim kName As String
    
    kName = QueryValue(HKEY_CLASSES_ROOT, fExt)
    GetItemDescription = QueryValue(HKEY_CLASSES_ROOT, kName)

End Function

Private Sub SetItemAppearance(nItem As Object, fType As LCType)
    
    On Error Resume Next
    
    Dim ItemColor As Long
    Dim TypeStr As String
    Dim IsGhosted As Boolean
    
    Select Case fType
        Case lctInvalid
            ItemColor = DimmedColor
            TypeStr = GetLocalizedStr(779)
            IsGhosted = True
        Case lctNone
            ItemColor = vbBlack
            TypeStr = GetLocalizedStr(782)
        Case lctFrames
            ItemColor = &H8000&
            TypeStr = GetLocalizedStr(781)
        Case lctStandard
            ItemColor = vbBlue
            TypeStr = GetLocalizedStr(780)
        Case lctReadOnly
            ItemColor = DimmedColor
            TypeStr = "[File is Read Only]"
            IsGhosted = True
    End Select
    
    nItem.ForeColor = ItemColor
    If TypeOf nItem Is ListItem Then
        nItem.SubItems(lvFiles.ColumnHeaders("chLC").Index - 1) = TypeStr
        nItem.Ghosted = IsGhosted
    End If

End Sub

Private Function GetRealFileName(f As String) As String

    GetRealFileName = Replace(Replace(f, tvFolders.Nodes(1).Text, RootWeb), "\\", "\")
    
    If Left(RootWeb, 2) = "\\" Then GetRealFileName = "\" + GetRealFileName

End Function

Private Function GetImage(FileName As String) As Integer

    Dim iIdx As Integer
    Dim fExt As String
    
    If LenB(FileName) = 0 Then
        FileName = TempPath
    Else
        fExt = GetFileExtension(FileName)
    End If
    
    On Error Resume Next
    iIdx = ilIcons.ListImages(fExt).Index
    
    If iIdx = 0 Then
        GetIcon picIcon, FileName
        picIcon.Picture = picIcon.Image
        ilIcons.ListImages.Add , fExt, picIcon
        iIdx = ilIcons.ListImages.Count
    End If
    
    GetImage = iIdx

End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    SaveSetting App.EXEName, "LCMWinPos", "X", Left
    SaveSetting App.EXEName, "LCMWinPos", "Y", Top
    SaveSetting App.EXEName, "LCMWinPos", "W", Width
    SaveSetting App.EXEName, "LCMWinPos", "H", Height
    SaveSetting App.EXEName, "LCMWinPos", "S0", picFFSplit.Top
    SaveSetting App.EXEName, "LCMWinPos", "S1", picSplit.Left
    If UsesFrames Then
        SaveSetting App.EXEName, "LCMWinPos", "S2", picLCSplit.Top
    End If
    SaveSetting App.EXEName, "LCMWinPos", "Filter", chkFilter.Value
    SaveSetting App.EXEName, "LCMWinPos", "IncludeFPSubWebs", chkIncludeFPSubWebs.Value

    If Not IsDebug Then SetupSubclassing False

End Sub

Private Sub Form_Resize()

    Dim cTop As Long
    Dim cLeft As Long
    Dim midHeight As Long
    Const PanesHeight = 176 * 15
    
    On Error Resume Next
    
    If Me.WindowState = vbMinimized Then Exit Sub
    
    IsResizing = True
    
    cTop = GetClientTop(hwnd)
    cLeft = GetClientTop(hwnd)
    
    midHeight = (Height - sbInfo.Height - cTop - 1170) / 2
    
    lblFolders.Move 195, 135
    With tvFolders
        .Move 90, lblFolders.Top + lblFolders.Height + 60, picSplit.Left - 135, picFFSplit.Top - 450
        lblFiles.Move lblFolders.Left, picFFSplit.Top + picFFSplit.Height + 90, .Width
        lvFiles.Move .Left, lblFiles.Top + lblFiles.Height + 60, .Width, Height - .Height - PanesHeight
        
        picFFSplit.Move .Left, picFFSplit.Top, .Width, 120
    End With
    
    If Not UsesFrames Then
        picLCSplit.Top = lvFiles.Top + lvFiles.Height + 60
    End If
    
    lblLC.Move picSplit.Left + picSplit.Width + 150, lblFolders.Top
    tvLC.Move lblLC.Left - (lblFolders.Left - tvFolders.Left), tvFolders.Top, Width - lblLC.Left - (picSplit.Width + 30), picLCSplit.Top - 450
    
    lblFLC.Move lblLC.Left, picLCSplit.Top + picLCSplit.Height + 90
    tvFLC.Move tvLC.Left, lblFLC.Top + lblFLC.Height + 60, tvLC.Width, Height - tvLC.Height - PanesHeight
    
    picSplit.Move picSplit.Left, lblFolders.Top, 120, midHeight * 2
    picLCSplit.Move tvLC.Left, picLCSplit.Top, tvLC.Width, 120
    
    With cmdClose
        .Move Width - (.Width + 225), Height - (sbInfo.Height + .Height + cTop + 45)
        cmdInstall.Move .Left - (cmdInstall.Width + 165), .Top
        cmdAuto.Move 75, .Top
        chkFilter.Move cmdAuto.Left, lvFiles.Top + lvFiles.Height + 90
        chkIncludeFPSubWebs.Move cmdAuto.Left, chkFilter.Top + chkFilter.Height + 30
    End With
    
    txtDummy.Width = tvFolders.Width + tvLC.Width / 2
    
    UpdateFilesLabel
    
    AutoSizeCols
    
    frmLCManMinWidth = (lblLC.Left + lblLC.Width) / Screen.TwipsPerPixelX + 25
    frmLCManMinHeight = picFFSplit.Top / Screen.TwipsPerPixelY + 200
    
    IsResizing = False
    
    Refresh

End Sub

Private Sub lvFiles_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    SortFilesList ColumnHeader.Index

End Sub

Private Sub SortFilesList(Optional ByVal ColumnHeaderIndex As Integer = -1)

    Dim c As Integer
    Dim oa As Integer
    Dim sIdx As Integer
    Static SortIndex As Integer

    With lvFiles
        For c = 1 To .ColumnHeaders.Count
            .ColumnHeaders(c).Icon = 3
            oa = .ColumnHeaders(c).Alignment
            .ColumnHeaders(c).Alignment = 0
            .ColumnHeaders(c).Alignment = oa
        Next c
    End With
    
    If ColumnHeaderIndex = -1 Then
        If SortIndex = 0 Then SortIndex = 1
        ColumnHeaderIndex = SortIndex
        lvFiles.SortOrder = IIf(lvFiles.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        SortIndex = ColumnHeaderIndex
    End If
    
    Select Case ColumnHeaderIndex
        Case 1 'Name
            sIdx = -1
        Case 2 'Size
            sIdx = 1
        Case 3 'Type
            sIdx = -1
        Case 4 'Date
            sIdx = 0
        Case 5 'Loader Code
            sIdx = -1
    End Select
    
    SortListView lvFiles, sIdx, ColumnHeaderIndex
    lvFiles.ColumnHeaders(ColumnHeaderIndex + 1).Icon = IIf(lvFiles.SortOrder = lvwAscending, 2, 1)

End Sub

Private Sub lvFiles_ItemClick(ByVal item As MSComctlLib.ListItem)

    Dim lcInfo As String
    Dim f As CFile
    Dim lStr As Integer
    Dim i As Integer
    Dim ipStr As String
    Dim spInfo() As String
    Dim HasSecPrjs As Boolean
    
    f = Files(Val(item.tag))
    
    Select Case f.fType
        Case lctFrames
            lStr = 783
        Case lctNone
            lStr = 784
        Case lctStandard
            lStr = 785
        Case lctInvalid
            lStr = 786
    End Select
    
    lcInfo = ": " + GetLocalizedStr(lStr)
    
    With Project
        ipStr = " for the '" + Project.Name
        If UBound(.SecondaryProjects) > 0 Then
            For i = 1 To UBound(.SecondaryProjects)
                spInfo = GetPrjInfo(.SecondaryProjects(i))
                If InStr(f.fSource, spInfo(1)) Then
                    HasSecPrjs = True
                    ipStr = ipStr + " + " + spInfo(2)
                End If
            Next i
            ipStr = ipStr + "' project" + IIf(HasSecPrjs, "s", "")
        Else
            ipStr = ipStr + "' project"
        End If
    End With
    
    UpdateStatusbar EllipseText(txtDummy, GetFileName(f.fName), DT_PATH_ELLIPSIS) + lcInfo + ipStr

End Sub

Private Function GetPrjInfo(f As String) As String()

    Dim sPrj As ProjectDef
    Dim sStr(1 To 2) As String
    
    sPrj = GetProjectProperties(f, False)
    Close #ff
    sStr(1) = sPrj.JSFileName
    sStr(2) = sPrj.Name
    
    GetPrjInfo = sStr

End Function

Private Sub UpdateStatusbar(sStr As String)

    sbInfo.Panels("pInfo").Text = sStr

End Sub

Private Sub lvFiles_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim sItem As ListItem
    Dim nItem As ListItem
    
    On Error Resume Next
    
    Set sItem = lvFiles.HitTest(x, y)
    If Not sItem Is Nothing Then
        If Button = vbRightButton Then
        
            If Not sItem.Selected Then
                For Each nItem In lvFiles.ListItems
                    nItem.Selected = False
                Next nItem
                sItem.Selected = True
                sItem.EnsureVisible
            End If
        
            LoadFilesForFolder sItem
            mnuContextILC.Enabled = Not sItem.Ghosted
            mnuContextIFLC.Enabled = Not sItem.Ghosted
            mnuContextRemove.caption = GetLocalizedStr(769)
            mnuContextRemove.Enabled = Files(sItem.tag).fType = lctFrames Or Files(sItem.tag).fType = lctStandard
            PopupMenu mnuContext, vbRightButton, lvFiles.Left + x, lvFiles.Top + y
        Else
            'sItem.Selected = True
            sItem.EnsureVisible
        End If
    End If

End Sub

Private Sub lvFiles_OLEStartDrag(data As MSComctlLib.DataObject, AllowedEffects As Long)

    PopulateDragData
    
    AllowedEffects = vbDropEffectCopy

End Sub

Private Sub PopulateDragData()

    Dim nItem As ListItem
    
    ReDim DragData(0)
    
    For Each nItem In lvFiles.ListItems
        If nItem.Selected Then
            ReDim Preserve DragData(UBound(DragData) + 1)
            DragData(UBound(DragData)) = Files(Val(nItem.tag)).fName
        End If
    Next nItem

End Sub

Private Sub mnuContextIFLC_Click()
    
    If ActiveControl.Name = "tvFolders" Then
        ReDim DragData(1)
        DragData(1) = AddTrailingSlash(GetRealFileName(tvFolders.SelectedItem.FullPath), "\")
    Else
        PopulateDragData
    End If
    EnQueueFiles tvFLC

End Sub

Private Sub mnuContextILC_Click()

    If ActiveControl.Name = "tvFolders" Then
        ReDim DragData(1)
        DragData(1) = AddTrailingSlash(GetRealFileName(tvFolders.SelectedItem.FullPath), "\")
    Else
        PopulateDragData
    End If
    EnQueueFiles tvLC

End Sub

Private Sub mnuContextRemove_Click()

    Dim sNode As Node
    Dim sItem As ListItem
    Dim DoAll As Boolean
    Dim lNode As Node
    
    SetDlgState False
    
    Select Case ActiveControl.Name
        Case "tvLC", "tvFLC"
            With ActiveControl
                If .SelectedItem.Index = 1 Then
                    While .Nodes.Count > 1
                        .Nodes.Remove 2
                    Wend
                Else
                    .Nodes.Remove .SelectedItem.Index
                End If
            End With
        Case Else
            DoAll = (ActiveControl.Name = "tvFolders")
            Set lNode = tvFolders.SelectedItem
            For Each sNode In tvFolders.Nodes
                If Left(sNode.FullPath, Len(lNode.FullPath)) = lNode.FullPath Then
                    If DoAll Then LoadFilesForFolder sNode
                    For Each sItem In lvFiles.ListItems
                        If sItem.Selected Or DoAll Then
                            With Files(sItem.tag)
                                If LenB(.fSource) <> 0 Then
                                    .fSource = RemoveLoaderCode(.fSource, .fName, .fYahooSiteBuilderSupport)
                                    SaveFile .fName, .fSource
                                    .fType = lctNone
                                End If
                            End With
                        End If
                    Next sItem
                End If
            Next sNode
            lNode.EnsureVisible
            lNode.Selected = True
            LoadFilesForFolder lNode
    End Select
    
    SetDlgState True
    
End Sub

Private Sub tmrInit_Timer()

    tmrInit.Enabled = False
    
    If SecProjMode = spmcUndefined Then
        Unload Me
    Else
        Form_Resize
        
        DoEvents
        
        LoadFiles
    End If
    
End Sub

Private Sub OptimizeFoldersDisplay(tv As TreeView, Optional RemoveEmptyFolders As Boolean = True)

    Dim sNode As Node

    With tv.Nodes(1)
        .Selected = True
        .EnsureVisible
        .Expanded = True
    End With
    
    If Not RemoveEmptyFolders Then Exit Sub
    
ReStart:
    For Each sNode In tv.Nodes
        If sNode.children = 0 And sNode.Image = 1 Then
            tv.Nodes.Remove sNode.Index
            GoTo ReStart
        End If
    Next sNode

End Sub

Private Sub tvFLC_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim sItem As Node
    
    On Error Resume Next
    
    With tvFLC
        Set sItem = .HitTest(x, y)
        If Not sItem Is Nothing And Button = vbRightButton Then
            sItem.Selected = True
            sItem.EnsureVisible
            mnuContextILC.Enabled = False
            mnuContextIFLC.Enabled = False
            mnuContextRemove.caption = GetLocalizedStr(787)
            mnuContextRemove.Enabled = True
            PopupMenu mnuContext, vbRightButton, .Left + x, .Top + y
        End If
    End With
    
End Sub

Private Sub tvFolders_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim sItem As Node
    
    Set sItem = tvFolders.HitTest(x, y)
    If sItem Is Nothing Then Exit Sub

    ' -----------------------------------------------------------------

    LoadFolderContents sItem
    
    ' -------------------------------------------------------------------

    On Error Resume Next
    
    sItem.Selected = True
    sItem.EnsureVisible
    If Button = vbRightButton Then
        LoadFilesForFolder sItem
        mnuContextILC.Enabled = True
        mnuContextIFLC.Enabled = True
        mnuContextRemove.caption = GetLocalizedStr(769)
        mnuContextRemove.Enabled = True
        PopupMenu mnuContext, vbRightButton, tvFolders.Left + x, tvFolders.Top + y
    End If

End Sub

Private Sub LoadFolderContents(sItem As Node)

    If sItem.tag = 1 Then Exit Sub
    sItem.tag = 1
    
    Dim Path As String
    Path = AddTrailingSlash(Replace(sItem.FullPath, "Root Web", RootWeb), "\")
    If Left(Path, 2) = "\\" Then
        Path = "\" + Replace(Path, "\\", "\")
    Else
        Path = Replace(Path, "\\", "\")
    End If
    AddFolder Path
    
    Dim i As Integer
    For i = 1 To UBound(Files)
        If InStr(Files(i).fName, Path) > 0 Then
            UpdateStatusbar GetLocalizedStr(774) + ": " + Files(i).fName
            AddFile Files(i), False, tvFolders, lctAll
        End If
    Next i

End Sub

Private Sub msgSubClass_NewMessage(ByVal hwnd As Long, uMsg As Long, wParam As Long, lParam As Long, Cancel As Boolean, UseRetVal As Boolean, RetVal As Long)

    If IsResizing Then Exit Sub

    Cancel = HandleSubclassMsg(hwnd, uMsg, wParam, lParam)

End Sub

Private Sub SetupSubclassing(scState As Boolean)

    If msgSubClass Is Nothing Then Set msgSubClass = New xfxSC
    
    frmLCManHWND = Me.hwnd
    msgSubClass.SubClassHwnd Me.hwnd, scState

End Sub

Private Sub tvFolders_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim sItem As Node
    
    On Error Resume Next
    
    Set sItem = tvFolders.HitTest(x, y)
    If Not sItem Is Nothing Then LoadFilesForFolder sItem

End Sub

Private Sub LoadFilesForFolder(fNode As Node)

    Dim i As Long
    Dim nItem As ListItem
    Dim sItem As ListSubItem
    Dim subPath As String
    
    subPath = AddTrailingSlash(GetRealFileName(fNode.FullPath), "\")

    lvFiles.ListItems.Clear
    lvFiles.MultiSelect = False
    For i = 1 To UBound(Files)
        With Files(i)
            If GetFilePath(.fName) = subPath Then
                Set nItem = lvFiles.ListItems.Add(, , GetFileName(.fName))
                nItem.SmallIcon = GetImage(.fName)
                nItem.tag = i
                
                nItem.SubItems(lvFiles.ColumnHeaders("chSize").Index - 1) = NiceBytes(FileLen(.fName), False, True)
                nItem.SubItems(lvFiles.ColumnHeaders("chType").Index - 1) = GetItemDescription("." + GetFileExtension(.fName))
                nItem.SubItems(lvFiles.ColumnHeaders("chDate").Index - 1) = FileDateTime(.fName)
                
                SetItemAppearance nItem, .fType
                
                For Each sItem In nItem.ListSubItems
                    With sItem
                        .ForeColor = nItem.ForeColor
                        '.Bold = nItem.Bold
                    End With
                Next sItem
            End If
        End With
    Next i
    lvFiles.MultiSelect = True
    AutoSizeCols
    SortFilesList
    
    UpdateFilesLabel
    
    If lvFiles.ListItems.Count > 0 Then
        lvFiles_ItemClick lvFiles.ListItems(1)
    Else
        UpdateStatusbar GetLocalizedStr(789)
    End If

End Sub

Private Sub AutoSizeCols()

    Dim ch As ColumnHeader
    
    LockWindowUpdate Me.hwnd
    
    For Each ch In lvFiles.ColumnHeaders
        ch.Text = String(10, " ") + ch.Text
    Next ch

    CoolListView lvFiles
    
    For Each ch In lvFiles.ColumnHeaders
        ch.Text = Trim(ch.Text)
    Next ch
    
    LockWindowUpdate 0

End Sub

Private Sub UpdateFilesLabel()

    Dim subPath As String
    Dim oWidth As Long
    
    subPath = AddTrailingSlash(GetRealFileName(tvFolders.SelectedItem.FullPath), "\")

    oWidth = txtDummy.Width
    txtDummy.Width = tvFolders.Width - TextWidth(GetLocalizedStr(764)) - 225
    lblFiles.caption = GetLocalizedStr(764) + " " + EllipseText(txtDummy, subPath, DT_PATH_ELLIPSIS)
    txtDummy.Width = oWidth

End Sub

Private Sub tvFolders_OLEStartDrag(data As MSComctlLib.DataObject, AllowedEffects As Long)

    LoadFilesForFolder tvFolders.SelectedItem
    
    ReDim DragData(1)
    DragData(1) = AddTrailingSlash(GetRealFileName(tvFolders.SelectedItem.FullPath), "\")
    
    AllowedEffects = vbDropEffectCopy

End Sub

Private Sub tvLC_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim sItem As Node
    
    On Error Resume Next
    
    With tvLC
        Set sItem = .HitTest(x, y)
        If Not sItem Is Nothing And Button = vbRightButton Then
            sItem.Selected = True
            sItem.EnsureVisible
            mnuContextILC.Enabled = False
            mnuContextIFLC.Enabled = False
            mnuContextRemove.caption = GetLocalizedStr(787)
            mnuContextRemove.Enabled = (.Nodes.Count > 1)
            PopupMenu mnuContext, vbRightButton, .Left + x, .Top + y
        End If
    End With

End Sub

Private Sub EnQueueFiles(TargetTV As TreeView)

          Dim i As Integer
          Dim f As Long
          Dim IsFrames As Boolean
          Dim OkToAdd As Boolean
          Dim InvFiles() As String
          Dim txtDummyWidth As Long
          Dim dtLen As Integer
          
10        On Error Resume Next
20        dtLen = UBound(DragData)
30        If dtLen = 0 Then Exit Sub
          
40        On Error GoTo EnQueueFiles_Error
          
50        ReDim InvFiles(0)
          
60        SetDlgState False
          
70        For i = 1 To dtLen
80            For f = 1 To UBound(Files)
90                With Files(f)
100                   If (.fName = DragData(i)) Or (Left(GetFilePath(.fName), Len(DragData(i))) = DragData(i)) Then
110                       If LenB(.fSource) = 0 Then
120                           ReDim Preserve InvFiles(UBound(InvFiles) + 1)
130                           InvFiles(UBound(InvFiles)) = .fName
140                       Else
150                           If IsFrames Then
160                               OkToAdd = True
170                           Else
180                               OkToAdd = (((lctStandard Or lctFrames Or lctNone) And .fType) = .fType)
190                           End If
200                           If OkToAdd Then
210                               UpdateStatusbar GetLocalizedStr(788) + ": " + .fName
220                               AddFile Files(f), True, TargetTV, lctAll, True
230                           End If
240                       End If
250                   End If
260               End With
270           Next f
280       Next i
          
290       TargetTV.Nodes(1).Expanded = True
          
300       If ShowAllFiles Then
310           If UBound(InvFiles) > 0 Then
320               txtDummyWidth = txtDummy.Width
330               txtDummy.Width = 4500
              
340               With TipsSys
350                   .CanDisable = False
360                   .TipTitle = "Some files could not be added"
370                   .Tip = "Some of the selected files could not be added because their type was not recognized." + vbCrLf + vbCrLf + "Files List:" + vbCrLf
380                   For i = 1 To UBound(InvFiles)
390                       .Tip = .Tip + EllipseText(txtDummy, InvFiles(i), DT_PATH_ELLIPSIS) + vbCrLf
400                   Next i
410                   .Tip = .Tip + vbCrLf + "To avoid receiving this warning in the future enable the option 'Hide Unsupported Documents'"
420                   .Show
430               End With
                  
440               txtDummy.Width = txtDummyWidth
450           End If
460       End If
          
ExitSub:
470       SetDlgState True

480       On Error GoTo 0
490       Exit Sub

EnQueueFiles_Error:

500       MsgBox "Error " & Err.number & " in line " & Erl & ": " & vbCrLf & Err.Description & " in Form.frmLCMan.EnQueueFiles"
    GoTo ExitSub:
          
End Sub

Private Sub tvLC_OLEDragDrop(data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)

    EnQueueFiles tvLC

End Sub

Private Sub tvFLC_OLEDragDrop(data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)

    EnQueueFiles tvFLC

End Sub

Private Sub picFFSplit_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    IsSplittingFF = True

End Sub

Private Sub picFFSplit_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim newVal As Long
    Static LastVal As Long

    If IsSplittingFF Then
    
        newVal = picFFSplit.Top + y
        
        If newVal = LastVal Then Exit Sub
        If newVal < 750 Then newVal = 750
        If (Height - newVal) < 3652 Then newVal = Height - 3652
        
        LastVal = newVal
    
        picFFSplit.Top = newVal
        Form_Resize
    End If

End Sub

Private Sub picFFSplit_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    IsSplittingFF = False

End Sub

Private Sub picSplit_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    IsSplitting = True

End Sub

Private Sub picSplit_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim newVal As Long
    Static LastVal As Long

    If IsSplitting Then
    
        newVal = picSplit.Left + x
        
        If newVal = LastVal Then Exit Sub
        If (newVal + lblLC.Width + 25 * Screen.TwipsPerPixelX > Width - 2 * GetClientLeft(Me.hwnd)) Then
            newVal = Width - 2 * GetClientLeft(Me.hwnd) - lblLC.Width - 25 * Screen.TwipsPerPixelX
        End If
        If newVal < lblFolders.Width + lblFolders.Left + 10 * Screen.TwipsPerPixelX Then
            newVal = lblFolders.Width + lblFolders.Left + 10 * Screen.TwipsPerPixelX
        End If
        
        LastVal = newVal
    
        picSplit.Left = newVal
        Form_Resize
    End If

End Sub

Private Sub picSplit_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    IsSplitting = False

End Sub

Private Sub picLCSplit_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    IsSplittingLC = True

End Sub

Private Sub picLCSplit_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim newVal As Long
    Static LastVal As Long

    If IsSplittingLC Then
    
        newVal = picLCSplit.Top + y
        
        If newVal = LastVal Then Exit Sub
        If newVal < 645 Then Exit Sub
        If (Height - newVal) < 2105 Then Exit Sub
        
        LastVal = newVal
    
        picLCSplit.Top = newVal
        Form_Resize
    End If

End Sub

Private Sub picLCSplit_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    IsSplittingLC = False

End Sub

Private Sub LocalizeUI()

    Dim ch() As String
    
    lblFolders.caption = GetLocalizedStr(763)
    lblFiles.caption = GetLocalizedStr(764)
    lblLC.caption = IIf(UBound(Project.Toolbars) > 1, GetLocalizedStr(767), GetLocalizedStr(766))
    lblFLC.caption = GetLocalizedStr(765)
    
    cmdInstall.caption = GetLocalizedStr(468)
    cmdClose.caption = GetLocalizedStr(424)
    
    On Error Resume Next
    ch = Split(GetLocalizedStr(790), "|")
    If UBound(ch) = 4 Then
        lvFiles.ColumnHeaders("chName").Text = ch(0)
        lvFiles.ColumnHeaders("chSize").Text = ch(1)
        lvFiles.ColumnHeaders("chType").Text = ch(2)
        lvFiles.ColumnHeaders("chDate").Text = ch(3)
        lvFiles.ColumnHeaders("chLC").Text = ch(4)
    End If

End Sub
