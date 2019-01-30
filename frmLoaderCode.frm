VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{9A2947B4-87EE-40EE-A3EF-32BDC32D5726}#1.0#0"; "xfxline3d.ocx"
Begin VB.Form frmLCInstall 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Install Loader Code"
   ClientHeight    =   3405
   ClientLeft      =   5835
   ClientTop       =   5055
   ClientWidth     =   6255
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   6255
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtDummy 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   90
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1275
      Visible         =   0   'False
      Width           =   4800
   End
   Begin VB.CommandButton cmdFromLinks 
      Caption         =   "Add Links"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5040
      TabIndex        =   6
      Top             =   1815
      Width           =   1155
   End
   Begin VB.CommandButton cmdAddFiles 
      Caption         =   "Add Files..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5040
      TabIndex        =   3
      Top             =   990
      Width           =   1155
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove Files"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5040
      TabIndex        =   7
      Top             =   2340
      Width           =   1155
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5040
      TabIndex        =   10
      Top             =   2970
      Width           =   1155
   End
   Begin VB.CommandButton cmdInstall 
      Caption         =   "Install"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5040
      TabIndex        =   2
      Top             =   300
      Width           =   1155
   End
   Begin VB.CommandButton cmdAddFolder 
      Caption         =   "Add Folder..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5040
      TabIndex        =   5
      Top             =   1395
      Width           =   1155
   End
   Begin MSComDlg.CommonDialog cDlg 
      Left            =   195
      Top             =   2310
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ilIcons 
      Left            =   750
      Top             =   2220
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
            Picture         =   "frmLoaderCode.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLoaderCode.frx":06D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLoaderCode.frx":0DA4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin xfxLine3D.ucLine3D uc3DLine1 
      Height          =   30
      Left            =   15
      TabIndex        =   8
      Top             =   2820
      Width           =   6165
      _ExtentX        =   10874
      _ExtentY        =   53
   End
   Begin MSComctlLib.TreeView tvFiles 
      Height          =   2400
      Left            =   45
      TabIndex        =   1
      Top             =   300
      Width           =   4905
      _ExtentX        =   8652
      _ExtentY        =   4233
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   529
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      ImageList       =   "ilIcons"
      Appearance      =   1
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
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      Caption         =   "Installing Loader Code:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   45
      TabIndex        =   9
      Top             =   2910
      UseMnemonic     =   0   'False
      Width           =   4905
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Files to Install the Loader Code"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   45
      TabIndex        =   0
      Top             =   75
      Width           =   2235
   End
End
Attribute VB_Name = "frmLCInstall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OriginalConfig As Integer

Private Sub cmdAddFiles_Click()

    On Error GoTo ExitSub
    
    Dim f() As String
    Dim i As Integer
    
    With cDlg
        .DialogTitle = GetLocalizedStr(754)
        .InitDir = Project.UserConfigs(Project.DefaultConfig).RootWeb
        .filter = SupportedHTMLDocs
        .CancelError = True
        .Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly + cdlOFNLongNames + cdlOFNPathMustExist + cdlOFNAllowMultiselect + cdlOFNExplorer
        .ShowOpen
        If InStr(.FileName, Chr(0)) Then
            f = Split(.FileName, Chr(0))
            For i = 1 To UBound(f) - 1
                AddFile f(0) + "\" + f(i)
            Next i
        Else
            AddFile .FileName
        End If
    End With
    
ExitSub:
    Exit Sub

End Sub

Private Sub cmdAddFolder_Click()
    
    Dim Path As String
    
    Me.Enabled = False
    
    Path = UnqualifyPath(Project.UserConfigs(Project.DefaultConfig).RootWeb)
    Path = BrowseForFolderByPath(Path, GetLocalizedStr(752), Me)
    
    DoEvents
    
    MousePointer = vbHourglass
    If LenB(Path) <> 0 Then AddFolder AddTrailingSlash(Path, "\")
    MousePointer = vbDefault
    
    Me.Enabled = True
    Me.SetFocus

End Sub

Private Sub cmdClose_Click()

    RestoreCommandsLinks
    Unload Me

End Sub

Private Sub cmdFromLinks_Click()

    Dim i As Integer
    Dim RootWeb As String
    
    RootWeb = Project.UserConfigs(Project.DefaultConfig).RootWeb
    
    For i = 1 To UBound(MenuCmds)
        With MenuCmds(i).Actions
            If .onclick.Type = atcURL Then
                If Left(.onclick.url, Len(RootWeb)) = RootWeb Then
                    AddFile .onclick.url
                End If
            End If
            If .onmouseover.Type = atcURL Then
                If Left(.onmouseover.url, Len(RootWeb)) = RootWeb Then
                    AddFile .onmouseover.url
                End If
            End If
            If .OnDoubleClick.Type = atcURL Then
                If Left(.OnDoubleClick.url, Len(RootWeb)) = RootWeb Then
                    AddFile .OnDoubleClick.url
                End If
            End If
        End With
    Next i
    
    For i = 1 To UBound(MenuGrps)
        With MenuGrps(i).Actions
            If .onclick.Type = atcURL Then
                If Left(.onclick.url, Len(RootWeb)) = RootWeb Then
                    AddFile .onclick.url
                End If
            End If
            If .onmouseover.Type = atcURL Then
                If Left(.onmouseover.url, Len(RootWeb)) = RootWeb Then
                    AddFile .onmouseover.url
                End If
            End If
            If .OnDoubleClick.Type = atcURL Then
                If Left(.OnDoubleClick.url, Len(RootWeb)) = RootWeb Then
                    AddFile .OnDoubleClick.url
                End If
            End If
        End With
    Next i

End Sub

Private Sub cmdInstall_Click()

    DoInstall
    cmdClose_Click

End Sub

Private Sub DoInstall()

    Dim sNode As Node
    Dim LoaderCode As String
    Dim nLoaderCode As String
    Dim ff As Integer
    Dim sCode As String
    Dim dFile As String
    
    On Error GoTo chkError
    
    Me.Enabled = False
    
    RestoreCommandsLinks
    CompileProject MenuGrps, MenuCmds, Project, Preferences, params, False, True
    LoaderCode = GenLoaderCode(False, False)
    ForceCommandsLinks2Local
    
ReStart:
    For Each sNode In tvFiles.Nodes
        If sNode.children = 0 Then

            dFile = Replace(sNode.FullPath, "Root Web", Project.UserConfigs(Project.DefaultConfig).RootWeb)
            dFile = RemoveDoubleSlashes(dFile)
            
            If Project.UserConfigs(Project.DefaultConfig).Type = ctcCDROM Then
                LoaderCode = GenLoaderCode(False, False, dFile)
            End If
            
            lblStatus.caption = "Installing Code:" + vbCrLf + EllipseText(txtDummy, dFile, DT_PATH_ELLIPSIS)
            DoEvents

            If Not FileExists(dFile) Then
                MsgBox "The file " + dFile + " could not be found" + vbCrLf + "Please check your Configurations", vbInformation + vbOKOnly, "Error Installing Loader Code"
                tvFiles.Nodes.Remove sNode.Index
                GoTo ReStart
            Else
                sCode = LoadFile(dFile)
                
                If Project.UserConfigs(Project.DefaultConfig).Type = ctcCDROM Then
                    nLoaderCode = Replace(LoaderCode, "%TOROOTRELPATH%", SetSlashDir(GetSmartRelPath(dFile, Project.UserConfigs(Project.DefaultConfig).RootWeb), sdFwd))
                    nLoaderCode = Replace(nLoaderCode, "%JSRELPATH%", SetSlashDir(GetSmartRelPath(dFile, Project.UserConfigs(Project.DefaultConfig).CompiledPath), sdFwd))
                    nLoaderCode = Replace(nLoaderCode, "%IMGRELPATH%", SetSlashDir(GetSmartRelPath(dFile, Project.UserConfigs(Project.DefaultConfig).ImagesPath), sdFwd))
                Else
                    nLoaderCode = LoaderCode
                End If
                
                sCode = RemoveLoaderCode(sCode, dFile)
                sCode = AttachLoaderCode(sCode, nLoaderCode)
                
                ff = FreeFile
                Open dFile For Output As #ff
                    Print #ff, sCode
                Close #ff
            End If
        End If
    Next sNode
    
    SaveLCFilesList tvFiles, False
    
ExitSub:
    Me.Enabled = True
    Exit Sub
    
chkError:
    MsgBox "An unexpected error has occurred while installing the loader code" + vbCrLf + "Error " & Err.number & ": " + Err.Description, vbInformation + vbOKOnly, GetLocalizedStr(660)
    GoTo ExitSub

End Sub

Private Sub cmdRemove_Click()

    Dim sNode As Node
    Dim sNodes() As Node
    Dim i As Integer
    
    ReDim sNodes(0)
    
    If Not tvFiles.SelectedItem Is Nothing Then
        For Each sNode In tvFiles.Nodes
            If sNode.Selected And sNode.key <> "[ROOT]" Then
                ReDim Preserve sNodes(UBound(sNodes) + 1)
                Set sNodes(UBound(sNodes)) = sNode
            End If
        Next sNode
        
        For i = 1 To UBound(sNodes)
            tvFiles.Nodes.Remove sNodes(i).Index
        Next i
    End If

End Sub

Private Sub Form_Load()

    CenterForm Me
    
    lblStatus.caption = ""
    
    ForceCommandsLinks2Local
    AddHSFile
    
    If Project.UserConfigs(Project.DefaultConfig).Frames.UseFrames Then
        cmdFromLinks.Enabled = False
    Else
        LoadLCFilesList tvFiles, False
    End If
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF1 Then showHelp "dialogs/installloadercode.htm"

End Sub

Private Sub AddHSFile()

    On Error Resume Next
    
    tvFiles.Nodes.Clear
    tvFiles.Nodes.Add , , "[ROOT]", "Root Web", 3
    
    If FileExists(Project.UserConfigs(Project.DefaultConfig).HotSpotEditor.HotSpotsFile) Then
        AddFile Project.UserConfigs(Project.DefaultConfig).HotSpotEditor.HotSpotsFile
    End If
    
    With tvFiles.Nodes("[ROOT]")
        .Selected = True
        .EnsureVisible
        .Expanded = True
    End With
    
End Sub

Private Sub ForceCommandsLinks2Local()

    OriginalConfig = Project.DefaultConfig
    If Project.UserConfigs(OriginalConfig).Type = ctcRemote Then
        Project.DefaultConfig = GetConfigID(Project.UserConfigs(OriginalConfig).LocalInfo4RemoteConfig)
        UpdateItemsLinks
    End If

End Sub

Private Sub RestoreCommandsLinks()

    If Project.UserConfigs(OriginalConfig).Type = ctcRemote Then
        Project.DefaultConfig = OriginalConfig
        UpdateItemsLinks
    End If

End Sub

Private Sub AddFolder(FolderName As String)

    Dim dFile As String
    Dim Folders() As String
    Dim i As Integer
    
    ReDim Folders(0)
    
    dFile = Dir(FolderName, vbDirectory)
    Do While LenB(dFile) <> 0
        If dFile <> "." And dFile <> ".." And ((GetAttr(FolderName + dFile) And vbDirectory) = vbDirectory) Then
            ReDim Preserve Folders(UBound(Folders) + 1)
            Folders(UBound(Folders)) = FolderName + dFile
        End If
        dFile = Dir
    Loop
    
    For i = 1 To UBound(Folders)
        AddFolder AddTrailingSlash(Folders(i), "\")
    Next i
    
    dFile = Dir(FolderName + "*.*")
    Do While LenB(dFile) <> 0
        If InStr(dFile, ".") > 0 Then
            If InStr(SupportedHTMLDocs, LCase(Right(dFile, Len(dFile) - InStrRev(dFile, ".") + 1)) + ";") > 0 Then
                AddFile FolderName + dFile
            End If
        End If
        dFile = Dir
    Loop

End Sub

Private Sub AddFile(ByVal FileName As String)

    Dim f() As String
    Dim i As Integer
    Dim pNode As Node
    Dim RootWeb As String
    Dim fp As String
    
    On Error Resume Next
    
    RootWeb = Project.UserConfigs(Project.DefaultConfig).RootWeb
    
    If LCase(Left(FileName, Len(RootWeb))) <> LCase(RootWeb) Then
        MsgBox "The file " + GetFileName(FileName) + " cannot be added to the list" + vbCrLf + "because it is outside your root web", vbOKOnly + vbInformation, "Invalid File Name"
    Else
        FileName = Mid(FileName, Len(RootWeb) + 1)
        f = Split(FileName, "\")
        
        Set pNode = tvFiles.Nodes("[ROOT]")
        For i = 0 To UBound(f) - 1
            Err.Clear
            fp = fp + f(i)
            Set pNode = tvFiles.Nodes.Add(pNode.Index, tvwChild, "K" + fp, f(i), 1)
            If Err.number <> 0 Then
                Set pNode = tvFiles.Nodes("K" + fp)
            End If
        Next i
        
        tvFiles.Nodes.Add pNode.Index, tvwChild, pNode.FullPath + "\" + "K" + fp + f(i), f(i), 2
    End If
    
End Sub

Private Sub tvFiles_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = 46 Then
        cmdRemove_Click
    End If

End Sub
