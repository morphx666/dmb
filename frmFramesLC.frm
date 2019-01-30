VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{9A2947B4-87EE-40EE-A3EF-32BDC32D5726}#1.0#0"; "xfxline3d.ocx"
Begin VB.Form frmLCFramesInstall 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Install Frames Loader Code"
   ClientHeight    =   3405
   ClientLeft      =   6420
   ClientTop       =   4905
   ClientWidth     =   6270
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
   ScaleWidth      =   6270
   ShowInTaskbar   =   0   'False
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
      TabIndex        =   4
      Top             =   1395
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
      TabIndex        =   5
      Top             =   1815
      Width           =   1155
   End
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
      Left            =   105
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1875
      Visible         =   0   'False
      Width           =   4800
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
   Begin MSComDlg.CommonDialog cDlg 
      Left            =   195
      Top             =   2565
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ilIcons 
      Left            =   750
      Top             =   2475
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
            Picture         =   "frmFramesLC.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFramesLC.frx":06D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFramesLC.frx":0DA4
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
      Caption         =   "Files to Install the Frames Loader Code"
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
      Width           =   2805
   End
End
Attribute VB_Name = "frmLCFramesInstall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAddFiles_Click()

    On Error GoTo ExitSub
    
    Dim f() As String
    Dim i As Integer
    
    With cDlg
        .DialogTitle = GetLocalizedStr(753)
        .InitDir = Project.UserConfigs(Project.DefaultConfig).RootWeb
        .FileName = ""
        .filter = SupportedHTMLDocs
        .CancelError = True
        .Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly + cdlOFNLongNames + cdlOFNPathMustExist + cdlOFNAllowMultiselect + cdlOFNExplorer
        .ShowOpen
        If InStr(.FileName, Chr(0)) Then
            f = Split(.FileName, Chr(0))
            For i = 1 To UBound(f)
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
    Dim dFile As String
    
    Me.Enabled = False
    
    RestoreCommandsLinks
    CompileProject MenuGrps, MenuCmds, Project, Preferences, params, False, True
    LoaderCode = GenLoaderCode(True, False)
    ForceCommandsLinks2Local
    
ReStart:
    For Each sNode In tvFiles.Nodes
        If sNode.children = 0 Then
            dFile = Replace(sNode.FullPath, "Root Web", Project.UserConfigs(Project.DefaultConfig).RootWeb)
            dFile = RemoveDoubleSlashes(dFile)
            
            If Project.UserConfigs(Project.DefaultConfig).Type = ctcCDROM Then
                LoaderCode = GenLoaderCode(True, False, dFile)
            End If
            
            lblStatus.caption = "Installing Code:" + vbCrLf + EllipseText(txtDummy, dFile, DT_PATH_ELLIPSIS)
            DoEvents
            
            If RemoveFramesLCode(dFile) Then
            
                If Project.UserConfigs(Project.DefaultConfig).Type = ctcCDROM Then
                    LoaderCode = GenLoaderCode(True, False, dFile)
                    nLoaderCode = Replace(LoaderCode, "%TOROOTRELPATH%", SetSlashDir(GetSmartRelPath(dFile, Project.UserConfigs(Project.DefaultConfig).RootWeb), sdFwd))
                    nLoaderCode = Replace(nLoaderCode, "%JSRELPATH%", SetSlashDir(GetSmartRelPath(dFile, Project.UserConfigs(Project.DefaultConfig).CompiledPath), sdFwd))
                    nLoaderCode = Replace(nLoaderCode, "%IMGRELPATH%", SetSlashDir(GetSmartRelPath(dFile, Project.UserConfigs(Project.DefaultConfig).ImagesPath), sdFwd))
                Else
                    nLoaderCode = LoaderCode
                End If
            
                If Not AttachFramesLCode(dFile, nLoaderCode) Then
                    tvFiles.Nodes.Remove sNode.Index
                    GoTo ReStart
                End If
            End If
        End If
    Next sNode
    
    SaveLCFilesList tvFiles, True
    
    Me.Enabled = True

End Sub

Private Function AttachFramesLCode(ByVal dFile As String, LoaderCode As String) As Boolean

    Dim ff As Integer
    Dim sCode As String
    Dim tCode As String
    Dim p1 As Long
    Dim p2 As Long
    
    On Error GoTo ReportError
    
    dFile = RemoveDoubleSlashes(dFile)
    sCode = LoadFile(dFile)
    tCode = LCase$(sCode)
    
    p1 = InStr(LCase$(tCode), "<body")
    p2 = InStr(p1, LCase$(tCode), ">")
    
    If p1 = 0 Or p2 = 0 Then
        MsgBox "The frames loader code could not be installed on the " + dFile + " file", vbCritical + vbOKOnly, GetLocalizedStr(660)
        AttachFramesLCode = False
    Else
        sCode = Left$(sCode, p2) + LoaderCode + Mid$(sCode, p2 + 1)
        ff = FreeFile
        Open dFile For Output As #ff
            Print #ff, sCode
        Close #ff
        AttachFramesLCode = True
    End If
    
    Exit Function
    
ReportError:
    MsgBox "Error " & Err.number & ": " + Err.Description + vbCrLf + "File: " + dFile, vbInformation + vbOKOnly, GetLocalizedStr(660)
    
End Function

Private Function RemoveFramesLCode(dFile As String) As Boolean

    Dim ff As Integer
    Dim sCode As String
    Dim p1 As Long
    Dim p2 As Long
    
    On Error GoTo ReportError
    
    dFile = RemoveDoubleSlashes(dFile)
    
    sCode = LoadFile(dFile)
    
    ff = FreeFile
    If InStr(sCode, LoaderCodeSTART) Then
        sCode = Left(sCode, InStr(sCode, LoaderCodeSTART) - 1) + Mid(sCode, InStr(sCode, LoaderCodeEND) + Len(LoaderCodeEND))
        Open dFile For Output As #ff
            Print #ff, sCode
        Close #ff
    Else
        'For compatibility with older versions of DMB
        If InStr(LCase(sCode), "frames.js") Then
            p1 = InStrRev(LCase(sCode), "<script", InStr(LCase(sCode), "frames.js")) - 1
            p2 = InStr(InStr(LCase(sCode), "frames.js"), LCase(sCode), "</script>") + Len("</script>")
            sCode = Left$(sCode, p1) + Mid$(sCode, p2)
            Open dFile For Output As #ff
                Print #ff, sCode
            Close #ff
        End If
    End If
    
    RemoveFramesLCode = True
    
    Exit Function
    
ReportError:
    MsgBox "Error " & Err.number & ": " + Err.Description + vbCrLf + "File: " + dFile, vbInformation + vbOKOnly, GetLocalizedStr(661)
    
End Function

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
    AddLinksFromCommands
    LoadLCFilesList tvFiles, True

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF1 Then showHelp "dialogs/installloadercode.htm"

End Sub

Private Sub AddLinksFromCommands()

    Dim c As Integer
    Dim dFile As String
    
    On Error Resume Next
    
    tvFiles.Nodes.Clear
    tvFiles.Nodes.Add , , "[ROOT]", "Root Web", 3
    
    For c = 1 To UBound(MenuCmds)
        With MenuCmds(c).Actions
            If .onmouseover.Type = atcURL And Len(.onmouseover.url) > 0 Then
                If LCase$(Left$(.onmouseover.url, 10)) <> "javascript" And _
                    LCase$(Left$(.onmouseover.url, 8)) <> "vbscript" And _
                    LCase$(Left$(.onmouseover.url, 4)) <> "http" And _
                    LCase$(Left$(.onmouseover.url, 8)) <> "ftp" Then
                    dFile = .onmouseover.url
                    If LenB(Dir(dFile)) <> 0 And InStr(dFile, ".") > 0 Then
                        Select Case LCase$(Mid$(dFile, InStrRev(dFile, ".") + 1))
                            Case "htm", "html", "asp", "php", "php3", "shtml"
                                AddFile dFile
                        End Select
                    End If
                End If
            End If
            If .onclick.Type = atcURL And Len(.onclick.url) > 0 Then
                If LCase$(Left$(.onclick.url, 10)) <> "javascript" And _
                    LCase$(Left$(.onclick.url, 8)) <> "vbscript" And _
                    LCase$(Left$(.onclick.url, 4)) <> "http" And _
                    LCase$(Left$(.onclick.url, 8)) <> "ftp" Then
                    dFile = .onclick.url
                    If LenB(Dir(dFile)) <> 0 And InStr(dFile, ".") > 0 Then
                        Select Case LCase$(Mid$(dFile, InStrRev(dFile, ".") + 1))
                            Case "htm", "html", "asp", "php", "php3", "shtml"
                                AddFile dFile
                        End Select
                    End If
                End If
            End If
            If .OnDoubleClick.Type = atcURL And Len(.OnDoubleClick.url) > 0 Then
                If LCase$(Left$(.OnDoubleClick.url, 10)) <> "javascript" And _
                    LCase$(Left$(.OnDoubleClick.url, 8)) <> "vbscript" And _
                    LCase$(Left$(.OnDoubleClick.url, 4)) <> "http" And _
                    LCase$(Left$(.OnDoubleClick.url, 8)) <> "ftp" Then
                    dFile = .OnDoubleClick.url
                    If LenB(Dir(dFile)) <> 0 And InStr(dFile, ".") > 0 Then
                        Select Case LCase$(Mid$(dFile, InStrRev(dFile, ".") + 1))
                            Case "htm", "html", "asp", "php", "php3", "shtml"
                                AddFile dFile
                        End Select
                    End If
                End If
            End If
        End With
    Next c
    
    With tvFiles.Nodes("[ROOT]")
        .Selected = True
        .EnsureVisible
        .Expanded = True
    End With
    
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
    Dim fName As String
    Dim fp As String
    
    On Error Resume Next
    
    RootWeb = Project.UserConfigs(Project.DefaultConfig).RootWeb
    
    fName = GetFileName(FileName)
    If LenB(fName) = 0 Then fName = FileName
    
    If LCase(Left(FileName, Len(RootWeb))) <> LCase(RootWeb) Then
        MsgBox "The file " + fName + " cannot be added to the list" + vbCrLf + "because it is outside your root web", vbOKOnly + vbInformation, "Invalid File Name"
    ElseIf LCase(FileName) = LCase(Project.UserConfigs(Project.DefaultConfig).HotSpotEditor.HotSpotsFile) Then
        MsgBox "The file " + fName + " cannot be added to the list" + vbCrLf + "because it is the document containing the hotspots", vbOKOnly + vbInformation, "Invalid File Name"
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
