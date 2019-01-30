VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{9A2947B4-87EE-40EE-A3EF-32BDC32D5726}#1.0#0"; "xfxline3d.ocx"
Begin VB.Form frmLCRemove 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Remove Loader Code"
   ClientHeight    =   3405
   ClientLeft      =   6270
   ClientTop       =   5280
   ClientWidth     =   6255
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
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
      Height          =   315
      Left            =   120
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1275
      Visible         =   0   'False
      Width           =   4800
   End
   Begin MSComDlg.CommonDialog cDlg 
      Left            =   1410
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdAddFolder 
      Caption         =   "Add Folder..."
      Height          =   360
      Left            =   5055
      TabIndex        =   5
      Top             =   1800
      Width           =   1155
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove Files"
      Height          =   360
      Left            =   5055
      TabIndex        =   6
      Top             =   2325
      Width           =   1155
   End
   Begin VB.CommandButton cmdAddFiles 
      Caption         =   "Add Files..."
      Height          =   360
      Left            =   5055
      TabIndex        =   4
      Top             =   1365
      Width           =   1155
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   360
      Left            =   5055
      TabIndex        =   2
      Top             =   285
      Width           =   1155
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   5055
      TabIndex        =   9
      Top             =   2955
      Width           =   1155
   End
   Begin MSComctlLib.ImageList ilIcons 
      Left            =   765
      Top             =   2460
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
            Picture         =   "frmLCRemove.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLCRemove.frx":06D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLCRemove.frx":0DA4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin xfxLine3D.ucLine3D uc3DLine1 
      Height          =   30
      Left            =   30
      TabIndex        =   7
      Top             =   2805
      Width           =   6165
      _ExtentX        =   10874
      _ExtentY        =   53
   End
   Begin MSComctlLib.TreeView tvFiles 
      Height          =   2400
      Left            =   60
      TabIndex        =   1
      Top             =   285
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
      Height          =   405
      Left            =   60
      TabIndex        =   8
      Top             =   2910
      UseMnemonic     =   0   'False
      Width           =   4905
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Files containing the Loader Code"
      Height          =   195
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   2340
   End
End
Attribute VB_Name = "frmLCRemove"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RootWeb As String

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
    
    lblStatus.caption = ""
    
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
    
    lblStatus.caption = ""
    
    Me.Enabled = True
    Me.SetFocus

End Sub

Private Sub cmdClose_Click()

    Unload Me

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

Private Sub cmdStart_Click()

    DoRemove
    cmdClose_Click

End Sub

Private Sub DoRemove()

    Dim sNode As Node
    Dim ff As Integer
    Dim sCode As String
    Dim dFile As String
    
    On Error GoTo chkError
    
    Me.Enabled = False
    
    For Each sNode In tvFiles.Nodes
        If sNode.children = 0 Then

            dFile = Replace(sNode.FullPath, "Root Web", RootWeb)
            dFile = RemoveDoubleSlashes(dFile)

            If LenB(dFile) = 0 Or LenB(Dir(dFile)) = 0 Then
                MsgBox "The file " + dFile + " could not be found" + vbCrLf + "Please check your Configurations", vbInformation + vbOKOnly, "Error Removing Loader Code"
            Else
            
                lblStatus.caption = "Removing Code:" + vbCrLf + EllipseText(txtDummy, dFile, DT_PATH_ELLIPSIS)
                DoEvents
            
                sCode = LoadFile(dFile)
                sCode = RemoveLoaderCode(sCode, dFile)

                ff = FreeFile
                Open dFile For Output As #ff
                    Print #ff, sCode
                Close #ff
            End If
        End If
    Next sNode
    
ExitSub:
    Me.Enabled = True
    Exit Sub
    
chkError:
    MsgBox "An unexpected error has occurred while installing the loader code" + vbCrLf + "Error " & Err.number & ": " + Err.Description, vbInformation + vbOKOnly, GetLocalizedStr(660)
    GoTo ExitSub

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF1 Then showHelp "dialogs/removeloadercode.htm"

End Sub

Private Sub Form_Load()

    CenterForm Me
    
    lblStatus.caption = ""
    
    tvFiles.Nodes.Clear
    tvFiles.Nodes.Add , , "[ROOT]", "Root Web", 3
    
    ForceCommandsLinks2Local
    
    RootWeb = GetRealLocal.RootWeb
    
    Me.Enabled = False
    
    LoadLCFilesList tvFiles, False
    LoadLCFilesList tvFiles, True
    
    tvFiles.Nodes(1).Expanded = True
    Me.Enabled = True

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
    Dim fp As String
    
    On Error Resume Next
    
    If Not HasLC(FileName) Then Exit Sub
    
    If LCase(Left(FileName, Len(RootWeb))) <> LCase(RootWeb) Then
        MsgBox "The file " + GetFileName(FileName) + " cannot be added to the list" + vbCrLf + "because it is outside your root web", vbOKOnly + vbInformation, "Invalid File Name"
    Else
    
        lblStatus.caption = "Adding File:" + vbCrLf + EllipseText(txtDummy, FileName, DT_PATH_ELLIPSIS)
        DoEvents
    
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
    End If
    
    tvFiles.Nodes.Add pNode.Index, tvwChild, pNode.FullPath + "\" + "K" + fp + f(i), f(i), 2
    
End Sub

Private Function HasLC(dFile As String) As Boolean

    Dim sCode As String

    sCode = LoadFile(dFile)
    HasLC = InStr(1, sCode, LoaderCodeSTART, vbTextCompare) > 0

End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    RestoreCommandsLinks

End Sub
