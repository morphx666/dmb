VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmItemHighlightWizard 
   Caption         =   "Item Highlight Wizard"
   ClientHeight    =   6330
   ClientLeft      =   5490
   ClientTop       =   4410
   ClientWidth     =   6495
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmItemHighlightWizard.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6330
   ScaleWidth      =   6495
   Begin VB.CheckBox chkHideDocsNoMenus 
      Caption         =   "Hide documents without the menus' loader code"
      Height          =   225
      Left            =   135
      TabIndex        =   12
      Top             =   5280
      Width           =   6210
   End
   Begin VB.CheckBox chkRemove 
      Height          =   405
      Left            =   2250
      Picture         =   "frmItemHighlightWizard.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5625
      Width           =   225
   End
   Begin VB.CheckBox chkDropDownMenu 
      Height          =   405
      Left            =   1020
      Picture         =   "frmItemHighlightWizard.frx":05CC
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5625
      Width           =   225
   End
   Begin VB.PictureBox picIcon 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   4350
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   5595
      Visible         =   0   'False
      Width           =   240
   End
   Begin MSComctlLib.ImageList ilIcons 
      Left            =   3420
      Top             =   5325
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
            Picture         =   "frmItemHighlightWizard.frx":060E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmItemHighlightWizard.frx":0CE0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   405
      Left            =   5475
      TabIndex        =   8
      Top             =   5625
      Width           =   885
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove"
      Enabled         =   0   'False
      Height          =   405
      Left            =   1365
      TabIndex        =   7
      Top             =   5625
      Width           =   885
   End
   Begin VB.CommandButton cmdInstall 
      Caption         =   "Install"
      Enabled         =   0   'False
      Height          =   405
      Left            =   135
      TabIndex        =   6
      Top             =   5625
      Width           =   885
   End
   Begin VB.TextBox txtJSCode 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   1110
      Left            =   135
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   4065
      Width           =   6225
   End
   Begin MSComctlLib.TreeView tvMenus 
      Height          =   3345
      Left            =   3270
      TabIndex        =   2
      Top             =   375
      Width           =   3090
      _ExtentX        =   5450
      _ExtentY        =   5900
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   353
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      Appearance      =   0
   End
   Begin MSComctlLib.TreeView tvBrowser 
      Height          =   3345
      Left            =   135
      TabIndex        =   0
      Top             =   375
      Width           =   3090
      _ExtentX        =   5450
      _ExtentY        =   5900
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   353
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "ilIcons"
      Appearance      =   0
   End
   Begin MSComctlLib.StatusBar sbDummy 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   13
      Top             =   6060
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
      EndProperty
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Information"
      Height          =   195
      Left            =   135
      TabIndex        =   4
      Top             =   3840
      Width           =   840
   End
   Begin VB.Label lblMenus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Menus"
      Height          =   195
      Left            =   3270
      TabIndex        =   3
      Top             =   150
      Width           =   465
   End
   Begin VB.Label lblFiles 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Files"
      Height          =   195
      Left            =   135
      TabIndex        =   1
      Top             =   150
      Width           =   315
   End
   Begin VB.Menu mnuOptionsInstall 
      Caption         =   "mnuOptionsInstall"
      Begin VB.Menu mnuOptionsInstallSel 
         Caption         =   "Install"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuOptionsInstall2All 
         Caption         =   "Install to all links in Menus"
      End
   End
   Begin VB.Menu mnuOptionsRemove 
      Caption         =   "mnuOptionsRemove"
      Begin VB.Menu mnuOptionsRemoveSel 
         Caption         =   "Remove"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuOptionsRemoveAll 
         Caption         =   "Remove All"
      End
   End
End
Attribute VB_Name = "frmItemHighlightWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents msgSubClass As xfxSC
Attribute msgSubClass.VB_VarHelpID = -1

Private Const dmbHLWM = "<!-- DMB Highlight START --><script language=""JavaScript"">%%CODE%%</script><!-- DMB Highlight END -->"

Dim needRecompile As Boolean
Dim IsLoading As Boolean
Dim IsResizing As Boolean

Private xMenu As CMenu

Private Sub chkDropDownMenu_Click()

    With chkDropDownMenu
        If .Value = vbChecked Then
            .Value = vbUnchecked
            tvBrowser.SetFocus
            DoEvents
            
            PopupMenu mnuOptionsInstall, , .Left - cmdInstall.Width - 15, .Top + .Height
        Else
            .Value = vbUnchecked
        End If
    End With

End Sub

Private Sub chkHideDocsNoMenus_Click()

    If IsLoading Then Exit Sub
    LoadWebFiles
    
    tvBrowser.SetFocus
    
    UpdateJSCode

End Sub

Private Sub chkRemove_Click()

    With chkRemove
        If .Value = vbChecked Then
            .Value = vbUnchecked
            tvBrowser.SetFocus
            DoEvents
            
            PopupMenu mnuOptionsRemove, , .Left - cmdRemove.Width - 15, .Top + .Height
        Else
            .Value = vbUnchecked
        End If
    End With

End Sub

Private Sub cmdClose_Click()

    Unload Me

End Sub

Private Sub cmdInstall_Click()

    InstallDMBHighlightCode tvBrowser.SelectedItem

End Sub

Private Sub InstallDMBHighlightCode(fn As Node)
    
    If fn.ForeColor = vbBlue Then RemoveDMBHighlightCode
    
    InstallCode fn.tag
    
    fn.ForeColor = vbBlue
    cmdRemove.Enabled = True
    mnuOptionsRemove.Enabled = True
    
    If Not Project.AutoSelFunction Then
        Project.HasChanged = True
        needRecompile = True
    End If
    Project.AutoSelFunction = True

End Sub

Private Sub InstallCode(fn As String)

    Dim sCode As String
    Dim hCode As String
    Dim p As Long

    sCode = RemoveCodeFromString(LoadFile(fn))
    hCode = Replace(dmbHLWM, "%%CODE%%", GetJSCode)
    
    p = InStr(1, sCode, "</body>", vbTextCompare)
    If p > 0 Then
        sCode = Left(sCode, p - 1) + hCode + Mid(sCode, p)
    Else
        sCode = sCode + hCode
    End If
    
    SaveFile fn, sCode

End Sub

Private Function GetJSCode() As String

    Dim jsCode As String
    
    jsCode = txtJSCode.tag
    If GetRealLocal.Frames.UseFrames Then
        jsCode = "function doDMBHL() {if(!cFrame) {window.setTimeout('doDMBHL()',100);return false;}cFrame.dmbUnHighlightGroupItems();" + jsCode + " return true;} doDMBHL();"
        jsCode = Replace(jsCode, "dmbHighlightTBItem(", "cFrame.dmbHighlightTBItem(")
        jsCode = Replace(jsCode, "dmbHighlightGroupItem(", "cFrame.dmbHighlightGroupItem(")
    End If
    
    GetJSCode = jsCode

End Function

Private Sub RemoveDMBHighlightCode()

    Dim fn As Node
    
    On Error GoTo ExitSub
    
    Set fn = tvBrowser.SelectedItem
    SaveFile fn.tag, RemoveCodeFromString(LoadFile(fn.tag))
    
    fn.ForeColor = vbBlack
    
    cmdRemove.Enabled = False
    mnuOptionsRemove.Enabled = False
    
ExitSub:

End Sub

Private Function RemoveCodeFromString(ByVal sCode As String) As String

    Dim p1 As Long
    Dim p2 As Long
    Dim s1 As String
    Dim s2 As String
    
    s1 = Split(dmbHLWM, "%%CODE%%")(0)
    s1 = Left(s1, InStr(s1, ">"))
    p1 = InStr(1, sCode, s1, vbTextCompare) + Len(s1)
    s2 = Split(dmbHLWM, "%%CODE%%")(1)
    s2 = Mid(s2, InStr(s2, ">") + 1)
    p2 = InStrRev(sCode, s2, , vbTextCompare)
    
    If p1 <> 0 And p2 <> 0 Then
        RemoveCodeFromString = Replace(sCode, s1 + Mid(sCode, p1, p2 - p1) + s2, "")
    Else
        RemoveCodeFromString = sCode
    End If

End Function

Private Sub cmdRemove_Click()

    RemoveDMBHighlightCode

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF1 Then showHelp "dialogs/autosel.htm"

End Sub

Private Sub Form_Load()

    Dim n As Node
    
    IsLoading = True
    
    If Val(GetSetting(App.EXEName, "IHWWinPos", "X")) = 0 Then
        CenterForm Me
    Else
        Left = GetSetting(App.EXEName, "IHWWinPos", "X")
        Top = GetSetting(App.EXEName, "IHWWinPos", "Y")
        Width = GetSetting(App.EXEName, "IHWWinPos", "W")
        Height = GetSetting(App.EXEName, "IHWWinPos", "H")
        
        If Left + Width / 2 > Screen.Width Or Top + Height / 2 > Screen.Height Then
            Left = Screen.Width / 2 - Width / 2
            Top = Screen.Height / 2 - Height / 2
        End If
    End If
    
    mnuOptionsInstall.Visible = False
    mnuOptionsRemove.Visible = False
    
    CenterForm Me
    SetupCharset Me
    LocalizeUI
    
    chkHideDocsNoMenus.Value = GetSetting(App.EXEName, "Preferences", "IHW_HideDocsNoMenus", vbChecked)
    
    If Not FolderExists(GetRealLocal.RootWeb) Then
        MsgBox "Please configure your project before using this wizard", vbInformation + vbOKOnly, "Unable to run Wizard"
        chkDropDownMenu.Enabled = False
        chkRemove.Enabled = False
        tvBrowser.Enabled = False
        tvMenus.Enabled = False
        chkHideDocsNoMenus.Enabled = False
    Else
        LoadWebFiles
    End If
    
    Set tvMenus.ImageList = frmMain.tvMapView.ImageList
    For Each n In frmMain.tvMapView.Nodes
        If n.parent Is Nothing Then PopulateMenus n, Nothing
    Next n
    
    UpdateJSCode
    
    IsLoading = False
    
    If Not IsDebug Then
        Set xMenu = New CMenu
        xMenu.Initialize Me
        
        SetupSubclassing True
    End If

End Sub

Private Sub LoadWebFiles()

    Dim n As Node

    tvBrowser.Nodes.Clear
    Set n = tvBrowser.Nodes.Add(, , , "Root Web", 2)
    n.Expanded = True
    PopulateFiles n, GetRealLocal.RootWeb, True

End Sub

Private Sub PopulateMenus(n As Node, p As Node)

    Dim nc As Node
    Dim nt As Node
    Dim nn As Node
    
    If p Is Nothing Then
        Set nt = tvMenus.Nodes.Add(, , n.key, n.Text, n.Image)
    Else
        Set nt = tvMenus.Nodes.Add(p, tvwChild, , n.Text, n.Image)
    End If
    CopyNodeStyles n, nt
    
    Set nc = n.Child
    Do Until nc Is Nothing
        If nc.children Then
            PopulateMenus nc, nt
        Else
            If nc.ForeColor <> Preferences.DisabledItem And nc.Image <> 21 Then
                Set nn = tvMenus.Nodes.Add(nt, tvwChild, nc.key, nc.Text, nc.Image)
                CopyNodeStyles nc, nn
            End If
        End If
        Set nc = nc.Next
    Loop

End Sub

Private Sub CopyNodeStyles(ns As Node, nt As Node)

    nt.Expanded = ns.Expanded
    nt.ForeColor = ns.ForeColor
    nt.Bold = ns.Bold
    nt.tag = ns.tag

End Sub

Private Sub PopulateFiles(n As Node, rp As String, Optional IsRoot As Boolean = False)

    Dim sFile As String
    Dim nItem As Node
    Dim Ok2Add As Boolean
    
    If Right(rp, 1) <> "\" Then rp = rp + "\"
    
    sFile = Dir(rp, vbDirectory Or vbHidden Or vbSystem)
    Do While LenB(sFile) <> 0
        If sFile <> "." And sFile <> ".." And ((GetAttr(rp + sFile) And vbDirectory) = vbDirectory) Then
            Ok2Add = True
            If InStr(1, sFile, "_vti_", vbTextCompare) > 0 Then Ok2Add = False
            If InStr(1, sFile, "_private", vbTextCompare) > 0 Then Ok2Add = False
            If InStr(1, sFile, "_themes", vbTextCompare) > 0 Then Ok2Add = False
            If InStr(1, sFile, "_scriptlibrary", vbTextCompare) > 0 Then Ok2Add = False
            If InStr(1, sFile, "_derived", vbTextCompare) > 0 Then Ok2Add = False
            If InStr(1, sFile, "_overlay", vbTextCompare) > 0 Then Ok2Add = False
            If InStr(1, sFile, "_fpclass", vbTextCompare) > 0 Then Ok2Add = False
            If InStr(1, sFile, "_colab", vbTextCompare) > 0 Then Ok2Add = False
            
            If (sFile = "_vti_pvt") And Not IsRoot Then
                Do
                    Set nItem = n.Child
                    If nItem Is Nothing Then Exit Sub
                    tvBrowser.Nodes.Remove nItem.Index
                Loop
            End If
            
            If Ok2Add Then
                Set nItem = tvBrowser.Nodes.Add(n, tvwChild, , sFile)
                nItem.tag = rp + sFile + "\"
                nItem.Image = GetImage("")
                tvBrowser.Nodes.Add nItem, tvwChild, , "<dummy>"
            End If
        End If
        sFile = Dir
    Loop
    
    sFile = Dir(rp + "*.*")
    Do While LenB(sFile) <> 0
         If (InStr(SupportedHTMLDocs, GetFileExtension(sFile) + ";") > 0) Or _
            (InStr(SupportedHTMLDocs, GetFileExtension(sFile) + "|") > 0) Then
            If HasLoaderCode(rp + sFile) Then
                Set nItem = tvBrowser.Nodes.Add(n, tvwChild, , sFile)
                nItem.tag = rp + sFile
                nItem.ForeColor = IIf(HasDMBHLCode(rp + sFile), vbBlue, vbBlack)
                nItem.Image = GetImage(rp + sFile)
            End If
        End If
        sFile = Dir
    Loop

End Sub

Private Function HasDMBHLCode(f As String) As Boolean

    Dim s1 As String
    Dim s2 As String
    Dim s As String
    
    s1 = Split(dmbHLWM, "%%CODE%%")(0)
    s1 = Left(s1, InStr(s1, ">"))
    s2 = Split(dmbHLWM, "%%CODE%%")(1)
    s2 = Mid(s2, InStr(s2, ">") + 1)
    
    s = LoadFile(f)
    HasDMBHLCode = ((InStr(1, s, s1, vbTextCompare) > 0) And (InStr(1, s, s2, vbTextCompare) > 0))

End Function

Private Function HasLoaderCode(f As String) As Boolean

    If chkHideDocsNoMenus.Value = vbUnchecked Then
        HasLoaderCode = True
    Else
        HasLoaderCode = (InStr(LoadFile(f), LoaderCodeSTART) > 0)
    End If

End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    SaveSetting App.EXEName, "IHWWinPos", "X", Left
    SaveSetting App.EXEName, "IHWWinPos", "Y", Top
    SaveSetting App.EXEName, "IHWWinPos", "W", Width
    SaveSetting App.EXEName, "IHWWinPos", "H", Height

    SaveSetting App.EXEName, "Preferences", "IHW_HideDocsNoMenus", chkHideDocsNoMenus.Value
    If needRecompile Then
        If MsgBox("This projects need to be recompiled." + vbCrLf + "Do you want to recompile it now?", vbQuestion + vbYesNo, "Recompile Required") = vbYes Then
            frmMain.ToolsCompile
        End If
    End If
    
    If Not IsDebug Then SetupSubclassing False
    
End Sub

Private Sub Form_Resize()

    Dim w As Integer
    Dim h As Integer
    Dim tx As Integer
    Dim ty As Integer
    
    On Error Resume Next
    
    IsResizing = True
    
    w = Width - GetClientLeft(Me.hwnd)
    h = Height - GetClientTop(Me.hwnd)
    tx = Screen.TwipsPerPixelX
    ty = Screen.TwipsPerPixelY

    tvBrowser.Width = w / 2 - 10 * tx
    tvBrowser.Height = h - 198 * ty
    
    tvMenus.Left = w / 2 + 6 * tx
    tvMenus.Width = w - tvMenus.Left - 8 * tx
    tvMenus.Height = tvBrowser.Height
    lblMenus.Left = tvMenus.Left
    
    lblInfo.Top = tvBrowser.Top + tvBrowser.Height + 6 * ty
    txtJSCode.Top = lblInfo.Top + lblInfo.Height + 2 * ty
    txtJSCode.Width = tvMenus.Left + tvMenus.Width - txtJSCode.Left
    
    chkHideDocsNoMenus.Top = txtJSCode.Top + txtJSCode.Height + 6 * ty
    
    cmdInstall.Top = h - cmdInstall.Height - 8 * ty
    cmdRemove.Top = cmdInstall.Top
    cmdClose.Top = cmdInstall.Top
    cmdClose.Left = txtJSCode.Left + txtJSCode.Width - cmdClose.Width
    chkDropDownMenu.Top = cmdInstall.Top
    chkRemove.Top = cmdInstall.Top
    
    IsResizing = False

End Sub

Private Sub SetupSubclassing(scState As Boolean)

    If msgSubClass Is Nothing Then Set msgSubClass = New xfxSC
    
    frmIHWHWND = Me.hwnd
    msgSubClass.SubClassHwnd Me.hwnd, scState

End Sub

Private Sub mnuOptionsInstallSel_Click()

    InstallDMBHighlightCode tvBrowser.SelectedItem

End Sub

Private Sub mnuOptionsInstall2All_Click()

    Dim i As Integer
    Dim a As ActionEvents
    Dim p As String
    Dim rw As String
    Dim lrw As String
    
    rw = Project.UserConfigs(Project.DefaultConfig).RootWeb
    lrw = GetRealLocal.RootWeb
    
    For i = 1 To UBound(MenuGrps)
        If Not IsSubMenu(i) And BelongsToToolbar(i, True) Then
            a = MenuGrps(i).Actions
            If a.onclick.Type = atcURL Then
                p = SetSlashDir(Replace(a.onclick.url, rw, lrw, , , vbTextCompare), sdBack)
                If FileExists(p) Then
                    txtJSCode.Text = "Installing code: " + GetFileName(p)
                    DoEvents
                    txtJSCode.tag = GetJSCode4Grp(i)
                    InstallCode p
                End If
            End If
        End If
    Next i
    
    For i = 1 To UBound(MenuCmds)
        a = MenuCmds(i).Actions
        If a.onclick.Type = atcURL Then
            p = SetSlashDir(Replace(a.onclick.url, rw, lrw, , , vbTextCompare), sdBack)
            If FileExists(p) Then
                txtJSCode.Text = "Installing code: " + GetFileName(p)
                DoEvents
                txtJSCode.tag = GetJSCode4Cmd(i)
                InstallCode p
            End If
        End If
    Next i
    
    If Not Project.AutoSelFunction Then
        Project.HasChanged = True
        needRecompile = True
    End If
    Project.AutoSelFunction = True
    
    LoadWebFiles
    
    txtJSCode.Text = "Done..."

End Sub

Private Sub mnuOptionsRemoveAll_Click()

    Dim n As Node
    
    LockWindowUpdate tvBrowser.hwnd
    
    txtJSCode.Text = "Please wait..."
    
ReStart:
    DoEvents
    For Each n In tvBrowser.Nodes
        If Not n.Expanded And InStr(n.tag, ".") = 0 Then
            txtJSCode.Text = "Analyzing: " + n.Text
            DoEvents
            n.Expanded = True
            GoTo ReStart
        End If
    Next n
    
    For Each n In tvBrowser.Nodes
        If n.ForeColor = vbBlue Then
            txtJSCode.Text = "Analyzing: " + n.Text
            DoEvents
            n.Selected = True
            RemoveDMBHighlightCode
        End If
    Next n
    
    LockWindowUpdate 0
    
    With tvBrowser.Nodes(1)
        .Selected = True
        .EnsureVisible
    End With
    
    txtJSCode.Text = "Done..."

End Sub

Private Sub mnuOptionsRemoveSel_Click()

    RemoveDMBHighlightCode

End Sub

Private Sub msgSubClass_NewMessage(ByVal hwnd As Long, uMsg As Long, wParam As Long, lParam As Long, Cancel As Boolean, UseRetVal As Boolean, RetVal As Long)

    If IsResizing Then Exit Sub

    Cancel = HandleSubclassMsg(hwnd, uMsg, wParam, lParam)

End Sub

Private Sub tvBrowser_Expand(ByVal Node As MSComctlLib.Node)

    If Node.Child.FirstSibling.Text = "<dummy>" Then
        tvBrowser.Nodes.Remove Node.Child.FirstSibling.Index
        PopulateFiles Node, Node.tag
    End If

End Sub

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

Private Sub tvBrowser_NodeClick(ByVal Node As MSComctlLib.Node)

    UpdateJSCode

End Sub

Private Sub UpdateJSCode()

    Dim fn As Node
    Dim mn As Node
    Dim m As Node
    Dim id As Integer
    
    Set fn = tvBrowser.SelectedItem
    Set mn = tvMenus.SelectedItem
    
    cmdInstall.Enabled = False
    mnuOptionsInstallSel.Enabled = False
    cmdRemove.Enabled = False
    mnuOptionsRemove.Enabled = False
    
    If fn Is Nothing Then
        If mn Is Nothing Then
            txtJSCode.Text = GetLocalizedStr(975)
        Else
            txtJSCode.Text = Replace(GetLocalizedStr(976), "%%MENU_NAME%%", mn.Text)
        End If
        Exit Sub
    End If
    
    If LenB(GetFileName(fn.tag)) = 0 Then
        If mn Is Nothing Then
            txtJSCode.Text = GetLocalizedStr(977) + " " + GetLocalizedStr(975)
        Else
            txtJSCode.Text = GetLocalizedStr(977) + " " + Replace(GetLocalizedStr(976), "%%MENU_NAME%%", mn.Text)
        End If
        Exit Sub
    End If
    
    cmdRemove.Enabled = (fn.ForeColor = vbBlue)
    mnuOptionsRemove.Enabled = cmdRemove.Enabled
    
    If fn.ForeColor = &H80000011 Then
        txtJSCode.Text = GetLocalizedStr(978)
        Exit Sub
    End If
    
    If mn Is Nothing Then
        txtJSCode.Text = Replace(GetLocalizedStr(979), "%%FILE_NAME%%", fn.Text)
        Exit Sub
    End If
    
    If Left(mn.key, 3) = "TBK" Then
        txtJSCode.Text = GetLocalizedStr(980)
        Exit Sub
    End If
    
    frmMain.NodeSelectedInMapView mn
    Set m = frmMain.tvMenus.SelectedItem
    id = GetID
    
    If IsCommand(m.key) Then
        If MenuCmds(id).disabled Then
            txtJSCode.Text = GetLocalizedStr(981)
            Exit Sub
        End If
        txtJSCode.tag = GetJSCode4Cmd(id)
    Else
        If IsGroup(m.key) Then
            If Not mn.parent.parent Is Nothing Or BelongsToToolbar(id, True) = 0 Then
                txtJSCode.Text = GetLocalizedStr(982)
                Exit Sub
            Else
                If MenuGrps(id).disabled Then
                    txtJSCode.Text = GetLocalizedStr(981)
                    Exit Sub
                End If
                
                txtJSCode.tag = GetJSCode4Grp(id)
            End If
        End If
    End If
    
    txtJSCode.Text = GetLocalizedStr(983) + vbCrLf + txtJSCode.tag
    cmdInstall.Enabled = True
    mnuOptionsInstallSel.Enabled = True
    Exit Sub

End Sub

Private Function GetJSCode4Cmd(ByVal i As Integer, Optional DoRecurse As Boolean = True) As String

    Dim s As String
    Dim c As MenuCmd
    
    c = MenuCmds(i)
    s = "dmbHighlightGroupItem('" + MenuGrps(c.parent).Name + "', '" + IIf(c.caption = "", "N" & (i - 1), EscapeCaption(c.caption)) + "');" + vbCrLf
    If DoRecurse Then
        Do
            If IsSubMenu(c.parent) Then
                s = s + GetJSCode4Cmd(SubMenuOf(c.parent), False)
                c = MenuCmds(SubMenuOf(c.parent))
            Else
                s = s + GetJSCode4Grp(c.parent) + vbCrLf
                Exit Do
            End If
        Loop
    End If
    
    GetJSCode4Cmd = s

End Function

Private Function GetJSCode4Grp(ByVal i As Integer) As String

    Dim tbi As Integer
    Dim id As Integer
    Dim t As Integer
    Dim g As Integer
    Dim nid As Integer
    
    tbi = BelongsToToolbar(i, True)
    If tbi > 0 Then
        For t = 1 To UBound(Project.Toolbars)
            For g = 1 To UBound(Project.Toolbars(t).Groups)
                id = id + 1
                If MenuGrps(i).Name = Project.Toolbars(t).Groups(g) Then
                    nid = id
                    Exit For
                End If
            Next g
            If nid = id Then Exit For
        Next t
        GetJSCode4Grp = "dmbHighlightTBItem(" & tbi & ", '" + IIf(MenuGrps(i).caption = "", "N" & (nid + 1000), EscapeCaption(MenuGrps(i).caption)) + "');"
    End If

End Function

Private Function EscapeCaption(ByVal c As String) As String

    c = Replace(c, "\", "\\")
    c = Replace(c, "'", "\'")
    c = Replace(c, "&", "&amp;")
    c = Replace(c, ", ", ",&nbsp;")
    c = Replace(c, " .", "&nbsp;.")
    EscapeCaption = c

End Function

Private Sub tvMenus_NodeClick(ByVal Node As MSComctlLib.Node)

    UpdateJSCode

End Sub

Private Sub LocalizeUI()

    cmdClose.caption = GetLocalizedStr(424)
    
    lblFiles.caption = GetLocalizedStr(970)
    lblMenus.caption = GetLocalizedStr(971)
    lblInfo.caption = GetLocalizedStr(972)
    
    cmdInstall.caption = GetLocalizedStr(468)
    mnuOptionsInstallSel.caption = GetLocalizedStr(468)
    mnuOptionsInstall2All.caption = GetLocalizedStr(973)
    
    cmdRemove.caption = GetLocalizedStr(201)
    mnuOptionsRemoveSel.caption = GetLocalizedStr(201)
    mnuOptionsRemoveAll.caption = GetLocalizedStr(974)
    
End Sub
