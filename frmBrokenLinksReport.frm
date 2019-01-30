VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{E1C6DB9D-BD4A-4A61-A759-0CED75D034BF}#43.0#0"; "SmartButton.ocx"
Object = "{9A2947B4-87EE-40EE-A3EF-32BDC32D5726}#1.0#0"; "xfxline3d.ocx"
Begin VB.Form frmBrokenLinksReport 
   Caption         =   "Broken Links Report"
   ClientHeight    =   5670
   ClientLeft      =   4770
   ClientTop       =   5025
   ClientWidth     =   8040
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBrokenLinksReport.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5670
   ScaleWidth      =   8040
   Begin VB.Frame frameButtons 
      BorderStyle     =   0  'None
      Height          =   3090
      Left            =   6540
      TabIndex        =   8
      Top             =   315
      Width           =   1410
      Begin VB.CheckBox chkFilter 
         Height          =   330
         Index           =   0
         Left            =   0
         Picture         =   "frmBrokenLinksReport.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   2760
         Value           =   1  'Checked
         Width           =   375
      End
      Begin VB.CheckBox chkFilter 
         Height          =   330
         Index           =   2
         Left            =   518
         Picture         =   "frmBrokenLinksReport.frx":0B14
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   2760
         Value           =   1  'Checked
         Width           =   375
      End
      Begin VB.CheckBox chkFilter 
         Height          =   330
         Index           =   1
         Left            =   1035
         Picture         =   "frmBrokenLinksReport.frx":109E
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   2760
         Value           =   1  'Checked
         Width           =   375
      End
      Begin SmartButtonProject.SmartButton cmdNextBroken 
         Height          =   315
         Left            =   0
         TabIndex        =   9
         Top             =   2100
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   556
         Caption         =   "       Next"
         Picture         =   "frmBrokenLinksReport.frx":1628
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionLayout   =   3
         PictureLayout   =   3
         OffsetLeft      =   3
      End
      Begin xfxLine3D.ucLine3D uc3dline1 
         Height          =   30
         Left            =   0
         TabIndex        =   10
         Top             =   495
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   53
      End
      Begin SmartButtonProject.SmartButton cmdEdit 
         Height          =   315
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   556
         Caption         =   "       Edit..."
         Picture         =   "frmBrokenLinksReport.frx":1782
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionLayout   =   3
         PictureLayout   =   3
         OffsetLeft      =   3
      End
      Begin SmartButtonProject.SmartButton cmdTestAll 
         Height          =   315
         Left            =   0
         TabIndex        =   12
         Top             =   1065
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   556
         Caption         =   "       Test All"
         Picture         =   "frmBrokenLinksReport.frx":18DC
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionLayout   =   3
         PictureLayout   =   3
         OffsetLeft      =   3
      End
      Begin SmartButtonProject.SmartButton cmdTestSel 
         Height          =   315
         Left            =   0
         TabIndex        =   13
         Top             =   675
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   556
         Caption         =   "       Test Selected"
         Picture         =   "frmBrokenLinksReport.frx":1E76
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionLayout   =   3
         PictureLayout   =   3
         OffsetLeft      =   3
         OffsetRight     =   3
      End
      Begin SmartButtonProject.SmartButton cmdPrevBroken 
         Height          =   315
         Left            =   0
         TabIndex        =   14
         Top             =   1740
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   556
         Caption         =   "       Previous"
         Picture         =   "frmBrokenLinksReport.frx":2410
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionLayout   =   3
         PictureLayout   =   3
         OffsetLeft      =   3
      End
      Begin xfxLine3D.ucLine3D uc3dline2 
         Height          =   30
         Left            =   0
         TabIndex        =   15
         Top             =   1560
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   53
      End
      Begin xfxLine3D.ucLine3D uc3DLine3 
         Height          =   30
         Left            =   0
         TabIndex        =   19
         Top             =   2595
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   53
      End
   End
   Begin VB.TextBox txtDummy 
      Height          =   315
      Left            =   1845
      TabIndex        =   7
      Top             =   4005
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   6540
      TabIndex        =   2
      Top             =   4845
      Width           =   1410
   End
   Begin VB.PictureBox picFlood 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   60
      ScaleHeight     =   330
      ScaleWidth      =   2280
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   4530
      Visible         =   0   'False
      Width           =   2280
   End
   Begin MSComctlLib.StatusBar sbInfo 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   6
      Top             =   5355
      Width           =   8040
      _ExtentX        =   14182
      _ExtentY        =   556
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
      EndProperty
   End
   Begin VB.Timer tmrTimeout 
      Enabled         =   0   'False
      Interval        =   15000
      Left            =   3465
      Top             =   4095
   End
   Begin VB.Timer tmrAutoClose 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   4005
      Top             =   4095
   End
   Begin MSComctlLib.ImageList ilIcons 
      Left            =   4800
      Top             =   4050
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
            Picture         =   "frmBrokenLinksReport.frx":256A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrokenLinksReport.frx":2B04
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrokenLinksReport.frx":309E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Height          =   375
      Left            =   6540
      TabIndex        =   4
      Top             =   4260
      Visible         =   0   'False
      Width           =   1410
   End
   Begin SHDocVwCtl.WebBrowser wBrowser 
      Height          =   1350
      Left            =   180
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   3945
      Width           =   1470
      ExtentX         =   2593
      ExtentY         =   2381
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
   Begin MSComctlLib.ListView lvLinks 
      Height          =   3600
      Left            =   90
      TabIndex        =   1
      Top             =   285
      Width           =   6360
      _ExtentX        =   11218
      _ExtentY        =   6350
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ilIcons"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
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
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "chItem"
         Text            =   "Item"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "chLink"
         Text            =   "Link"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label lblListTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Links Report"
      Height          =   195
      Left            =   105
      TabIndex        =   0
      Top             =   75
      Width           =   885
   End
End
Attribute VB_Name = "frmBrokenLinksReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

#If LITE = 0 Then

Dim IsPorcessing As Boolean
Dim LinkResult As Boolean
Dim BrowserReportedError As Boolean
Dim UserCancelled As Boolean
Dim IsTesting As Boolean
Private FloodPanel As New clsFlood

Private Enum lnkTypeConstants
    ltcUntested = 1
    ltcValid = 3
    ltcBroken = 2
End Enum
Private Type lnkDef
    item As String
    key As String
    Link As String
    Type As lnkTypeConstants
End Type
Dim Links() As lnkDef

Dim IsResizing As Boolean
Private WithEvents msgSubClass As xfxSC
Attribute msgSubClass.VB_VarHelpID = -1

Private Sub chkFilter_Click(Index As Integer)

    ApplyFilter
    lvLinks.SetFocus
    
    DoEvents
    
    lvLinks.Refresh

End Sub

Private Sub ApplyFilter(Optional idx As Integer = -1)

    Dim nItem As ListItem
    Dim i As Integer
    Dim f As Integer
    Dim t As Integer
    
    If idx = -1 Then
        f = 1
        t = UBound(Links)
    Else
        f = idx
        t = idx
    End If
    
    For i = f To t
        Set nItem = lvLinks.FindItem(i, lvwTag, , lvwWhole)
        
        If chkFilter(Links(i).Type - 1).Value = vbChecked Then
            If nItem Is Nothing Then
                Set nItem = lvLinks.ListItems.Add(, , Links(i).item)
            End If
            
            With nItem
                .SubItems(1) = Links(i).Link
                .tag = i
                .SmallIcon = Links(i).Type
                .ForeColor = SetItemColor(Links(i).Type)
            End With
        ElseIf Not nItem Is Nothing Then
            lvLinks.ListItems.Remove nItem.Index
        End If
    Next i

    CoolListView lvLinks

End Sub

Private Sub cmdEdit_Click()

    Dim id As Integer
    Dim purl As String
    Dim idx As Integer
    
    If IsTesting Then Exit Sub

    purl = lvLinks.SelectedItem.SubItems(1)
    With frmEditLink
        .lblItemName.caption = lvLinks.SelectedItem.Text
        .txtLink.Text = purl
        .Show vbModal
    End With

    If purl = lvLinks.SelectedItem.SubItems(1) Then Exit Sub
    
    SetCtrlsState False
    
    idx = Val(lvLinks.SelectedItem.tag)
    Links(idx).Link = lvLinks.SelectedItem.SubItems(1)
    id = Val(Mid(Links(idx).key, 2))
    If Left(Links(idx).key, 1) = "G" Then
        MenuGrps(id).Actions.onclick.url = Links(idx).Link
    Else
        MenuCmds(id).Actions.onclick.url = Links(idx).Link
    End If
        
    TestExtLinks False
    
    SetCtrlsState True

End Sub

Private Sub cmdNextBroken_Click()

    Dim nItem As ListItem
    Dim i As Integer
    
    If SelCount = 0 Then Exit Sub
    
    For i = lvLinks.SelectedItem.Index + 1 To lvLinks.ListItems.Count
        Set nItem = lvLinks.ListItems(i)
        If nItem.SmallIcon = 2 Then
            lvLinks.MultiSelect = False
            lvLinks_ItemClick nItem
            lvLinks.MultiSelect = True
            Exit Sub
        End If
    Next i
    
    For i = 1 To lvLinks.SelectedItem.Index - 1
        Set nItem = lvLinks.ListItems(i)
        If nItem.SmallIcon = 2 Then
            lvLinks.MultiSelect = False
            lvLinks_ItemClick nItem
            lvLinks.MultiSelect = True
            Exit Sub
        End If
    Next i
    
    PlaySound QueryValue(HKEY_CURRENT_USER, "AppEvents\Schemes\Apps\.Default\.Default\.Current"), 0&, SND_ASYNC Or SND_FILENAME Or SND_NOWAIT

End Sub

Private Sub cmdClose_Click()

    Unload Me

End Sub

Private Sub cmdPrevBroken_Click()

    Dim nItem As ListItem
    Dim i As Integer
    
    If SelCount = 0 Then Exit Sub
    
    For i = lvLinks.SelectedItem.Index - 1 To 1 Step -1
        Set nItem = lvLinks.ListItems(i)
        If nItem.SmallIcon = 2 Then
            lvLinks.MultiSelect = False
            lvLinks_ItemClick nItem
            lvLinks.MultiSelect = True
            Exit Sub
        End If
    Next i
    
    For i = lvLinks.ListItems.Count To lvLinks.SelectedItem.Index + 1 Step -1
        Set nItem = lvLinks.ListItems(i)
        If nItem.SmallIcon = 2 Then
            lvLinks.MultiSelect = False
            lvLinks_ItemClick nItem
            lvLinks.MultiSelect = True
            Exit Sub
        End If
    Next i
    
    PlaySound QueryValue(HKEY_CURRENT_USER, "AppEvents\Schemes\Apps\.Default\.Default\.Current"), 0&, SND_ASYNC Or SND_FILENAME Or SND_NOWAIT

End Sub

Private Sub cmdStop_Click()

    UserCancelled = True
    
End Sub

Private Sub cmdTestAll_Click()

    TestExtLinks True

End Sub

Private Sub TestExtLinks(TestAll As Boolean)

    Dim nNode As ListItem
    Dim a As ActionEvents
    Dim id As Integer
    Dim sc As Integer
    Dim i As Integer
    Dim k As Integer
    Dim idx As Integer
    
    IsTesting = True
    
    Form_Resize
    
    ConfigIE
    SetCtrlsState False
    
    If TestAll Then
        lvLinks.MultiSelect = False
        sc = lvLinks.ListItems.Count
    Else
        sc = SelCount
    End If
    
    For i = 1 To lvLinks.ListItems.Count
        idx = 0
        Set nNode = lvLinks.ListItems(i)
        If TestAll Or nNode.Selected Then
        
            idx = Val(nNode.tag)
        
            k = k + 1
            FloodPanel.caption = EllipseText(txtDummy, GetLocalizedStr(934) + ": " + Links(idx).Link, DT_WORD_ELLIPSIS)
            FloodPanel.Value = (k / sc) * 100
            
            If NagScreenIsVisible Then frmNag.lblInfo.caption = GetLocalizedStr(935) + " " & Int(k / sc * 100) & "%"
        
            If TestAll Then nNode.Selected = True
            nNode.EnsureVisible
            If IsExternalLink(nNode.SubItems(1)) Then
                Links(idx).Type = IIf(IsExternalLinkValid(Links(idx).Link), ltcValid, ltcBroken)
            Else
                id = Val(Mid(Links(idx).key, 2))
                If Left(Links(idx).key, 1) = "G" Then
                    a = MenuGrps(id).Actions
                Else
                    a = MenuCmds(id).Actions
                End If
                Links(idx).Type = IIf(IsLinkValid(a), ltcValid, ltcBroken)
            End If
        End If
        If UserCancelled Then Exit For
        If idx <> 0 Then ApplyFilter idx
        If (i Mod 25) = 0 Then DoEvents
    Next i
    If TestAll Then lvLinks.MultiSelect = True
    
    FloodPanel.Value = 0
    SetCtrlsState True
    
    IsTesting = False

End Sub

Private Function SelCount() As Integer

    Dim nItem As ListItem
    Dim sc As Integer
    
    For Each nItem In lvLinks.ListItems
        sc = sc + Abs(nItem.Selected)
    Next nItem
    
    SelCount = sc

End Function

Private Function IsExternalLinkValid(url As String) As Boolean

    UserCancelled = False
    wBrowser.Stop
    wBrowser.Navigate2 "about:blank"
    Do
        DoEvents
    Loop While (wBrowser.Busy And Not UserCancelled)
    
    If Not UserCancelled Then
        BrowserReportedError = False
        IsPorcessing = True
        LinkResult = False
        
        wBrowser.Navigate url
        
        tmrTimeout.Enabled = True
        Do
            DoEvents
        Loop While (IsPorcessing And Not UserCancelled)
        tmrTimeout.Enabled = False
    End If
    
    If UserCancelled Then
        IsExternalLinkValid = True
    Else
        IsExternalLinkValid = LinkResult
    End If
    
    DoEvents

End Function

Private Sub cmdTestSel_Click()

    TestExtLinks False

End Sub

Private Function SetCtrlsState(State As Boolean)

    Dim sc As Integer

    UserCancelled = False
    
    sc = SelCount

    cmdEdit.Enabled = State And sc = 1
    cmdTestAll.Enabled = State And sc > 0
    cmdTestSel.Enabled = State And sc > 0
    cmdClose.Enabled = State
    cmdNextBroken.Enabled = State And sc > 0
    cmdPrevBroken.Enabled = State And sc > 0
    
    chkFilter(0).Enabled = State
    chkFilter(1).Enabled = State
    chkFilter(2).Enabled = State
    
    cmdStop.Visible = Not State
    
    Screen.MousePointer = IIf(State, vbDefault, vbHourglass)

End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF1 Then showHelp "dialogs/blr.htm"

End Sub

Private Sub Form_Load()

    Dim nItem As ListItem

    LocalizeUI
    SetupCharset Me
    
    If Not IsDebug Then SetupSubclassing True
    
    If Val(GetSetting(App.EXEName, "BLRWinPos", "X")) = 0 Then
        CenterForm Me
    Else
        Left = GetSetting(App.EXEName, "BLRWinPos", "X")
        Top = GetSetting(App.EXEName, "BLRWinPos", "Y")
        Width = GetSetting(App.EXEName, "BLRWinPos", "W")
        Height = GetSetting(App.EXEName, "BLRWinPos", "H")
        
        If Left + Width / 2 > Screen.Width Or Top + Height / 2 > Screen.Height Then
            Left = Screen.Width / 2 - Width / 2
            Top = Screen.Height / 2 - Height / 2
        End If
    End If
    
    wBrowser.Silent = True
    
    Set FloodPanel.PictureControl = IIf(LinkVerifyMode = spmcManual, picFlood, frmMain.picFlood)
    
    ReDim Links(0)
    GenerateReport
    
    If LinkVerifyMode = spmcAuto Then
        If Preferences.VerifyLinksOptions.VerifyExternalLinks Then
            TestExtLinks True
        End If
        For Each nItem In lvLinks.ListItems
            If nItem.SmallIcon = 2 Then
                lvLinks.MultiSelect = False
                lvLinks_ItemClick nItem
                lvLinks.MultiSelect = True
                Exit Sub
            End If
        Next nItem
        
        tmrAutoClose.Enabled = True
    End If

End Sub

Private Function HasLink(a As ActionEvents) As Boolean

    HasLink = (a.onclick.Type = atcNewWindow Or a.onclick.Type = atcURL)

End Function

Private Sub GenerateReport()

    Dim gc As Integer
    Dim cc As Integer
    Dim i As Integer
    Dim t As Integer
    
    gc = UBound(MenuGrps)
    cc = UBound(MenuCmds)
    t = gc + cc
    
    FloodPanel.caption = GetLocalizedStr(936)

    If gc Then
        For i = 1 To gc
            FloodPanel.Value = i / t
            With MenuGrps(i)
                If Not IsSubMenu(i) Then
                    If HasLink(.Actions) Then
                        If IsExt(.Actions) Then
                            AddItem NiceGrpCaption(i), .Actions.onclick.url, "G" & i, 1
                        Else
                            AddItem NiceGrpCaption(i), .Actions.onclick.url, "G" & i, IIf(IsLinkValid(.Actions), 3, 2)
                        End If
                    End If
                End If
            End With
        Next i
    End If
    
    If cc Then
        For i = 1 To cc
            FloodPanel.Value = (i + gc) / t
            With MenuCmds(i)
                If HasLink(.Actions) Then
                    If IsExt(.Actions) Then
                        AddItem NiceCmdCaption(i), .Actions.onclick.url, "C" & i, 1
                    Else
                        AddItem NiceCmdCaption(i), .Actions.onclick.url, "C" & i, IIf(IsLinkValid(.Actions), 3, 2)
                    End If
                End If
            End With
        Next i
    End If
    
    ApplyFilter
    
    FloodPanel.Value = 0
    SetCtrlsState True
    
End Sub

Private Function IsExt(a As ActionEvents) As Boolean

    If a.onclick.Type = atcNewWindow Or a.onclick.Type = atcURL Then
        IsExt = IsExternalLink(a.onclick.url)
    Else
        IsExt = True
    End If

End Function

Private Sub AddItem(ItemName As String, url As String, ItemID As String, Optional iconIdx As Integer)

    Dim idx As Integer
    Dim i As Integer

    If LenB(url) <> 0 And Not UsesProtocol(url) Then
        For i = 1 To UBound(Links)
            If Links(i).key = ItemID Then
                idx = i
                Exit For
            End If
        Next i
        
        If idx = 0 Then
            ReDim Preserve Links(UBound(Links) + 1)
            idx = UBound(Links)
        End If
        
        With Links(idx)
            .item = ItemName
            .key = ItemID
            .Link = url
            .Type = iconIdx
        End With
    End If
    
End Sub

Private Function SetItemColor(iconIdx) As Long

    Dim c As Long
    
    Select Case iconIdx
        Case 1  ' Untested
            c = Preferences.DisabledItem
        Case 2  ' Broken
            c = Preferences.BrokenLink
        Case 3  ' Valid
            c = &H80000012
    End Select
    
    SetItemColor = c

End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Not IsDebug Then SetupSubclassing False

    SaveSetting App.EXEName, "BLRWinPos", "X", Left
    SaveSetting App.EXEName, "BLRWinPos", "Y", Top
    SaveSetting App.EXEName, "BLRWinPos", "W", Width
    SaveSetting App.EXEName, "BLRWinPos", "H", Height

End Sub

Private Sub Form_Resize()

    Dim ct As Long
    Dim cl As Long
    
    On Error Resume Next
    
    IsResizing = True
    
    cl = GetClientLeft(Me.hwnd)
    ct = GetClientTop(Me.hwnd)

    With cmdClose
        .Move Width - .Width - 105 - cl, Height - sbInfo.Height - .Height - ct - 120
        lvLinks.Move 90, 285, .Left - 90 * 2, .Top - 285 * 2 + 120
        frameButtons.Move .Left, lvLinks.Top + 240
        cmdStop.Move .Left, lvLinks.Top + lvLinks.Height - cmdStop.Height
    End With
    
    wBrowser.Left = Width + cl * 2
    
    With sbInfo
        picFlood.Move 30, .Top + 60, .Width - 480, .Height - 75
        txtDummy.Width = .Width - 900
    End With
    
    CoolListView lvLinks
    
    IsResizing = False

End Sub

Private Sub lvLinks_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    With lvLinks
        If .SortKey = ColumnHeader.Index - 1 Then
            If .SortOrder = lvwAscending Then
                .SortOrder = lvwDescending
            Else
                .SortOrder = lvwAscending
            End If
        Else
            .SortKey = ColumnHeader.Index - 1
            .SortOrder = lvwAscending
        End If
        .Sorted = True
    End With

End Sub

Private Sub lvLinks_DblClick()

    If IsTesting Then Exit Sub
    If Not lvLinks.SelectedItem Is Nothing Then cmdEdit_Click

End Sub

Private Sub lvLinks_ItemClick(ByVal item As MSComctlLib.ListItem)

    If IsTesting Then Exit Sub
    item.Selected = True
    item.EnsureVisible
    SetCtrlsState True

End Sub

Private Sub tmrAutoClose_Timer()

    Unload Me

End Sub

Private Function IsLinkValid(a As ActionEvents) As Boolean

    Dim p1 As Long
    Dim p2 As Long
    Dim c As ConfigDef
    Dim url As String
    
    url = a.onclick.url
    
    If LenB(url) = 0 Then
        IsLinkValid = True
    Else
        If FolderExists(GetRealLocal.RootWeb) Then
            If a.onclick.Type = atcURL Or a.onclick.Type = atcNewWindow Then
                If UsesProtocol(url) Then
                    IsLinkValid = True
                Else
                    c = Project.UserConfigs(Project.DefaultConfig)
                    If c.Type = ctcRemote Then
                        If LCase(Left(url, Len(c.RootWeb))) <> c.RootWeb Then
                            IsLinkValid = False
                            Exit Function
                        End If
                        url = Replace(SetSlashDir(GetRealLocal.RootWeb + FixURL(url), sdBack), "\\", "\")
                    ElseIf IsExternalLink(url) Then
                        IsLinkValid = True
                        Exit Function
                    End If
                    p1 = InStr(url, "#")
                    p2 = InStr(url, "?")
                    If p1 = 0 And p2 <> 0 Then p1 = p2
                    If p1 > 0 Then url = Left(url, p1 - 1)
                    IsLinkValid = FileExists(url)
                End If
            Else
                IsLinkValid = True
            End If
        Else
            IsLinkValid = False
        End If
    End If

End Function

Private Sub tmrTimeout_Timer()

    tmrTimeout.Enabled = False
    IsPorcessing = False

End Sub

Private Sub wBrowser_DocumentComplete(ByVal pDisp As Object, url As Variant)

    Dim d As HTMLDocument
    Dim np As String

    On Error Resume Next

    If BrowserReportedError Or Not IsPorcessing Then Exit Sub

    DoEvents

    If LenB(url) <> 0 And url <> "http:///" And url <> "about:blank" Then
        Set d = pDisp.Document
        np = LCase(d.nameProp)

        LinkResult = InStr(np, "not found") = 0 And _
                    (np <> "cannot find server") And _
                    (np <> "the page cannot be found") And _
                    (np <> "no se puede encontrar el servidor") And _
                    Left(url, 6) <> "res://"

        If LinkResult Then
            LinkResult = (InStr(d.documentElement.innerHTML, "doNetDetect(") = 0)
        End If
    End If

    IsPorcessing = False

End Sub

Private Sub wBrowser_NavigateComplete2(ByVal pDisp As Object, url As Variant)

    Dim d As HTMLDocument
    Dim np As String
    
    On Error Resume Next

    If LenB(url) <> 0 And url <> "http:///" And url <> "about:blank" Then
    
        ' This is only required because IE4 and IE5 do not support the cool "NavigateError"
        ' event. Otherwise, LinkResult = True would be enough here...
        Set d = pDisp.Document
        np = LCase(d.nameProp)

        LinkResult = InStr(np, "not found") = 0 And _
                    (np <> "cannot find server") And _
                    (np <> "the page cannot be found") And _
                    (np <> "no se puede encontrar el servidor") And _
                    Left(url, 6) <> "res://"

        If LinkResult Then
            LinkResult = (InStr(d.documentElement.innerHTML, "doNetDetect(") = 0)
        End If
        
        wBrowser.Stop
    End If
    
    IsPorcessing = False

End Sub

Private Sub wBrowser_NavigateError(ByVal pDisp As Object, url As Variant, Frame As Variant, StatusCode As Variant, Cancel As Boolean)

    BrowserReportedError = (StatusCode >= &H800C0005 And StatusCode <= &H800C0010) Or _
                            (StatusCode >= &H800C0014 And StatusCode <= &H800C0300)
    LinkResult = False
    IsPorcessing = Not BrowserReportedError
    Cancel = BrowserReportedError

End Sub

Private Sub msgSubClass_NewMessage(ByVal hwnd As Long, uMsg As Long, wParam As Long, lParam As Long, Cancel As Boolean, UseRetVal As Boolean, RetVal As Long)

    If IsResizing Or IsTesting Then Exit Sub

    Cancel = HandleSubclassMsg(hwnd, uMsg, wParam, lParam)

End Sub

Private Sub SetupSubclassing(scState As Boolean)

    If msgSubClass Is Nothing Then Set msgSubClass = New xfxSC
    
    frmBLReportHWND = Me.hwnd
    msgSubClass.SubClassHwnd Me.hwnd, scState

End Sub

Private Sub LocalizeUI()

    caption = GetLocalizedStr(925)
    
    lblListTitle.caption = GetLocalizedStr(926)
    
    cmdEdit.caption = "       " + GetLocalizedStr(339)
    cmdTestSel.caption = "       " + GetLocalizedStr(927)
    cmdTestAll.caption = "       " + GetLocalizedStr(928)
    cmdPrevBroken.caption = "       " + GetLocalizedStr(929)
    cmdNextBroken.caption = "       " + GetLocalizedStr(616)
    
    chkFilter(0).ToolTipText = GetLocalizedStr(930)
    chkFilter(2).ToolTipText = GetLocalizedStr(931)
    chkFilter(1).ToolTipText = GetLocalizedStr(932)
    
    cmdClose.caption = GetLocalizedStr(424)
    cmdStop.caption = GetLocalizedStr(933)
    
End Sub

#End If
