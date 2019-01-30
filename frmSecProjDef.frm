VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{9A2947B4-87EE-40EE-A3EF-32BDC32D5726}#1.0#0"; "xfxline3d.ocx"
Begin VB.Form frmSecProjDef 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Secondary Projects"
   ClientHeight    =   3390
   ClientLeft      =   5580
   ClientTop       =   5655
   ClientWidth     =   7020
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
   ScaleHeight     =   3390
   ScaleWidth      =   7020
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSelAll 
      Caption         =   "Select All"
      Height          =   360
      Left            =   5850
      TabIndex        =   5
      Top             =   1500
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   360
      Left            =   5850
      TabIndex        =   10
      Top             =   2940
      Width           =   1095
   End
   Begin VB.CommandButton cmdCompile 
      Caption         =   "Compile"
      Height          =   360
      Left            =   90
      TabIndex        =   7
      Top             =   2940
      Width           =   1095
   End
   Begin VB.CommandButton cmdInstall 
      Caption         =   "Install..."
      Height          =   360
      Left            =   1305
      TabIndex        =   8
      Top             =   2940
      Width           =   1095
   End
   Begin VB.TextBox dummyText 
      Height          =   315
      Left            =   3375
      MultiLine       =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   135
      Visible         =   0   'False
      Width           =   2340
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   360
      Left            =   4635
      TabIndex        =   9
      Top             =   2940
      Width           =   1095
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add..."
      Height          =   360
      Left            =   5850
      TabIndex        =   3
      Top             =   315
      Width           =   1095
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   360
      Left            =   5850
      TabIndex        =   4
      Top             =   765
      Width           =   1095
   End
   Begin MSComctlLib.ListView lvProjects 
      Height          =   2415
      Left            =   90
      TabIndex        =   2
      Top             =   315
      Width           =   5640
      _ExtentX        =   9948
      _ExtentY        =   4260
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
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
         Key             =   "chName"
         Text            =   "Title"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "chLocation"
         Text            =   "Location"
         Object.Width           =   2540
      EndProperty
   End
   Begin xfxLine3D.ucLine3D uc3DLine1 
      Height          =   30
      Left            =   0
      TabIndex        =   6
      Top             =   2835
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   53
   End
   Begin VB.Label lblProjects 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Projects"
      Height          =   195
      Left            =   90
      TabIndex        =   0
      Top             =   90
      UseMnemonic     =   0   'False
      Width           =   585
   End
End
Attribute VB_Name = "frmSecProjDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()

    Dim i As Integer
    Dim fs() As String

    With frmMain.cDlg
        .DialogTitle = GetLocalizedStr(255)
        .Flags = cdlOFNCreatePrompt + cdlOFNHideReadOnly + cdlOFNLongNames + cdlOFNNoReadOnlyReturn + cdlOFNAllowMultiselect + cdlOFNExplorer
        .filter = GetLocalizedStr(256) + "|*.dmb"
        .FileName = ""
        On Error Resume Next
        .ShowOpen
        On Error GoTo 0
        If Err.number > 0 Or LenB(.FileName) = 0 Then GoTo ExitSub
        If InStr(.FileName, Chr(0)) Then
            fs = Split(.FileName, Chr(0))
            For i = 1 To UBound(fs)
                AddItem fs(0) + "\" + fs(i)
            Next i
        Else
            AddItem .FileName
        End If
    End With
    
    UpdateSecondaryProjectsConfig
    
ExitSub:

End Sub

Private Sub cmdCancel_Click()

    SecProjMode = spmcUndefined
    Unload Me

End Sub

Private Sub cmdCompile_Click()

    Me.Enabled = False

    SaveProjects
    frmMain.ToolsCompile
    
    Me.Enabled = True

End Sub

Private Sub cmdDelete_Click()

    Dim sIdx As Integer

    If Not lvProjects.SelectedItem Is Nothing Then
        sIdx = lvProjects.SelectedItem.Index
        lvProjects.ListItems.Remove sIdx
        
        If lvProjects.ListItems.Count > 0 Then
            If sIdx > lvProjects.ListItems.Count Then sIdx = lvProjects.ListItems.Count
            With lvProjects.ListItems(sIdx)
                .EnsureVisible
                .Selected = True
            End With
        End If
    End If
    
    UpdateSecondaryProjectsConfig

End Sub

Private Sub cmdOK_Click()

    UpdateSecondaryProjectsConfig
    
    Unload Me

End Sub

Private Sub UpdateSecondaryProjectsConfig()

    Dim i As Integer
    
    If SecProjMode = spmcFromInstallMenus Then
        SaveProjects
    Else
        With lvProjects.ListItems
            ReDim Project.SecondaryProjects(.Count)
            For i = 1 To .Count
                Project.SecondaryProjects(i) = .item(i).tag
            Next i
        End With
    End If

End Sub

Private Sub SaveProjects()

    Dim sItem As ListItem
    Dim sPrj As ProjectDef
    
    ReDim SelSecProjects(0)
    ReDim SelSecProjectsTitles(0)
    For Each sItem In lvProjects.ListItems
        If sItem.Checked Then
            ReDim Preserve SelSecProjects(UBound(SelSecProjects) + 1)
            ReDim Preserve SelSecProjectsTitles(UBound(SelSecProjectsTitles) + 1)
            sPrj = GetProjectProperties(sItem.tag)
            Close #ff
            SelSecProjects(UBound(SelSecProjects)) = sPrj.JSFileName
            SelSecProjectsTitles(UBound(SelSecProjectsTitles)) = sPrj.Name
        End If
    Next sItem

End Sub

Private Sub cmdInstall_Click()

    SaveProjects
    
    If SecProjMode = spmcFromStdDlg Then
        If UBound(SelSecProjects) = 0 Then
            'MsgBox GetLocalizedStr(827), vbInformation + vbOKOnly, GetLocalizedStr(828)
        Else
            SecProjMode = spmcFromStdDlg
        End If
    End If
    
    frmLCMan.Show vbModal

End Sub

Private Sub cmdSelAll_Click()

    Dim sItem As ListItem
    
    For Each sItem In lvProjects.ListItems
        sItem.Checked = Not sItem.Ghosted
    Next sItem

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF1 Then showHelp "dialogs/sec_proj.htm"

End Sub

Private Sub Form_Load()

    Dim i As Integer

    SetupCharset Me
    LocalizeUI
    CenterForm Me
    
    With lvProjects
        Set .SmallIcons = frmMain.ilIcons
        .ColumnHeaders(1).Width = .Width - dummyText.Width - 30
        .ColumnHeaders(2).Width = .Width - .ColumnHeaders(1).Width - 60
    End With
    
    With Project
        For i = 1 To UBound(.SecondaryProjects)
            AddItem .SecondaryProjects(i)
        Next i
    End With
    
    If SecProjMode = spmcFromInstallMenus Then
        cmdInstall.Visible = False
        cmdCompile.Visible = False
    End If
    
    CoolListView lvProjects

End Sub

Private Sub AddItem(FileName As String)

    Dim nItem As ListItem
    Dim cPrj As ProjectDef
    Dim Exists As Boolean
    Dim UsesFrames As Boolean
    
    If FileName = Project.FileName Then
        MsgBox "The master project cannot be used as a secondary project of itself", vbInformation + vbOKOnly, "Error Adding Secondary Project"
        Exit Sub
    End If
    
    For Each nItem In lvProjects.ListItems
        If nItem.tag = FileName Then
            nItem.EnsureVisible
            nItem.Selected = True
            Exit Sub
        End If
    Next nItem

    Exists = FileExists(FileName)
    
    If Exists Then
        cPrj = GetProjectProperties(FileName, False)
        Close #ff
        cSep = Chr(255) + Chr(255)
        
        UsesFrames = cPrj.UserConfigs(cPrj.DefaultConfig).Frames.UseFrames
        If UsesFrames Then
            MsgBox "This project cannot be added because it uses Frames Support", vbInformation + vbOKOnly, "Error Adding Secondary Project"
            Exit Sub
        End If
    End If

    Set nItem = lvProjects.ListItems.Add(, , cPrj.Name, , IconIndex("DMBProjectIcon"))
    nItem.Bold = True
    If Exists And Not UsesFrames Then
        nItem.Ghosted = (Val(cPrj.version) < 400000)
        nItem.ForeColor = IIf(Val(cPrj.version) < 400000, &H80000011, vbBlack)
        'nItem.SubItems(1) = EllipseText(dummyText, FileName, DT_PATH_ELLIPSIS)
        nItem.SubItems(1) = FileName
    Else
        If LenB(nItem.Text) = 0 Then nItem.Text = "(unknown)"
        nItem.Ghosted = True
        nItem.ForeColor = vbRed
        'nItem.SubItems(1) = EllipseText(dummyText, FileName, DT_PATH_ELLIPSIS)
        nItem.SubItems(1) = FileName
    End If
    nItem.Checked = False
    nItem.ListSubItems(1).ForeColor = nItem.ForeColor
    nItem.tag = FileName

End Sub

Private Sub lvProjects_ItemCheck(ByVal item As MSComctlLib.ListItem)

    If item.Checked Then
        If item.Ghosted Then
            MsgBox GetLocalizedStr(829), vbInformation + vbOKOnly, GetLocalizedStr(828)
            item.Checked = False
        End If
    End If

End Sub

Private Sub LocalizeUI()

    Me.caption = GetLocalizedStr(825)
    
    lblProjects.caption = GetLocalizedStr(826)

    cmdAdd.caption = GetLocalizedStr(338)
    cmdDelete.caption = GetLocalizedStr(808)
    cmdSelAll.caption = GetLocalizedStr(503)
    cmdDelete.caption = GetLocalizedStr(808)
    
    cmdOK.caption = GetLocalizedStr(186)
    cmdCancel.caption = GetLocalizedStr(187)
    
    cmdCompile.caption = Replace(GetLocalizedStr(161), "&", "")
    cmdInstall.caption = GetLocalizedStr(468)
    
    lvProjects.ColumnHeaders(1).Text = GetLocalizedStr(918)
    lvProjects.ColumnHeaders(2).Text = GetLocalizedStr(916)

End Sub
