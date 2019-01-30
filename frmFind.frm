VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFind 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find"
   ClientHeight    =   2970
   ClientLeft      =   5385
   ClientTop       =   6105
   ClientWidth     =   6555
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFind.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   6555
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2145
      Left            =   45
      ScaleHeight     =   2145
      ScaleWidth      =   6495
      TabIndex        =   4
      Top             =   795
      Width           =   6495
      Begin VB.CommandButton cmdReplaceAll 
         Caption         =   "Replace &All"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   5280
         TabIndex        =   11
         Top             =   915
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.Frame frameScope 
         Caption         =   "Scope"
         Height          =   915
         Left            =   0
         TabIndex        =   5
         Top             =   -45
         Width           =   2580
         Begin VB.CheckBox chkScopeCommand 
            Caption         =   "Commands"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   255
            TabIndex        =   7
            Top             =   570
            Value           =   1  'Checked
            Width           =   2115
         End
         Begin VB.CheckBox chkScopeGroup 
            Caption         =   "Groups"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   255
            TabIndex        =   6
            Top             =   315
            Value           =   1  'Checked
            Width           =   2115
         End
      End
      Begin VB.Frame frameOptions 
         Caption         =   "Options"
         Height          =   1200
         Left            =   0
         TabIndex        =   12
         Top             =   930
         Width           =   2580
         Begin VB.CheckBox chkOnlyEnabled 
            Caption         =   "Only Enabled Items"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   255
            TabIndex        =   15
            Top             =   810
            Value           =   1  'Checked
            Width           =   2265
         End
         Begin VB.CheckBox chkOPCase 
            Caption         =   "Match Case"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   255
            TabIndex        =   14
            Top             =   570
            Width           =   2265
         End
         Begin VB.CheckBox chkOPWholeWord 
            Caption         =   "Find Whole Word Only"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   255
            TabIndex        =   13
            Top             =   315
            Width           =   2265
         End
      End
      Begin VB.CommandButton cmdCancel 
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
         Height          =   345
         Left            =   5280
         TabIndex        =   16
         Top             =   1785
         Width           =   1200
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "Find &Next"
         Default         =   -1  'True
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   5280
         TabIndex        =   9
         Top             =   60
         Width           =   1200
      End
      Begin VB.CommandButton cmdReplace 
         Caption         =   "&Replace..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   5280
         TabIndex        =   10
         Top             =   480
         Width           =   1200
      End
      Begin MSComctlLib.TreeView tvProperties 
         Height          =   2070
         Left            =   2655
         TabIndex        =   8
         Top             =   60
         Width           =   2550
         _ExtentX        =   4498
         _ExtentY        =   3651
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   353
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         Checkboxes      =   -1  'True
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
   End
   Begin VB.TextBox txtReplace 
      Height          =   315
      Left            =   1275
      TabIndex        =   3
      Top             =   435
      Width           =   3975
   End
   Begin VB.TextBox txtString 
      Height          =   315
      Left            =   1275
      TabIndex        =   1
      Top             =   68
      Width           =   3975
   End
   Begin VB.Label lblReplace 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Replace"
      Height          =   210
      Left            =   540
      TabIndex        =   2
      Top             =   480
      Width           =   630
   End
   Begin VB.Label lblFind 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Find"
      Height          =   210
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   330
   End
   Begin VB.Menu mnuSel 
      Caption         =   "mnuSel"
      Begin VB.Menu mnuSelAll 
         Caption         =   "Select All"
      End
      Begin VB.Menu mnuSelNone 
         Caption         =   "Select None"
      End
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum FindModeConstants
    fmcFind = 0
    fmcReplace = 1
    fmcReplaceAll = 2
End Enum

Private FindMode As FindModeConstants

Dim lg As Integer
Dim lc As Integer
Dim MatchIsSel As Boolean
Dim MatchedSection As String
Dim ci() As String

Private Sub UpdateButtons()

    Dim Ok2Find As Boolean
    Dim nNode As Node
        
    For Each nNode In tvProperties.Nodes
        If nNode.Checked Then
            Ok2Find = True
            Exit For
        End If
    Next nNode
    
    Ok2Find = Ok2Find And LenB(txtString.Text) <> 0 And _
                ((chkScopeGroup.Value = vbChecked) Or _
                    (chkScopeCommand.Value = vbChecked))
    
    cmdFind.Enabled = Ok2Find
    cmdReplace.Enabled = Ok2Find Or cmdReplace.Caption = GetLocalizedStr(142)
    cmdReplaceAll.Enabled = cmdReplace.Enabled

End Sub

Private Sub CreateNodes()

    Dim nNode As Node
    
    lc = 1
    lg = 1
    
    With tvProperties.Nodes
        .Clear
        Set nNode = .Add(, , "ALL", "All Properties", IconIndex("All Properties"))
        nNode.Expanded = True
            Set nNode = .Add("ALL", tvwChild, "Name", "Name", IconIndex("Font Name"))
            Set nNode = .Add("ALL", tvwChild, "Caption", "Caption", IconIndex("Caption"))
            Set nNode = .Add("ALL", tvwChild, "Actions", "Actions", IconIndex("Events"))
                Set nNode = .Add("Actions", tvwChild, "OnClick", "Click", IconIndex("EventClick"))
                    Set nNode = .Add("OnClick", tvwChild, "OnClickURL", "URL / Script", IconIndex("URL"))
                    Set nNode = .Add("OnClick", tvwChild, "OnClickTargetFrame", "Target Frame", IconIndex("Target Frame"))
                Set nNode = .Add("Actions", tvwChild, "OnMouseOver", "Mouse Over", IconIndex("EventOver"))
                    Set nNode = .Add("OnMouseOver", tvwChild, "OnMouseOverURL", "URL / Script", IconIndex("URL"))
                    Set nNode = .Add("OnMouseOver", tvwChild, "OnMouseOverTargetFrame", "Target Frame", IconIndex("Target Frame"))
                Set nNode = .Add("Actions", tvwChild, "OnDoubleClick", "Double Click", IconIndex("EventDoubleClick"))
                    Set nNode = .Add("OnDoubleClick", tvwChild, "OnDoubleClickURL", "URL / Script", IconIndex("URL"))
                    Set nNode = .Add("OnDoubleClick", tvwChild, "OnDoubleClickTargetFrame", "Target Frame", IconIndex("Target Frame"))
            Set nNode = .Add("ALL", tvwChild, "StatusText", "Status Text", IconIndex("Status Text"))
    End With
    
    With tvProperties
        If (chkScopeCommand.Value = vbUnchecked) And (chkScopeGroup.Value = vbUnchecked) Then
            .Nodes.Clear
        End If
    End With
    
End Sub

Private Sub chkScopeCommand_Click()

    CreateNodes
    UpdateButtons

End Sub

Private Sub chkScopeGroup_Click()

    CreateNodes
    UpdateButtons

End Sub

Private Sub cmdCancel_Click()

    Hide

End Sub

Private Sub MatchFound(id As Integer, oType As DataOnClibboardConstants, PropertyStr As String)

    Dim iKey As String

    Select Case oType
        Case docGroup
            iKey = "G"
            Add2CIArray MenuGrps(id).Name
        Case docCommand
            iKey = "C"
            Add2CIArray MenuCmds(id).Name
    End Select
    
    iKey = iKey & id
    
    With frmMain
        .tvMenus.Nodes(iKey).EnsureVisible
        .tvMenus.Nodes(iKey).Selected = True
        .SynchViews
        .UpdateControls
        
        MatchedSection = PropertyStr
        Select Case PropertyStr
            Case "Name"
                MatchIsSel = True
            Case "Caption"
                MakeSelection .txtCaption
            Case "OnMouseOverURL"
                .tsCmdType.Tabs("tsOver").Selected = True
                MakeSelection .txtURL
            Case "OnMouseOverTargetFrame"
                .tsCmdType.Tabs("tsOver").Selected = True
                MakeSelection .cmbTargetFrame
            Case "OnClickURL"
                .tsCmdType.Tabs("tsClick").Selected = True
                MakeSelection .txtURL
            Case "OnClickTargetFrame"
                .tsCmdType.Tabs("tsClick").Selected = True
                MakeSelection .cmbTargetFrame
            Case "OnDoubleClickURL"
                .tsCmdType.Tabs("tsDoubleClick").Selected = True
                MakeSelection .txtURL
            Case "OnDoubleClickTargetFrame"
                .tsCmdType.Tabs("tsDoubleClick").Selected = True
                MakeSelection .cmbTargetFrame
            Case "StatusText"
                MakeSelection .txtStatus
        End Select
    End With
    
    If FindMode <> fmcFind Then DoFind

End Sub

Private Sub MakeSelection(SelOn As Object)

    On Error Resume Next

    SelOn.SetFocus
    SelOn.SelStart = InStr(SelOn.Text, txtString.Text) - 1
    SelOn.SelLength = Len(txtString.Text)
    
    DoEvents
    
    MatchIsSel = True

End Sub

Private Sub cmdFind_Click()

    SetDlgState False

    ReDim ci(0)
    
    FindMode = fmcFind
    DoFind
    
    SetDlgState True

End Sub

Private Sub SetDlgState(State As Boolean)

    cmdCancel.Enabled = State
    cmdFind.Enabled = State
    cmdReplace.Enabled = State
    cmdReplaceAll.Enabled = State
    frmMain.Enabled = State
    Me.Enabled = State
    
    Screen.MousePointer = IIf(State, vbDefault, vbHourglass)

End Sub

Friend Sub DoFind()

    Dim i As Integer
    Dim mf As Boolean
    Static RecCount As Integer
    
    RecCount = RecCount + 1
    
    If FindMode <> fmcFind And MatchIsSel Then
        Select Case MatchedSection
            Case "Name"
                frmMain.tvMenus_AfterLabelEdit 0, Replace(frmMain.tvMenus.SelectedItem.Text, txtString.Text, txtReplace.Text, , , vbTextCompare)
                frmMain.tvMenus.SelectedItem.Text = Replace(frmMain.tvMenus.SelectedItem.Text, txtString.Text, txtReplace.Text, , , vbTextCompare)
            Case "Caption"
                frmMain.txtCaption.Text = Replace(frmMain.txtCaption.Text, txtString.Text, txtReplace.Text, , , vbTextCompare)
            Case "OnMouseOverTargetFrame"
                frmMain.tsCmdType.Tabs("tsOver").Selected = True
                frmMain.UpdateControls
                frmMain.cmbTargetFrame.Text = Replace(frmMain.cmbTargetFrame.Text, txtString.Text, txtReplace.Text, , , vbTextCompare)
            Case "OnClickTargetFrame"
                frmMain.tsCmdType.Tabs("tsClick").Selected = True
                frmMain.UpdateControls
                frmMain.cmbTargetFrame.Text = Replace(frmMain.cmbTargetFrame.Text, txtString.Text, txtReplace.Text, , , vbTextCompare)
            Case "OnDoubleClickTargetFrame"
                frmMain.tsCmdType.Tabs("tsDoubleClick").Selected = True
                frmMain.UpdateControls
                frmMain.cmbTargetFrame.Text = Replace(frmMain.cmbTargetFrame.Text, txtString.Text, txtReplace.Text, , , vbTextCompare)
            Case "OnClickURL"
                frmMain.tsCmdType.Tabs("tsClick").Selected = True
                frmMain.UpdateControls
                frmMain.txtURL.Text = Replace(frmMain.txtURL.Text, txtString.Text, txtReplace.Text, , , vbTextCompare)
            Case "OnMouseOverURL"
                frmMain.tsCmdType.Tabs("tsOver").Selected = True
                frmMain.UpdateControls
                frmMain.txtURL.Text = Replace(frmMain.txtURL.Text, txtString.Text, txtReplace.Text, , , vbTextCompare)
            Case "OnDoubleClickURL"
                frmMain.tsCmdType.Tabs("tsDoubleClick").Selected = True
                frmMain.UpdateControls
                frmMain.txtURL.Text = Replace(frmMain.txtURL.Text, txtString.Text, txtReplace.Text, , , vbTextCompare)
            Case "StatusText"
                frmMain.txtStatus.Text = Replace(frmMain.txtStatus.Text, txtString.Text, txtReplace.Text, , , vbTextCompare)
        End Select
        MatchIsSel = False
        If FindMode = fmcReplace Then GoTo AbortSub
    End If
    
    If chkScopeGroup.Value = vbChecked Then
        For i = lg To UBound(MenuGrps)
            If (((chkOnlyEnabled.Value = vbChecked) And (Not MenuGrps(i).disabled)) Or _
                (chkOnlyEnabled.Value = vbUnchecked)) And _
                (Not HasChanged(MenuGrps(i).Name)) Then
                If tvProperties.Nodes("Name").Checked Then
                    If IsMatch(MenuGrps(i).Name) Then
                        'lg = i + 1
                        MatchFound i, docGroup, "Name"
                        mf = True
                    End If
                End If
                If tvProperties.Nodes("Caption").Checked Then
                    If IsMatch(MenuGrps(i).Caption) Then
                        'lg = i + 1
                        MatchFound i, docGroup, "Caption"
                        mf = True
                    End If
                End If
                If tvProperties.Nodes("OnMouseOverURL").Checked Then
                    If IsMatch(MenuGrps(i).Actions.onmouseover.url) Then
                        'lg = i + 1
                        MatchFound i, docGroup, "OnMouseOverURL"
                        mf = True
                    End If
                End If
                If tvProperties.Nodes("OnMouseOverTargetFrame").Checked Then
                    If IsMatch(MenuGrps(i).Actions.onmouseover.TargetFrame) Then
                        'lg = i + 1
                        MatchFound i, docGroup, "OnMouseOverTargetFrame"
                        mf = True
                    End If
                End If
                If tvProperties.Nodes("OnClickURL").Checked Then
                    If IsMatch(MenuGrps(i).Actions.onclick.url) Then
                        'lg = i + 1
                        MatchFound i, docGroup, "OnClickURL"
                        mf = True
                    End If
                End If
                If tvProperties.Nodes("OnClickTargetFrame").Checked Then
                    If IsMatch(MenuGrps(i).Actions.onclick.TargetFrame) Then
                        'lg = i + 1
                        MatchFound i, docGroup, "OnClickTargetFrame"
                        mf = True
                    End If
                End If
                If tvProperties.Nodes("OnDoubleClickURL").Checked Then
                    If IsMatch(MenuGrps(i).Actions.OnDoubleClick.url) Then
                        'lg = i + 1
                        MatchFound i, docGroup, "OnDoubleClickURL"
                        mf = True
                    End If
                End If
                If tvProperties.Nodes("OnDoubleClickTargetFrame").Checked Then
                    If IsMatch(MenuGrps(i).Actions.OnDoubleClick.TargetFrame) Then
                        'lg = i + 1
                        MatchFound i, docGroup, "OnDoubleClickTargetFrame"
                        mf = True
                    End If
                End If
                If tvProperties.Nodes("StatusText").Checked Then
                    If IsMatch(MenuGrps(i).WinStatus) Then
                        'lg = i + 1
                        MatchFound i, docGroup, "StatusText"
                        mf = True
                    End If
                End If
            End If
            If mf Then
                lg = i + 1
                GoTo AbortSub
            End If
        Next i
    End If
    
    If chkScopeCommand.Value = vbChecked Then
        For i = lc To UBound(MenuCmds)
            If (((chkOnlyEnabled.Value = vbChecked) And (Not MenuCmds(i).disabled)) Or _
                (chkOnlyEnabled.Value = vbUnchecked)) And _
                (Not HasChanged(MenuCmds(i).Name)) Then
                If tvProperties.Nodes("Name").Checked Then
                    If IsMatch(MenuCmds(i).Name) Then
                        'lc = i + 1
                        MatchFound i, docCommand, "Name"
                        mf = True
                    End If
                End If
                If tvProperties.Nodes("Caption").Checked Then
                    If IsMatch(MenuCmds(i).Caption) Then
                        'lc = i + 1
                        MatchFound i, docCommand, "Caption"
                        mf = True
                    End If
                End If
                If tvProperties.Nodes("OnMouseOverURL").Checked Then
                    If IsMatch(MenuCmds(i).Actions.onmouseover.url) Then
                        'lc = i + 1
                        MatchFound i, docCommand, "OnMouseOverURL"
                        mf = True
                    End If
                End If
                If tvProperties.Nodes("OnMouseOverTargetFrame").Checked Then
                    If IsMatch(MenuCmds(i).Actions.onmouseover.TargetFrame) Then
                        'lc = i + 1
                        MatchFound i, docCommand, "OnMouseOverTargetFrame"
                        mf = True
                    End If
                End If
                If tvProperties.Nodes("OnClickURL").Checked Then
                    If IsMatch(MenuCmds(i).Actions.onclick.url) Then
                        'lc = i + 1
                        MatchFound i, docCommand, "OnClickURL"
                        mf = True
                    End If
                End If
                If tvProperties.Nodes("OnClickTargetFrame").Checked Then
                    If IsMatch(MenuCmds(i).Actions.onclick.TargetFrame) Then
                        'lc = i + 1
                        MatchFound i, docCommand, "OnClickTargetFrame"
                        mf = True
                    End If
                End If
                If tvProperties.Nodes("OnDoubleClickURL").Checked Then
                    If IsMatch(MenuCmds(i).Actions.OnDoubleClick.url) Then
                        'lc = i + 1
                        MatchFound i, docCommand, "OnDoubleClickURL"
                        mf = True
                    End If
                End If
                If tvProperties.Nodes("OnDoubleClickTargetFrame").Checked Then
                    If IsMatch(MenuCmds(i).Actions.OnDoubleClick.TargetFrame) Then
                        'lc = i + 1
                        MatchFound i, docCommand, "OnDoubleClickTargetFrame"
                        mf = True
                    End If
                End If
                If tvProperties.Nodes("StatusText").Checked Then
                    If IsMatch(MenuCmds(i).WinStatus) Then
                        'lc = i + 1
                        MatchFound i, docCommand, "StatusText"
                        mf = True
                    End If
                End If
            End If
            If mf Then
                lc = i + 1
                GoTo AbortSub
            End If
        Next i
    End If
    
    lg = 1
    lc = 1
    
AbortSub:
    RecCount = RecCount - 1
    
    If RecCount = 0 And ((lg = 1 And lc = 1) Or (FindMode = fmcReplaceAll)) Then
        MsgBox "No more matches found", vbInformation + vbOKOnly, "Find/Replace"
    End If
    
End Sub

Private Function IsMatch(FindIn As String) As Boolean

    Dim s1 As String
    Dim s2 As String
    
    s1 = FindIn
    s2 = txtString.Text
    
    If chkOPCase.Value = vbUnchecked Then
        s1 = LCase$(s1)
        s2 = LCase$(s2)
    End If
    
    If chkOPWholeWord.Value = vbChecked Then
        IsMatch = (InStr(s1, " " + s2 + " ") > 0) Or _
                    Left$(s1, Len(s2) + 1) = s2 + " " Or _
                    Right$(s1, Len(s2) + 1) = " " + s2 Or _
                    s1 = s2
    Else
        IsMatch = (InStr(s1, s2) > 0)
    End If

End Function

Private Function HasChanged(iName As String)

    HasChanged = (InStr(Join(ci, "|") + "|", "|" + iName + "|") > 0)

End Function

Private Sub Add2CIArray(s As String)

    ReDim Preserve ci(UBound(ci) + 1)
    ci(UBound(ci)) = s

End Sub

Private Sub cmdReplace_Click()

    If cmdReplace.Caption = GetLocalizedStr(142) Then
        Top = Top - 200
        SwitchToReplaceMode
    Else
        FindMode = fmcReplace
        StartReplaceOperation
    End If

End Sub

Private Sub StartReplaceOperation()

    SetDlgState False
    
    ReDim ci(0)
        
    IsReplacing = True
    MatchedSection = vbNullString
    DoFind
    IsReplacing = False
    
    SetDlgState True
    
    frmMain.SaveState GetLocalizedStr(612)
    
    If InMapMode Then
        DontRefreshMap = False
        frmMain.RefreshMap
    End If

End Sub

Private Sub cmdReplaceAll_Click()

    lg = 1
    lc = 1
    FindMode = fmcReplaceAll
    StartReplaceOperation

End Sub

Private Sub Form_Load()

    LocalizeUI
    
    tvProperties.ImageList = frmMain.ilIcons
    
    mnuSel.Visible = False
    DoEvents
    CenterForm Me
    FixCtrls4Skin Me
    
    DoEvents
    
    CreateNodes

End Sub

Friend Sub SwitchToReplaceMode()

    Caption = GetLocalizedStr(606)

    cmdReplace.Caption = "&" + GetLocalizedStr(606)
    picContainer.Top = 800
    Height = 2955 + GetClientTop(Me.hWnd)
    If LenB(txtString.Text) <> 0 Then
        txtReplace.SetFocus
    Else
        txtString.SetFocus
    End If
    cmdReplaceAll.Visible = True
    
    UpdateButtons

End Sub

Friend Sub SwitchToFindMode()

    Caption = GetLocalizedStr(605)

    cmdReplace.Caption = GetLocalizedStr(142)
    picContainer.Top = 435
    Height = 2600 + GetClientTop(Me.hWnd)
    txtString.SetFocus
    cmdReplaceAll.Visible = False
    cmdReplace.Enabled = True

End Sub

Private Sub mnuSelAll_Click()

    Dim nNode As Node
    
    For Each nNode In tvProperties.Nodes
        nNode.Checked = True
    Next nNode

End Sub

Private Sub mnuSelNone_Click()

    Dim nNode As Node
    
    For Each nNode In tvProperties.Nodes
        nNode.Checked = False
    Next nNode

End Sub

Private Sub tvProperties_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = vbRightButton Then
        If Not tvProperties.HitTest(x, y) Is Nothing Then
            tvProperties.HitTest(x, y).Selected = True
        End If
        PopupMenu mnuSel, vbRightButton, tvProperties.Left + x, tvProperties.Top + y
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
        Set Node = Node.Parent
        Do Until Node Is Nothing
            Set cNode = Node.Child
            Node.Checked = False
            Do Until cNode Is Nothing
                If cNode.Checked Then
                    cNode.Parent.Checked = True
                    Exit Do
                End If
                Set cNode = cNode.Next
            Loop
            Set Node = Node.Parent
        Loop
    End If
    
    UpdateButtons

End Sub

Private Sub txtReplace_Change()

    UpdateButtons

End Sub

Private Sub txtString_Change()

    UpdateButtons

End Sub

Friend Sub LocalizeUI()

    frameScope.Caption = GetLocalizedStr(607)
    chkScopeGroup.Caption = GetLocalizedStr(505)
    chkScopeCommand.Caption = GetLocalizedStr(506)
    
    frameOptions.Caption = GetLocalizedStr(337)
    chkOPWholeWord.Caption = GetLocalizedStr(608)
    chkOPCase.Caption = GetLocalizedStr(609)
    chkOnlyEnabled.Caption = GetLocalizedStr(610)
    
    cmdFind.Caption = GetLocalizedStr(141)
    cmdReplace.Caption = GetLocalizedStr(142)
    cmdReplaceAll.Caption = GetLocalizedStr(611)
    cmdCancel.Caption = GetLocalizedStr(424)
    
    mnuSelAll.Caption = GetLocalizedStr(503)
    mnuSelNone.Caption = GetLocalizedStr(504)

    lblFind.Caption = GetLocalizedStr(605)
    lblReplace.Caption = GetLocalizedStr(606)

    FixContolsWidth Me

End Sub
