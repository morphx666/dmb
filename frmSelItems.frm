VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{9A2947B4-87EE-40EE-A3EF-32BDC32D5726}#1.0#0"; "xfxline3d.ocx"
Begin VB.Form frmSelItems 
   Caption         =   "Select Items"
   ClientHeight    =   6345
   ClientLeft      =   5820
   ClientTop       =   5310
   ClientWidth     =   7620
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
   ScaleHeight     =   6345
   ScaleWidth      =   7620
   Begin VB.CommandButton cmdSelOp 
      Caption         =   "Inv"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   4845
      TabIndex        =   18
      Top             =   105
      Width           =   510
   End
   Begin VB.CommandButton cmdSelOp 
      Caption         =   "None"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   4320
      TabIndex        =   17
      Top             =   105
      Width           =   510
   End
   Begin VB.CommandButton cmdSelOp 
      Caption         =   "All"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   3810
      TabIndex        =   16
      Top             =   105
      Width           =   510
   End
   Begin MSComctlLib.StatusBar sbInfo 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   15
      Top             =   6075
      Width           =   7620
      _ExtentX        =   13441
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12912
            Key             =   "pInfo"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilIcons 
      Left            =   645
      Top             =   4830
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.Frame frameFilter 
      Caption         =   "Filter"
      Height          =   5175
      Left            =   75
      TabIndex        =   1
      Top             =   30
      Width           =   3645
      Begin MSComctlLib.ListView lvFont 
         Height          =   570
         Left            =   150
         TabIndex        =   14
         Top             =   2715
         Width           =   3240
         _ExtentX        =   5715
         _ExtentY        =   1005
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         Enabled         =   0   'False
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Font"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lvImages 
         Height          =   1275
         Left            =   150
         TabIndex        =   13
         Top             =   3705
         Width           =   3240
         _ExtentX        =   5715
         _ExtentY        =   2249
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ilIcons"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         Enabled         =   0   'False
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Image"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.CheckBox chkImage 
         Caption         =   "Uses the Image"
         Height          =   195
         Left            =   150
         TabIndex        =   12
         Top             =   3465
         Width           =   3405
      End
      Begin VB.CheckBox chkFontType 
         Caption         =   "Uses the Font"
         Height          =   195
         Left            =   150
         TabIndex        =   11
         Top             =   2475
         Width           =   3405
      End
      Begin xfxLine3D.ucLine3D uc3DLine1 
         Height          =   30
         Left            =   75
         TabIndex        =   10
         Top             =   2295
         Width           =   3420
         _ExtentX        =   6033
         _ExtentY        =   53
      End
      Begin VB.TextBox txtText 
         Height          =   285
         Left            =   150
         TabIndex        =   9
         Top             =   1860
         Width           =   2145
      End
      Begin VB.ComboBox cmbToolbars 
         Height          =   315
         Left            =   150
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1185
         Width           =   2145
      End
      Begin VB.CheckBox chkCascade 
         Caption         =   "Cascade"
         Height          =   240
         Left            =   150
         TabIndex        =   3
         Top             =   622
         Width           =   3405
      End
      Begin VB.CheckBox chkEnabled 
         Caption         =   "Enabled"
         Height          =   240
         Left            =   150
         TabIndex        =   2
         Top             =   300
         Value           =   2  'Grayed
         Width           =   3405
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Caption Contains"
         Height          =   195
         Left            =   150
         TabIndex        =   8
         Top             =   1620
         Width           =   1230
      End
      Begin VB.Label lblBelongTB 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Belonging to the Toolbar"
         Height          =   195
         Left            =   150
         TabIndex        =   4
         Top             =   945
         Width           =   1755
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   360
      Left            =   5505
      TabIndex        =   6
      Top             =   5445
      Width           =   930
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   360
      Left            =   6570
      TabIndex        =   7
      Top             =   5445
      Width           =   930
   End
   Begin MSComctlLib.ListView lvItems 
      Height          =   4785
      Left            =   3810
      TabIndex        =   0
      Top             =   420
      Width           =   3675
      _ExtentX        =   6482
      _ExtentY        =   8440
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   4233
      EndProperty
   End
   Begin VB.Menu mnuSel 
      Caption         =   "Selection"
      Begin VB.Menu mnuSelAll 
         Caption         =   "Select All"
      End
      Begin VB.Menu mnuSelNone 
         Caption         =   "Select None"
      End
   End
End
Attribute VB_Name = "frmSelItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim LastEnabledState As Integer
Dim CusSelBack() As String
Dim WithEvents msgSubClass As xfxSC
Attribute msgSubClass.VB_VarHelpID = -1

Private Sub chkImage_Click()

    lvImages.Enabled = (chkImage.Value = vbChecked)
    UpdateItemsList

End Sub

Private Sub chkCascade_Click()

    UpdateItemsList

End Sub

Private Sub chkEnabled_Click()

    Static IgnoreEvent As Boolean
    
    If IgnoreEvent Then Exit Sub
    IgnoreEvent = True

    If LastEnabledState = vbChecked Then
        chkEnabled.Value = vbGrayed
    Else
        If LastEnabledState = vbUnchecked Then
            chkEnabled.Value = vbChecked
        Else
            If LastEnabledState = vbGrayed Then
                chkEnabled.Value = vbUnchecked
            End If
        End If
    End If

    UpdateItemsList
    LastEnabledState = chkEnabled.Value
    
    IgnoreEvent = False

End Sub

Private Sub chkFontType_Click()

    lvFont.Enabled = (chkFontType.Value = vbChecked)
    UpdateItemsList

End Sub

Private Sub cmbToolbars_Click()

    UpdateItemsList

End Sub

Private Sub cmdCancel_Click()

    dmbClipboard.CustomSel = CusSelBack
    Unload Me

End Sub

Private Sub cmdOK_Click()
    
    Unload Me

End Sub

Private Sub ApplyCustomSel()

    Dim nItem As ListItem
    Dim i As Integer

    With dmbClipboard
        Erase .CustomSel
        For Each nItem In lvItems.ListItems
            If nItem.Checked Then
                i = i + 1
                ReDim Preserve .CustomSel(i)
                .CustomSel(i) = nItem.tag
            End If
        Next nItem
    End With
    
    UpdateOKButtonState

End Sub

Private Sub cmdSelOp_Click(Index As Integer)

    Dim nItem As ListItem
    
    For Each nItem In lvItems.ListItems
        Select Case Index
            Case 0: nItem.Checked = True
            Case 1: nItem.Checked = False
            Case 2: nItem.Checked = Not nItem.Checked
        End Select
    Next nItem
    
    ApplyCustomSel

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF1 Then showHelp "dialogs/ss_ais.htm"

End Sub

Private Sub Form_Load()

    CenterForm Me
    SetupCharset Me
    LocalizeUI
    
    LastEnabledState = chkEnabled.Value
    
    CusSelBack = dmbClipboard.CustomSel
    mnuSel.Visible = False
    CenterForm Me
    
    Set msgSubClass = New xfxSC
    msgSubClass.SubClassHwnd Me.hwnd, True
    
    InitDlg
    FixCtrls4Skin Me

    UpdateItemsList

End Sub

Private Sub InitDlg()

    Dim i As Integer

    With cmbToolbars
        .Clear
        .AddItem "(any)"
        For i = 1 To UBound(Project.Toolbars)
            .AddItem Project.Toolbars(i).Name
        Next i
        .ListIndex = 0
    End With
    
    caption = "Select the items to receive the selected styles from the '"
    
    If dmbClipboard.ObjSrc = docGroup Then
        caption = caption + NiceGrpCaption(GetIDByName(dmbClipboard.GrpContents.Name)) + "' Group"
        For i = 1 To UBound(MenuGrps)
            Add2LV lvFont, GetFontStr(MenuGrps(i).DefNormalFont)
            Add2LV lvFont, GetFontStr(MenuGrps(i).DefHoverFont)
            With MenuGrps(i).tbiLeftImage
                If LenB(.NormalImage) <> 0 Then Add2LV lvImages, GetFileName(.NormalImage), AddNewIcon(.NormalImage)
                If LenB(.HoverImage) <> 0 Then Add2LV lvImages, GetFileName(.HoverImage), AddNewIcon(.HoverImage)
            End With
            With MenuGrps(i).tbiRightImage
                If LenB(.NormalImage) <> 0 Then Add2LV lvImages, GetFileName(.NormalImage), AddNewIcon(.NormalImage)
                If LenB(.HoverImage) <> 0 Then Add2LV lvImages, GetFileName(.HoverImage), AddNewIcon(.HoverImage)
            End With
            With MenuGrps(i).tbiBackImage
                If LenB(.NormalImage) <> 0 Then Add2LV lvImages, GetFileName(.NormalImage), AddNewIcon(.NormalImage)
                If LenB(.HoverImage) <> 0 Then Add2LV lvImages, GetFileName(.HoverImage), AddNewIcon(.HoverImage)
            End With
        Next i
    End If
    If dmbClipboard.ObjSrc = docCommand Then
        caption = caption + NiceCmdCaption(GetIDByName(dmbClipboard.CmdContents.Name)) + "' Command"
        For i = 1 To UBound(MenuCmds)
            Add2LV lvFont, GetFontStr(MenuCmds(i).NormalFont)
            Add2LV lvFont, GetFontStr(MenuCmds(i).HoverFont)
            With MenuCmds(i).LeftImage
                If LenB(.NormalImage) <> 0 Then Add2LV lvImages, GetFileName(.NormalImage), AddNewIcon(.NormalImage)
                If LenB(.HoverImage) <> 0 Then Add2LV lvImages, GetFileName(.HoverImage), AddNewIcon(.HoverImage)
            End With
            With MenuCmds(i).RightImage
                If LenB(.NormalImage) <> 0 Then Add2LV lvImages, GetFileName(.NormalImage), AddNewIcon(.NormalImage)
                If LenB(.HoverImage) <> 0 Then Add2LV lvImages, GetFileName(.HoverImage), AddNewIcon(.HoverImage)
            End With
            With MenuCmds(i).BackImage
                If LenB(.NormalImage) <> 0 Then Add2LV lvImages, GetFileName(.NormalImage), AddNewIcon(.NormalImage)
                If LenB(.HoverImage) <> 0 Then Add2LV lvImages, GetFileName(.HoverImage), AddNewIcon(.HoverImage)
            End With
        Next i
    End If
    
    CoolListView lvFont
    CoolListView lvImages

End Sub

Private Function GetFontStr(f As tFont) As String

    With f
        GetFontStr = .FontName + ", " & .FontSize & IIf(.FontBold, ", Bold ", vbNullString) + IIf(.FontItalic, ", Italic ", vbNullString) + IIf(.FontUnderline, ", Underline ", vbNullString)
    End With

End Function

Private Function AddNewIcon(ByVal fn As String) As Integer

    With frmMain.picItemIcon
        .Picture = LoadPictureRes(fn)
        Set .Picture = .Image
    End With
    With ilIcons.ListImages.Add(, , frmMain.picItemIcon.Picture)
        AddNewIcon = .Index
    End With

End Function

Private Sub Add2LV(lv As ListView, ByVal s As String, Optional ByVal iconIdx As Integer = -1)

    If lv.FindItem(s, lvwText, , lvwWhole) Is Nothing Then
        With lv.ListItems.Add(, , s)
            If iconIdx <> -1 Then .SmallIcon = iconIdx
        End With
    End If

End Sub

Private Sub UpdateItemsList()

    Dim i As Integer
    Dim nItem As ListItem
    Dim Include As Boolean
    
    lvItems.ListItems.Clear
    
    sbInfo.Panels("pInfo").Text = GetLocalizedStr(499) + " "

    If IsGroup(frmMain.tvMenus.SelectedItem.key) Then
        chkCascade.Enabled = False
        For i = 1 To UBound(MenuGrps)
            If dmbClipboard.GrpContents.Name <> MenuGrps(i).Name Then
                Include = True
                
                If cmbToolbars.ListIndex <> 0 Then
                    Include = (BelongsToToolbar(i, True) = cmbToolbars.ListIndex) And Include
                End If
                
                Select Case chkEnabled.Value
                    Case vbChecked
                        Include = (Not MenuGrps(i).disabled) And Include
                    Case vbUnchecked
                        Include = MenuGrps(i).disabled And Include
                    Case vbGrayed
                End Select
                
                If LenB(txtText.Text) <> 0 Then
                    Include = (InStr(1, MenuGrps(i).caption, txtText.Text, vbTextCompare) > 0) And Include
                End If
                
                If chkFontType.Value = vbChecked Then
                    For Each nItem In lvFont.ListItems
                        If nItem.Checked Then
                            Include = ((GetFontStr(MenuGrps(i).DefNormalFont) = nItem.Text) Or (GetFontStr(MenuGrps(i).DefHoverFont) = nItem.Text)) And Include
                        End If
                    Next nItem
                End If
                
                If chkImage.Value = vbChecked Then
                    For Each nItem In lvImages.ListItems
                        If nItem.Checked Then
                            Include = ((GetFileName(MenuGrps(i).tbiLeftImage.NormalImage) = nItem.Text) Or _
                                        (GetFileName(MenuGrps(i).tbiLeftImage.HoverImage) = nItem.Text) Or _
                                        (GetFileName(MenuGrps(i).tbiRightImage.NormalImage) = nItem.Text) Or _
                                        (GetFileName(MenuGrps(i).tbiRightImage.HoverImage) = nItem.Text) Or _
                                        (GetFileName(MenuGrps(i).tbiBackImage.NormalImage) = nItem.Text) Or _
                                        (GetFileName(MenuGrps(i).tbiBackImage.HoverImage) = nItem.Text)) And _
                                        Include
                        End If
                    Next nItem
                End If
                
                If Include Then
                    Set nItem = lvItems.ListItems.Add(, , NiceGrpCaption(i))
                    nItem.tag = MenuGrps(i).Name
                    nItem.Checked = IsOnList(nItem.tag)
                End If
            End If
        Next i
        sbInfo.Panels("pInfo").Text = sbInfo.Panels("pInfo").Text & lvItems.ListItems.Count & " " + GetLocalizedStr(505) + " " + GetLocalizedStr(500)
    End If
    
    If IsCommand(frmMain.tvMenus.SelectedItem.key) Then
        For i = 1 To UBound(MenuCmds)
            If dmbClipboard.CmdContents.Name <> MenuCmds(i).Name Then
                Include = True
                
                If cmbToolbars.ListIndex <> 0 Then
                    Include = (BelongsToToolbar(i, False) = cmbToolbars.ListIndex) And Include
                End If
                
                Select Case chkEnabled.Value
                    Case vbChecked
                        Include = (Not MenuCmds(i).disabled) And Include
                    Case vbUnchecked
                        Include = MenuCmds(i).disabled And Include
                    Case vbGrayed
                End Select
                
                If chkCascade.Value = vbChecked Then
                    Include = ((MenuCmds(i).Actions.onclick.Type = atcCascade) Or _
                                (MenuCmds(i).Actions.onmouseover.Type = atcCascade) Or _
                                (MenuCmds(i).Actions.OnDoubleClick.Type = atcCascade)) And Include
                End If
                
                If chkFontType.Value = vbChecked Then
                    For Each nItem In lvFont.ListItems
                        If nItem.Checked Then
                            Include = ((GetFontStr(MenuCmds(i).NormalFont) = nItem.Text) Or (GetFontStr(MenuCmds(i).HoverFont) = nItem.Text)) And Include
                        End If
                    Next nItem
                End If
                
                If chkImage.Value = vbChecked Then
                    For Each nItem In lvImages.ListItems
                        If nItem.Checked Then
                            Include = ((GetFileName(MenuCmds(i).LeftImage.NormalImage) = nItem.Text) Or _
                                        (GetFileName(MenuCmds(i).LeftImage.HoverImage) = nItem.Text) Or _
                                        (GetFileName(MenuCmds(i).RightImage.NormalImage) = nItem.Text) Or _
                                        (GetFileName(MenuCmds(i).RightImage.HoverImage) = nItem.Text) Or _
                                        (GetFileName(MenuCmds(i).BackImage.NormalImage) = nItem.Text) Or _
                                        (GetFileName(MenuCmds(i).BackImage.HoverImage) = nItem.Text)) And _
                                        Include
                        End If
                    Next nItem
                End If
                
                If LenB(txtText.Text) <> 0 Then
                    Include = (InStr(1, MenuCmds(i).caption, txtText.Text, vbTextCompare) > 0) And Include
                End If
            End If
            
            If Include Then
                If dmbClipboard.CmdContents.Name <> MenuCmds(i).Name And MenuCmds(i).Name <> "[SEP]" Then
                    Set nItem = lvItems.ListItems.Add(, , NiceGrpCaption(MenuCmds(i).parent) + " / " + NiceCmdCaption(i))
                    nItem.tag = MenuCmds(i).Name
                    nItem.Checked = IsOnList(nItem.tag)
                End If
            End If
        Next i
        sbInfo.Panels("pInfo").Text = sbInfo.Panels("pInfo").Text & lvItems.ListItems.Count & " " + GetLocalizedStr(506) + " " + GetLocalizedStr(500)
    End If
    
    If lvItems.ListItems.Count > 0 Then
        lvItems.ListItems(1).Selected = True
    End If
    
    ApplyCustomSel
    
    CoolListView lvItems

End Sub

Private Function IsOnList(ItemName As String) As Boolean

    Dim i As Integer
    
    IsOnList = False
    
    On Error Resume Next
    i = UBound(dmbClipboard.CustomSel)
    If Err.number Then
        Exit Function
    End If
    
    For i = i To 1 Step -1
        If dmbClipboard.CustomSel(i) = ItemName Then
            IsOnList = True
            Exit For
        End If
    Next i

End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    msgSubClass.SubClassHwnd Me.hwnd, False

End Sub

Private Sub Form_Resize()

    On Error Resume Next
    
    With lvItems
        .Move .Left, .Top, Width - .Left - GetClientLeft(Me.hwnd) - lvFont.Left, Height - GetClientTop(Me.hwnd) - sbInfo.Height - cmdOK.Height - .Top - 8 * 15
        frameFilter.Height = .Height + .Top - 2 * 15
        cmdCancel.Move .Left + .Width - cmdCancel.Width, .Top + .Height + 6 * 15
        cmdOK.Move cmdCancel.Left - cmdOK.Width - 8 * 15, cmdCancel.Top
    End With
    CoolListView lvItems
    
End Sub

Private Sub lvFont_ItemCheck(ByVal item As MSComctlLib.ListItem)

    UpdateItemsList

End Sub

Private Sub lvImages_ItemCheck(ByVal item As MSComctlLib.ListItem)

    UpdateItemsList

End Sub

Private Sub lvItems_ItemCheck(ByVal item As MSComctlLib.ListItem)

    ApplyCustomSel

End Sub

Private Sub UpdateOKButtonState()

    Dim nItem As ListItem

    cmdOK.Enabled = False
    For Each nItem In lvItems.ListItems
        If nItem.Checked Then
            cmdOK.Enabled = True
            Exit For
        End If
    Next nItem

End Sub

Private Sub LocalizeUI()

    frameFilter.caption = GetLocalizedStr(501)

    mnuSelAll.caption = GetLocalizedStr(503)
    mnuSelNone.caption = GetLocalizedStr(504)
    
    chkEnabled.caption = GetLocalizedStr(104)
    chkCascade.caption = GetLocalizedStr(502)
     
    lblBelongTB.caption = GetLocalizedStr(794)

    cmdOK.caption = GetLocalizedStr(186)
    cmdCancel.caption = GetLocalizedStr(187)
    If Preferences.language <> "eng" Then
        cmdOK.Width = SetCtrlWidth(cmdOK)
        cmdCancel.Width = SetCtrlWidth(cmdCancel)
    End If
    
    FixContolsWidth Me

End Sub

Private Sub msgSubClass_NewMessage(ByVal hwnd As Long, uMsg As Long, wParam As Long, lParam As Long, Cancel As Boolean, UseRetVal As Boolean, RetVal As Long)

    Dim MMI As MINMAXINFO

    If hwnd = Me.hwnd Then
        If uMsg = WM_GETMINMAXINFO Then
            CopyMemory MMI, ByVal lParam, LenB(MMI)
            With MMI
                .ptMinTrackSize.x = 516
                .ptMinTrackSize.y = 459
                .ptMaxTrackSize.x = Screen.Width / Screen.TwipsPerPixelX
                .ptMaxTrackSize.y = Screen.Height / Screen.TwipsPerPixelY
            End With
            CopyMemory ByVal lParam, MMI, LenB(MMI)
            Cancel = True
        End If
    End If

End Sub

Private Sub txtText_Change()

    UpdateItemsList

End Sub
