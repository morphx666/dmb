VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmWindowOpenParams 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New Window Parameters"
   ClientHeight    =   4725
   ClientLeft      =   6555
   ClientTop       =   5205
   ClientWidth     =   5190
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmWindowOpenParams.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   5190
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   3210
      TabIndex        =   13
      Top             =   4320
      Width           =   885
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   345
      Left            =   4245
      TabIndex        =   14
      Top             =   4320
      Width           =   885
   End
   Begin VB.Frame framePosSize 
      Caption         =   "Position and Size"
      Height          =   1260
      Left            =   45
      TabIndex        =   2
      Top             =   750
      Width           =   3180
      Begin VB.CheckBox chkCenter 
         Caption         =   "Center on screen"
         Height          =   195
         Left            =   495
         TabIndex        =   15
         Top             =   945
         Width           =   2250
      End
      Begin VB.TextBox txtHeight 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2220
         TabIndex        =   10
         Text            =   "000"
         Top             =   600
         WhatsThisHelpID =   20330
         Width           =   540
      End
      Begin VB.TextBox txtWidth 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2220
         TabIndex        =   6
         Text            =   "000"
         Top             =   255
         WhatsThisHelpID =   20330
         Width           =   540
      End
      Begin VB.TextBox txtTop 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   870
         TabIndex        =   8
         Text            =   "000"
         Top             =   600
         WhatsThisHelpID =   20330
         Width           =   540
      End
      Begin VB.TextBox txtLeft 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   870
         TabIndex        =   4
         Text            =   "000"
         Top             =   255
         WhatsThisHelpID =   20330
         Width           =   540
      End
      Begin VB.Label lblHeight 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Height"
         Height          =   195
         Left            =   1650
         TabIndex        =   9
         Top             =   645
         Width           =   465
      End
      Begin VB.Label lblWidth 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Width"
         Height          =   195
         Left            =   1695
         TabIndex        =   5
         Top             =   300
         Width           =   420
      End
      Begin VB.Label lblTop 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Top"
         Height          =   195
         Left            =   495
         TabIndex        =   7
         Top             =   645
         Width           =   270
      End
      Begin VB.Label lblLeft 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Left"
         Height          =   195
         Left            =   480
         TabIndex        =   3
         Top             =   300
         Width           =   285
      End
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   45
      TabIndex        =   1
      Top             =   330
      Width           =   1740
   End
   Begin MSComctlLib.ListView lvOptions 
      Height          =   1860
      Left            =   45
      TabIndex        =   12
      Top             =   2325
      Width           =   5100
      _ExtentX        =   8996
      _ExtentY        =   3281
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
         Key             =   "chOptionName"
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "chDescription"
         Text            =   "Description"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label lblOptions 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Options"
      Height          =   195
      Left            =   45
      TabIndex        =   11
      Top             =   2085
      Width           =   555
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      Height          =   195
      Left            =   45
      TabIndex        =   0
      Top             =   105
      Width           =   405
   End
End
Attribute VB_Name = "frmWindowOpenParams"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkCenter_Click()

    If chkCenter.Value = vbChecked Then
        txtLeft.Text = "*"
        txtLeft.Enabled = False
        txtTop.Text = "*"
        txtTop.Enabled = False
    Else
        txtLeft.Text = "0"
        txtLeft.Enabled = True
        txtTop.Text = "0"
        txtTop.Enabled = True
    End If

End Sub

Private Sub cmdCancel_Click()

    Unload Me

End Sub

Private Sub cmdOK_Click()

    Dim par As String
    Dim i As Integer
    
    par = txtName.Text & _
          cSep & txtLeft.Text & _
          cSep & txtTop.Text & _
          cSep & txtWidth.Text & _
          cSep & txtHeight.Text
    For i = 1 To lvOptions.ListItems.Count
        par = par & cSep & Abs(lvOptions.ListItems.item(i).Checked)
    Next i
    
    If IsGroup(frmMain.tvMenus.SelectedItem.key) Then
        Select Case frmMain.tsCmdType.SelectedItem.key
            Case "tsOver"
                MenuGrps(GetID).Actions.onmouseover.WindowOpenParams = par
            Case "tsClick"
                MenuGrps(GetID).Actions.onclick.WindowOpenParams = par
            Case "tsDoubleClick"
                MenuGrps(GetID).Actions.OnDoubleClick.WindowOpenParams = par
        End Select
    Else
        Select Case frmMain.tsCmdType.SelectedItem.key
            Case "tsOver"
                MenuCmds(GetID).Actions.onmouseover.WindowOpenParams = par
            Case "tsClick"
                MenuCmds(GetID).Actions.onclick.WindowOpenParams = par
            Case "tsDoubleClick"
                MenuCmds(GetID).Actions.OnDoubleClick.WindowOpenParams = par
        End Select
    End If
    
    Unload Me

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF1 Then showHelp "dialogs/group_and_command_nwparams.htm"

End Sub

Private Sub Form_Load()

    Dim par As String
    Dim i As Integer
    
    CenterForm Me
    LocalizeUI
    
    If IsGroup(frmMain.tvMenus.SelectedItem.key) Then
        caption = NiceGrpCaption(GetID) + " - " + GetLocalizedStr(475)
    Else
        caption = NiceGrpCaption(MenuCmds(GetID).parent) + "/" + NiceCmdCaption(GetID) + " - " + GetLocalizedStr(475)
    End If
    
    If IsGroup(frmMain.tvMenus.SelectedItem.key) Then
        Select Case frmMain.tsCmdType.SelectedItem.key
            Case "tsOver"
                par = MenuGrps(GetID).Actions.onmouseover.WindowOpenParams
            Case "tsClick"
                par = MenuGrps(GetID).Actions.onclick.WindowOpenParams
            Case "tsDoubleClick"
                par = MenuGrps(GetID).Actions.OnDoubleClick.WindowOpenParams
        End Select
    Else
        Select Case frmMain.tsCmdType.SelectedItem.key
            Case "tsOver"
                par = MenuCmds(GetID).Actions.onmouseover.WindowOpenParams
            Case "tsClick"
                par = MenuCmds(GetID).Actions.onclick.WindowOpenParams
            Case "tsDoubleClick"
                par = MenuCmds(GetID).Actions.OnDoubleClick.WindowOpenParams
        End Select
    End If
    
    If LenB(par) = 0 Then par = nwdPar
    
    If LenB(par) <> 0 Then
        txtName.Text = GetParam(par, 1)
        
        chkCenter.Value = IIf(GetParam(par, 2) = "*", vbChecked, vbUnchecked)
        
        txtLeft.Text = GetParam(par, 2)
        txtTop.Text = GetParam(par, 3)
        txtWidth.Text = GetParam(par, 4)
        txtHeight.Text = GetParam(par, 5)
        
        With lvOptions.ListItems
            For i = 1 To .Count
                .item(i).Checked = -Val(GetParam(par, i + 5))
            Next i
        End With
    End If
    
    FixCtrls4Skin Me
    
End Sub

Private Sub txtHeight_GotFocus()

    SelAll txtHeight

End Sub

Private Sub txtHeight_KeyPress(KeyAscii As Integer)

    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 37 Then
        KeyAscii = 0
    End If

End Sub

Private Sub txtLeft_GotFocus()

    SelAll txtLeft

End Sub

Private Sub txtLeft_KeyPress(KeyAscii As Integer)

    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 37 Then
        KeyAscii = 0
    End If

End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)

    If (KeyAscii < 97 Or KeyAscii > 122) And _
       (KeyAscii < 65 Or KeyAscii > 90) And _
       KeyAscii <> 8 And KeyAscii <> 95 Then
        KeyAscii = 0
    End If

End Sub

Private Sub txtTop_GotFocus()

    SelAll txtTop

End Sub

Private Sub txtTop_KeyPress(KeyAscii As Integer)

    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 37 Then
        KeyAscii = 0
    End If

End Sub

Private Sub txtWidth_GotFocus()

    SelAll txtWidth

End Sub

Private Sub txtWidth_KeyPress(KeyAscii As Integer)

    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 37 Then
        KeyAscii = 0
    End If

End Sub

Private Sub LocalizeUI()

    Dim nItem As ListItem
    Dim i As Integer

    framePosSize.caption = GetLocalizedStr(476)
    lblOptions.caption = GetLocalizedStr(337)
    
    lblName.caption = GetLocalizedStr(409)
    lblLeft.caption = GetLocalizedStr(190)
    lblTop.caption = GetLocalizedStr(497)
    lblWidth.caption = GetLocalizedStr(428)
    lblHeight.caption = GetLocalizedStr(429)
    
    For i = 477 To 495 Step 2
        Set nItem = lvOptions.ListItems.Add(, , GetLocalizedStr(i))
        'nItem.ToolTipText = GetLocalizedStr(i + 1)
        nItem.SubItems(1) = GetLocalizedStr(i + 1)
    Next i
    
    CoolListView lvOptions
    
    cmdOK.caption = GetLocalizedStr(186)
    cmdCancel.caption = GetLocalizedStr(187)
    
    If Preferences.language <> "eng" Then
        cmdOK.Width = SetCtrlWidth(cmdOK)
        cmdCancel.Width = SetCtrlWidth(cmdCancel)
    End If
    
    FixContolsWidth Me

End Sub
