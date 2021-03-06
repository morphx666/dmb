VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{E1C6DB9D-BD4A-4A61-A759-0CED75D034BF}#43.0#0"; "SmartButton.ocx"
Begin VB.Form frmGrpSelectionFX 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   6120
   ClientLeft      =   11580
   ClientTop       =   5220
   ClientWidth     =   5475
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmGrpSelectionFX.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   5475
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frameSelFX 
      Caption         =   "Selection Effects"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1890
      Left            =   120
      TabIndex        =   11
      Top             =   585
      Width           =   5235
      Begin VB.TextBox txtRadiusTR 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4515
         TabIndex        =   24
         Text            =   "123"
         Top             =   765
         Width           =   420
      End
      Begin VB.TextBox txtRadiusTL 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3840
         TabIndex        =   23
         Text            =   "123"
         Top             =   765
         Width           =   420
      End
      Begin VB.TextBox txtRadiusBR 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4515
         TabIndex        =   22
         Text            =   "123"
         Top             =   1380
         Width           =   420
      End
      Begin VB.TextBox txtRadiusBL 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3840
         TabIndex        =   21
         Text            =   "123"
         Top             =   1380
         Width           =   420
      End
      Begin VB.ComboBox cmbFXO 
         Height          =   315
         ItemData        =   "frmGrpSelectionFX.frx":058A
         Left            =   225
         List            =   "frmGrpSelectionFX.frx":0597
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   1365
         Width           =   1065
      End
      Begin VB.ComboBox cmbFXN 
         Height          =   315
         ItemData        =   "frmGrpSelectionFX.frx":05B3
         Left            =   225
         List            =   "frmGrpSelectionFX.frx":05C0
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   585
         Width           =   1065
      End
      Begin VB.TextBox txtBorderSize 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2265
         TabIndex        =   12
         Text            =   "123"
         Top             =   600
         Width           =   420
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   285
         Left            =   2685
         TabIndex        =   15
         Top             =   600
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   393216
         BuddyControl    =   "txtBorderSize"
         BuddyDispid     =   196616
         OrigLeft        =   1305
         OrigTop         =   1320
         OrigRight       =   1500
         OrigBottom      =   1575
         Max             =   20
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin SmartButtonProject.SmartButton cmdColor 
         Height          =   240
         Index           =   0
         Left            =   1350
         TabIndex        =   16
         Top             =   615
         Width           =   270
         _ExtentX        =   476
         _ExtentY        =   423
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowFocus       =   -1  'True
      End
      Begin SmartButtonProject.SmartButton cmdColor 
         Height          =   240
         Index           =   1
         Left            =   1350
         TabIndex        =   17
         Top             =   1395
         Width           =   270
         _ExtentX        =   476
         _ExtentY        =   423
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowFocus       =   -1  'True
      End
      Begin VB.Label lblBorderRadius 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Radius"
         Height          =   195
         Left            =   3180
         TabIndex        =   25
         Top             =   1080
         Width           =   480
      End
      Begin VB.Label lblOver 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mouse Over"
         Height          =   195
         Left            =   225
         TabIndex        =   20
         Top             =   1155
         Width           =   870
      End
      Begin VB.Label lblNormal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Normal"
         Height          =   195
         Left            =   225
         TabIndex        =   19
         Top             =   375
         Width           =   495
      End
      Begin VB.Label lblBorderSize 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Border Size"
         Height          =   195
         Left            =   1965
         TabIndex        =   18
         Top             =   375
         Width           =   810
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   1965
         Picture         =   "frmGrpSelectionFX.frx":05DC
         Top             =   615
         Width           =   240
      End
   End
   Begin VB.Frame frameMargins 
      Caption         =   "Margins"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   4
      Top             =   2655
      Width           =   5235
      Begin VB.TextBox txtMarginY 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2265
         TabIndex        =   6
         Text            =   "123"
         Top             =   555
         Width           =   420
      End
      Begin VB.TextBox txtMarginX 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   525
         TabIndex        =   5
         Text            =   "123"
         Top             =   555
         Width           =   420
      End
      Begin MSComCtl2.UpDown UpDown3 
         Height          =   285
         Left            =   2685
         TabIndex        =   7
         Top             =   555
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   393216
         BuddyControl    =   "txtMarginY"
         BuddyDispid     =   196623
         OrigLeft        =   3180
         OrigTop         =   2190
         OrigRight       =   3375
         OrigBottom      =   2475
         Max             =   20
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown UpDown2 
         Height          =   285
         Left            =   945
         TabIndex        =   8
         Top             =   555
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   393216
         BuddyControl    =   "txtMarginX"
         BuddyDispid     =   196624
         OrigLeft        =   870
         OrigTop         =   2190
         OrigRight       =   1065
         OrigBottom      =   2475
         Max             =   20
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.Label lblVM 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vertical Margin"
         Height          =   195
         Left            =   1980
         TabIndex        =   10
         Top             =   330
         Width           =   1050
      End
      Begin VB.Image Image3 
         Height          =   240
         Left            =   1980
         Picture         =   "frmGrpSelectionFX.frx":0966
         Top             =   570
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   240
         Left            =   225
         Picture         =   "frmGrpSelectionFX.frx":0CF0
         Top             =   570
         Width           =   240
      End
      Begin VB.Label lblHM 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Horizontal Margin"
         Height          =   195
         Left            =   225
         TabIndex        =   9
         Top             =   330
         Width           =   1500
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4455
      TabIndex        =   3
      Top             =   5640
      Width           =   900
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3405
      TabIndex        =   2
      Top             =   5640
      Width           =   900
   End
   Begin VB.Frame frmLiveSample 
      Caption         =   "Live Sample"
      ForeColor       =   &H00000000&
      Height          =   1335
      Left            =   120
      TabIndex        =   1
      Top             =   4155
      Width           =   5235
   End
   Begin SmartButtonProject.SmartButton sbApplyOptions 
      Height          =   345
      Left            =   135
      TabIndex        =   0
      Top             =   105
      Width           =   5220
      _ExtentX        =   9208
      _ExtentY        =   609
      Caption         =   "Options"
      Picture         =   "frmGrpSelectionFX.frx":107A
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
   End
End
Attribute VB_Name = "frmGrpSelectionFX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim BackGrp As MenuGrp
Dim SelId As Integer
Dim IsUpdating As Boolean

Private WithEvents msgSubClass As xfxSC
Attribute msgSubClass.VB_VarHelpID = -1

Private Sub cmdCancel_Click()

    MenuGrps(SelId) = BackGrp
    Unload Me

End Sub

Private Sub cmdOK_Click()

    ApplyStyleOptions
    frmMain.SaveState "Change " + MenuGrps(SelId).Name + " " + GetLocalizedStr(984)
    
    Unload Me

End Sub

Private Sub ApplyStyleOptions()

    Dim i As Integer
    Dim c As Integer
    Dim t As Integer
    Dim sId As Integer
    
    sId = GetID
    
    For c = 0 To frmMain.mnuStyleOptionsOP.Count - 1
        If frmMain.mnuStyleOptionsOP(c).Checked Then
            t = Val(frmMain.mnuStyleOptionsOP(c).tag)
            Select Case c
                Case 0: ' do nothing
                Case 2:
                    For i = 1 To UBound(MenuGrps)
                        If BelongsToToolbar(i, True) = t Then CopyStyle sId, i
                    Next i
                Case 3:
                    For i = 1 To UBound(MenuGrps)
                        CopyStyle sId, i
                    Next i
            End Select
            Exit Sub
        End If
    Next c
    
    With dmbClipboard
        For i = 1 To UBound(.CustomSel)
            CopyStyle sId, GetIDByName(.CustomSel(i))
        Next i
    End With

End Sub

Private Sub CopyStyle(sId As Integer, tID As Integer)

    With MenuGrps(tID)
        .CmdsFXNormal = MenuGrps(sId).CmdsFXNormal
        .CmdsFXOver = MenuGrps(sId).CmdsFXOver
        .CmdsFXSize = MenuGrps(sId).CmdsFXSize
        .CmdsMarginX = MenuGrps(sId).CmdsMarginX
        .CmdsMarginY = MenuGrps(sId).CmdsMarginY
        .CmdsFXnColor = MenuGrps(sId).CmdsFXnColor
        .CmdsFXhColor = MenuGrps(sId).CmdsFXhColor
        .tbiRadius.topLeft = MenuGrps(sId).tbiRadius.topLeft
        .tbiRadius.topRight = MenuGrps(sId).tbiRadius.topRight
        .tbiRadius.bottomLeft = MenuGrps(sId).tbiRadius.bottomLeft
        .tbiRadius.bottomRight = MenuGrps(sId).tbiRadius.bottomRight
    End With

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF1 Then showHelp "dialogs/group_selfx.htm"

End Sub

Private Sub Form_Load()

    CenterForm Me
    LocalizeUI
    
    Set msgSubClass = New xfxSC
    msgSubClass.SubClassHwnd Me.hwnd, True
    
    SelId = GetID
    BackGrp = MenuGrps(SelId)
    
    IsUpdating = True
    
    With MenuGrps(SelId)
        cmbFXN.ListIndex = .CmdsFXNormal
        cmbFXO.ListIndex = .CmdsFXOver
        txtBorderSize.Text = .CmdsFXSize
        txtMarginX.Text = .CmdsMarginX
        txtMarginY.Text = .CmdsMarginY
        SetColor .CmdsFXnColor, cmdColor(0)
        SetColor .CmdsFXhColor, cmdColor(1)
        
        txtRadiusTL.Text = CStr(.tbiRadius.topLeft)
        txtRadiusTR.Text = CStr(.tbiRadius.topRight)
        txtRadiusBL.Text = CStr(.tbiRadius.bottomLeft)
        txtRadiusBR.Text = CStr(.tbiRadius.bottomRight)
        
        caption = NiceGrpCaption(SelId) + " - " + GetLocalizedStr(984)
    End With
    
    IsUpdating = False
    
    UpdateSample True

End Sub

Private Sub sbApplyOptions_Click()

    PopupMenu frmMain.mnuStyleOptions, , sbApplyOptions.Left, sbApplyOptions.Top + sbApplyOptions.Height

End Sub

Private Sub txtBorderSize_GotFocus()

    SelAll txtBorderSize

End Sub

Private Sub txtBorderSize_KeyPress(KeyAscii As Integer)

    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If

End Sub

Private Sub txtMarginX_Change()

    UpdateSample

End Sub

Private Sub txtMarginX_GotFocus()

    SelAll txtMarginX

End Sub

Private Sub cmdColor_Click(Index As Integer)

    BuildUsedColorsArray
    
    With cmdColor(Index)
        SelColor = .tag
        SelColor_CanBeTransparent = True
        frmColorPicker.Show vbModal
        SetColor SelColor, cmdColor(Index)
    End With
    
    UpdateSample

End Sub

Private Sub txtMarginX_KeyPress(KeyAscii As Integer)

    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If

End Sub

Private Sub txtMarginY_Change()

    UpdateSample

End Sub

Private Sub txtMarginY_GotFocus()

    SelAll txtMarginY

End Sub

Private Sub txtMarginY_KeyPress(KeyAscii As Integer)

    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If

End Sub

Private Sub cmbFXN_Click()

    UpdateSample

End Sub

Private Sub cmbFXO_Click()

    UpdateSample

End Sub

Private Sub txtBorderSize_Change()

    UpdateSample

End Sub

Private Sub UpdateSample(Optional IsLoading As Boolean)

    On Error Resume Next
    
    If IsUpdating Then Exit Sub

    With MenuGrps(SelId)
        .CmdsFXNormal = cmbFXN.ListIndex
        .CmdsFXOver = cmbFXO.ListIndex
        .CmdsFXSize = Abs(Val(txtBorderSize.Text))
        .CmdsMarginX = Abs(Val(txtMarginX.Text))
        .CmdsMarginY = Abs(Val(txtMarginY.Text))
        .CmdsFXnColor = cmdColor(0).tag
        .CmdsFXhColor = cmdColor(1).tag
        
        .tbiRadius.topLeft = Val(txtRadiusTL.Text)
        .tbiRadius.topRight = Val(txtRadiusTR.Text)
        .tbiRadius.bottomLeft = Val(txtRadiusBL.Text)
        .tbiRadius.bottomRight = Val(txtRadiusBR.Text)
        
        cmdColor(0).Visible = True
        cmdColor(1).Visible = True
    End With
    
'    If CreateToolbar Then
'        ucDMBSC.DefItems SelId, MenuCmds, MenuGrps, True
'    Else
'        If frmMain.tvMenus.SelectedItem.Child Is Nothing Then
'            ucDMBSC.DefItems 1, MenuCmds, MenuGrps
'        Else
'            ucDMBSC.DefItems GetID(frmMain.tvMenus.SelectedItem.Child), MenuCmds, MenuGrps
'        End If
'    End If
    If Not IsLoading Then frmMain.DoLivePreview wbLivePreview

End Sub

Private Sub LocalizeUI()

    frameSelFX.caption = GetLocalizedStr(984)
    lblNormal.caption = GetLocalizedStr(179)
    lblOver.caption = GetLocalizedStr(180)
    lblBorderSize.caption = GetLocalizedStr(206)
    
    frameMargins.caption = GetLocalizedStr(216)
    lblHM.caption = GetLocalizedStr(219)
    lblVM.caption = GetLocalizedStr(220)
    
    PopulateBorderStyleCombo cmbFXN
    PopulateBorderStyleCombo cmbFXO
    
    frmLiveSample.caption = GetLocalizedStr(188)
    
    cmdOK.caption = GetLocalizedStr(186)
    cmdCancel.caption = GetLocalizedStr(187)
    
    FixContolsWidth Me
        
    If Preferences.language <> "eng" Then
        cmdOK.Width = SetCtrlWidth(cmdOK)
        cmdCancel.Width = SetCtrlWidth(cmdCancel)
    End If

End Sub

Private Sub msgSubClass_NewMessage(ByVal hwnd As Long, uMsg As Long, wParam As Long, lParam As Long, Cancel As Boolean, UseRetVal As Boolean, RetVal As Long)

    Select Case hwnd
        Case Me.hwnd
            Select Case uMsg
                Case WM_CTLCOLORSTATIC, WM_CTLCOLOREDIT, WM_PAINT, 78
                    DrawColorBoxes Me
            End Select
    End Select
    
End Sub

Private Sub txtRadiusBL_Change()

    UpdateSample

End Sub

Private Sub txtRadiusBL_GotFocus()

    SelAll txtRadiusBL

End Sub

Private Sub txtRadiusBL_KeyPress(KeyAscii As Integer)

    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If

End Sub

Private Sub txtRadiusBR_Change()

    UpdateSample

End Sub

Private Sub txtRadiusBR_GotFocus()

    SelAll txtRadiusBR

End Sub

Private Sub txtRadiusBR_KeyPress(KeyAscii As Integer)

    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If

End Sub

Private Sub txtRadiusTL_Change()

    UpdateSample

End Sub

Private Sub txtRadiusTL_GotFocus()

    SelAll txtRadiusTL

End Sub

Private Sub txtRadiusTL_KeyPress(KeyAscii As Integer)

    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If

End Sub

Private Sub txtRadiusTR_Change()

    UpdateSample

End Sub

Private Sub txtRadiusTR_GotFocus()

    SelAll txtRadiusTR

End Sub

Private Sub txtRadiusTR_KeyPress(KeyAscii As Integer)

    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If

End Sub
