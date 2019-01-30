VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{E1C6DB9D-BD4A-4A61-A759-0CED75D034BF}#43.0#0"; "SmartButton.ocx"
Begin VB.Form frmGrpMargins 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Margins"
   ClientHeight    =   3615
   ClientLeft      =   11580
   ClientTop       =   6435
   ClientWidth     =   3810
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmGrpMargins.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   3810
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmLiveSample 
      Caption         =   "Live Sample"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1335
      Left            =   120
      TabIndex        =   7
      Top             =   1665
      Width           =   3600
   End
   Begin SmartButtonProject.SmartButton sbApplyOptions 
      Height          =   345
      Left            =   135
      TabIndex        =   6
      Top             =   120
      Width           =   3570
      _ExtentX        =   6297
      _ExtentY        =   609
      Caption         =   "Options"
      Picture         =   "frmGrpMargins.frx":038A
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
   Begin VB.Frame frameGroupMargins 
      Caption         =   "Group Contents' Margins"
      Height          =   1050
      Left            =   120
      TabIndex        =   0
      Top             =   555
      Width           =   3600
      Begin VB.TextBox txtMarginV 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2205
         TabIndex        =   9
         Text            =   "123"
         Top             =   525
         Width           =   420
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   285
         Left            =   1140
         TabIndex        =   3
         Top             =   510
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         BuddyControl    =   "txtMarginH"
         BuddyDispid     =   196612
         OrigLeft        =   1260
         OrigTop         =   600
         OrigRight       =   1455
         OrigBottom      =   900
         Max             =   20
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtMarginH 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   720
         TabIndex        =   2
         Text            =   "123"
         Top             =   510
         Width           =   420
      End
      Begin MSComCtl2.UpDown UpDown2 
         Height          =   285
         Left            =   2640
         TabIndex        =   8
         Top             =   525
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         BuddyControl    =   "txtMarginV"
         BuddyDispid     =   196611
         OrigLeft        =   2520
         OrigTop         =   630
         OrigRight       =   2715
         OrigBottom      =   885
         Max             =   20
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.Label lblV 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vertical"
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
         Left            =   1905
         TabIndex        =   10
         Top             =   300
         Width           =   525
      End
      Begin VB.Image Image2 
         Height          =   240
         Left            =   1905
         Picture         =   "frmGrpMargins.frx":04E4
         Top             =   540
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   420
         Picture         =   "frmGrpMargins.frx":086E
         Top             =   525
         Width           =   240
      End
      Begin VB.Label lblH 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Horizontal"
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
         Left            =   420
         TabIndex        =   1
         Top             =   300
         Width           =   720
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2835
      TabIndex        =   5
      Top             =   3150
      Width           =   900
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1845
      TabIndex        =   4
      Top             =   3150
      Width           =   900
   End
End
Attribute VB_Name = "frmGrpMargins"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim BackGrp As MenuGrp
Dim IsUpdating As Boolean

Private WithEvents msgSubClass As xfxSC
Attribute msgSubClass.VB_VarHelpID = -1

Private Sub cmdCancel_Click()

    MenuGrps(GetID) = BackGrp
    Unload Me

End Sub

Private Sub cmdOK_Click()

    With MenuGrps(GetID)
        .ContentsMarginH = txtMarginH.Text
        .ContentsMarginV = txtMarginV.Text
    End With
    
    Project.HasChanged = True
    
    ApplyStyleOptions

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
        .ContentsMarginH = MenuGrps(sId).ContentsMarginH
        .ContentsMarginV = MenuGrps(sId).ContentsMarginV
    End With

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF1 Then showHelp "dialogs/group_margins.htm"

End Sub

Private Sub Form_Load()

    CenterForm Me
    LocalizeUI
    
    Set msgSubClass = New xfxSC
    msgSubClass.SubClassHwnd Me.hwnd, True
    
    BackGrp = MenuGrps(GetID)
    
    IsUpdating = True
    With MenuGrps(GetID)
        txtMarginH.Text = .ContentsMarginH
        txtMarginV.Text = .ContentsMarginV
        
        caption = NiceGrpCaption(GetID) + " - " + GetLocalizedStr(216)
    End With
    IsUpdating = False

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    msgSubClass.SubClassHwnd Me.hwnd, False

End Sub

Private Sub sbApplyOptions_Click()

    PopupMenu frmMain.mnuStyleOptions, , sbApplyOptions.Left, sbApplyOptions.Top + sbApplyOptions.Height

End Sub

Private Sub txtMarginH_Change()

    UpdateSample

End Sub

Private Sub txtMarginH_GotFocus()

    SelAll txtMarginH

End Sub

Private Sub txtMarginH_KeyPress(KeyAscii As Integer)

    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If

End Sub

Private Sub txtMarginV_Change()

    UpdateSample

End Sub

Private Sub txtMarginV_GotFocus()

    SelAll txtMarginV

End Sub

Private Sub txtMarginV_KeyPress(KeyAscii As Integer)

    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If

End Sub

Private Sub UpdateSample()

    If IsUpdating Then Exit Sub

    With MenuGrps(GetID)
        .ContentsMarginH = Val(txtMarginH.Text)
        .ContentsMarginV = Val(txtMarginV.Text)
        
        caption = NiceGrpCaption(GetID) + " - " + GetLocalizedStr(216)
    End With

    frmMain.DoLivePreview wbLivePreview, True

End Sub

Private Sub msgSubClass_NewMessage(ByVal hwnd As Long, uMsg As Long, wParam As Long, lParam As Long, Cancel As Boolean, UseRetVal As Boolean, RetVal As Long)

    Select Case hwnd
        Case Me.hwnd
            Select Case uMsg
                Case WM_CTLCOLORSTATIC, WM_CTLCOLOREDIT, WM_PAINT
                    DrawColorBoxes Me
            End Select
    End Select
    
End Sub

Private Sub LocalizeUI()

    frameGroupMargins.caption = GetLocalizedStr(209)
    
    lblH.caption = GetLocalizedStr(211)
    lblV.caption = GetLocalizedStr(210)
    
    cmdOK.caption = GetLocalizedStr(186)
    cmdCancel.caption = GetLocalizedStr(187)
    
    If Preferences.language <> "eng" Then
        cmdOK.Width = SetCtrlWidth(cmdOK)
        cmdCancel.Width = SetCtrlWidth(cmdCancel)
    End If

End Sub
