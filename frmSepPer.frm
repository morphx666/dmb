VERSION 5.00
Object = "{E1C6DB9D-BD4A-4A61-A759-0CED75D034BF}#43.0#0"; "SmartButton.ocx"
Object = "{DB06EC30-01E1-485F-A3C7-CE80CA0D7D37}#2.0#0"; "xFXSlider.ocx"
Object = "{9A2947B4-87EE-40EE-A3EF-32BDC32D5726}#1.0#0"; "xfxline3d.ocx"
Begin VB.Form frmSepPer 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Separator Length"
   ClientHeight    =   2865
   ClientLeft      =   5955
   ClientTop       =   6765
   ClientWidth     =   3780
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSepPer.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2865
   ScaleWidth      =   3780
   ShowInTaskbar   =   0   'False
   Begin SmartButtonProject.SmartButton sbApplyOptions 
      Height          =   345
      Left            =   135
      TabIndex        =   7
      Top             =   120
      Width           =   3525
      _ExtentX        =   6218
      _ExtentY        =   609
      Caption         =   "Options"
      Picture         =   "frmSepPer.frx":058A
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
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2775
      TabIndex        =   6
      Top             =   2370
      Width           =   900
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1755
      TabIndex        =   5
      Top             =   2370
      Width           =   900
   End
   Begin VB.Frame frameNormal 
      Caption         =   "Separator Length"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   555
      Width           =   3555
      Begin xFXSlider.ucSlider sldLen 
         Height          =   270
         Left            =   180
         TabIndex        =   2
         Top             =   675
         Width           =   2445
         _ExtentX        =   820
         _ExtentY        =   476
         TickStyle       =   0
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomKnobImage =   "frmSepPer.frx":06E4
         CustomSelKnobImage=   "frmSepPer.frx":09BE
      End
      Begin VB.Label lblVal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "100%"
         Height          =   195
         Left            =   2760
         TabIndex        =   3
         Top             =   720
         Width           =   435
      End
      Begin VB.Label lblPerc 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Percentage"
         Height          =   195
         Left            =   180
         TabIndex        =   1
         Top             =   390
         Width           =   825
      End
   End
   Begin xfxLine3D.ucLine3D uc3DLine1 
      Height          =   30
      Left            =   120
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2235
      Width           =   3555
      _ExtentX        =   6271
      _ExtentY        =   53
   End
End
Attribute VB_Name = "frmSepPer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents msgSubClass As xfxSC
Attribute msgSubClass.VB_VarHelpID = -1

Private Sub cmdCancel_Click()

    Unload Me

End Sub

Private Sub cmdOK_Click()

    ApplyStyleOptions
    MenuCmds(GetID).SeparatorPercent = sldLen.Value
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
                Case 1:
                    For i = 1 To UBound(MenuCmds)
                        If MenuCmds(i).parent = t Then CopyStyle sId, i
                    Next i
                Case 2:
                    For i = 1 To UBound(MenuCmds)
                        If BelongsToToolbar(i, False) = t Then CopyStyle sId, i
                    Next i
                Case 3:
                    For i = 1 To UBound(MenuCmds)
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

    With MenuCmds(tID)
        .SeparatorPercent = MenuCmds(sId).SeparatorPercent
    End With

End Sub

Private Sub Form_Load()

    CenterForm Me
    SetupCharset Me
    LocalizeUI
    
    Set msgSubClass = New xfxSC
    msgSubClass.SubClassHwnd Me.hwnd, True
    
    With MenuCmds(GetID)
        sldLen.Value = .SeparatorPercent
        
        caption = NiceGrpCaption(.parent) + "/" + NiceCmdCaption(GetID) + " - " + GetLocalizedStr(791)
    End With
    
    sldLen_Change

End Sub

Private Sub LocalizeUI()

    frameNormal.caption = GetLocalizedStr(791)
    lblPerc.caption = GetLocalizedStr(792)
    
    cmdOK.caption = GetLocalizedStr(186)
    cmdCancel.caption = GetLocalizedStr(187)
    If Preferences.language <> "eng" Then
        cmdOK.Width = SetCtrlWidth(cmdOK)
        cmdCancel.Width = SetCtrlWidth(cmdCancel)
    End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF1 Then showHelp "dialogs/sep_size.htm"

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    msgSubClass.SubClassHwnd Me.hwnd, False

End Sub

Private Sub sbApplyOptions_Click()

    PopupMenu frmMain.mnuStyleOptions, , sbApplyOptions.Left, sbApplyOptions.Top + sbApplyOptions.Height

End Sub

Private Sub sldLen_Change()

    lblVal.caption = sldLen.Value & "%"

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
