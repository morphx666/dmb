VERSION 5.00
Object = "{E1C6DB9D-BD4A-4A61-A759-0CED75D034BF}#43.0#0"; "SmartButton.ocx"
Object = "{9A2947B4-87EE-40EE-A3EF-32BDC32D5726}#1.0#0"; "xfxline3d.ocx"
Begin VB.Form frmColor 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Color"
   ClientHeight    =   4035
   ClientLeft      =   6015
   ClientTop       =   5625
   ClientWidth     =   5475
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmColor.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   5475
   ShowInTaskbar   =   0   'False
   Begin SmartButtonProject.SmartButton sbApplyOptions 
      Height          =   345
      Left            =   135
      TabIndex        =   18
      Top             =   135
      Width           =   5190
      _ExtentX        =   9155
      _ExtentY        =   609
      Caption         =   "Options"
      Picture         =   "frmColor.frx":014A
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
   Begin VB.Frame frameHover 
      Caption         =   "Mouse Over"
      Height          =   1380
      Left            =   2820
      TabIndex        =   8
      Top             =   555
      Width           =   2535
      Begin SmartButtonProject.SmartButton cmdColor 
         Height          =   240
         Index           =   2
         Left            =   1800
         TabIndex        =   10
         Top             =   240
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
         Index           =   3
         Left            =   1800
         TabIndex        =   12
         Top             =   570
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
      Begin SmartButtonProject.SmartButton cmdAuto 
         Height          =   300
         Left            =   1320
         TabIndex        =   14
         Top             =   990
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   529
         Caption         =   "Auto"
         Picture         =   "frmColor.frx":02A4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionLayout   =   4
         PictureLayout   =   3
      End
      Begin xfxLine3D.ucLine3D uc3DLine2 
         Height          =   30
         Left            =   45
         TabIndex        =   13
         Top             =   900
         Width           =   2400
         _ExtentX        =   4233
         _ExtentY        =   53
      End
      Begin VB.Label lblBackColorO 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Back Color"
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
         Left            =   855
         TabIndex        =   11
         Top             =   600
         Width           =   750
      End
      Begin VB.Label lblTextColorO 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Text Color"
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
         Left            =   855
         TabIndex        =   9
         Top             =   270
         Width           =   750
      End
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
      Left            =   3405
      TabIndex        =   16
      Top             =   3540
      Width           =   900
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
      Left            =   4455
      TabIndex        =   17
      Top             =   3540
      Width           =   900
   End
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
      TabIndex        =   15
      Top             =   2100
      Width           =   5235
   End
   Begin VB.Frame frameNormal 
      Caption         =   "Normal"
      Height          =   1380
      Left            =   120
      TabIndex        =   0
      Top             =   555
      Width           =   2535
      Begin xfxLine3D.ucLine3D uc3DLine1 
         Height          =   30
         Left            =   75
         TabIndex        =   5
         Top             =   900
         Width           =   2400
         _ExtentX        =   4233
         _ExtentY        =   53
      End
      Begin SmartButtonProject.SmartButton cmdColor 
         Height          =   240
         Index           =   0
         Left            =   1800
         TabIndex        =   2
         Top             =   240
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
         Left            =   1800
         TabIndex        =   4
         Top             =   570
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
      Begin SmartButtonProject.SmartButton cmdDefault 
         Height          =   300
         Left            =   135
         TabIndex        =   6
         Top             =   990
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   529
         Caption         =   "Reset"
         Picture         =   "frmColor.frx":063E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionLayout   =   4
         PictureLayout   =   3
      End
      Begin SmartButtonProject.SmartButton cmdRevert 
         Height          =   300
         Left            =   1320
         TabIndex        =   7
         Top             =   990
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   529
         Caption         =   "Revert"
         Picture         =   "frmColor.frx":09D8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionLayout   =   4
         PictureLayout   =   3
      End
      Begin VB.Label lblBackColorN 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Back Color"
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
         Left            =   855
         TabIndex        =   3
         Top             =   600
         Width           =   750
      End
      Begin VB.Label lblTextColorN 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Text Color"
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
         Left            =   855
         TabIndex        =   1
         Top             =   270
         Width           =   750
      End
   End
End
Attribute VB_Name = "frmColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim BackCmd As MenuCmd
Dim IsSep As Boolean

Private WithEvents msgSubClass As xfxSC
Attribute msgSubClass.VB_VarHelpID = -1

Private Sub cmdAuto_Click()

    With MenuCmds(GetID)
        SetColor .nTextColor, cmdColor(3)
        SetColor .nBackColor, cmdColor(2)
    End With
    
    UpdateSample

    cmdOK.SetFocus

End Sub

Private Sub cmdCancel_Click()

    MenuCmds(GetID) = BackCmd
    Unload Me

End Sub

Private Sub cmdColor_Click(Index As Integer)

    BuildUsedColorsArray

    With cmdColor(Index)
        SelColor = .tag
        SelColor_CanBeTransparent = (Index = 1 Or Index = 3)
        frmColorPicker.Show vbModal
        SetColor SelColor, cmdColor(Index)
    End With
    
    UpdateSample

End Sub

Private Sub cmdDefault_Click()

    With MenuCmds(GetID)
        SetColor MenuGrps(.parent).nTextColor, cmdColor(0)
        If CreateToolbar Then
            SetColor MenuGrps(.parent).nBackColor, cmdColor(1)
        Else
            SetColor MenuGrps(.parent).bColor, cmdColor(1)
        End If
    End With
    
    UpdateSample
    
    cmdOK.SetFocus

End Sub

Private Sub UpdateSample()

    With MenuCmds(GetID)
        .nTextColor = cmdColor(0).tag
        .nBackColor = cmdColor(1).tag
        .hTextColor = cmdColor(2).tag
        .hBackColor = cmdColor(3).tag
    End With
    
    frmMain.DoLivePreview wbLivePreview

End Sub

Private Sub cmdOK_Click()
    
    ApplyStyleOptions
    frmMain.SaveState "Change " + MenuCmds(GetID).Name + " Color"
    
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
        If IsSep And .Name <> "[SEP]" Then Exit Sub
        If Not IsSep And .Name = "[SEP]" Then Exit Sub
        .hBackColor = MenuCmds(sId).hBackColor
        .hTextColor = MenuCmds(sId).hTextColor
        .nBackColor = MenuCmds(sId).nBackColor
        .nTextColor = MenuCmds(sId).nTextColor
    End With

End Sub

Private Sub cmdRevert_Click()

    MenuCmds(GetID) = BackCmd
    SetDefaults
    
    cmdOK.SetFocus

End Sub

Private Sub SetDefaults()

    With MenuCmds(GetID)
        SetColor .nTextColor, cmdColor(0)
        SetColor .nBackColor, cmdColor(1)
        SetColor .hTextColor, cmdColor(2)
        SetColor .hBackColor, cmdColor(3)
        
        caption = NiceGrpCaption(.parent) + "/" + NiceCmdCaption(GetID) + " - " + GetLocalizedStr(212)
    End With
    
    frameHover.Enabled = Not IsSep
    lblTextColorO.Enabled = Not IsSep
    lblBackColorO.Enabled = Not IsSep
    cmdAuto.Enabled = Not IsSep
    cmdDefault.Enabled = Not IsSep

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF1 Then showHelp "dialogs/command_color.htm"

End Sub

Private Sub Form_Load()

    CenterForm Me

    Set msgSubClass = New xfxSC
    msgSubClass.SubClassHwnd Me.hwnd, True
    
    IsSep = IsSeparator(frmMain.tvMenus.SelectedItem.key)
    
    LocalizeUI
    BackCmd = MenuCmds(GetID)

    SetDefaults
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    msgSubClass.SubClassHwnd Me.hwnd, False

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

Private Sub sbApplyOptions_Click()

    PopupMenu frmMain.mnuStyleOptions, , sbApplyOptions.Left, sbApplyOptions.Top + sbApplyOptions.Height

End Sub

Private Sub LocalizeUI()

    frameNormal.caption = GetLocalizedStr(179)
    frameHover.caption = GetLocalizedStr(180)

    If IsSep Then
        lblTextColorN.caption = GetLocalizedStr(669)
    Else
        lblTextColorN.caption = GetLocalizedStr(181)
    End If
    lblBackColorN.caption = GetLocalizedStr(182)
    lblTextColorO.caption = GetLocalizedStr(181)
    lblBackColorO.caption = GetLocalizedStr(182)
    
    cmdDefault.caption = GetLocalizedStr(183)
    cmdRevert.caption = GetLocalizedStr(184)
    cmdAuto.caption = GetLocalizedStr(185)
    
    frmLiveSample.caption = GetLocalizedStr(188)
    
    cmdOK.caption = GetLocalizedStr(186)
    cmdCancel.caption = GetLocalizedStr(187)
    
    If Preferences.language <> "eng" Then
        cmdOK.Width = SetCtrlWidth(cmdOK)
        cmdCancel.Width = SetCtrlWidth(cmdCancel)
    End If

End Sub
