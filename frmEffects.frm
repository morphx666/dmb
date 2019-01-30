VERSION 5.00
Object = "{E1C6DB9D-BD4A-4A61-A759-0CED75D034BF}#43.0#0"; "SmartButton.ocx"
Object = "{9A2947B4-87EE-40EE-A3EF-32BDC32D5726}#1.0#0"; "xfxline3d.ocx"
Object = "{DB06EC30-01E1-485F-A3C7-CE80CA0D7D37}#2.0#0"; "xFXSlider.ocx"
Begin VB.Form frmStyleEffects 
   Caption         =   "Effects"
   ClientHeight    =   5190
   ClientLeft      =   3945
   ClientTop       =   6405
   ClientWidth     =   7740
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5190
   ScaleWidth      =   7740
   Begin VB.CheckBox chkContextMenu 
      Caption         =   "Context Menu"
      Height          =   225
      Left            =   345
      TabIndex        =   1
      Top             =   2205
      Width           =   2625
   End
   Begin xFXSlider.ucSlider sldTransparency 
      Height          =   270
      Left            =   345
      TabIndex        =   0
      Top             =   1245
      Width           =   4200
      _ExtentX        =   820
      _ExtentY        =   476
      Value           =   0
      TickStyle       =   0
      SmallChange     =   1
      LargeChange     =   1
      HighlightColor  =   4210752
      HighlightColorEnd=   -2147483633
      HighlightPaintMode=   1
      BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomKnobImage =   "frmEffects.frx":0000
      CustomSelKnobImage=   "frmEffects.frx":02DA
   End
   Begin xfxLine3D.ucLine3D uc3DLineDiv 
      Height          =   30
      Left            =   15
      TabIndex        =   2
      Top             =   1965
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   53
   End
   Begin SmartButtonProject.SmartButton cmdShadowColor 
      Height          =   240
      Left            =   4665
      TabIndex        =   3
      Top             =   330
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
   Begin xFXSlider.ucSlider sldDropShadow 
      Height          =   270
      Left            =   345
      TabIndex        =   4
      Top             =   315
      Width           =   4200
      _ExtentX        =   820
      _ExtentY        =   476
      Max             =   10
      Value           =   0
      TickStyle       =   0
      TickFrequency   =   1
      SmallChange     =   1
      LargeChange     =   1
      HighlightColor  =   14737632
      HighlightColorEnd=   4210752
      HighlightPaintMode=   1
      BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomKnobImage =   "frmEffects.frx":05B4
      CustomSelKnobImage=   "frmEffects.frx":088E
   End
   Begin VB.Label lblTInvisible 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Invisible"
      Height          =   195
      Left            =   4170
      TabIndex        =   11
      Top             =   1560
      Width           =   585
   End
   Begin VB.Label lblTOFF 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "OFF"
      Height          =   195
      Left            =   375
      TabIndex        =   10
      Top             =   1560
      Width           =   300
   End
   Begin VB.Label lblDSDarker 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Larger"
      Height          =   195
      Left            =   4200
      TabIndex        =   9
      Top             =   630
      Width           =   465
   End
   Begin VB.Label lblDSOFF 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "OFF"
      Height          =   195
      Left            =   390
      TabIndex        =   8
      Top             =   630
      Width           =   300
   End
   Begin VB.Image Image4 
      Height          =   240
      Left            =   90
      Picture         =   "frmEffects.frx":0B68
      Top             =   1260
      Width           =   240
   End
   Begin VB.Image Image5 
      Height          =   240
      Left            =   90
      Picture         =   "frmEffects.frx":0CB2
      Top             =   330
      Width           =   240
   End
   Begin VB.Label lblTransparency 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Transparency"
      Height          =   195
      Left            =   1950
      TabIndex        =   7
      Top             =   1020
      Width           =   990
   End
   Begin VB.Label lblDropShadow 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Drop Shadow Size"
      Height          =   195
      Left            =   1800
      TabIndex        =   6
      Top             =   90
      Width           =   1290
   End
   Begin VB.Label lblShadowColor 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Color"
      Height          =   195
      Left            =   4620
      TabIndex        =   5
      Top             =   90
      Width           =   375
   End
End
Attribute VB_Name = "frmStyleEffects"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private IsUpdating As Boolean

Private Sub chkContextMenu_Click()

    Dim g As Integer
    
    For g = 1 To UBound(MenuGrps)
        MenuGrps(g).IsContext = False
    Next g
    
    UpdateMenuItem

End Sub

Friend Sub UpdateUI(g As MenuGrp)

    IsUpdating = True

    With g
        sldDropShadow.Value = .DropShadowSize
        SetColor .DropShadowColor, cmdShadowColor
        sldTransparency.Value = .Transparency
        chkContextMenu.Value = IIf(.IsContext, vbChecked, vbUnchecked)
    End With
    
    IsUpdating = False

End Sub

Private Sub cmdShadowColor_Click()

    BuildUsedColorsArray
    
    With cmdShadowColor
        SelColor = .Tag
        SelColor_CanBeTransparent = False
        frmColorPicker.Show vbModal
        SetColor SelColor, cmdShadowColor
    End With
    
    UpdateMenuItem

End Sub

Private Sub Form_Load()

    SetupCharset Me
    LocalizeUI

End Sub

Friend Sub Form_Resize()

    On Error GoTo ExitSub
    
    sldDropShadow.Width = Width - sldDropShadow.Left * 2 - cmdShadowColor.Width
    cmdShadowColor.Left = sldDropShadow.Left + sldDropShadow.Width + 5 * 15
    lblShadowColor.Left = cmdShadowColor.Left + (cmdShadowColor.Width - lblShadowColor.Width) / 2
    lblDropShadow.Left = sldDropShadow.Left + (sldDropShadow.Width - lblDropShadow.Width) / 2
    lblDSDarker.Left = sldDropShadow.Left + sldDropShadow.Width - lblDSDarker.Width / 2 - 4 * 15
    
    sldTransparency.Width = sldDropShadow.Width
    lblTransparency.Left = sldTransparency.Left + (sldTransparency.Width - lblTransparency.Width) / 2
    lblTInvisible.Left = sldTransparency.Left + sldTransparency.Width - lblTInvisible.Width / 2 - 4 * 15
    
    uc3DLineDiv.Left = 15
    uc3DLineDiv.Width = Width - 30
    
ExitSub:

End Sub

Private Sub sldDropShadow_Change()

    UpdateMenuItem

End Sub

Private Sub sldTransparency_Change()

    UpdateMenuItem

End Sub

Private Sub UpdateMenuItem()

    If IsUpdating Then Exit Sub
    frmMain.UpdateItemData GetLocalizedStr(189) + cSep + GetLocalizedStr(231), True, True

End Sub

Private Sub LocalizeUI()

    lblDropShadow.Caption = GetLocalizedStr(221): lblDropShadow.Left = sldDropShadow.Left + (sldDropShadow.Width - lblDropShadow.Width) / 2
    lblTransparency.Caption = GetLocalizedStr(222): lblTransparency.Left = sldTransparency.Left + (sldTransparency.Width - lblTransparency.Width) / 2
    lblDSOFF.Caption = GetLocalizedStr(228)
    lblTOFF.Caption = GetLocalizedStr(228)
    lblDSDarker.Caption = GetLocalizedStr(229)
    lblTInvisible.Caption = GetLocalizedStr(230)
    lblShadowColor.Caption = GetLocalizedStr(212)
    chkContextMenu.Caption = GetLocalizedStr(223)

End Sub
