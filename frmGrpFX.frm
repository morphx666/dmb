VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{E1C6DB9D-BD4A-4A61-A759-0CED75D034BF}#43.0#0"; "SmartButton.ocx"
Object = "{DB06EC30-01E1-485F-A3C7-CE80CA0D7D37}#2.0#0"; "xFXSlider.ocx"
Object = "{9A2947B4-87EE-40EE-A3EF-32BDC32D5726}#1.0#0"; "xfxline3d.ocx"
Begin VB.Form frmGrpFX 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Special Effects"
   ClientHeight    =   6660
   ClientLeft      =   4320
   ClientTop       =   5085
   ClientWidth     =   10005
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmGrpFX.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   10005
   ShowInTaskbar   =   0   'False
   Begin SmartButtonProject.SmartButton sbApplyOptions 
      Height          =   345
      Left            =   135
      TabIndex        =   32
      Top             =   120
      Width           =   5220
      _ExtentX        =   9208
      _ExtentY        =   609
      Caption         =   "Options"
      Picture         =   "frmGrpFX.frx":014A
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
   Begin VB.Frame frameGroupSize 
      BorderStyle     =   0  'None
      Caption         =   "Group Size"
      Height          =   3315
      Left            =   135
      TabIndex        =   15
      Top             =   885
      Width           =   5190
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Height          =   1395
         Left            =   60
         TabIndex        =   22
         Top             =   1635
         Width           =   3450
         Begin VB.CommandButton cmdScrolling 
            Caption         =   "Scrolling..."
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
            Height          =   285
            Left            =   1470
            TabIndex        =   29
            Top             =   1110
            Width           =   1170
         End
         Begin VB.TextBox txtHeight 
            Alignment       =   1  'Right Justify
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
            Height          =   285
            Left            =   1020
            TabIndex        =   27
            Text            =   "0"
            Top             =   780
            Width           =   405
         End
         Begin VB.OptionButton opHeight 
            Caption         =   "Manual"
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
            Index           =   2
            Left            =   765
            TabIndex        =   26
            Top             =   540
            Width           =   1920
         End
         Begin VB.OptionButton opHeight 
            Caption         =   "Background Image"
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
            Index           =   1
            Left            =   765
            TabIndex        =   24
            Top             =   300
            Width           =   1920
         End
         Begin VB.OptionButton opHeight 
            Caption         =   "Auto"
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
            Index           =   0
            Left            =   765
            TabIndex        =   23
            Top             =   60
            Value           =   -1  'True
            Width           =   1920
         End
         Begin SmartButtonProject.SmartButton cmdAutoHeight 
            Height          =   300
            Left            =   1470
            TabIndex        =   28
            Top             =   780
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   529
            Caption         =   "Calculate"
            Picture         =   "frmGrpFX.frx":02A4
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
            Enabled         =   0   'False
         End
         Begin VB.Label lblGSHeight 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Height"
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
            Left            =   165
            TabIndex        =   25
            Top             =   435
            Width           =   465
         End
         Begin VB.Image imgGSHeight 
            Appearance      =   0  'Flat
            Height          =   240
            Left            =   270
            Picture         =   "frmGrpFX.frx":063E
            Top             =   210
            Width           =   240
         End
      End
      Begin VB.OptionButton opWidth 
         Caption         =   "Manual"
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
         Index           =   2
         Left            =   855
         TabIndex        =   19
         Top             =   780
         Width           =   1935
      End
      Begin VB.TextBox txtWidth 
         Alignment       =   1  'Right Justify
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
         Height          =   285
         Left            =   1110
         TabIndex        =   20
         Text            =   "0"
         Top             =   1020
         Width           =   405
      End
      Begin VB.OptionButton opWidth 
         Caption         =   "Background Image"
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
         Index           =   1
         Left            =   855
         TabIndex        =   17
         Top             =   525
         Width           =   1935
      End
      Begin VB.OptionButton opWidth 
         Caption         =   "Auto"
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
         Index           =   0
         Left            =   855
         TabIndex        =   16
         Top             =   285
         Value           =   -1  'True
         Width           =   1935
      End
      Begin SmartButtonProject.SmartButton cmdAutoWidth 
         Height          =   300
         Left            =   1560
         TabIndex        =   21
         Top             =   1020
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   529
         Caption         =   "Calculate"
         Picture         =   "frmGrpFX.frx":09C8
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
         Enabled         =   0   'False
      End
      Begin VB.Label lblGSWidth 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Width"
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
         Left            =   270
         TabIndex        =   18
         Top             =   675
         Width           =   420
      End
      Begin VB.Image imgGSWidth 
         Appearance      =   0  'Flat
         Height          =   240
         Left            =   360
         Picture         =   "frmGrpFX.frx":0D62
         Top             =   450
         Width           =   240
      End
   End
   Begin VB.Frame frameGE 
      BorderStyle     =   0  'None
      Caption         =   "Special Group Effects"
      Height          =   3315
      Left            =   4335
      TabIndex        =   0
      Top             =   1305
      Width           =   5190
      Begin xFXSlider.ucSlider sldTransparency 
         Height          =   270
         Left            =   420
         TabIndex        =   8
         Top             =   1305
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
         CustomKnobImage =   "frmGrpFX.frx":10EC
         CustomSelKnobImage=   "frmGrpFX.frx":13C6
      End
      Begin VB.CheckBox chkContextMenu 
         Caption         =   "Context Menu"
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
         Left            =   420
         TabIndex        =   12
         Top             =   2265
         Width           =   2625
      End
      Begin xfxLine3D.ucLine3D uc3DLine2 
         Height          =   30
         Left            =   90
         TabIndex        =   11
         Top             =   2025
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   53
      End
      Begin SmartButtonProject.SmartButton cmdShadowColor 
         Height          =   240
         Left            =   4740
         TabIndex        =   4
         Top             =   390
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
         Left            =   420
         TabIndex        =   3
         Top             =   375
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
         CustomKnobImage =   "frmGrpFX.frx":16A0
         CustomSelKnobImage=   "frmGrpFX.frx":197A
      End
      Begin VB.Label lblShadowColor 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Color"
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
         Left            =   4688
         TabIndex        =   2
         Top             =   150
         Width           =   375
      End
      Begin VB.Label lblDropShadow 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Drop Shadow Size"
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
         Left            =   1875
         TabIndex        =   1
         Top             =   150
         Width           =   1290
      End
      Begin VB.Label lblTransparency 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transparency"
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
         Left            =   2025
         TabIndex        =   7
         Top             =   1080
         Width           =   990
      End
      Begin VB.Image Image5 
         Height          =   240
         Left            =   165
         Picture         =   "frmGrpFX.frx":1C54
         Top             =   390
         Width           =   240
      End
      Begin VB.Image Image4 
         Height          =   240
         Left            =   165
         Picture         =   "frmGrpFX.frx":1D9E
         Top             =   1320
         Width           =   240
      End
      Begin VB.Label lblDSOFF 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "OFF"
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
         Left            =   465
         TabIndex        =   5
         Top             =   690
         Width           =   300
      End
      Begin VB.Label lblDSDarker 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Larger"
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
         Left            =   4275
         TabIndex        =   6
         Top             =   690
         Width           =   465
      End
      Begin VB.Label lblTOFF 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "OFF"
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
         Left            =   450
         TabIndex        =   9
         Top             =   1620
         Width           =   300
      End
      Begin VB.Label lblTInvisible 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Invisible"
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
         Left            =   4245
         TabIndex        =   10
         Top             =   1620
         Width           =   585
      End
   End
   Begin MSComctlLib.TabStrip tsOptions 
      Height          =   3675
      Left            =   120
      TabIndex        =   31
      Top             =   555
      Width           =   5235
      _ExtentX        =   9234
      _ExtentY        =   6482
      TabWidthStyle   =   1
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Group Size"
            Key             =   "tsGroupSize"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Group Effects"
            Key             =   "tsSGFX"
            ImageVarType    =   2
         EndProperty
      EndProperty
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
      TabIndex        =   13
      Top             =   4305
      Width           =   5235
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
      Left            =   3420
      TabIndex        =   14
      Top             =   5730
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
      TabIndex        =   30
      Top             =   5730
      Width           =   900
   End
End
Attribute VB_Name = "frmGrpFX"
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

Private Sub chkContextMenu_Click()

    Dim g As Integer
    
    For g = 1 To UBound(MenuGrps)
        MenuGrps(g).IsContext = False
    Next g
    
    UpdateSample

End Sub

Private Sub cmdAutoHeight_Click()

    txtHeight.Text = GetOptimalHeight
    txtHeight.SetFocus
    
    UpdateSample

End Sub

Private Function GetOptimalHeight() As Integer

    Dim sId As Integer
    Dim lfHeight As Integer
    
    sId = GetID

    lfHeight = MenuGrps(sId).fHeight
    MenuGrps(sId).fHeight = 0
    GetOptimalHeight = GetDivHeight(sId)
    MenuGrps(sId).fHeight = lfHeight

End Function

Private Sub cmdAutoWidth_Click()

    Dim lfWidth As Integer

    lfWidth = MenuGrps(GetID).fWidth
    MenuGrps(GetID).fWidth = 0
    txtWidth.Text = GetDivWidth(GetID)
    txtWidth.SetFocus
    MenuGrps(GetID).fWidth = lfWidth
    
    UpdateSample

End Sub

Private Sub cmdCancel_Click()

    MenuGrps(SelId) = BackGrp
    Unload Me

End Sub

Private Sub UpdateSample(Optional IsLoading As Boolean)

    On Error Resume Next
    
    If IsUpdating Then Exit Sub

    With MenuGrps(SelId)
        .DropShadowSize = sldDropShadow.Value
        .DropShadowColor = cmdShadowColor.tag
        cmdShadowColor.Enabled = .DropShadowSize > 0
        .Transparency = sldTransparency.Value
        .IsContext = (chkContextMenu.Value = vbChecked)
        
        If opWidth(0).Value Then .fWidth = 0
        If opWidth(1).Value Then .fWidth = -1
        If opWidth(2).Value Then .fWidth = Abs(Val(txtWidth.Text))
        
        If opHeight(0).Value Then
            .fHeight = 0
            .scrolling.maxHeight = 0
        End If
        If opHeight(1).Value Then .fHeight = -1
        If opHeight(2).Value Then .fHeight = Abs(Val(txtHeight.Text))
        
        cmdScrolling.Enabled = opHeight(2).Value And (.fHeight < GetOptimalHeight) And (Val(txtHeight.Text) > 0)
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
    If Not IsLoading Then frmMain.DoLivePreview wbLivePreview, True

End Sub

Private Sub cmdOK_Click()

    If MenuGrps(GetID).scrolling.maxHeight > 0 Then
        MenuGrps(GetID).fHeight = 0
    End If
    
    ApplyStyleOptions
    frmMain.SaveState "Change " + MenuGrps(GetID).Name + " " + GetLocalizedStr(231)
    
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
        .DropShadowSize = MenuGrps(sId).DropShadowSize
        .DropShadowColor = MenuGrps(sId).DropShadowColor
        .Transparency = MenuGrps(sId).Transparency
        .IsContext = MenuGrps(sId).IsContext
        .fWidth = MenuGrps(sId).fWidth
        .fHeight = MenuGrps(sId).fHeight
        .scrolling = MenuGrps(sId).scrolling
    End With

End Sub

Private Sub cmdScrolling_Click()

    frmGrpScrolling.Show vbModal
    txtHeight.SetFocus

End Sub

Private Sub cmdShadowColor_Click()

    BuildUsedColorsArray
    
    With cmdShadowColor
        SelColor = .tag
        SelColor_CanBeTransparent = False
        frmColorPicker.Show vbModal
        SetColor SelColor, cmdShadowColor
    End With
    
    UpdateSample

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case tsOptions.SelectedItem.key
        Case "tsGroupSize"
            If KeyCode = vbKeyF1 Then showHelp "dialogs/group_sfx.htm"
        Case "tsSGFX"
            If KeyCode = vbKeyF1 Then showHelp "dialogs/group_sfx_sgfx.htm"
    End Select

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    msgSubClass.SubClassHwnd Me.hwnd, False

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

Private Sub Form_Load()

    #If LITE = 1 Then
        cmdScrolling.Visible = False
    #End If

    Width = 5550
    Height = cmdOK.Top + cmdOK.Height + GetClientTop(Me.hwnd) + 105

    CenterForm Me
    LocalizeUI
    
    Set msgSubClass = New xfxSC
    msgSubClass.SubClassHwnd Me.hwnd, True
    
    With frameGroupSize
        frameGE.Move .Left, .Top, .Width, .Height
        .ZOrder 0
    End With
    
    SelId = GetID
    BackGrp = MenuGrps(SelId)
    
    IsUpdating = True
    
    With MenuGrps(GetID)
        sldDropShadow.Value = .DropShadowSize
        sldTransparency.Value = .Transparency
        SetColor .DropShadowColor, cmdShadowColor
        chkContextMenu.Value = Abs(.IsContext)
        
        caption = NiceGrpCaption(GetID) + " - " + GetLocalizedStr(231)
        
        If LenB(.Image) = 0 Then
            opWidth(1).Enabled = False
            opHeight(1).Enabled = False
        End If
        Select Case .fWidth
            Case -1
                opWidth(IIf(LenB(.Image) = 0, 0, 1)).Value = True
            Case 0
                opWidth(0).Value = True
            Case Else
                opWidth(2).Value = True
                txtWidth.Text = .fWidth
        End Select
        If .scrolling.maxHeight > 0 Then .fHeight = .scrolling.maxHeight
        Select Case .fHeight
            Case -1
                opHeight(IIf(LenB(.Image) = 0, 0, 1)).Value = True
            Case 0
                opHeight(0).Value = True
            Case Else
                opHeight(2).Value = True
                txtHeight.Text = .fHeight
        End Select
    End With
    FixCtrls4Skin Me
    
    IsUpdating = False
    
    UpdateSample True

End Sub

Private Sub opHeight_Click(Index As Integer)

    txtHeight.Enabled = opHeight(2).Value
    cmdAutoHeight.Enabled = opHeight(2).Value
    
    If IsUpdating Then Exit Sub
    
    If Index = 2 And Val(txtHeight.Text) = 0 Then
        cmdAutoHeight_Click
    Else
        UpdateSample
    End If

End Sub

Private Sub opWidth_Click(Index As Integer)

    txtWidth.Enabled = opWidth(2).Value
    cmdAutoWidth.Enabled = opWidth(2).Value
    
    If IsUpdating Then Exit Sub
    
    If Index = 2 And Val(txtWidth.Text) = 0 Then
        cmdAutoWidth_Click
    Else
        UpdateSample
    End If

End Sub

Private Sub sbApplyOptions_Click()

    PopupMenu frmMain.mnuStyleOptions, , sbApplyOptions.Left, sbApplyOptions.Top + sbApplyOptions.Height
    
End Sub

Private Sub sldDropShadow_Change()

    UpdateSample

End Sub

Private Sub sldTransparency_Change()

    UpdateSample

End Sub

Private Sub tsOptions_Click()

    Select Case tsOptions.SelectedItem.key
        Case "tsGroupSize"
            frameGE.Visible = False
            frameGroupSize.Visible = True
        Case "tsSGFX"
            #If LITE = 0 Then
                frameGE.Visible = True
                frameGroupSize.Visible = False
                sldDropShadow.SetFocus
            #Else
                frmMain.ShowLITELImitationInfo 2
                tsOptions.Tabs(1).Selected = True
            #End If
    End Select
    
    Refresh

End Sub

Private Sub txtHeight_Change()

    UpdateSample

End Sub

Private Sub txtHeight_GotFocus()

    SelAll txtHeight

End Sub

Private Sub txtHeight_KeyPress(KeyAscii As Integer)

    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If

End Sub

Private Sub txtWidth_Change()

    UpdateSample

End Sub

Private Sub txtWidth_GotFocus()

    SelAll txtWidth

End Sub

Private Sub txtWidth_KeyPress(KeyAscii As Integer)

    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If

End Sub

Private Sub LocalizeUI()
    
    'frameGE.Caption = GetLocalizedStr(218)
    tsOptions.Tabs("tsSGFX").caption = GetLocalizedStr(218)
    lblDropShadow.caption = GetLocalizedStr(221): lblDropShadow.Left = sldDropShadow.Left + (sldDropShadow.Width - lblDropShadow.Width) / 2
    lblTransparency.caption = GetLocalizedStr(222): lblTransparency.Left = sldTransparency.Left + (sldTransparency.Width - lblTransparency.Width) / 2
    lblDSOFF.caption = GetLocalizedStr(228)
    lblTOFF.caption = GetLocalizedStr(228)
    lblDSDarker.caption = GetLocalizedStr(229)
    lblTInvisible.caption = GetLocalizedStr(230)
    lblShadowColor.caption = GetLocalizedStr(212)
    chkContextMenu.caption = GetLocalizedStr(223)
    
    'frameGroupSize.Caption = GetLocalizedStr(227)
    tsOptions.Tabs("tsGroupSize").caption = GetLocalizedStr(227)
    opWidth(0).caption = GetLocalizedStr(185)
    opWidth(1).caption = GetLocalizedStr(199)
    opWidth(2).caption = GetLocalizedStr(224)
    opHeight(0).caption = GetLocalizedStr(185)
    opHeight(1).caption = GetLocalizedStr(199)
    opHeight(2).caption = GetLocalizedStr(224)
    lblGSWidth.caption = GetLocalizedStr(428): lblGSWidth.Left = (imgGSWidth.Width - lblGSWidth.Width) / 2 + imgGSWidth.Left
    lblGSHeight.caption = GetLocalizedStr(429): lblGSHeight.Left = (imgGSHeight.Width - lblGSHeight.Width) / 2 + imgGSHeight.Left
    cmdAutoWidth.caption = GetLocalizedStr(225)
    cmdAutoHeight.caption = GetLocalizedStr(225)
    
    frmLiveSample.caption = GetLocalizedStr(188)
    
    cmdOK.caption = GetLocalizedStr(186)
    cmdCancel.caption = GetLocalizedStr(187)
    
    FixContolsWidth Me
    
    cmdAutoWidth.Width = SetCtrlWidth(cmdAutoWidth)
    cmdAutoHeight.Width = SetCtrlWidth(cmdAutoHeight)
    
    If Preferences.language <> "eng" Then
        cmdOK.Width = SetCtrlWidth(cmdOK)
        cmdCancel.Width = SetCtrlWidth(cmdCancel)
    End If

End Sub
