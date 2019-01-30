VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{E1C6DB9D-BD4A-4A61-A759-0CED75D034BF}#43.0#0"; "SmartButton.ocx"
Object = "{9A2947B4-87EE-40EE-A3EF-32BDC32D5726}#1.0#0"; "xfxline3d.ocx"
Begin VB.Form frmGrpColor 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Color"
   ClientHeight    =   5895
   ClientLeft      =   3030
   ClientTop       =   4575
   ClientWidth     =   5850
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmGrpColor.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   5850
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frameToolbarItem 
      BorderStyle     =   0  'None
      Caption         =   "Toolbar Item"
      Height          =   2505
      Left            =   5415
      TabIndex        =   0
      Top             =   -15
      Width           =   5520
      Begin VB.Frame frameHover 
         Caption         =   "Mouse Over"
         Height          =   1095
         Left            =   2610
         TabIndex        =   6
         Top             =   255
         Width           =   2355
         Begin SmartButtonProject.SmartButton cmdColor 
            Height          =   240
            Index           =   3
            Left            =   1590
            TabIndex        =   8
            Top             =   315
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
            Index           =   4
            Left            =   1590
            TabIndex        =   10
            Top             =   660
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
            Left            =   675
            TabIndex        =   9
            Top             =   690
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
            Left            =   675
            TabIndex        =   7
            Top             =   345
            Width           =   750
         End
      End
      Begin VB.Frame frameNormal 
         Caption         =   "Normal"
         Height          =   1095
         Left            =   165
         TabIndex        =   1
         Top             =   255
         Width           =   2355
         Begin SmartButtonProject.SmartButton cmdColor 
            Height          =   240
            Index           =   1
            Left            =   1620
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
         Begin SmartButtonProject.SmartButton cmdColor 
            Height          =   240
            Index           =   2
            Left            =   1620
            TabIndex        =   5
            Top             =   660
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
            Left            =   675
            TabIndex        =   2
            Top             =   360
            Width           =   750
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
            Left            =   675
            TabIndex        =   4
            Top             =   690
            Width           =   750
         End
      End
   End
   Begin SmartButtonProject.SmartButton sbApplyOptions 
      Height          =   345
      Left            =   135
      TabIndex        =   11
      Top             =   120
      Width           =   5595
      _ExtentX        =   9869
      _ExtentY        =   609
      Caption         =   "Options"
      Picture         =   "frmGrpColor.frx":014A
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
   Begin VB.Frame frameGroupProp 
      BorderStyle     =   0  'None
      Caption         =   "Group Properties"
      Height          =   2940
      Left            =   165
      TabIndex        =   13
      Top             =   900
      Width           =   5520
      Begin VB.TextBox txtRadiusBL 
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
         Left            =   3525
         TabIndex        =   25
         Text            =   "123"
         Top             =   1335
         Width           =   420
      End
      Begin VB.TextBox txtRadiusBR 
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
         Left            =   4200
         TabIndex        =   26
         Text            =   "123"
         Top             =   1335
         Width           =   420
      End
      Begin VB.TextBox txtRadiusTL 
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
         Left            =   3525
         TabIndex        =   17
         Text            =   "123"
         Top             =   720
         Width           =   420
      End
      Begin VB.TextBox txtRadiusTR 
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
         Left            =   4200
         TabIndex        =   18
         Text            =   "123"
         Top             =   720
         Width           =   420
      End
      Begin VB.ComboBox cmbFX 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmGrpColor.frx":02A4
         Left            =   1583
         List            =   "frmGrpColor.frx":02B1
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   2310
         Width           =   1065
      End
      Begin SmartButtonProject.SmartButton cmdColor 
         Height          =   180
         Index           =   7
         Left            =   1710
         TabIndex        =   21
         Top             =   1440
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   318
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
      Begin xfxLine3D.ucLine3D uc3DLine1 
         Height          =   30
         Left            =   60
         TabIndex        =   27
         Top             =   1740
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   53
      End
      Begin VB.ComboBox cmbBorder 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmGrpColor.frx":02CD
         Left            =   1583
         List            =   "frmGrpColor.frx":02F2
         TabIndex        =   29
         Text            =   "cmbBorder"
         Top             =   1920
         WhatsThisHelpID =   20240
         Width           =   1065
      End
      Begin SmartButtonProject.SmartButton cmdColor 
         Height          =   240
         Index           =   0
         Left            =   1845
         TabIndex        =   15
         Top             =   270
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
         PictureLayout   =   6
         ShowFocus       =   -1  'True
      End
      Begin SmartButtonProject.SmartButton cmdColor 
         Height          =   540
         Index           =   6
         Left            =   2250
         TabIndex        =   20
         Top             =   900
         Width           =   180
         _ExtentX        =   318
         _ExtentY        =   953
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
         Height          =   540
         Index           =   8
         Left            =   1530
         TabIndex        =   22
         Top             =   900
         Width           =   180
         _ExtentX        =   318
         _ExtentY        =   953
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
         Height          =   180
         Index           =   5
         Left            =   1710
         TabIndex        =   16
         Top             =   720
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   318
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PictureLayout   =   6
         ShowFocus       =   -1  'True
      End
      Begin SmartButtonProject.SmartButton cmdColor 
         Height          =   240
         Index           =   9
         Left            =   1845
         TabIndex        =   23
         Top             =   1050
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
         PictureLayout   =   6
         ShowFocus       =   -1  'True
      End
      Begin VB.Label lblBorderRadius 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Radius"
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
         Left            =   2865
         TabIndex        =   24
         Top             =   1050
         Width           =   480
      End
      Begin VB.Label lblBorderStyle 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Border Style"
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
         Left            =   495
         TabIndex        =   30
         Top             =   2370
         Width           =   885
      End
      Begin VB.Label lblBorderSize 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Border Size"
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
         Left            =   570
         TabIndex        =   28
         Top             =   1980
         Width           =   810
      End
      Begin VB.Label lblCorners 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Corners"
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
         Left            =   810
         TabIndex        =   19
         Top             =   1050
         Width           =   570
      End
      Begin VB.Label lblGBackColor 
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
         Left            =   630
         TabIndex        =   14
         Top             =   300
         Width           =   750
      End
   End
   Begin MSComctlLib.TabStrip tsOptions 
      Height          =   3315
      Left            =   120
      TabIndex        =   12
      Top             =   555
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   5847
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Group Properties"
            Key             =   "tsGroupProperties"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Toolbar Item"
            Key             =   "tsToolbarItem"
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
      TabIndex        =   32
      Top             =   3975
      Width           =   5625
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
      Left            =   3780
      TabIndex        =   33
      Top             =   5415
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
      Left            =   4845
      TabIndex        =   34
      Top             =   5415
      Width           =   900
   End
End
Attribute VB_Name = "frmGrpColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim BackGrp As MenuGrp
Dim IsLoading As Boolean

Private WithEvents msgSubClass As xfxSC
Attribute msgSubClass.VB_VarHelpID = -1

Private Sub cmbBorder_Change()

    UpdateSample

End Sub

Private Sub cmbBorder_Click()

    UpdateSample

End Sub

Private Sub cmbFX_Click()

    UpdateSample

End Sub

Private Sub cmdCancel_Click()

    MenuGrps(GetID) = BackGrp
    Unload Me

End Sub

Private Sub cmdColor_Click(Index As Integer)

    BuildUsedColorsArray

    With cmdColor(Index)
        SelColor = .tag
        SelColor_CanBeTransparent = Not (Index = 1 Or Index = 3)
        frmColorPicker.Show vbModal
        SetColor SelColor, cmdColor(Index)
    End With
    
    cmdColor(Index).ZOrder 0
    
    If Index = 9 Then
        Dim i As Integer
        For i = 5 To 8
            SetColor SelColor, cmdColor(i)
        Next i
    End If
    
    UpdateSample

End Sub

Private Sub UpdateSample()

    If IsLoading Then Exit Sub

    With MenuGrps(GetID)
        .bColor = cmdColor(0).tag
        .nTextColor = cmdColor(1).tag
        .nBackColor = cmdColor(2).tag
        .hTextColor = cmdColor(3).tag
        .hBackColor = cmdColor(4).tag
        
        .Corners.topCorner = cmdColor(5).tag
        .Corners.rightCorner = cmdColor(6).tag
        .Corners.bottomCorner = cmdColor(7).tag
        .Corners.leftCorner = cmdColor(8).tag
        
        .frameBorder = Val(cmbBorder.Text)
        
        .Radius.topLeft = Val(txtRadiusTL.Text)
        .Radius.topRight = Val(txtRadiusTR.Text)
        .Radius.bottomLeft = Val(txtRadiusBL.Text)
        .Radius.bottomRight = Val(txtRadiusBR.Text)
        
        .BorderStyle = cmbFX.ListIndex
    End With
    
    frmMain.DoLivePreview wbLivePreview, (tsOptions.SelectedItem.key = "tsGroupProperties")

End Sub

Private Sub cmdOK_Click()

    ApplyStyleOptions
    frmMain.SaveState "Change " + MenuGrps(GetID).Name + " Color"
    
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
        .bColor = MenuGrps(sId).bColor
        .BorderStyle = MenuGrps(sId).BorderStyle
        .Corners = MenuGrps(sId).Corners
        .frameBorder = MenuGrps(sId).frameBorder
        .hBackColor = MenuGrps(sId).hBackColor
        .hTextColor = MenuGrps(sId).hTextColor
        .nBackColor = MenuGrps(sId).nBackColor
        .nTextColor = MenuGrps(sId).nTextColor
        .Radius.topLeft = MenuGrps(sId).Radius.topLeft
        .Radius.topRight = MenuGrps(sId).Radius.topRight
        .Radius.bottomLeft = MenuGrps(sId).Radius.bottomLeft
        .Radius.bottomRight = MenuGrps(sId).Radius.bottomRight
    End With

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case tsOptions.SelectedItem.key
        Case "tsGroupProperties"
            If KeyCode = vbKeyF1 Then showHelp "dialogs/group_color.htm"
        Case "tsToolbarItem"
            If KeyCode = vbKeyF1 Then showHelp "dialogs/group_color_tbi.htm"
    End Select

End Sub

Private Sub Form_Load()
    
    CenterForm Me
    LocalizeUI
    
    Set msgSubClass = New xfxSC
    msgSubClass.SubClassHwnd Me.hwnd, True
    
    BackGrp = MenuGrps(GetID)
    
    IsLoading = True
    With MenuGrps(GetID)
        SetColor .bColor, cmdColor(0)
        SetColor .nTextColor, cmdColor(1)
        SetColor .nBackColor, cmdColor(2)
        SetColor .hTextColor, cmdColor(3)
        SetColor .hBackColor, cmdColor(4)
        
        SetColor .Corners.topCorner, cmdColor(5)
        SetColor .Corners.rightCorner, cmdColor(6)
        SetColor .Corners.bottomCorner, cmdColor(7)
        SetColor .Corners.leftCorner, cmdColor(8)
        SetColor .Corners.topCorner, cmdColor(9)
        
        cmbBorder.Text = .frameBorder
        
        caption = NiceGrpCaption(GetID) + " - " + GetLocalizedStr(212)
        
        cmbFX.ListIndex = .BorderStyle
        
        txtRadiusTL.Text = CStr(.Radius.topLeft)
        txtRadiusTR.Text = CStr(.Radius.topRight)
        txtRadiusBL.Text = CStr(.Radius.bottomLeft)
        txtRadiusBR.Text = CStr(.Radius.bottomRight)
    End With
    IsLoading = False
    
    frameToolbarItem.Enabled = CreateToolbar
    frameHover.Enabled = CreateToolbar
    frameNormal.Enabled = CreateToolbar
    
    frameGroupProp.ZOrder 0
    frameToolbarItem.Move frameGroupProp.Left, frameGroupProp.Top
    
    FixCtrls4Skin Me
    
    If BelongsToToolbar(GetID, True) > 0 Then
        If IsSubMenu(GetID) Then tsOptions.Tabs.Remove 2
    Else
        tsOptions.Tabs.Remove 2
    End If
    
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

Private Sub tsOptions_Click()

    Select Case tsOptions.SelectedItem.key
        Case "tsGroupProperties"
            frameToolbarItem.Visible = False
            frameGroupProp.Visible = True
        Case "tsToolbarItem"
            frameToolbarItem.Visible = True
            frameGroupProp.Visible = False
    End Select
    
    UpdateSample

End Sub

Private Sub LocalizeUI()

    'frameGroupProp.Caption = GetLocalizedStr(124)
    'frameToolbarItem.Caption = GetLocalizedStr(204)
    tsOptions.Tabs("tsGroupProperties").caption = GetLocalizedStr(124)
    tsOptions.Tabs("tsToolbarItem").caption = GetLocalizedStr(204)

    frameNormal.caption = GetLocalizedStr(179)
    frameHover.caption = GetLocalizedStr(180)

    lblTextColorN.caption = GetLocalizedStr(181)
    lblBackColorN.caption = GetLocalizedStr(182)
    lblTextColorO.caption = GetLocalizedStr(181)
    lblBackColorO.caption = GetLocalizedStr(182)
    
    PopulateBorderStyleCombo cmbFX
    
    lblBorderSize.caption = GetLocalizedStr(206)
    
    lblCorners.caption = GetLocalizedStr(667)
    
    lblGBackColor.caption = GetLocalizedStr(182)
    
    frmLiveSample.caption = GetLocalizedStr(188)
    
    cmdOK.caption = GetLocalizedStr(186)
    cmdCancel.caption = GetLocalizedStr(187)
    
    If Preferences.language <> "eng" Then
        cmdOK.Width = SetCtrlWidth(cmdOK)
        cmdCancel.Width = SetCtrlWidth(cmdCancel)
    End If

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
