VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{E1C6DB9D-BD4A-4A61-A759-0CED75D034BF}#43.0#0"; "SmartButton.ocx"
Object = "{9A2947B4-87EE-40EE-A3EF-32BDC32D5726}#1.0#0"; "xfxline3d.ocx"
Begin VB.Form frmGrpScrolling 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Group Scrolling"
   ClientHeight    =   6030
   ClientLeft      =   5910
   ClientTop       =   4590
   ClientWidth     =   4980
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmGrpScrolling.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   4980
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdDefaults 
      Caption         =   "Defaults"
      Height          =   375
      Left            =   105
      TabIndex        =   42
      Top             =   5520
      Width           =   900
   End
   Begin VB.TextBox txtBorderSize 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   1575
      TabIndex        =   39
      Text            =   "123"
      Top             =   4530
      Width           =   420
   End
   Begin VB.CheckBox chkScrollOnMouseOver 
      Caption         =   "Scroll when the mouse is over the scroll buttons"
      Height          =   195
      Left            =   285
      TabIndex        =   27
      Top             =   4995
      Width           =   4080
   End
   Begin VB.PictureBox picSize 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   855
      ScaleHeight     =   675
      ScaleWidth      =   735
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   5265
      Visible         =   0   'False
      Width           =   735
   End
   Begin xfxLine3D.ucLine3D uc3DLine3 
      Height          =   30
      Left            =   30
      TabIndex        =   28
      Top             =   5355
      Width           =   4980
      _ExtentX        =   8784
      _ExtentY        =   53
   End
   Begin VB.TextBox txtMargin 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   1575
      TabIndex        =   25
      Text            =   "123"
      Top             =   4170
      Width           =   420
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3975
      TabIndex        =   30
      Top             =   5520
      Width           =   900
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2970
      TabIndex        =   29
      Top             =   5520
      Width           =   900
   End
   Begin VB.CheckBox chkEnable 
      Caption         =   "Enable Scrolling"
      Height          =   225
      Left            =   60
      TabIndex        =   0
      Top             =   180
      Width           =   1980
   End
   Begin VB.Frame frameNormal 
      Caption         =   "Normal"
      Height          =   3375
      Left            =   60
      TabIndex        =   2
      Top             =   615
      Width           =   2340
      Begin VB.ComboBox cmbFXN 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmGrpScrolling.frx":038A
         Left            =   210
         List            =   "frmGrpScrolling.frx":0397
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   2895
         Width           =   1065
      End
      Begin xfxLine3D.ucLine3D uc3DLine1 
         Height          =   30
         Left            =   75
         TabIndex        =   5
         Top             =   720
         Width           =   2160
         _ExtentX        =   3810
         _ExtentY        =   53
      End
      Begin VB.PictureBox picArrow 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   0
         Left            =   210
         ScaleHeight     =   450
         ScaleWidth      =   450
         TabIndex        =   7
         Top             =   1080
         Width           =   480
      End
      Begin VB.PictureBox picArrow 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   2
         Left            =   210
         ScaleHeight     =   450
         ScaleWidth      =   450
         TabIndex        =   11
         Top             =   1935
         Width           =   480
      End
      Begin SmartButtonProject.SmartButton cmdColor 
         Height          =   240
         Index           =   0
         Left            =   1177
         TabIndex        =   4
         Top             =   300
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
         Enabled         =   0   'False
         ShowFocus       =   -1  'True
      End
      Begin SmartButtonProject.SmartButton cmdChange 
         Height          =   240
         Index           =   0
         Left            =   750
         TabIndex        =   8
         Top             =   1080
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   423
         Caption         =   "Change"
         Picture         =   "frmGrpScrolling.frx":03B3
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
      Begin SmartButtonProject.SmartButton cmdChange 
         Height          =   240
         Index           =   2
         Left            =   750
         TabIndex        =   12
         Top             =   1935
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   423
         Caption         =   "Change"
         Picture         =   "frmGrpScrolling.frx":074D
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
      Begin SmartButtonProject.SmartButton cmdRemove 
         Height          =   240
         Index           =   0
         Left            =   750
         TabIndex        =   9
         Top             =   1320
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   423
         Caption         =   "Remove"
         Picture         =   "frmGrpScrolling.frx":0AE7
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
      Begin SmartButtonProject.SmartButton cmdColor 
         Height          =   240
         Index           =   2
         Left            =   1380
         TabIndex        =   32
         Top             =   2932
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
         Enabled         =   0   'False
         ShowFocus       =   -1  'True
      End
      Begin xfxLine3D.ucLine3D uc3DLine4 
         Height          =   30
         Left            =   75
         TabIndex        =   34
         Top             =   2565
         Width           =   2160
         _ExtentX        =   3810
         _ExtentY        =   53
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Border Effect"
         Height          =   195
         Left            =   210
         TabIndex        =   33
         Top             =   2670
         Width           =   960
      End
      Begin VB.Label lblTextColorN 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Back Color"
         Height          =   195
         Left            =   210
         TabIndex        =   3
         Top             =   330
         Width           =   1005
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Arrow Up"
         Height          =   195
         Left            =   210
         TabIndex        =   6
         Top             =   855
         Width           =   675
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Arrow Down"
         Height          =   195
         Left            =   210
         TabIndex        =   10
         Top             =   1710
         Width           =   885
      End
   End
   Begin VB.Frame frameHover 
      Caption         =   "Mouse Over"
      Height          =   3375
      Left            =   2535
      TabIndex        =   13
      Top             =   615
      Width           =   2340
      Begin VB.ComboBox cmbFXO 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmGrpScrolling.frx":0E81
         Left            =   210
         List            =   "frmGrpScrolling.frx":0E8E
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   2895
         Width           =   1065
      End
      Begin VB.PictureBox picArrow 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   1
         Left            =   210
         ScaleHeight     =   450
         ScaleWidth      =   450
         TabIndex        =   18
         Top             =   1080
         Width           =   480
      End
      Begin VB.PictureBox picArrow 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   3
         Left            =   210
         ScaleHeight     =   450
         ScaleWidth      =   450
         TabIndex        =   22
         Top             =   1935
         Width           =   480
      End
      Begin SmartButtonProject.SmartButton cmdColor 
         Height          =   240
         Index           =   1
         Left            =   1282
         TabIndex        =   15
         Top             =   300
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
         Enabled         =   0   'False
         ShowFocus       =   -1  'True
      End
      Begin SmartButtonProject.SmartButton cmdChange 
         Height          =   240
         Index           =   1
         Left            =   855
         TabIndex        =   19
         Top             =   1080
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   423
         Caption         =   "Change"
         Picture         =   "frmGrpScrolling.frx":0EAA
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
      Begin SmartButtonProject.SmartButton cmdChange 
         Height          =   240
         Index           =   3
         Left            =   855
         TabIndex        =   23
         Top             =   1935
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   423
         Caption         =   "Change"
         Picture         =   "frmGrpScrolling.frx":1244
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
      Begin SmartButtonProject.SmartButton cmdRemove 
         Height          =   240
         Index           =   1
         Left            =   855
         TabIndex        =   20
         Top             =   1320
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   423
         Caption         =   "Remove"
         Picture         =   "frmGrpScrolling.frx":15DE
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
      Begin xfxLine3D.ucLine3D uc3DLine2 
         Height          =   30
         Left            =   75
         TabIndex        =   16
         Top             =   720
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   53
      End
      Begin SmartButtonProject.SmartButton cmdColor 
         Height          =   240
         Index           =   3
         Left            =   1380
         TabIndex        =   36
         Top             =   2932
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
         Enabled         =   0   'False
         ShowFocus       =   -1  'True
      End
      Begin xfxLine3D.ucLine3D uc3DLine5 
         Height          =   30
         Left            =   75
         TabIndex        =   38
         Top             =   2565
         Width           =   2160
         _ExtentX        =   3810
         _ExtentY        =   53
      End
      Begin VB.Label lblOver 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Border Effect"
         Height          =   195
         Left            =   210
         TabIndex        =   37
         Top             =   2670
         Width           =   960
      End
      Begin VB.Label lblTextColorO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Back Color"
         Height          =   195
         Left            =   210
         TabIndex        =   14
         Top             =   330
         Width           =   750
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Arrow Up"
         Height          =   195
         Left            =   210
         TabIndex        =   17
         Top             =   855
         Width           =   675
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Arrow Down"
         Height          =   195
         Left            =   210
         TabIndex        =   21
         Top             =   1710
         Width           =   885
      End
   End
   Begin MSComCtl2.UpDown udMargin 
      Height          =   285
      Left            =   1995
      TabIndex        =   26
      Top             =   4170
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   503
      _Version        =   393216
      BuddyControl    =   "txtMargin"
      BuddyDispid     =   196613
      OrigLeft        =   1335
      OrigTop         =   3555
      OrigRight       =   1575
      OrigBottom      =   3840
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   0   'False
   End
   Begin MSComCtl2.UpDown udBorderSize 
      Height          =   285
      Left            =   1995
      TabIndex        =   40
      Top             =   4530
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   503
      _Version        =   393216
      AutoBuddy       =   -1  'True
      BuddyControl    =   "txtBorderSize"
      BuddyDispid     =   196610
      OrigLeft        =   1305
      OrigTop         =   1320
      OrigRight       =   1500
      OrigBottom      =   1575
      Max             =   20
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   0   'False
   End
   Begin VB.Label lblBorderSize 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Border Size"
      Height          =   195
      Left            =   600
      TabIndex        =   41
      Top             =   4575
      Width           =   810
   End
   Begin VB.Image Image2 
      Height          =   240
      Left            =   285
      Picture         =   "frmGrpScrolling.frx":1978
      Top             =   4545
      Width           =   240
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Margin"
      Height          =   195
      Left            =   600
      TabIndex        =   24
      Top             =   4208
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   285
      Picture         =   "frmGrpScrolling.frx":1D02
      Top             =   4185
      Width           =   240
   End
End
Attribute VB_Name = "frmGrpScrolling"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

#If LITE = 0 Then

Dim bGrpS As GrpScrollDef
Dim sId As Integer
Dim IsGlobal As Boolean

Private Sub chkEnable_Click()

    cmdColor(0).Enabled = (chkEnable.Value = vbChecked)
    cmdColor(1).Enabled = (chkEnable.Value = vbChecked)
    cmdColor(2).Enabled = (chkEnable.Value = vbChecked)
    cmdColor(3).Enabled = (chkEnable.Value = vbChecked)
    
    cmdChange(0).Enabled = (chkEnable.Value = vbChecked)
    cmdChange(1).Enabled = (chkEnable.Value = vbChecked)
    cmdChange(2).Enabled = (chkEnable.Value = vbChecked)
    cmdChange(3).Enabled = (chkEnable.Value = vbChecked)
    
    cmdRemove(0).Enabled = (chkEnable.Value = vbChecked)
    cmdRemove(1).Enabled = (chkEnable.Value = vbChecked)
    
    cmbFXN.Enabled = (chkEnable.Value = vbChecked)
    cmbFXO.Enabled = (chkEnable.Value = vbChecked)
    
    txtBorderSize.Enabled = (chkEnable.Value = vbChecked)
    udBorderSize.Enabled = (chkEnable.Value = vbChecked)
    udMargin.Enabled = (chkEnable.Value = vbChecked)
    txtMargin.Enabled = (chkEnable.Value = vbChecked)
    
    If chkEnable.Value = vbChecked Then
        MenuGrps(sId).scrolling.maxHeight = MenuGrps(sId).fHeight
    Else
        MenuGrps(sId).scrolling.maxHeight = 0
    End If

End Sub

Private Sub chkScrollOnMouseOver_Click()

    MenuGrps(sId).scrolling.onmouseover = (chkScrollOnMouseOver.Value = vbChecked)

End Sub

Private Sub cmbFXN_Click()

    MenuGrps(sId).scrolling.FXNormal = cmbFXN.ListIndex

End Sub

Private Sub cmbFXO_Click()

    MenuGrps(sId).scrolling.FXOver = cmbFXO.ListIndex

End Sub

Private Sub cmdCancel_Click()

    MenuGrps(sId).scrolling = bGrpS
    Unload Me

End Sub

Private Sub cmdChange_Click(Index As Integer)

    SelImage.FileName = picArrow(Index).tag
    frmRscImages.Show vbModal
    
    With SelImage
        If .IsValid Then
            Select Case Index
                Case 0
                    MenuGrps(sId).scrolling.UpImage.NormalImage = .FileName
                Case 1
                    MenuGrps(sId).scrolling.UpImage.HoverImage = .FileName
                Case 2
                    MenuGrps(sId).scrolling.DnImage.NormalImage = .FileName
                Case 3
                    MenuGrps(sId).scrolling.DnImage.HoverImage = .FileName
            End Select
        End If
    End With
    
    UpdateImages

End Sub

Private Sub UpdateImages()

    With MenuGrps(sId).scrolling
        picArrow(0).Picture = LoadPictureRes(.UpImage.NormalImage)
        picArrow(1).Picture = LoadPictureRes(.UpImage.HoverImage)
        picArrow(2).Picture = LoadPictureRes(.DnImage.NormalImage)
        picArrow(3).Picture = LoadPictureRes(.DnImage.HoverImage)
    End With

End Sub

Private Sub cmdColor_Click(Index As Integer)

    BuildUsedColorsArray

    With cmdColor(Index)
        SelColor = .tag
        SelColor_CanBeTransparent = False
        frmColorPicker.Show vbModal
        SetColor SelColor, cmdColor(Index)
    End With
    
    cmdColor(Index).ZOrder 0
    
    UpdateGroupData

End Sub

Private Sub UpdateGroupData()

    With MenuGrps(sId).scrolling
        .nColor = cmdColor(0).tag
        .hColor = cmdColor(1).tag
        .FXnColor = cmdColor(2).tag
        .FXhColor = cmdColor(3).tag
    End With

End Sub

Private Sub cmdDefaults_Click()

    With MenuGrps(sId).scrolling
        .UpImage.NormalImage = AppPath + "exhtml\aup_b.gif"
        .UpImage.HoverImage = AppPath + "exhtml\aup_w.gif"
        .DnImage.NormalImage = AppPath + "exhtml\adn_b.gif"
        .DnImage.HoverImage = AppPath + "exhtml\adn_w.gif"
        
        .FXhColor = &H0
        .FXnColor = &H0
        .FXNormal = cfxcNone
        .FXOver = cfxcNone
        .FXSize = 1
        .hColor = &H202080
        .margin = 4
        .nColor = &H808080
        .onmouseover = True
    End With
    
    SetupUI

End Sub

Private Sub cmdOK_Click()

    With MenuGrps(sId).scrolling.DnImage
        Set picSize.Picture = picArrow(0).Picture
        DoEvents
        .w = picSize.Width / Screen.TwipsPerPixelX
        .h = picSize.Height / Screen.TwipsPerPixelY
    End With

    Unload Me

End Sub

Private Sub cmdRemove_Click(Index As Integer)

    With MenuGrps(sId).scrolling
        Select Case Index
            Case 0
                .UpImage.NormalImage = ""
                .DnImage.NormalImage = ""
            Case 1
                .UpImage.HoverImage = ""
                .DnImage.HoverImage = ""
        End Select
    End With
    
    UpdateImages

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF1 Then showHelp "dialogs/group_sfx_gs_src.htm"

End Sub

Private Sub Form_Load()

    CenterForm Me
    SetupCharset Me
    LocalizeUI
    
    sId = GetID
    IsGlobal = (MenuGrps(sId).Name = "mnuGrpGIS_771248")
    
    bGrpS = MenuGrps(sId).scrolling
    
    caption = NiceGrpCaption(sId) + IIf(IsGlobal, "", " - " + GetLocalizedStr(830))
    
    SetupUI
    
    If IsGlobal Then
        chkEnable.Visible = False
        chkEnable.Value = vbChecked
    End If

End Sub

Private Sub SetupUI()

    With MenuGrps(sId).scrolling
        chkEnable.Value = IIf(.maxHeight > 0 Or IsGlobal, vbChecked, vbUnchecked)
        
        SetColor .nColor, cmdColor(0)
        SetColor .hColor, cmdColor(1)
        SetColor .FXnColor, cmdColor(2)
        SetColor .FXhColor, cmdColor(3)
        
        txtMargin.Text = .margin
        txtBorderSize.Text = .FXSize
        
        cmbFXN.ListIndex = .FXNormal
        cmbFXO.ListIndex = .FXOver
        
        chkScrollOnMouseOver.Value = IIf(.onmouseover, vbChecked, vbUnchecked)
    End With
    
    UpdateImages

End Sub

Private Sub LocalizeUI()

    PopulateBorderStyleCombo cmbFXN
    PopulateBorderStyleCombo cmbFXO
    
    cmdOK.caption = GetLocalizedStr(186)
    cmdCancel.caption = GetLocalizedStr(187)
    
    If Preferences.language <> "eng" Then
        cmdOK.Width = SetCtrlWidth(cmdOK)
        cmdCancel.Width = SetCtrlWidth(cmdCancel)
    End If

End Sub

Private Sub txtBorderSize_Change()

    MenuGrps(sId).scrolling.FXSize = Val(txtBorderSize.Text)

End Sub

Private Sub txtMargin_Change()

    MenuGrps(sId).scrolling.margin = Val(txtMargin.Text)

End Sub

Private Sub txtMargin_GotFocus()

    SelAll txtMargin

End Sub

#End If
