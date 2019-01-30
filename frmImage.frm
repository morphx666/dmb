VERSION 5.00
Object = "{E1C6DB9D-BD4A-4A61-A759-0CED75D034BF}#43.0#0"; "SmartButton.ocx"
Begin VB.Form frmImage 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Image"
   ClientHeight    =   5910
   ClientLeft      =   4005
   ClientTop       =   2940
   ClientWidth     =   7260
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmImage.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   7260
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picSize 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   330
      ScaleHeight     =   675
      ScaleWidth      =   735
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   5085
      Visible         =   0   'False
      Width           =   735
   End
   Begin SmartButtonProject.SmartButton sbApplyOptions 
      Height          =   345
      Left            =   135
      TabIndex        =   39
      Top             =   120
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   609
      Caption         =   "Options"
      Picture         =   "frmImage.frx":014A
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
      TabIndex        =   36
      Top             =   3975
      Width           =   7020
   End
   Begin VB.Frame frameBack 
      Caption         =   "Background Image"
      Height          =   3330
      Left            =   2505
      TabIndex        =   13
      Top             =   555
      Width           =   2235
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   510
         Left            =   195
         ScaleHeight     =   510
         ScaleWidth      =   1800
         TabIndex        =   40
         Top             =   2685
         Width           =   1800
         Begin VB.CheckBox chkTile 
            Caption         =   "Tile"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   0
            TabIndex        =   42
            Top             =   0
            Width           =   1710
         End
         Begin VB.CheckBox chkAllowCrop 
            Caption         =   "Allow Cropping"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   0
            TabIndex        =   41
            Top             =   255
            Width           =   1710
         End
      End
      Begin VB.Frame frameBackO 
         Caption         =   "Mouse Over"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   195
         TabIndex        =   18
         Top             =   1470
         Width           =   1800
         Begin VB.PictureBox picBackOver 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   480
            Left            =   105
            ScaleHeight     =   450
            ScaleWidth      =   450
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   375
            Width           =   480
         End
         Begin SmartButtonProject.SmartButton cmdChange 
            Height          =   240
            Index           =   5
            Left            =   630
            TabIndex        =   20
            Top             =   375
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   423
            Caption         =   "Change"
            Picture         =   "frmImage.frx":02A4
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
         Begin SmartButtonProject.SmartButton cmdSameBack 
            Height          =   240
            Left            =   630
            TabIndex        =   21
            Top             =   615
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   423
            Caption         =   "Same"
            Picture         =   "frmImage.frx":063E
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
      End
      Begin VB.Frame frameBackN 
         Caption         =   "Normal"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   195
         TabIndex        =   14
         Top             =   315
         Width           =   1800
         Begin VB.PictureBox picBackNormal 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   480
            Left            =   105
            ScaleHeight     =   450
            ScaleWidth      =   450
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   375
            Width           =   480
         End
         Begin SmartButtonProject.SmartButton cmdChange 
            Height          =   240
            Index           =   4
            Left            =   630
            TabIndex        =   16
            Top             =   375
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   423
            Caption         =   "Change"
            Picture         =   "frmImage.frx":09D8
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
         Begin SmartButtonProject.SmartButton cmdRemoveBack 
            Height          =   240
            Left            =   630
            TabIndex        =   17
            Top             =   615
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   423
            Caption         =   "Remove"
            Picture         =   "frmImage.frx":0D72
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
      End
   End
   Begin VB.Frame frameRight 
      Caption         =   "Right Image"
      Height          =   3330
      Left            =   4905
      TabIndex        =   22
      Top             =   555
      Width           =   2235
      Begin VB.TextBox txtRightM 
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
         Left            =   1545
         TabIndex        =   45
         Text            =   "000"
         Top             =   2880
         Width           =   420
      End
      Begin VB.Frame frameRightN 
         Caption         =   "Normal"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   195
         TabIndex        =   23
         Top             =   315
         Width           =   1800
         Begin VB.PictureBox picRightNormal 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   480
            Left            =   105
            ScaleHeight     =   450
            ScaleWidth      =   450
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   375
            Width           =   480
         End
         Begin SmartButtonProject.SmartButton cmdChange 
            Height          =   240
            Index           =   2
            Left            =   630
            TabIndex        =   25
            Top             =   375
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   423
            Caption         =   "Change"
            Picture         =   "frmImage.frx":110C
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
         Begin SmartButtonProject.SmartButton cmdRemoveRight 
            Height          =   240
            Left            =   630
            TabIndex        =   26
            Top             =   615
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   423
            Caption         =   "Remove"
            Picture         =   "frmImage.frx":14A6
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
      End
      Begin VB.TextBox txtRightH 
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
         Left            =   810
         TabIndex        =   34
         Text            =   "000"
         Top             =   2880
         WhatsThisHelpID =   20340
         Width           =   420
      End
      Begin VB.TextBox txtRightW 
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
         Left            =   195
         TabIndex        =   32
         Text            =   "000"
         Top             =   2880
         Width           =   420
      End
      Begin VB.Frame frameRightO 
         Caption         =   "Mouse Over"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   195
         TabIndex        =   27
         Top             =   1470
         Width           =   1800
         Begin VB.PictureBox picRightOver 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   480
            Left            =   105
            ScaleHeight     =   450
            ScaleWidth      =   450
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   375
            Width           =   480
         End
         Begin SmartButtonProject.SmartButton cmdChange 
            Height          =   240
            Index           =   3
            Left            =   630
            TabIndex        =   29
            Top             =   375
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   423
            Caption         =   "Change"
            Picture         =   "frmImage.frx":1840
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
         Begin SmartButtonProject.SmartButton cmdSameRight 
            Height          =   240
            Left            =   630
            TabIndex        =   30
            Top             =   615
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   423
            Caption         =   "Same"
            Picture         =   "frmImage.frx":1BDA
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
      End
      Begin VB.Label lblRightM 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Margin"
         Height          =   210
         Left            =   1440
         TabIndex        =   46
         Top             =   2655
         Width           =   525
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "x"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   660
         TabIndex        =   33
         Top             =   2910
         Width           =   105
      End
      Begin VB.Label lblRightS 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Size"
         Height          =   210
         Left            =   195
         TabIndex        =   31
         Top             =   2655
         Width           =   315
      End
   End
   Begin VB.Frame frameLeft 
      Caption         =   "Left Image"
      Height          =   3330
      Left            =   120
      TabIndex        =   0
      Top             =   555
      Width           =   2235
      Begin VB.TextBox txtLeftM 
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
         Left            =   1650
         TabIndex        =   43
         Text            =   "000"
         Top             =   2880
         Width           =   420
      End
      Begin VB.TextBox txtLeftW 
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
         Left            =   195
         TabIndex        =   10
         Text            =   "000"
         Top             =   2880
         Width           =   420
      End
      Begin VB.TextBox txtLeftH 
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
         Left            =   810
         TabIndex        =   12
         Text            =   "000"
         Top             =   2880
         WhatsThisHelpID =   20340
         Width           =   420
      End
      Begin VB.Frame frameLeftO 
         Caption         =   "Mouse Over"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   195
         TabIndex        =   5
         Top             =   1470
         Width           =   1800
         Begin VB.PictureBox picLeftOver 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   480
            Left            =   105
            ScaleHeight     =   450
            ScaleWidth      =   450
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   375
            Width           =   480
         End
         Begin SmartButtonProject.SmartButton cmdChange 
            Height          =   240
            Index           =   1
            Left            =   630
            TabIndex        =   7
            Top             =   375
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   423
            Caption         =   "Change"
            Picture         =   "frmImage.frx":1F74
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
         Begin SmartButtonProject.SmartButton cmdSameLeft 
            Height          =   240
            Left            =   630
            TabIndex        =   8
            Top             =   615
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   423
            Caption         =   "Same"
            Picture         =   "frmImage.frx":230E
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
      End
      Begin VB.Frame frameLeftN 
         Caption         =   "Normal"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   195
         TabIndex        =   1
         Top             =   315
         Width           =   1800
         Begin VB.PictureBox picLeftNormal 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   480
            Left            =   105
            ScaleHeight     =   450
            ScaleWidth      =   450
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   375
            Width           =   480
         End
         Begin SmartButtonProject.SmartButton cmdChange 
            Height          =   240
            Index           =   0
            Left            =   630
            TabIndex        =   3
            Top             =   375
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   423
            Caption         =   "Change"
            Picture         =   "frmImage.frx":26A8
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
         Begin SmartButtonProject.SmartButton cmdRemoveLeft 
            Height          =   240
            Left            =   630
            TabIndex        =   4
            Top             =   615
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   423
            Caption         =   "Remove"
            Picture         =   "frmImage.frx":2A42
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
      End
      Begin VB.Label lblLeftM 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Margin"
         Height          =   210
         Left            =   1545
         TabIndex        =   44
         Top             =   2655
         Width           =   525
      End
      Begin VB.Label lblLeftS 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Size"
         Height          =   210
         Left            =   195
         TabIndex        =   9
         Top             =   2655
         Width           =   315
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "x"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   660
         TabIndex        =   11
         Top             =   2910
         Width           =   105
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
      Left            =   6240
      TabIndex        =   38
      Top             =   5415
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
      Left            =   5250
      TabIndex        =   37
      Top             =   5415
      Width           =   900
   End
End
Attribute VB_Name = "frmImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim BackCmd As MenuCmd
Dim SelId As Integer

Private WithEvents msgSubClass As xfxSC
Attribute msgSubClass.VB_VarHelpID = -1

Private Sub chkAllowCrop_Click()

    On Error Resume Next
    
    If MenuCmds(SelId).BackImage.NormalImage <> "" Then
        picSize.Picture = LoadPictureRes(MenuCmds(SelId).BackImage.NormalImage)
        MenuCmds(SelId).BackImage.w = picSize.Width / Screen.TwipsPerPixelX
        MenuCmds(SelId).BackImage.h = picSize.Height / Screen.TwipsPerPixelY
    End If
    
    MenuCmds(SelId).BackImage.AllowCrop = (chkAllowCrop.Value = vbChecked)
    UpdateSample

End Sub

Private Sub chkTile_Click()

    On Error Resume Next
    MenuCmds(SelId).BackImage.Tile = (chkTile.Value = vbChecked)
    UpdateSample

End Sub

Private Sub cmdCancel_Click()

    MenuCmds(SelId) = BackCmd
    Unload Me

End Sub

Private Sub cmdChange_Click(Index As Integer)

    Select Case Index
        Case 0
            SelImage.FileName = MenuCmds(SelId).LeftImage.NormalImage
        Case 1
            SelImage.FileName = MenuCmds(SelId).LeftImage.HoverImage
        Case 2
            SelImage.FileName = MenuCmds(SelId).RightImage.NormalImage
        Case 3
            SelImage.FileName = MenuCmds(SelId).RightImage.HoverImage
        Case 4
            SelImage.FileName = MenuCmds(SelId).BackImage.NormalImage
            SelImage.SupportsFlash = True
        Case 5
            SelImage.FileName = MenuCmds(SelId).BackImage.HoverImage
            SelImage.SupportsFlash = True
    End Select
    frmRscImages.Show vbModal
    
    With SelImage
        If .IsValid Then
            Set picSize.Picture = .Picture
            DoEvents
            Select Case Index
                Case 0
                    txtLeftW.Text = picSize.Width / Screen.TwipsPerPixelX
                    txtLeftH.Text = picSize.Height / Screen.TwipsPerPixelY
                    With MenuCmds(SelId).LeftImage
                        .NormalImage = SelImage.FileName
                        .w = picSize.Width / Screen.TwipsPerPixelX
                        .h = picSize.Height / Screen.TwipsPerPixelY
                    End With
                Case 1
                    MenuCmds(SelId).LeftImage.HoverImage = .FileName
                Case 2
                    txtRightW.Text = picSize.Width / Screen.TwipsPerPixelX
                    txtRightH.Text = picSize.Height / Screen.TwipsPerPixelY
                    With MenuCmds(SelId).RightImage
                        .NormalImage = SelImage.FileName
                        .w = picSize.Width / Screen.TwipsPerPixelX
                        .h = picSize.Height / Screen.TwipsPerPixelY
                    End With
                Case 3
                    MenuCmds(SelId).RightImage.HoverImage = .FileName
                Case 4
                    With MenuCmds(SelId).BackImage
                        .NormalImage = SelImage.FileName
                        .w = picSize.Width / Screen.TwipsPerPixelX
                        .h = picSize.Height / Screen.TwipsPerPixelY
                    End With
                Case 5
                    MenuCmds(SelId).BackImage.HoverImage = .FileName
            End Select
        End If
    End With
    
    Me.SetFocus
    UpdateSample

End Sub

Private Sub UpdateSample(Optional IsLoading As Boolean)

    Static IsBusy As Boolean
    
    If IsBusy Then Exit Sub
    IsBusy = True

    With MenuCmds(SelId)
        picLeftNormal.Picture = LoadPictureRes(.LeftImage.NormalImage)
        picLeftOver.Picture = LoadPictureRes(.LeftImage.HoverImage)
        txtLeftW.Text = .LeftImage.w
        txtLeftH.Text = .LeftImage.h
        txtLeftM.Text = .LeftImage.margin
        txtLeftW.Enabled = LenB(.LeftImage.NormalImage) <> 0
        txtLeftH.Enabled = LenB(.LeftImage.NormalImage) <> 0
        txtLeftM.Enabled = LenB(.LeftImage.NormalImage) <> 0
        
        chkTile.Value = IIf(.BackImage.Tile, vbChecked, vbUnchecked)
        chkAllowCrop.Value = IIf(.BackImage.AllowCrop, vbChecked, vbUnchecked)
        TileImage .BackImage.NormalImage, picBackNormal
        TileImage .BackImage.HoverImage, picBackOver
        
        picRightNormal.Picture = LoadPictureRes(.RightImage.NormalImage)
        picRightOver.Picture = LoadPictureRes(.RightImage.HoverImage)
        txtRightW.Text = .RightImage.w
        txtRightH.Text = .RightImage.h
        txtRightM.Text = .RightImage.margin
        txtRightW.Enabled = LenB(.RightImage.NormalImage) <> 0
        txtRightH.Enabled = LenB(.RightImage.NormalImage) <> 0
        txtRightM.Enabled = LenB(.RightImage.NormalImage) <> 0
    End With

    If Not IsLoading Then frmMain.DoLivePreview wbLivePreview
    
    IsBusy = False

End Sub

Private Sub cmdOK_Click()

    ApplyStyleOptions
    frmMain.SaveState GetLocalizedStr(189) + " " + MenuCmds(SelId).Name + " " + GetLocalizedStr(214)
    
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
        .BackImage = MenuCmds(sId).BackImage
        .LeftImage = MenuCmds(sId).LeftImage
        .RightImage = MenuCmds(sId).RightImage
    End With

End Sub

Private Sub cmdRemoveBack_Click()

    With MenuCmds(SelId).BackImage
        .NormalImage = ""
        .HoverImage = ""
    End With
    
    Me.SetFocus
    
    UpdateSample

End Sub

Private Sub cmdRemoveLeft_Click()

    With MenuCmds(SelId).LeftImage
        .NormalImage = ""
        .HoverImage = ""
    End With

    txtLeftW.Text = 0
    txtLeftH.Text = 0
    
    Me.SetFocus
    
    UpdateSample

End Sub

Private Sub cmdSameBack_Click()

    With MenuCmds(SelId).BackImage
        .HoverImage = .NormalImage
    End With
    
    Me.SetFocus
    
    UpdateSample

End Sub

Private Sub cmdSameLeft_Click()

    With MenuCmds(SelId).LeftImage
        .HoverImage = .NormalImage
    End With
    
    Me.SetFocus
    
    UpdateSample

End Sub

Private Sub cmdRemoveRight_Click()

    With MenuCmds(SelId).RightImage
        .NormalImage = ""
        .HoverImage = ""
    End With
    txtRightW.Text = 0
    txtRightH.Text = 0
    
    Me.SetFocus
    
    UpdateSample

End Sub

Private Sub cmdSameRight_Click()

    With MenuCmds(SelId).RightImage
        .HoverImage = .NormalImage
    End With
    
    Me.SetFocus
    
    UpdateSample

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF1 Then showHelp "dialogs/command_image.htm"

End Sub

Private Sub Form_Load()

    CenterForm Me
    LocalizeUI
    
    Set msgSubClass = New xfxSC
    msgSubClass.SubClassHwnd Me.hwnd, True
    
    SelId = GetID
    BackCmd = MenuCmds(SelId)
    
    caption = NiceGrpCaption(MenuCmds(SelId).parent) + "/" + NiceCmdCaption(SelId) + " - " + GetLocalizedStr(214)
    FixCtrls4Skin Me
    
    UpdateSample True
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    msgSubClass.SubClassHwnd Me.hwnd, False

End Sub

Private Sub picBackNormal_DblClick()

    cmdChange_Click 4

End Sub

Private Sub picBackOver_DblClick()

    cmdChange_Click 5

End Sub

Private Sub picLeftNormal_DblClick()

    cmdChange_Click 0

End Sub

Private Sub picLeftOver_DblClick()

    cmdChange_Click 1

End Sub

Private Sub picRightNormal_DblClick()

    cmdChange_Click 2

End Sub

Private Sub picRightOver_DblClick()

    cmdChange_Click 3

End Sub

Private Sub sbApplyOptions_Click()

    PopupMenu frmMain.mnuStyleOptions, , sbApplyOptions.Left, sbApplyOptions.Top + sbApplyOptions.Height

End Sub

Private Sub txtLeftH_KeyPress(KeyAscii As Integer)

    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If

End Sub

Private Sub txtLeftM_Change()

    On Error Resume Next
    MenuCmds(SelId).LeftImage.margin = Val(txtLeftM.Text)
    UpdateSample

End Sub

Private Sub txtLeftM_GotFocus()

    SelAll txtLeftM

End Sub

Private Sub txtLeftW_Change()

    On Error Resume Next
    MenuCmds(SelId).LeftImage.w = Val(txtLeftW.Text)
    UpdateSample

End Sub

Private Sub txtLeftW_GotFocus()

    SelAll txtLeftW

End Sub

Private Sub txtLeftW_KeyPress(KeyAscii As Integer)

    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If

End Sub

Private Sub txtRightH_KeyPress(KeyAscii As Integer)

    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If

End Sub

Private Sub txtRightM_Change()

    On Error Resume Next
    MenuCmds(SelId).RightImage.margin = Val(txtRightM.Text)
    UpdateSample

End Sub

Private Sub txtRightM_GotFocus()

    SelAll txtRightM

End Sub

Private Sub txtRightW_Change()

    On Error Resume Next
    MenuCmds(SelId).RightImage.w = Val(txtRightW.Text)
    UpdateSample

End Sub

Private Sub txtRightW_GotFocus()

    SelAll txtRightW

End Sub

Private Sub txtLeftH_Change()

    On Error Resume Next
    MenuCmds(SelId).LeftImage.h = Val(txtLeftH.Text)
    UpdateSample

End Sub

Private Sub txtLeftH_GotFocus()

    SelAll txtLeftH

End Sub

Private Sub txtRightH_Change()

    On Error Resume Next
    MenuCmds(SelId).RightImage.h = Val(txtRightH.Text)
    UpdateSample

End Sub

Private Sub txtRightH_GotFocus()

    SelAll txtRightH

End Sub

Private Sub txtRightW_KeyPress(KeyAscii As Integer)

    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If

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

    Dim i As Integer

    frameLeft.caption = GetLocalizedStr(198)
    frameBack.caption = GetLocalizedStr(199)
    frameRight.caption = GetLocalizedStr(200)
    
    frameLeftN.caption = GetLocalizedStr(179)
    frameLeftO.caption = GetLocalizedStr(180)
    frameBackN.caption = GetLocalizedStr(179)
    frameBackO.caption = GetLocalizedStr(180)
    frameRightN.caption = GetLocalizedStr(179)
    frameRightO.caption = GetLocalizedStr(180)

    For i = 0 To cmdChange.Count - 1
        cmdChange(i).caption = GetLocalizedStr(189)
    Next i
    
    cmdRemoveLeft.caption = GetLocalizedStr(201)
    cmdRemoveBack.caption = GetLocalizedStr(201)
    cmdRemoveRight.caption = GetLocalizedStr(201)
    
    cmdSameLeft.caption = GetLocalizedStr(202)
    cmdSameBack.caption = GetLocalizedStr(202)
    cmdSameRight.caption = GetLocalizedStr(202)
    
    lblLeftS.caption = GetLocalizedStr(203)
    lblLeftM.caption = GetLocalizedStr(988)
    lblRightS.caption = GetLocalizedStr(203)
    lblRightM.caption = GetLocalizedStr(988)
    
    frmLiveSample.caption = GetLocalizedStr(188)
    
    cmdOK.caption = GetLocalizedStr(186)
    cmdCancel.caption = GetLocalizedStr(187)
    
    If Preferences.language <> "eng" Then
        cmdOK.Width = SetCtrlWidth(cmdOK)
        cmdCancel.Width = SetCtrlWidth(cmdCancel)
    End If

End Sub
