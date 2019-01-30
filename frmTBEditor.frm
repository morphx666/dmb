VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{E1C6DB9D-BD4A-4A61-A759-0CED75D034BF}#43.0#0"; "SmartButton.ocx"
Object = "{DB06EC30-01E1-485F-A3C7-CE80CA0D7D37}#2.0#0"; "xFXSlider.ocx"
Object = "{9A2947B4-87EE-40EE-A3EF-32BDC32D5726}#1.0#0"; "xfxline3d.ocx"
Begin VB.Form frmTBEditor 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Toolbars Editor"
   ClientHeight    =   10470
   ClientLeft      =   4305
   ClientTop       =   3090
   ClientWidth     =   13380
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTBEditor.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10470
   ScaleWidth      =   13380
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrUpdateSelTB 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   2520
      Top             =   6300
   End
   Begin VB.CommandButton cmdCompile 
      Caption         =   "Compile"
      Height          =   375
      Left            =   1215
      TabIndex        =   113
      Top             =   6120
      Width           =   1080
   End
   Begin VB.PictureBox picTBPositioning 
      Height          =   4545
      Left            =   6915
      ScaleHeight     =   4485
      ScaleWidth      =   6270
      TabIndex        =   66
      Top             =   4980
      Width           =   6330
      Begin xfxLine3D.ucLine3D uc3DLine7 
         Height          =   30
         Left            =   1545
         TabIndex        =   88
         Top             =   1410
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   53
      End
      Begin xfxLine3D.ucLine3D uc3DLine6 
         Height          =   30
         Left            =   1545
         TabIndex        =   87
         Top             =   900
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   53
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   2970
         Left            =   1575
         TabIndex        =   68
         Top             =   90
         Width           =   4380
         Begin VB.OptionButton opAlignment 
            Caption         =   "Free Flow"
            Enabled         =   0   'False
            Height          =   225
            Index           =   11
            Left            =   0
            TabIndex        =   114
            Top             =   2655
            Width           =   4320
         End
         Begin VB.CheckBox chkSyncObjSize 
            Caption         =   "Keep Object's size synchronized"
            Height          =   210
            Left            =   1410
            TabIndex        =   112
            Top             =   2370
            Width           =   3225
         End
         Begin SmartButtonProject.SmartButton cmdSelRefImage 
            Height          =   315
            Left            =   3465
            TabIndex        =   110
            Top             =   1635
            Width           =   360
            _ExtentX        =   635
            _ExtentY        =   556
            Picture         =   "frmTBEditor.frx":014A
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.TextBox txtAttachedTo 
            Height          =   285
            Left            =   1410
            TabIndex        =   84
            Top             =   1650
            Width           =   2025
         End
         Begin VB.OptionButton opAlignment 
            Caption         =   "Attached To"
            Enabled         =   0   'False
            Height          =   225
            Index           =   10
            Left            =   0
            TabIndex        =   82
            Top             =   1410
            Width           =   1275
         End
         Begin VB.TextBox txtACY 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   1935
            TabIndex        =   79
            Text            =   "000"
            Top             =   930
            Width           =   420
         End
         Begin VB.TextBox txtACX 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   1410
            TabIndex        =   78
            Text            =   "000"
            Top             =   930
            Width           =   420
         End
         Begin VB.OptionButton opAlignment 
            Caption         =   "Custom"
            Enabled         =   0   'False
            Height          =   240
            Index           =   9
            Left            =   0
            TabIndex        =   80
            Top             =   960
            Width           =   1965
         End
         Begin VB.OptionButton opAlignment 
            Enabled         =   0   'False
            Height          =   225
            Index           =   8
            Left            =   525
            TabIndex        =   77
            Top             =   510
            Width           =   225
         End
         Begin VB.OptionButton opAlignment 
            Enabled         =   0   'False
            Height          =   225
            Index           =   7
            Left            =   255
            TabIndex        =   76
            Top             =   510
            Width           =   225
         End
         Begin VB.OptionButton opAlignment 
            Enabled         =   0   'False
            Height          =   225
            Index           =   6
            Left            =   0
            TabIndex        =   75
            Top             =   510
            Width           =   225
         End
         Begin VB.OptionButton opAlignment 
            Enabled         =   0   'False
            Height          =   225
            Index           =   5
            Left            =   525
            TabIndex        =   74
            Top             =   270
            Width           =   225
         End
         Begin VB.OptionButton opAlignment 
            Enabled         =   0   'False
            Height          =   225
            Index           =   4
            Left            =   255
            TabIndex        =   73
            Top             =   270
            Width           =   225
         End
         Begin VB.OptionButton opAlignment 
            Enabled         =   0   'False
            Height          =   225
            Index           =   3
            Left            =   0
            TabIndex        =   72
            Top             =   270
            Width           =   225
         End
         Begin VB.OptionButton opAlignment 
            Enabled         =   0   'False
            Height          =   225
            Index           =   2
            Left            =   525
            TabIndex        =   71
            Top             =   30
            Width           =   225
         End
         Begin VB.OptionButton opAlignment 
            Enabled         =   0   'False
            Height          =   225
            Index           =   1
            Left            =   255
            TabIndex        =   70
            Top             =   30
            Width           =   225
         End
         Begin VB.OptionButton opAlignment 
            Enabled         =   0   'False
            Height          =   225
            Index           =   0
            Left            =   0
            TabIndex        =   69
            Top             =   30
            Value           =   -1  'True
            Width           =   225
         End
         Begin MSComctlLib.ImageCombo icmbAlignment 
            Height          =   330
            Left            =   1410
            TabIndex        =   86
            Top             =   1980
            Width           =   2025
            _ExtentX        =   3572
            _ExtentY        =   582
            _Version        =   393216
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Locked          =   -1  'True
         End
         Begin VB.Label lblTBPosInfo 
            BackStyle       =   0  'Transparent
            Caption         =   "Info"
            ForeColor       =   &H80000011&
            Height          =   660
            Left            =   870
            TabIndex        =   111
            Top             =   45
            Width           =   3420
         End
         Begin VB.Label lblAlignment 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Alignment"
            Height          =   195
            Left            =   645
            TabIndex        =   85
            Top             =   2055
            Width           =   705
         End
         Begin VB.Label lblObjName 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Object Name"
            Height          =   195
            Left            =   420
            TabIndex        =   83
            Top             =   1695
            Width           =   930
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ","
            Height          =   195
            Left            =   1845
            TabIndex        =   81
            Top             =   1050
            Width           =   60
         End
      End
      Begin VB.TextBox txtMarginH 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1575
         TabIndex        =   93
         Text            =   "888"
         Top             =   3840
         Width           =   480
      End
      Begin VB.TextBox txtMarginV 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1575
         TabIndex        =   96
         Text            =   "888"
         Top             =   4185
         Width           =   480
      End
      Begin VB.ComboBox cmbSpanning 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmTBEditor.frx":02A4
         Left            =   1575
         List            =   "frmTBEditor.frx":02AE
         Style           =   2  'Dropdown List
         TabIndex        =   91
         Top             =   3255
         Width           =   1620
      End
      Begin xfxLine3D.ucLine3D uc3DLine4 
         Height          =   30
         Left            =   0
         TabIndex        =   92
         Top             =   3720
         Width           =   6270
         _ExtentX        =   11060
         _ExtentY        =   53
      End
      Begin xfxLine3D.ucLine3D uc3DLine3 
         Height          =   30
         Left            =   0
         TabIndex        =   89
         Top             =   3090
         Width           =   6240
         _ExtentX        =   11007
         _ExtentY        =   53
      End
      Begin VB.Label lblTBOffsetH 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pixels Horizontally"
         Height          =   195
         Left            =   2115
         TabIndex        =   94
         Top             =   3885
         Width           =   1290
      End
      Begin VB.Label lblTBOffsetV 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pixels Vertically"
         Height          =   195
         Left            =   2115
         TabIndex        =   97
         Top             =   4230
         Width           =   1095
      End
      Begin VB.Label lblTBOffset 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Offset"
         Height          =   195
         Left            =   990
         TabIndex        =   95
         Top             =   4035
         Width           =   465
      End
      Begin VB.Label lblTBSpanning 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Spanning"
         Height          =   195
         Left            =   795
         TabIndex        =   90
         Top             =   3315
         Width           =   660
      End
      Begin VB.Label lblTBAlignment 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alignment"
         Height          =   195
         Left            =   750
         TabIndex        =   67
         Top             =   345
         Width           =   705
      End
   End
   Begin VB.PictureBox picTBFX 
      Height          =   2205
      Left            =   4770
      ScaleHeight     =   2145
      ScaleWidth      =   6270
      TabIndex        =   98
      Top             =   6390
      Width           =   6330
      Begin xFXSlider.ucSlider sldTransparency 
         Height          =   270
         Left            =   930
         TabIndex        =   106
         Top             =   1365
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
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomKnobImage =   "frmTBEditor.frx":02C5
         CustomSelKnobImage=   "frmTBEditor.frx":059F
      End
      Begin SmartButtonProject.SmartButton cmdShadowColor 
         Height          =   240
         Left            =   5250
         TabIndex        =   102
         Top             =   450
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
         Left            =   930
         TabIndex        =   101
         Top             =   435
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
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomKnobImage =   "frmTBEditor.frx":0879
         CustomSelKnobImage=   "frmTBEditor.frx":0B53
      End
      Begin VB.Label lblTInvisible 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Invisible"
         Height          =   195
         Left            =   4755
         TabIndex        =   108
         Top             =   1680
         Width           =   585
      End
      Begin VB.Label lblTOFF 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "OFF"
         Height          =   195
         Left            =   960
         TabIndex        =   107
         Top             =   1680
         Width           =   300
      End
      Begin VB.Label lblDSDarker 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Larger"
         Height          =   195
         Left            =   4785
         TabIndex        =   104
         Top             =   750
         Width           =   465
      End
      Begin VB.Label lblDSOFF 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "OFF"
         Height          =   195
         Left            =   975
         TabIndex        =   103
         Top             =   750
         Width           =   300
      End
      Begin VB.Image Image4 
         Height          =   240
         Left            =   675
         Picture         =   "frmTBEditor.frx":0E2D
         Top             =   1380
         Width           =   240
      End
      Begin VB.Image Image5 
         Height          =   240
         Left            =   675
         Picture         =   "frmTBEditor.frx":0F77
         Top             =   450
         Width           =   240
      End
      Begin VB.Label lblTransparency 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transparency"
         Height          =   195
         Left            =   2535
         TabIndex        =   105
         Top             =   1140
         Width           =   990
      End
      Begin VB.Label lblDropShadow 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Drop Shadow Size"
         Height          =   195
         Left            =   2385
         TabIndex        =   99
         Top             =   210
         Width           =   1290
      End
      Begin VB.Label lblShadowColor 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Color"
         Height          =   195
         Left            =   5205
         TabIndex        =   100
         Top             =   210
         Width           =   375
      End
   End
   Begin VB.PictureBox picTBAdvanced 
      Height          =   4545
      Left            =   270
      ScaleHeight     =   4485
      ScaleWidth      =   6270
      TabIndex        =   42
      Top             =   6855
      Width           =   6330
      Begin VB.CheckBox chkSmartScrolling 
         Caption         =   "Enable Smart Scrolling"
         Height          =   195
         Left            =   285
         TabIndex        =   115
         Top             =   900
         Width           =   2265
      End
      Begin VB.TextBox txtVisCond 
         BeginProperty Font 
            Name            =   "Monotype.com"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   0
         TabIndex        =   62
         Top             =   3330
         Width           =   5205
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   795
         Left            =   2670
         ScaleHeight     =   795
         ScaleWidth      =   2520
         TabIndex        =   53
         Top             =   1740
         Width           =   2520
         Begin VB.TextBox txtTBHeight 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   900
            TabIndex        =   57
            Text            =   "0"
            Top             =   495
            Width           =   405
         End
         Begin VB.OptionButton opTBHeight 
            Caption         =   "Manual"
            Height          =   225
            Index           =   1
            Left            =   615
            TabIndex        =   55
            Top             =   225
            Width           =   1620
         End
         Begin VB.OptionButton opTBHeight 
            Caption         =   "Auto"
            Height          =   225
            Index           =   0
            Left            =   615
            TabIndex        =   54
            Top             =   0
            Width           =   1620
         End
         Begin SmartButtonProject.SmartButton cmdTBAutoHeight 
            Height          =   300
            Left            =   1350
            TabIndex        =   58
            Top             =   495
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   529
            Caption         =   "Calculate"
            Picture         =   "frmTBEditor.frx":10C1
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
            Height          =   195
            Left            =   75
            TabIndex        =   56
            Top             =   360
            Width           =   705
         End
         Begin VB.Image imgGSHeight 
            Appearance      =   0  'Flat
            Height          =   240
            Left            =   180
            Picture         =   "frmTBEditor.frx":145B
            Top             =   135
            Width           =   240
         End
      End
      Begin VB.TextBox txtTBWidth 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   915
         TabIndex        =   59
         Text            =   "0"
         Top             =   2460
         Width           =   405
      End
      Begin VB.OptionButton opTBWidth 
         Caption         =   "Manual"
         Height          =   225
         Index           =   2
         Left            =   630
         TabIndex        =   52
         Top             =   2205
         Width           =   1935
      End
      Begin VB.OptionButton opTBWidth 
         Caption         =   "Match Group's Width"
         Height          =   210
         Index           =   1
         Left            =   630
         TabIndex        =   50
         Top             =   1980
         Width           =   1860
      End
      Begin VB.OptionButton opTBWidth 
         Caption         =   "Auto"
         Height          =   225
         Index           =   0
         Left            =   630
         TabIndex        =   49
         Top             =   1740
         Width           =   720
      End
      Begin VB.CheckBox chkFollowScrolling 
         Caption         =   "Follow Scrolling"
         Enabled         =   0   'False
         Height          =   345
         Left            =   0
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   180
         WhatsThisHelpID =   20410
         Width           =   4560
      End
      Begin VB.CheckBox chkHorizontal 
         Enabled         =   0   'False
         Height          =   210
         Left            =   285
         TabIndex        =   44
         Top             =   585
         Width           =   195
      End
      Begin VB.CheckBox chkVertical 
         Enabled         =   0   'False
         Height          =   210
         Left            =   1155
         TabIndex        =   46
         Top             =   585
         Width           =   195
      End
      Begin MSComCtl2.UpDown udV 
         Height          =   360
         Left            =   1395
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   510
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Enabled         =   0   'False
      End
      Begin MSComCtl2.UpDown udH 
         Height          =   240
         Left            =   525
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   570
         Width           =   390
         _ExtentX        =   688
         _ExtentY        =   423
         _Version        =   393216
         Orientation     =   1
         Enabled         =   0   'False
      End
      Begin SmartButtonProject.SmartButton cmdTBAutoWidth 
         Height          =   300
         Left            =   1365
         TabIndex        =   60
         Top             =   2460
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   529
         Caption         =   "Calculate"
         Picture         =   "frmTBEditor.frx":17E5
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
      Begin xfxLine3D.ucLine3D ucLine3D1 
         Height          =   30
         Left            =   0
         TabIndex        =   116
         Top             =   1305
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   53
      End
      Begin xfxLine3D.ucLine3D ucLine3D2 
         Height          =   30
         Left            =   0
         TabIndex        =   117
         Top             =   2940
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   53
      End
      Begin VB.Label lblVisCond 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Visibility Condition"
         Height          =   195
         Left            =   0
         TabIndex        =   61
         Top             =   3090
         Width           =   1275
      End
      Begin VB.Image imgGSWidth 
         Appearance      =   0  'Flat
         Height          =   240
         Left            =   165
         Picture         =   "frmTBEditor.frx":1B7F
         Top             =   1875
         Width           =   240
      End
      Begin VB.Label lblGSWidth 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Width"
         Height          =   195
         Left            =   75
         TabIndex        =   51
         Top             =   2115
         Width           =   420
      End
      Begin VB.Label lblTBSize 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Toolbar Size"
         Height          =   210
         Left            =   0
         TabIndex        =   48
         Top             =   1470
         Width           =   990
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   6105
      TabIndex        =   3
      Top             =   495
      Width           =   90
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   60
         X2              =   60
         Y1              =   0
         Y2              =   465
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         X1              =   45
         X2              =   45
         Y1              =   0
         Y2              =   465
      End
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "Preview..."
      Height          =   375
      Left            =   45
      TabIndex        =   63
      Top             =   6120
      Width           =   1080
   End
   Begin VB.TextBox txtTBName 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1545
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   210
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.PictureBox picTBAppearance 
      AutoRedraw      =   -1  'True
      Height          =   4545
      Left            =   6915
      ScaleHeight     =   4485
      ScaleWidth      =   6270
      TabIndex        =   5
      Top             =   105
      Width           =   6330
      Begin VB.TextBox txtRadiusTR 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   5535
         TabIndex        =   121
         Text            =   "123"
         Top             =   285
         Width           =   420
      End
      Begin VB.TextBox txtRadiusTL 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4860
         TabIndex        =   120
         Text            =   "123"
         Top             =   285
         Width           =   420
      End
      Begin VB.TextBox txtRadiusBR 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   5535
         TabIndex        =   119
         Text            =   "123"
         Top             =   900
         Width           =   420
      End
      Begin VB.TextBox txtRadiusBL 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4860
         TabIndex        =   118
         Text            =   "123"
         Top             =   900
         Width           =   420
      End
      Begin VB.ComboBox cmbFX 
         Height          =   315
         ItemData        =   "frmTBEditor.frx":1F09
         Left            =   1905
         List            =   "frmTBEditor.frx":1F16
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   990
         Width           =   1305
      End
      Begin VB.TextBox txtTBCMarginH 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2205
         TabIndex        =   19
         Text            =   "123"
         Top             =   1875
         Width           =   420
      End
      Begin VB.TextBox txtTBCMarginV 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3495
         TabIndex        =   17
         Text            =   "123"
         Top             =   1860
         Width           =   420
      End
      Begin VB.TextBox txtTBSeparation 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1905
         TabIndex        =   22
         Text            =   "888"
         Top             =   2325
         Width           =   420
      End
      Begin VB.PictureBox picTBImage 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   3015
         ScaleHeight     =   450
         ScaleWidth      =   450
         TabIndex        =   31
         Top             =   3660
         Width           =   480
      End
      Begin VB.ComboBox cmbBorder 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmTBEditor.frx":1F32
         Left            =   1905
         List            =   "frmTBEditor.frx":1F34
         TabIndex        =   9
         Text            =   "cmbBorder"
         Top             =   555
         Width           =   1305
      End
      Begin VB.CheckBox chkJustify 
         Height          =   210
         Left            =   1905
         TabIndex        =   24
         Top             =   2775
         Width           =   195
      End
      Begin VB.ComboBox cmbStyle 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmTBEditor.frx":1F36
         Left            =   1905
         List            =   "frmTBEditor.frx":1F40
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   120
         Width           =   1665
      End
      Begin MSComCtl2.UpDown udSep 
         Height          =   285
         Left            =   2340
         TabIndex        =   23
         Top             =   2325
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         BuddyControl    =   "txtTBSeparation"
         BuddyDispid     =   196666
         OrigLeft        =   1785
         OrigTop         =   1125
         OrigRight       =   2025
         OrigBottom      =   1440
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin SmartButtonProject.SmartButton cmdTBBackColor 
         Height          =   240
         Left            =   1905
         TabIndex        =   28
         Top             =   3300
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
      Begin SmartButtonProject.SmartButton cmdChange 
         Height          =   240
         Left            =   1905
         TabIndex        =   29
         Top             =   3660
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   423
         Caption         =   "Change"
         Picture         =   "frmTBEditor.frx":1F5A
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
      Begin SmartButtonProject.SmartButton cmdRemove 
         Height          =   240
         Left            =   1905
         TabIndex        =   32
         Top             =   3900
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   423
         Caption         =   "Remove"
         Picture         =   "frmTBEditor.frx":22F4
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
      Begin SmartButtonProject.SmartButton cmdTBBorderColor 
         Height          =   240
         Left            =   3270
         TabIndex        =   10
         Top             =   585
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
      Begin MSComCtl2.UpDown udTBCMV 
         Height          =   285
         Left            =   3930
         TabIndex        =   18
         Top             =   1860
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         BuddyControl    =   "txtTBCMarginV"
         BuddyDispid     =   196665
         OrigLeft        =   2520
         OrigTop         =   630
         OrigRight       =   2715
         OrigBottom      =   885
         Max             =   20
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udTBCMH 
         Height          =   285
         Left            =   2625
         TabIndex        =   20
         Top             =   1875
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         BuddyControl    =   "txtTBCMarginH"
         BuddyDispid     =   196664
         OrigLeft        =   1260
         OrigTop         =   600
         OrigRight       =   1455
         OrigBottom      =   900
         Max             =   20
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin xfxLine3D.ucLine3D uc3DLine1 
         Height          =   30
         Left            =   30
         TabIndex        =   13
         Top             =   1515
         Width           =   6195
         _ExtentX        =   10927
         _ExtentY        =   53
      End
      Begin xfxLine3D.ucLine3D uc3DLine2 
         Height          =   30
         Left            =   0
         TabIndex        =   26
         Top             =   3135
         Width           =   6195
         _ExtentX        =   10927
         _ExtentY        =   53
      End
      Begin VB.Label lblBorderRadius 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Radius"
         Height          =   195
         Left            =   4200
         TabIndex        =   122
         Top             =   615
         Width           =   480
      End
      Begin VB.Label lblBorderStyle 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Border Style"
         Height          =   195
         Left            =   30
         TabIndex        =   11
         Top             =   1050
         Width           =   1800
      End
      Begin VB.Label lblJustify 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Justify HotSpots"
         Height          =   390
         Left            =   30
         TabIndex        =   25
         Top             =   2685
         Width           =   1800
      End
      Begin VB.Label lblTBMargins 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Margins"
         Height          =   195
         Left            =   30
         TabIndex        =   16
         Top             =   1920
         Width           =   1800
      End
      Begin VB.Label lblH 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Horizontal"
         Height          =   195
         Left            =   1905
         TabIndex        =   14
         Top             =   1650
         Width           =   720
      End
      Begin VB.Label lblV 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vertical"
         Height          =   195
         Left            =   3195
         TabIndex        =   15
         Top             =   1650
         Width           =   525
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   1905
         Picture         =   "frmTBEditor.frx":268E
         Top             =   1890
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   240
         Left            =   3195
         Picture         =   "frmTBEditor.frx":2A18
         Top             =   1890
         Width           =   240
      End
      Begin VB.Label lblTBSeparation 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Separation"
         Height          =   195
         Left            =   30
         TabIndex        =   21
         Top             =   2355
         Width           =   1800
      End
      Begin VB.Label lblTBBorder 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Border"
         Height          =   195
         Left            =   30
         TabIndex        =   8
         Top             =   615
         Width           =   1800
      End
      Begin VB.Label lblTBBackImage 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Back Image"
         Height          =   195
         Left            =   30
         TabIndex        =   30
         Top             =   3780
         Width           =   1800
      End
      Begin VB.Label lblTBStyle 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Toolbar Style"
         Height          =   195
         Left            =   30
         TabIndex        =   6
         Top             =   180
         Width           =   1800
      End
      Begin VB.Label lblTBBackColor 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Back Color"
         Height          =   195
         Left            =   30
         TabIndex        =   27
         Top             =   3315
         Width           =   1800
      End
   End
   Begin SmartButtonProject.SmartButton cmdAddTB 
      Height          =   360
      Left            =   5220
      TabIndex        =   1
      Top             =   495
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   635
      Picture         =   "frmTBEditor.frx":2DA2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin SmartButtonProject.SmartButton cmdDelTB 
      Height          =   360
      Left            =   5670
      TabIndex        =   2
      Top             =   495
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   635
      Picture         =   "frmTBEditor.frx":2EFC
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin SmartButtonProject.SmartButton cmdRenameTB 
      Height          =   360
      Left            =   6225
      TabIndex        =   4
      Top             =   495
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   635
      Picture         =   "frmTBEditor.frx":3056
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4845
      TabIndex        =   64
      Top             =   6120
      Width           =   900
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5880
      TabIndex        =   65
      Top             =   6120
      Width           =   900
   End
   Begin VB.PictureBox picTBGeneral 
      Height          =   4545
      Left            =   225
      ScaleHeight     =   4485
      ScaleWidth      =   6270
      TabIndex        =   34
      Top             =   960
      Width           =   6330
      Begin xfxLine3D.ucLine3D uc3DLine5 
         Height          =   30
         Left            =   5940
         TabIndex        =   39
         Top             =   1185
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   53
      End
      Begin SmartButtonProject.SmartButton cmdUp 
         Height          =   360
         Left            =   5940
         TabIndex        =   40
         Top             =   1275
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   635
         Picture         =   "frmTBEditor.frx":33F0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin SmartButtonProject.SmartButton cmdDown 
         Height          =   360
         Left            =   5940
         TabIndex        =   41
         Top             =   1650
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   635
         Picture         =   "frmTBEditor.frx":354A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin SmartButtonProject.SmartButton cmdDelGroups 
         Height          =   360
         Left            =   5940
         TabIndex        =   38
         Top             =   765
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   635
         Picture         =   "frmTBEditor.frx":36A4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin SmartButtonProject.SmartButton cmdAddGroup 
         Height          =   360
         Left            =   5940
         TabIndex        =   37
         Top             =   390
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   635
         Picture         =   "frmTBEditor.frx":3C3E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComctlLib.ListView lvGroups 
         Height          =   4065
         Left            =   30
         TabIndex        =   36
         Top             =   390
         Width           =   5865
         _ExtentX        =   10345
         _ExtentY        =   7170
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         Enabled         =   0   'False
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
            Text            =   "Name"
            Object.Width           =   4815
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Caption"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label lblTBSelGrps 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select the groups to include in this Toolbar"
         Height          =   195
         Left            =   45
         TabIndex        =   35
         Top             =   150
         Width           =   3045
      End
   End
   Begin MSComctlLib.TabStrip tsToolbar 
      Height          =   5025
      Left            =   150
      TabIndex        =   33
      Top             =   885
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   8864
      TabWidthStyle   =   1
      Placement       =   1
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   5
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "General"
            Key             =   "tsGeneral"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Appearance"
            Key             =   "tsAppearance"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Positioning"
            Key             =   "tsPositioning"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Effects"
            Key             =   "tsEffects"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Advanced"
            Key             =   "tsAdvanced"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.TabStrip tsToolbars 
      Height          =   5910
      Left            =   45
      TabIndex        =   109
      Top             =   135
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   10425
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuGroups 
      Caption         =   "mnuGroups"
      Begin VB.Menu mnuGrp 
         Caption         =   ""
         Index           =   0
      End
      Begin VB.Menu mnuGrp 
         Caption         =   "-"
         Index           =   1
      End
   End
End
Attribute VB_Name = "frmTBEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ProjectBack As ProjectDef
Dim GrpsBack() As MenuGrp
Dim CmdsBack() As MenuCmd
Dim IsUpdating As Boolean
Dim CurSelTab As String
Dim lastSelTBIdx As Integer

Private WithEvents msgSubClass As xfxSC
Attribute msgSubClass.VB_VarHelpID = -1

Private Sub UpdateToolbarControls()

    Dim i As Integer
    Dim IsOn As Boolean
    
    IsOn = tsToolbars.Tabs.Count > 0

    For i = 0 To opAlignment.Count - 1
        opAlignment(i).Enabled = IsOn
    Next i
    chkFollowScrolling.Enabled = IsOn
    chkHorizontal.Enabled = IsOn And (chkFollowScrolling.Value = vbChecked)
    chkVertical.Enabled = IsOn And (chkFollowScrolling.Value = vbChecked)
    chkSmartScrolling.Enabled = IsOn And (chkFollowScrolling.Value = vbChecked)
    txtMarginH.Enabled = IsOn
    txtMarginV.Enabled = IsOn
    udH.Enabled = IsOn And (chkFollowScrolling.Value = vbChecked) And (chkHorizontal.Value = vbChecked)
    udV.Enabled = IsOn And (chkFollowScrolling.Value = vbChecked) And (chkVertical.Value = vbChecked)
    txtMarginH.Enabled = IsOn
    txtMarginV.Enabled = IsOn
    cmbStyle.Enabled = IsOn
    cmbSpanning.Enabled = IsOn
    cmbBorder.Enabled = IsOn
    cmbFX.Enabled = IsOn
    chkJustify.Enabled = IsOn
    cmdTBBackColor.Enabled = IsOn
    cmdTBBorderColor.Enabled = IsOn
    txtACX.Enabled = (opAlignment(9).Value = True) And IsOn
    txtACY.Enabled = (opAlignment(9).Value = True) And IsOn
    txtAttachedTo.Enabled = (opAlignment(10).Value = True) And IsOn
    cmdSelRefImage.Enabled = txtAttachedTo.Enabled
    icmbAlignment.Enabled = (opAlignment(10).Value = True) And IsOn
    chkSyncObjSize.Enabled = icmbAlignment.Enabled And (icmbAlignment.SelectedItem.tag = 0) And IsOn
    lvGroups.Enabled = IsOn And (lvGroups.ListItems.Count > 0)
    cmdChange.Enabled = IsOn
    cmdRemove.Enabled = IsOn
    txtTBSeparation.Enabled = IsOn
    udSep.Enabled = IsOn
    txtTBCMarginH.Enabled = IsOn
    txtTBCMarginV.Enabled = IsOn
    udTBCMH.Enabled = IsOn
    udTBCMV.Enabled = IsOn
    sldDropShadow.Enabled = IsOn
    cmdShadowColor.Enabled = sldDropShadow.Value > 0 And IsOn
    sldTransparency.Enabled = IsOn
    txtRadiusTL.Enabled = IsOn
    txtRadiusTR.Enabled = IsOn
    txtRadiusBL.Enabled = IsOn
    txtRadiusBR.Enabled = IsOn
    
    SetButtonsState
    
    opTBWidth(0).Enabled = IsOn
    opTBWidth(1).Enabled = IsOn
    opTBWidth(2).Enabled = IsOn
    txtTBWidth.Enabled = IsOn And opTBWidth(2).Value
    cmdTBAutoWidth.Enabled = txtTBWidth.Enabled
    
    opTBHeight(0).Enabled = IsOn
    opTBHeight(1).Enabled = IsOn
    txtTBHeight.Enabled = IsOn And opTBHeight(1).Value
    cmdTBAutoHeight.Enabled = txtTBHeight.Enabled
    
    If tsToolbars.SelectedItem Is Nothing Then
        opAlignment(11).caption = "Free Flow"
    Else
        opAlignment(11).caption = "Free Flow [dmbTB" + CStr(tsToolbars.SelectedItem.Index) + "ph]"
    End If
    opAlignment(11).Width = GetTextSize(opAlignment(11).caption + "XXXXXXXX")(1) * Screen.TwipsPerPixelX
    
    If Project.ToolBar.IsTemplate Then
        txtVisCond.Enabled = False
        txtVisCond.BackColor = &H80000000
    Else
        txtVisCond.Enabled = True
        txtVisCond.BackColor = &H80000005
    End If
    
End Sub

Private Sub ShowTBPosInfo()

    Dim i As Integer

    For i = 0 To opAlignment.Count - 1
        If opAlignment(i).Value Then
            Select Case i
                Case 0: lblTBPosInfo.caption = "Aligned to the Left and at the Top of the page"
                Case 1: lblTBPosInfo.caption = "Center aligned and at the Top of the page"
                Case 2: lblTBPosInfo.caption = "Aligned to the Right and at the Top of the Page"
                
                Case 3: lblTBPosInfo.caption = "Aligned to the Left and at the Center of the page"
                Case 4: lblTBPosInfo.caption = "Center aligned and at the Center of the page"
                Case 5: lblTBPosInfo.caption = "Aligned to the Right and at the Center of the page"
                
                Case 6: lblTBPosInfo.caption = "Aligned to the Left and at the Bottom of the page"
                Case 7: lblTBPosInfo.caption = "Center aligned and at the Bottom of the page"
                Case 8: lblTBPosInfo.caption = "Aligned to the Right and at the Bottom of the page"
                
                Case 9: lblTBPosInfo.caption = txtACX.Text + " pixels from the left and " + txtACY.Text + " pixels from the top"
                Case 10: lblTBPosInfo.caption = "Attached to the '" + txtAttachedTo.Text + "' image and aligned to the " + _
                                                Split(icmbAlignment.Text, "/")(0) + " and the " + Split(icmbAlignment.Text, "/")(1) + " of the image"
                Case 11: lblTBPosInfo.caption = "The toolbar will flow with the rest of the HTML contents and will be placed inside the element whose ID is dmbTB" + CStr(tsToolbars.SelectedItem.Index) + "ph"
            End Select
        End If
    Next i

End Sub

Private Sub SetButtonsState()

    Dim IsOn As Boolean
    
    IsOn = tsToolbars.Tabs.Count > 0

    cmdAddGroup.Enabled = IsOn And Not Project.ToolBar.IsTemplate
    cmdDelGroups.Enabled = IsOn And (lvGroups.ListItems.Count > 0) And Not Project.ToolBar.IsTemplate
    cmdUp.Enabled = IsOn And (lvGroups.ListItems.Count > 1) And Not Project.ToolBar.IsTemplate
    cmdDown.Enabled = cmdUp.Enabled And Not Project.ToolBar.IsTemplate

End Sub

Private Sub chkFollowScrolling_Click()

    UpdateTBInProject
    UpdateToolbarControls

End Sub

Private Sub chkHorizontal_Click()

    UpdateTBInProject
    UpdateToolbarControls

End Sub

Private Sub chkJustify_Click()

    UpdateTBInProject
    UpdateToolbarControls

End Sub

Private Sub chkSmartScrolling_Click()

    UpdateTBInProject

End Sub

Private Sub chkSyncObjSize_Click()

    UpdateTBInProject

End Sub

Private Sub chkVertical_Click()

    UpdateTBInProject
    UpdateToolbarControls

End Sub

Private Sub cmbBorder_Change()

    UpdateTBInProject

End Sub

Private Sub cmbBorder_Click()

    UpdateTBInProject

End Sub

Private Sub cmbFX_Click()

    UpdateTBInProject

End Sub

Private Sub cmbSpanning_Click()

    UpdateTBInProject

End Sub

Private Sub cmbStyle_Click()

    UpdateTBInProject
    
    If IsUpdating = False Then AdjustMenusAlignment tsToolbars.SelectedItem.Index

End Sub

Private Sub cmdAddGroup_Click()

    PopupMenu mnuGroups, , picTBGeneral.Left + cmdAddGroup.Left, picTBGeneral.Top + cmdAddGroup.Top + cmdAddGroup.Height, mnuGrp(0)

End Sub

Private Sub cmdAddTB_Click()

    ReDim Preserve Project.Toolbars(UBound(Project.Toolbars) + 1)
    With Project.Toolbars(UBound(Project.Toolbars))
        .Name = GetTBSecuenceName("Untitled")
        ReDim .Groups(0)
        .Compile = True
    End With
    
    IsUpdating = True
    InitDlg UBound(Project.Toolbars)
    IsUpdating = False
    
    RenameTB

End Sub

Private Sub RenameTB()

    Dim stab As MSComctlLib.Tab
    
    Set stab = tsToolbars.SelectedItem
    
    cmdCancel.Cancel = False
    cmdOK.Default = False

    With txtTBName
        .Left = tsToolbars.Left + stab.Left - 15
        .Top = stab.Top
        .Width = stab.Width - 60
        .Text = stab.caption
        .Visible = True
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
    End With

End Sub

Private Sub cmdCancel_Click()

    Project = ProjectBack
    MenuGrps = GrpsBack
    MenuCmds = CmdsBack

    Unload Me

End Sub

Private Sub cmdChange_Click()
    
    SelImage.FileName = picTBImage.tag
    SelImage.SupportsFlash = True
    frmRscImages.Show vbModal
    With SelImage
        If .IsValid Then
            SetTBPicture .FileName
        End If
    End With

End Sub

Private Sub cmdCompile_Click()

    Me.Enabled = False
    frmMain.ToolsCompile
    Me.Enabled = True

End Sub

Private Sub cmdDelGroups_Click()

    DelSelectedGroups

End Sub

Private Sub cmdDelTB_Click()

    Dim i As Integer
    
    For i = tsToolbars.SelectedItem.Index To tsToolbars.Tabs.Count - 1
        Project.Toolbars(i) = Project.Toolbars(i + 1)
    Next i
    ReDim Preserve Project.Toolbars(UBound(Project.Toolbars) - 1)
    
    IsUpdating = True
    
    InitDlg
    AutoSelToolbar
    
    IsUpdating = False

End Sub

Private Sub cmdDown_Click()

    MoveGrpDown

End Sub

Private Sub cmdOK_Click()

    If RequiresImageCode Then Project.RemoveImageAutoPosCode = False

    frmMain.SaveState GetLocalizedStr(793)
    Unload Me

End Sub

Private Sub cmdPreview_Click()

    frmMain.ShowPreview

End Sub

Private Sub cmdRemove_Click()

    SetTBPicture ""

End Sub

Private Sub cmdRenameTB_Click()

    RenameTB

End Sub

Private Sub cmdSelRefImage_Click()

    frmRefImage.Show vbModal
    txtAttachedTo.Text = Project.ToolBar.AttachTo

End Sub

Private Sub cmdShadowColor_Click()

    BuildUsedColorsArray
    
    With cmdShadowColor
        SelColor = .tag
        SelColor_CanBeTransparent = False
        frmColorPicker.Show vbModal
        SetColor SelColor, cmdShadowColor
    End With
    
    UpdateTBInProject

End Sub

Private Sub cmdTBAutoHeight_Click()

    Dim r() As Integer

    r = GetTBHeight(tsToolbars.SelectedItem.Index, True)
    txtTBHeight.Text = r(1)

End Sub

Private Sub cmdTBAutoWidth_Click()

    Dim r() As Integer

    r = GetTBWidth(tsToolbars.SelectedItem.Index, True)
    txtTBWidth.Text = r(1)

End Sub

Private Sub cmdTBBackColor_Click()
    
    BuildUsedColorsArray
    
    With cmdTBBackColor
        SelColor = .tag
        SelColor_CanBeTransparent = True
        frmColorPicker.Show vbModal
        SetColor SelColor, cmdTBBackColor
    End With
    
    UpdateTBInProject
    
End Sub

Private Sub cmdTBBorderColor_Click()

    BuildUsedColorsArray
    
    With cmdTBBorderColor
        SelColor = .tag
        SelColor_CanBeTransparent = True
        frmColorPicker.Show vbModal
        SetColor SelColor, cmdTBBorderColor
    End With
    
    UpdateTBInProject

End Sub

Private Sub cmdUp_Click()

    MoveGrpUp

End Sub

Private Sub Form_Load()
    
    mnuGroups.Visible = False
    
    CurSelTab = ""
    lastSelTBIdx = 0

    Width = 6945
    Height = 6540 + GetClientTop(Me.hwnd)
    
    CenterForm Me
    SetupCharset Me
    LocalizeUI
    
    ProjectBack = Project
    GrpsBack = MenuGrps
    CmdsBack = MenuCmds
    
    With picTBGeneral
        .BorderStyle = 0
        .ZOrder 0
        picTBAdvanced.Move .Left, .Top, .Width, .Height
        picTBAdvanced.BorderStyle = 0
        picTBAppearance.Move .Left, .Top, .Width, .Height
        picTBAppearance.BorderStyle = 0
        picTBPositioning.Move .Left, .Top, .Width, .Height
        picTBPositioning.BorderStyle = 0
        picTBFX.Move .Left, .Top, .Width, .Height
        picTBFX.BorderStyle = 0
    End With
    
    'Set cmdAddGroup.Picture = frmMain.ilIcons.ListImages("mnuMenuAddGrp").Picture
    'Set cmdDelGroups.Picture = frmMain.ilIcons.ListImages("mnuEditRem|mnuMenuEditRem").Picture
    'Set cmdUp.Picture = frmMain.ilIcons.ListImages("btnUp").Picture
    'Set cmdDown.Picture = frmMain.ilIcons.ListImages("btnDown").Picture
    
    CreateAlignCombo
    FixCtrls4Skin Me
    
    IsUpdating = True
    
    InitDlg
    CreateGroupsMenu
    AutoSelToolbar
    
    #If LITE = 1 Then
        cmdAddTB.Visible = False
    #End If
    
    IsUpdating = False

End Sub

Private Sub AutoSelToolbar()

    Dim tbIndex As Integer
    
    On Error Resume Next

    With frmMain
        If IsTBMapSel Then
            tbIndex = ToolbarIndexByKey(.tvMapView.SelectedItem.key)
        Else
            If Not .tvMenus.SelectedItem Is Nothing Then
                With .tvMenus.SelectedItem
                    If IsGroup(.key) Then
                        tbIndex = MemberOf(GetID)
                    Else
                        tbIndex = MemberOf(MenuCmds(GetID).parent)
                    End If
                End With
            End If
        End If
    End With
    
    If tbIndex = 0 Then tbIndex = 1
    
    If tbIndex <= (tsToolbars.Tabs.Count + 1) Then
        tsToolbars.Tabs(tbIndex).Selected = True
    End If
    
    Select Case TBEPage
        Case tbepcGeneral
            tsToolbar.Tabs("tsGeneral").Selected = True
        Case tbepcAppearance
            tsToolbar.Tabs("tsAppearance").Selected = True
        Case tbepcPositioning
            tsToolbar.Tabs("tsPositioning").Selected = True
        Case tbepcEffects
            tsToolbar.Tabs("tsEffects").Selected = True
        Case tbepcAdvanced
            tsToolbar.Tabs("tsAdvanced").Selected = True
    End Select
    CurSelTab = tsToolbar.SelectedItem.key
    
    UpdateTBButtonsStatus
    
End Sub

Private Sub CreateAlignCombo()

    Dim nItem As ComboItem
    
    icmbAlignment.ComboItems.Clear
    Set icmbAlignment.ImageList = frmMain.ilAlignment
    Set nItem = icmbAlignment.ComboItems.Add(, , GetLocalizedStr(116), 8)
    nItem.tag = 0
    Set nItem = icmbAlignment.ComboItems.Add(, , GetLocalizedStr(117), 2)
    nItem.tag = 1
    Set nItem = icmbAlignment.ComboItems.Add(, , GetLocalizedStr(118), 7)
    nItem.tag = 2
    Set nItem = icmbAlignment.ComboItems.Add(, , GetLocalizedStr(119), 1)
    nItem.tag = 3
    Set nItem = icmbAlignment.ComboItems.Add(, , GetLocalizedStr(120), 4)
    nItem.tag = 4
    Set nItem = icmbAlignment.ComboItems.Add(, , GetLocalizedStr(121), 3)
    nItem.tag = 5
    Set nItem = icmbAlignment.ComboItems.Add(, , GetLocalizedStr(122), 6)
    nItem.tag = 6
    Set nItem = icmbAlignment.ComboItems.Add(, , GetLocalizedStr(123), 5)
    nItem.tag = 7
    Set nItem = icmbAlignment.ComboItems.Add(, , GetLocalizedStr(818), 9)
    nItem.tag = 8
    Set nItem = icmbAlignment.ComboItems.Add(, , GetLocalizedStr(819), 10)
    nItem.tag = 9
    Set nItem = icmbAlignment.ComboItems.Add(, , GetLocalizedStr(820), 11)
    nItem.tag = 10
    Set nItem = icmbAlignment.ComboItems.Add(, , GetLocalizedStr(821), 12)
    nItem.tag = 11
    
    icmbAlignment.ComboItems(1).Selected = True

End Sub

Private Sub CreateGroupsMenu()

    Dim i As Integer
    
    Set msgSubClass = New xfxSC
    msgSubClass.SubClassHwnd picTBAppearance.hwnd, True
    
    mnuGrp(0).caption = GetLocalizedStr(103) + vbTab + GetLocalizedStr(409)
    mnuGrp(0).Enabled = False
    
    For i = 1 To UBound(MenuGrps)
        Load mnuGrp(i + 1)
        mnuGrp(i + 1).caption = xUNI2Unicode(MenuGrps(i).caption) + vbTab + "(" + MenuGrps(i).Name + ")"
        mnuGrp(i + 1).Enabled = True
        mnuGrp(i + 1).Checked = Not (lvGroups.FindItem(MenuGrps(i).Name, lvwText, , lvwWhole) Is Nothing)
    Next i

End Sub

Private Sub UpdateGroupMenuCheckStates()

    Dim i As Integer

    If mnuGrp.Count = 2 Then Exit Sub
    For i = 1 To UBound(MenuGrps)
        mnuGrp(i + 1).Checked = Not (lvGroups.FindItem(MenuGrps(i).Name, lvwText, , lvwWhole) Is Nothing)
    Next i
    lvGroups.Enabled = True

End Sub

Private Sub InitDlg(Optional SelTB As Integer = 1)

    Dim i As Integer
    
    While tsToolbars.Tabs.Count
        tsToolbars.Tabs.Remove 1
    Wend

    cmbBorder.AddItem GetLocalizedStr(110)
    For i = 1 To 10
        cmbBorder.AddItem CStr(i)
    Next i
    
    For i = 1 To UBound(Project.Toolbars)
        tsToolbars.Tabs.Add
        tsToolbars.Tabs(i).caption = Project.Toolbars(i).Name
    Next i
    
    If tsToolbars.Tabs.Count > 0 Then
        tsToolbars.Tabs(SelTB).Selected = True
        tsToolbars_Click
    Else
        CreateGroupsList
        UpdateToolbarControls
    End If
    
    UpdateTBButtonsStatus

End Sub

Private Sub UpdateTBButtonsStatus()

    cmdDelTB.Enabled = (tsToolbars.Tabs.Count > 0) And Not Project.ToolBar.IsTemplate
    cmdRenameTB.Enabled = cmdDelTB.Enabled

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    msgSubClass.SubClassHwnd picTBAppearance.hwnd, False

End Sub

Private Sub icmbAlignment_Click()

    If IsUpdating Then Exit Sub
    UpdateTBInProject
    UpdateToolbarControls

End Sub

Private Sub msgSubClass_NewMessage(ByVal hwnd As Long, uMsg As Long, wParam As Long, lParam As Long, Cancel As Boolean, UseRetVal As Boolean, RetVal As Long)

    Select Case hwnd
        Case picTBAppearance.hwnd
            Select Case uMsg
                Case WM_CTLCOLORSTATIC, WM_CTLCOLOREDIT, WM_PAINT
                    DrawColorBoxes Me
            End Select
    End Select
    
End Sub

Private Sub SetSelTBValues()

    On Error Resume Next

    With Project.ToolBar
        chkFollowScrolling.Value = Abs(.FollowHScroll Or .FollowVScroll)
        chkHorizontal.Value = Abs(.FollowHScroll)
        chkVertical.Value = Abs(.FollowVScroll)
        chkSmartScrolling.Value = Abs(.SmartScrolling)
        Select Case .Width
            Case 0
                opTBWidth(0).Value = True
            Case 1
                opTBWidth(1).Value = True
            Case Else
                opTBWidth(2).Value = True
                txtTBWidth.Text = Abs(.Width)
        End Select
        Select Case .Height
            Case 0
                opTBHeight(0).Value = True
            Case Else
                opTBHeight(1).Value = True
                txtTBHeight.Text = Abs(.Height)
        End Select
        txtMarginH.Text = .OffsetH
        txtMarginV.Text = .OffsetV
        txtTBCMarginH.Text = .ContentsMarginH
        txtTBCMarginV.Text = .ContentsMarginV
        opAlignment(.Alignment).Value = True
        cmbStyle.ListIndex = .Style
        cmbSpanning.ListIndex = .Spanning
        If .bOrder >= 0 And .bOrder <= 10 Then
            cmbBorder.ListIndex = .bOrder
        Else
            cmbBorder.Text = .bOrder
        End If
        cmbFX.ListIndex = .BorderStyle
        chkJustify.Value = Abs(.JustifyHotSpots)
        txtTBSeparation.Text = .Separation
        SetColor .BackColor, cmdTBBackColor
        SetColor .BorderColor, cmdTBBorderColor
        SetTBPicture .Image
        txtACX.Text = .CustX
        txtACY.Text = .CustY
        txtAttachedTo.Text = .AttachTo
        chkSyncObjSize.Value = IIf(.AttachToAutoResize, vbChecked, vbUnchecked)
        SelectAlignment .AttachToAlignment
        txtVisCond.Text = .Condition
        
        SetColor .DropShadowColor, cmdShadowColor
        sldDropShadow.Value = .DropShadowSize
        sldTransparency.Value = .Transparency
        
        txtRadiusTL.Text = .Radius.topLeft
        txtRadiusTR.Text = .Radius.topRight
        txtRadiusBL.Text = .Radius.bottomLeft
        txtRadiusBR.Text = .Radius.bottomRight
    End With

End Sub

Private Sub opAlignment_Click(Index As Integer)

    If Not IsUpdating Then
        If Index = 11 Then
            With TipsSys
                .CanDisable = True
                .DialogTitle = "Incompatibility Warning"
                .TipTitle = "Netscape 4 does not support this feature"
                .Tip = "Toolbars using the Free Flow alignment setting will not be displayed under Netscape Navigator 4"
                .Show
            End With
        End If
    End If

    UpdateTBInProject
    UpdateToolbarControls

End Sub

Private Sub opTBHeight_Click(Index As Integer)

    txtTBHeight.Enabled = (Index = 1)
    cmdTBAutoHeight.Enabled = (Index = 1)
    
    UpdateTBInProject

End Sub

Private Sub opTBWidth_Click(Index As Integer)

    txtTBWidth.Enabled = (Index = 2)
    cmdTBAutoWidth.Enabled = (Index = 2)
    
    UpdateTBInProject

End Sub

Private Sub DelSelectedGroups()

    Dim nItem As ListItem
    
ReStart:
    For Each nItem In lvGroups.ListItems
        If nItem.Selected Then
            lvGroups.ListItems.Remove nItem.Index
            GoTo ReStart
        End If
    Next nItem
    UpdateGroupMenuCheckStates
    GenerateGroupsArray
    
    CoolListView lvGroups

End Sub

Private Sub MoveGrpDown()

    Dim SelItem As String

    With lvGroups
        If .SelectedItem.Index = .ListItems.Count Then Exit Sub
        SelItem = .SelectedItem.Text
        
        .SelectedItem.Text = .ListItems(.SelectedItem.Index + 1).Text
        .ListItems(.SelectedItem.Index + 1).Text = SelItem
        
        GenerateGroupsArray
        CreateGroupsList
        
        .MultiSelect = False
        With .FindItem(SelItem, lvwText, , lvwWhole)
            .Selected = True
            .EnsureVisible
        End With
        .MultiSelect = True
    End With
    
    lvGroups.SetFocus

End Sub

Private Sub MoveGrpUp()

    Dim SelItem As String

    With lvGroups
        If .SelectedItem.Index = 1 Then Exit Sub
        SelItem = .SelectedItem.Text
        
        .SelectedItem.Text = .ListItems(.SelectedItem.Index - 1).Text
        .ListItems(.SelectedItem.Index - 1).Text = SelItem
        
        GenerateGroupsArray
        CreateGroupsList
        
        .MultiSelect = False
        With .FindItem(SelItem, lvwText, , lvwWhole)
            .Selected = True
            .EnsureVisible
        End With
        .MultiSelect = True
    End With
    
    lvGroups.SetFocus

End Sub

Private Sub mnuGrp_Click(Index As Integer)

    Dim nItem As ListItem

    With mnuGrp(Index)
        If .Checked Then
            lvGroups.ListItems.Remove lvGroups.FindItem(MenuGrps(Index - 1).Name, lvwText, , lvwWhole).Index
        Else
            Set nItem = lvGroups.ListItems.Add(, , MenuGrps(Index - 1).Name)
            nItem.SubItems(1) = MenuGrps(Index - 1).caption
        End If
        .Checked = Not .Checked
    End With
    
    GenerateGroupsArray
    CreateGroupsList
    
End Sub

Private Sub GenerateGroupsArray()

    Dim nItem As ListItem

    With Project.Toolbars(tsToolbars.SelectedItem.Index)
        ReDim .Groups(0)
        For Each nItem In lvGroups.ListItems
            If nItem.tag <> 1 Then
                ReDim Preserve .Groups(UBound(.Groups) + 1)
                .Groups(UBound(.Groups)) = nItem.Text
            End If
        Next nItem
    End With
    
    SetButtonsState

End Sub

Private Sub sldDropShadow_Change()

    UpdateTBInProject
    
    cmdShadowColor.Enabled = Project.Toolbars(tsToolbars.SelectedItem.Index).DropShadowSize > 0

End Sub

Private Sub sldTransparency_Change()

    UpdateTBInProject

End Sub

Private Sub tmrUpdateSelTB_Timer()
    
    If lastSelTBIdx = tsToolbars.SelectedItem.Index Then Exit Sub
    tmrUpdateSelTB.Enabled = False
    
    lastSelTBIdx = tsToolbars.SelectedItem.Index
    
    Project.ToolBar = Project.Toolbars(lastSelTBIdx)
    
    IsUpdating = True
    
    CreateGroupsList
    UpdateToolbarControls

    SetSelTBValues
    
    frmMain.DoLivePreview , , , lastSelTBIdx
    IsUpdating = False

End Sub

Private Sub tsToolbar_BeforeClick(Cancel As Integer)

    CurSelTab = tsToolbar.SelectedItem.key

End Sub

Private Sub tsToolbars_BeforeClick(Cancel As Integer)

    UpdateTBInProject

End Sub

Private Sub tsToolbars_Click()
    
    tmrUpdateSelTB.Enabled = False
    tmrUpdateSelTB.Enabled = True
    
End Sub

Private Sub tsToolbar_Click()

    On Error Resume Next

    Select Case tsToolbar.SelectedItem.key
        Case "tsGeneral"
            picTBGeneral.ZOrder 0
        Case "tsPositioning"
            picTBPositioning.ZOrder 0
            ShowTBPosInfo
        Case "tsAppearance"
            picTBAppearance.ZOrder 0
        Case "tsAdvanced"
'            #If LITE = 1 Then
'                frmMain.ShowLITELImitationInfo 2
'                tsToolbar.Tabs(CurSelTab).Selected = True
'            #Else
'                picTBAdvanced.ZOrder 0
'            #End If
            #If LITE = 1 Then
                chkFollowScrolling.Visible = False
                chkHorizontal.Visible = False
                chkVertical.Visible = False
                chkSmartScrolling.Visible = False
                udH.Visible = False
                udV.Visible = False
                ucLine3D1.Visible = False
                ucLine3D2.Visible = False
                lblVisCond.Visible = False
                txtVisCond.Visible = False
            #End If
            picTBAdvanced.ZOrder 0
        Case "tsEffects"
            #If LITE = 1 Then
                frmMain.ShowLITELImitationInfo 2
                tsToolbar.Tabs(CurSelTab).Selected = True
            #Else
                picTBFX.ZOrder 0
            #End If
    End Select

End Sub

Private Sub txtACX_Change()

    UpdateTBInProject

End Sub

Private Sub txtACX_GotFocus()

    SelAll txtACX

End Sub

Private Sub txtACX_KeyPress(KeyAscii As Integer)

    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 45 Then
        KeyAscii = 0
    End If

End Sub

Private Sub txtACY_Change()

    UpdateTBInProject

End Sub

Private Sub txtACY_GotFocus()

    SelAll txtACY

End Sub

Private Sub txtACY_KeyPress(KeyAscii As Integer)

    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 45 Then
        KeyAscii = 0
    End If

End Sub

Private Sub txtAttachedTo_Change()
    
    UpdateTBInProject

End Sub

Private Sub txtAttachedTo_GotFocus()

    SelAll txtAttachedTo

End Sub

Private Sub txtMarginH_Change()

    UpdateTBInProject

End Sub

Private Sub txtMarginH_GotFocus()

    SelAll txtMarginH

End Sub

Private Sub txtMarginH_KeyPress(KeyAscii As Integer)

    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 45 Then
        KeyAscii = 0
    End If

End Sub

Private Sub txtMarginV_Change()

    UpdateTBInProject

End Sub

Private Sub txtMarginV_GotFocus()

    SelAll txtMarginV

End Sub

Private Sub txtMarginV_KeyPress(KeyAscii As Integer)

    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 45 Then
        KeyAscii = 0
    End If

End Sub

Private Sub txtRadiusBL_Change()

    UpdateTBInProject

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

    UpdateTBInProject

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

    UpdateTBInProject

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
    
    UpdateTBInProject

End Sub

Private Sub txtRadiusTR_GotFocus()

    SelAll txtRadiusTR

End Sub

Private Sub txtRadiusTR_KeyPress(KeyAscii As Integer)

    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If

End Sub

Private Sub txtTBCMarginH_Change()

    UpdateTBInProject

End Sub

Private Sub txtTBCMarginH_GotFocus()

    SelAll txtTBCMarginH

End Sub

Private Sub txtTBCMarginH_KeyPress(KeyAscii As Integer)

    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If

End Sub

Private Sub txtTBCMarginV_Change()

    UpdateTBInProject

End Sub

Private Sub txtTBCMarginV_GotFocus()

    SelAll txtTBCMarginV

End Sub

Private Sub txtTBCMarginV_KeyPress(KeyAscii As Integer)

    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If

End Sub

Private Sub txtTBHeight_Change()

    UpdateTBInProject

End Sub

Private Sub txtTBName_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        tsToolbars.SelectedItem.caption = txtTBName.Text
        Project.Toolbars(tsToolbars.SelectedItem.Index).Name = txtTBName.Text
        txtTBName.Visible = False
        
        cmdCancel.Cancel = True
        cmdOK.Default = True
        
        KeyAscii = 0
    End If

End Sub

Private Sub txtTBName_LostFocus()

    txtTBName.Visible = False
    
    cmdCancel.Cancel = True
    cmdOK.Default = True

End Sub

Private Sub txtTBSeparation_Change()

    UpdateTBInProject

End Sub

Private Sub txtTBSeparation_GotFocus()

    SelAll txtTBSeparation

End Sub

Private Sub txtTBSeparation_KeyPress(KeyAscii As Integer)

    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> Asc("-") Then
        KeyAscii = 0
    End If

End Sub

Private Sub LocalizeUI()

    caption = "Toolbars Editor"
    
    'Toolbar
    lblTBSelGrps.caption = GetLocalizedStr(351)
    lvGroups.ColumnHeaders(1).Text = GetLocalizedStr(409)
    lvGroups.ColumnHeaders(2).Text = GetLocalizedStr(103)
    
    tsToolbar.Tabs(1).caption = GetLocalizedStr(321)
    tsToolbar.Tabs(2).caption = GetLocalizedStr(352)
    tsToolbar.Tabs(3).caption = GetLocalizedStr(353)
    tsToolbar.Tabs(4).caption = GetLocalizedStr(832)
    tsToolbar.Tabs(5).caption = GetLocalizedStr(325)
    
    lblTBStyle.caption = GetLocalizedStr(354)
    lblTBBorder.caption = GetLocalizedStr(355)
    lblTBSeparation.caption = GetLocalizedStr(356)
    lblTBBackColor.caption = GetLocalizedStr(358)
    lblTBBackImage.caption = GetLocalizedStr(359)
    lblTBMargins.caption = GetLocalizedStr(216)
    lblH.caption = GetLocalizedStr(211)
    lblV.caption = GetLocalizedStr(210)
    lblBorderStyle.caption = GetLocalizedStr(967)
    
    lblDropShadow.caption = GetLocalizedStr(221): lblDropShadow.Left = sldDropShadow.Left + (sldDropShadow.Width - lblDropShadow.Width) / 2
    lblTransparency.caption = GetLocalizedStr(222): lblTransparency.Left = sldTransparency.Left + (sldTransparency.Width - lblTransparency.Width) / 2
    lblDSOFF.caption = GetLocalizedStr(228)
    lblTOFF.caption = GetLocalizedStr(228)
    lblDSDarker.caption = GetLocalizedStr(229)
    lblTInvisible.caption = GetLocalizedStr(230)
    lblShadowColor.caption = GetLocalizedStr(212)
    
    cmdChange.caption = GetLocalizedStr(189)
    cmdRemove.caption = GetLocalizedStr(201)
    
    lblTBAlignment.caption = GetLocalizedStr(115)
    lblTBSpanning.caption = GetLocalizedStr(361)
    lblTBOffset.caption = GetLocalizedStr(362)
    lblTBOffsetH.caption = GetLocalizedStr(363)
    lblTBOffsetV.caption = GetLocalizedStr(364)
    
    cmbStyle.Clear
    cmbStyle.AddItem GetLocalizedStr(211)
    cmbStyle.AddItem GetLocalizedStr(210)
    
    cmbSpanning.Clear
    cmbSpanning.AddItem GetLocalizedStr(365)
    cmbSpanning.AddItem GetLocalizedStr(366)
    
    opAlignment(9).caption = GetLocalizedStr(367)
    chkFollowScrolling.caption = GetLocalizedStr(368)
    opTBWidth(0).caption = GetLocalizedStr(185)
    opTBWidth(1).caption = GetLocalizedStr(964)
    opTBWidth(2).caption = GetLocalizedStr(224)
    lblGSWidth.caption = GetLocalizedStr(428)
    cmdTBAutoWidth.caption = GetLocalizedStr(225)
    opTBHeight(0).caption = GetLocalizedStr(185)
    opTBHeight(1).caption = GetLocalizedStr(224)
    lblGSHeight.caption = GetLocalizedStr(429)
    cmdTBAutoHeight.caption = GetLocalizedStr(225)
    lblVisCond.caption = GetLocalizedStr(965)
    lblTBSize.caption = GetLocalizedStr(966)
    opAlignment(10).caption = GetLocalizedStr(968)
    lblObjName.caption = GetLocalizedStr(969)
    lblAlignment.caption = GetLocalizedStr(115)
    
    PopulateBorderStyleCombo cmbFX
    
    cmdOK.caption = GetLocalizedStr(186)
    cmdCancel.caption = GetLocalizedStr(187)
    cmdPreview.caption = GetLocalizedStr(158)
    
    FixContolsWidth Me
    
    If Preferences.language <> "eng" Then
        cmdOK.Width = SetCtrlWidth(cmdOK)
        cmdCancel.Width = SetCtrlWidth(cmdCancel)
        cmdPreview.Width = SetCtrlWidth(cmdPreview)
    End If

End Sub

Private Sub CreateGroupsList()

    Dim i As Integer
    Dim nItem As ListItem
    Dim mg As MenuGrp
    
    lvGroups.ListItems.Clear
    
    If tsToolbars.Tabs.Count = 0 Then Exit Sub
    
    On Error GoTo Try2Abort
    
    With Project.Toolbars(tsToolbars.SelectedItem.Index)
        For i = 1 To UBound(.Groups)
            mg = MenuGrps(GetIDByName(.Groups(i)))
            Set nItem = lvGroups.ListItems.Add(, , mg.Name)
            nItem.SubItems(1) = xUNI2Unicode(mg.caption)
        Next i
    End With
    
Try2Abort:
    If lvGroups.ListItems.Count > 0 Then
        lvGroups.ListItems(1).Selected = True
    End If
    
    UpdateGroupMenuCheckStates
    CoolListView lvGroups

End Sub

Private Sub SetTBPicture(FileName As String)

    On Error Resume Next

    TileImage FileName, picTBImage
    picTBImage.tag = FileName
    
    UpdateTBInProject

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF1 Then
        Select Case tsToolbar.SelectedItem.key
            Case "tsGeneral"
                showHelp "dialogs/pp_tb_general.htm"
            Case "tsAppearance"
                showHelp "dialogs/pp_tb_appearance.htm"
            Case "tsPositioning"
                showHelp "dialogs/pp_tb_positioning.htm"
            Case "tsAdvanced"
                showHelp "dialogs/pp_tb_advanced.htm"
            Case "tsEffects"
                showHelp "dialogs/pp_tb_fx.htm"
        End Select
    End If

End Sub

Private Sub UpdateTBInProject()

    Dim i As Integer
    
    SetJustifyHotSpotsLabelCaption
    If IsUpdating Then Exit Sub

    With Project.Toolbars(tsToolbars.SelectedItem.Index)
        .FollowHScroll = (chkHorizontal.Value = vbChecked) And (chkFollowScrolling.Value = vbChecked)
        .FollowVScroll = (chkVertical.Value = vbChecked) And (chkFollowScrolling.Value = vbChecked)
        .SmartScrolling = (chkSmartScrolling.Value = vbChecked)
        .OffsetH = Val(txtMarginH.Text)
        .OffsetV = Val(txtMarginV.Text)
        .ContentsMarginH = Val(txtTBCMarginH.Text)
        .ContentsMarginV = Val(txtTBCMarginV.Text)
        
        .DropShadowColor = cmdShadowColor.tag
        .DropShadowSize = sldDropShadow.Value
        .Transparency = sldTransparency.Value
        
        If opTBWidth(0).Value Then .Width = 0
        If opTBWidth(1).Value Then .Width = 1
        If opTBWidth(2).Value Then .Width = -Val(txtTBWidth.Text)
        
        If opTBHeight(0).Value Then .Height = 0
        If opTBHeight(1).Value Then .Height = -Val(txtTBHeight.Text)
        
        For i = 0 To opAlignment.Count - 1
            If opAlignment(i).Value = True Then
                .Alignment = i
                Exit For
            End If
        Next i
        
        .Style = cmbStyle.ListIndex
        .Spanning = cmbSpanning.ListIndex
        .bOrder = Val(cmbBorder.Text)
        .JustifyHotSpots = (chkJustify.Value = vbChecked)
        .BackColor = cmdTBBackColor.tag
        .BorderColor = cmdTBBorderColor.tag
        .BorderStyle = cmbFX.ListIndex
        .Separation = Val(txtTBSeparation.Text)
        .Image = picTBImage.tag
        .CustX = Val(txtACX.Text)
        .CustY = Val(txtACY.Text)
        
        .AttachTo = txtAttachedTo.Text
        .AttachToAlignment = Val(icmbAlignment.SelectedItem.tag)
        .AttachToAutoResize = (chkSyncObjSize.Value = vbChecked)
        
        .Radius.topLeft = Val(txtRadiusTL.Text)
        .Radius.topRight = Val(txtRadiusTR.Text)
        .Radius.bottomLeft = Val(txtRadiusBL.Text)
        .Radius.bottomRight = Val(txtRadiusBR.Text)
        
        .Condition = txtVisCond.Text
    End With
    
    ShowTBPosInfo
    UpdateTBButtonsStatus
    
    frmMain.DoLivePreview , , , tsToolbars.SelectedItem.Index
    
End Sub

Private Sub SetJustifyHotSpotsLabelCaption()

    lblJustify.caption = Replace(GetLocalizedStr(357), "%WH%", _
                        IIf(cmbStyle.ListIndex = 0, GetLocalizedStr(428), GetLocalizedStr(429)))

End Sub

Private Sub txtTBWidth_Change()

    UpdateTBInProject
    
End Sub

Private Sub SelectAlignment(v As GroupAlignmentConstants)

    Dim i As Integer
    
    For i = 1 To icmbAlignment.ComboItems.Count
        If Val(icmbAlignment.ComboItems(i).tag) = v Then
            icmbAlignment.ComboItems(i).Selected = True
            Exit For
        End If
    Next i

End Sub

Private Sub txtVisCond_Change()

    UpdateTBInProject

End Sub
