VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{E1C6DB9D-BD4A-4A61-A759-0CED75D034BF}#43.0#0"; "SmartButton.ocx"
Object = "{DB06EC30-01E1-485F-A3C7-CE80CA0D7D37}#2.0#0"; "xFXSlider.ocx"
Object = "{9A2947B4-87EE-40EE-A3EF-32BDC32D5726}#1.0#0"; "xfxline3d.ocx"
Begin VB.Form frmProjProp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Project Properties"
   ClientHeight    =   12900
   ClientLeft      =   2115
   ClientTop       =   2295
   ClientWidth     =   14700
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmProjProp.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   12900
   ScaleWidth      =   14700
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picGlobalSettings 
      Height          =   6765
      Left            =   510
      ScaleHeight     =   6705
      ScaleWidth      =   7080
      TabIndex        =   54
      Top             =   5985
      Width           =   7140
      Begin VB.PictureBox picGSTimers 
         Height          =   3390
         Left            =   555
         ScaleHeight     =   3330
         ScaleWidth      =   5625
         TabIndex        =   76
         Top             =   270
         Width           =   5685
         Begin xfxLine3D.ucLine3D uc3DLine2 
            Height          =   30
            Left            =   15
            TabIndex        =   109
            Top             =   2475
            Width           =   6405
            _ExtentX        =   11298
            _ExtentY        =   53
         End
         Begin xFXSlider.ucSlider sldSubMenusDelay 
            Height          =   270
            Left            =   120
            TabIndex        =   77
            Top             =   1080
            Width           =   4575
            _ExtentX        =   820
            _ExtentY        =   476
            Max             =   2000
            TickStyle       =   0
            TickFrequency   =   110
            SmallChange     =   10
            LargeChange     =   100
            BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomKnobImage =   "frmProjProp.frx":058A
            CustomSelKnobImage=   "frmProjProp.frx":0864
         End
         Begin xFXSlider.ucSlider sldHideDelay 
            Height          =   270
            Left            =   120
            TabIndex        =   78
            Top             =   2835
            Width           =   4575
            _ExtentX        =   820
            _ExtentY        =   476
            Min             =   10
            Max             =   2000
            Value           =   250
            TickStyle       =   0
            TickFrequency   =   110
            SmallChange     =   10
            LargeChange     =   100
            BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomKnobImage =   "frmProjProp.frx":0B3E
            CustomSelKnobImage=   "frmProjProp.frx":0E18
         End
         Begin xFXSlider.ucSlider sldRootMenusDelay 
            Height          =   270
            Left            =   120
            TabIndex        =   79
            Top             =   285
            Width           =   4575
            _ExtentX        =   820
            _ExtentY        =   476
            Max             =   2000
            Value           =   15
            TickStyle       =   0
            TickFrequency   =   110
            SmallChange     =   10
            LargeChange     =   100
            BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CaptionPosition =   0
            Caption         =   ""
            CustomKnobImage =   "frmProjProp.frx":10F2
            CustomSelKnobImage=   "frmProjProp.frx":13CC
         End
         Begin xFXSlider.ucSlider sldSelChangeDelay 
            Height          =   270
            Left            =   120
            TabIndex        =   111
            Top             =   1890
            Width           =   4575
            _ExtentX        =   820
            _ExtentY        =   476
            Min             =   10
            Max             =   2000
            Value           =   250
            TickStyle       =   0
            TickFrequency   =   110
            SmallChange     =   10
            LargeChange     =   100
            BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomKnobImage =   "frmProjProp.frx":16A6
            CustomSelKnobImage=   "frmProjProp.frx":1980
         End
         Begin VB.Label mhDSec 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0ms"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000011&
            Height          =   195
            Left            =   4890
            TabIndex        =   118
            Top             =   2873
            Width           =   285
         End
         Begin VB.Label scDSec 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0ms"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000011&
            Height          =   195
            Left            =   4890
            TabIndex        =   117
            Top             =   1928
            Width           =   285
         End
         Begin VB.Label smDSec 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0ms"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000011&
            Height          =   195
            Left            =   4890
            TabIndex        =   116
            Top             =   1118
            Width           =   285
         End
         Begin VB.Label rmDSec 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0ms"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000011&
            Height          =   195
            Left            =   4890
            TabIndex        =   115
            Top             =   330
            Width           =   285
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "More"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Left            =   4395
            TabIndex        =   114
            Top             =   2175
            Width           =   345
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Less"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Left            =   60
            TabIndex        =   113
            Top             =   2175
            Width           =   300
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Selection Change Delay"
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
            Left            =   75
            TabIndex        =   112
            Top             =   1650
            Width           =   1695
         End
         Begin VB.Label lblMore2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "More"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Left            =   4395
            TabIndex        =   88
            Top             =   1380
            Width           =   345
         End
         Begin VB.Label lblLess2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Less"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Left            =   60
            TabIndex        =   87
            Top             =   1380
            Width           =   300
         End
         Begin VB.Label lblSubmDispDelay 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Submenus Display Delay"
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
            Left            =   75
            TabIndex        =   86
            Top             =   840
            Width           =   1740
         End
         Begin VB.Label lblMore1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "More"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Left            =   4395
            TabIndex        =   85
            Top             =   3120
            Width           =   345
         End
         Begin VB.Label lblLess1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Less"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Left            =   60
            TabIndex        =   84
            Top             =   3120
            Width           =   300
         End
         Begin VB.Label lblMenusHideDelay 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Menus Hide Delay"
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
            Left            =   75
            TabIndex        =   83
            Top             =   2595
            Width           =   1275
         End
         Begin VB.Label lblRootmDispDelay 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Root Menus Display Delay"
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
            Left            =   75
            TabIndex        =   82
            Top             =   45
            Width           =   1860
         End
         Begin VB.Label lblLess3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Less"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Left            =   60
            TabIndex        =   81
            Top             =   585
            Width           =   300
         End
         Begin VB.Label lblMore3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "More"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Left            =   4395
            TabIndex        =   80
            Top             =   585
            Width           =   345
         End
      End
      Begin VB.PictureBox picGSFX 
         Height          =   4110
         Left            =   750
         ScaleHeight     =   4050
         ScaleWidth      =   4530
         TabIndex        =   89
         Top             =   1845
         Width           =   4590
         Begin VB.Frame Frame2 
            Caption         =   "Menu Items Blinking Effect"
            Height          =   1635
            Left            =   75
            TabIndex        =   100
            Top             =   2160
            Width           =   6270
            Begin xFXSlider.ucSlider sldBlinkEffect 
               Height          =   270
               Left            =   1380
               TabIndex        =   101
               Top             =   405
               Width           =   2430
               _ExtentX        =   820
               _ExtentY        =   476
               Max             =   8
               Value           =   5
               TickStyle       =   0
               TickFrequency   =   1
               SmallChange     =   1
               HighlightColorEnd=   128
               BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   ""
               CustomKnobImage =   "frmProjProp.frx":1C5A
               CustomSelKnobImage=   "frmProjProp.frx":1F34
            End
            Begin xFXSlider.ucSlider sldBlinkSpeed 
               Height          =   270
               Left            =   1380
               TabIndex        =   105
               Top             =   990
               Width           =   2430
               _ExtentX        =   820
               _ExtentY        =   476
               Max             =   500
               TickStyle       =   0
               TickFrequency   =   25
               SmallChange     =   1
               HighlightColorEnd=   128
               BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   ""
               CustomKnobImage =   "frmProjProp.frx":220E
               CustomSelKnobImage=   "frmProjProp.frx":24E8
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Fast"
               BeginProperty Font 
                  Name            =   "Small Fonts"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   165
               Left            =   3555
               TabIndex        =   108
               Top             =   1275
               Width           =   285
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Slow"
               BeginProperty Font 
                  Name            =   "Small Fonts"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   165
               Left            =   1290
               TabIndex        =   107
               Top             =   1275
               Width           =   300
            End
            Begin VB.Label Label4 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Speed"
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
               Left            =   240
               TabIndex        =   106
               Top             =   1065
               Width           =   450
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Many"
               BeginProperty Font 
                  Name            =   "Small Fonts"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   165
               Left            =   3510
               TabIndex        =   104
               Top             =   690
               Width           =   345
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Off"
               BeginProperty Font 
                  Name            =   "Small Fonts"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   165
               Left            =   1350
               TabIndex        =   103
               Top             =   690
               Width           =   195
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Amount"
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
               Left            =   240
               TabIndex        =   102
               Top             =   480
               Width           =   555
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "Unfolding Effects"
            Height          =   2040
            Left            =   75
            TabIndex        =   92
            Top             =   45
            Width           =   6270
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
               ItemData        =   "frmProjProp.frx":27C2
               Left            =   1380
               List            =   "frmProjProp.frx":27D5
               Style           =   2  'Dropdown List
               TabIndex        =   93
               Top             =   375
               Width           =   2430
            End
            Begin xFXSlider.ucSlider sldAnimSpeed 
               Height          =   270
               Left            =   1380
               TabIndex        =   94
               Top             =   840
               Width           =   2430
               _ExtentX        =   820
               _ExtentY        =   476
               Min             =   5
               Max             =   50
               Value           =   5
               TickStyle       =   0
               TickFrequency   =   5
               HighlightColorEnd=   128
               BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   ""
               CustomKnobImage =   "frmProjProp.frx":280E
               CustomSelKnobImage=   "frmProjProp.frx":2AE8
            End
            Begin SmartButtonProject.SmartButton cmdAUE 
               Height          =   360
               Left            =   240
               TabIndex        =   95
               Top             =   1500
               Width           =   2865
               _ExtentX        =   5054
               _ExtentY        =   635
               Caption         =   "        Advanced Unfolding Effects..."
               Picture         =   "frmProjProp.frx":2DC2
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
               OffsetLeft      =   4
            End
            Begin VB.Label lblEffectType 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Effect Type"
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
               Left            =   240
               TabIndex        =   99
               Top             =   435
               Width           =   840
            End
            Begin VB.Label lblSlow 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Slow"
               BeginProperty Font 
                  Name            =   "Small Fonts"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   165
               Left            =   1275
               TabIndex        =   98
               Top             =   1155
               Width           =   300
            End
            Begin VB.Label lblFast 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Fast"
               BeginProperty Font 
                  Name            =   "Small Fonts"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   165
               Left            =   3585
               TabIndex        =   97
               Top             =   1155
               Width           =   285
            End
            Begin VB.Label lblEffectSpeed 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Effect Speed"
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
               Left            =   240
               TabIndex        =   96
               Top             =   885
               Width           =   930
            End
         End
      End
      Begin VB.PictureBox picGSExtFeat 
         Height          =   945
         Left            =   60
         ScaleHeight     =   885
         ScaleWidth      =   1890
         TabIndex        =   90
         Top             =   4485
         Width           =   1950
         Begin MSComctlLib.ListView lvExtFeatures 
            Height          =   3855
            Left            =   -15
            TabIndex        =   91
            Top             =   -15
            Width           =   6405
            _ExtentX        =   11298
            _ExtentY        =   6800
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            HideColumnHeaders=   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Key             =   "chDesc"
               Text            =   "Description"
               Object.Width           =   2540
            EndProperty
         End
      End
      Begin VB.CommandButton cmdRestore 
         Caption         =   "Restore Defaults"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   4935
         TabIndex        =   55
         Top             =   4155
         Width           =   1575
      End
      Begin MSComctlLib.TabStrip tsGlobalOptions 
         Height          =   4425
         Left            =   0
         TabIndex        =   75
         Top             =   -15
         Width           =   6570
         _ExtentX        =   11589
         _ExtentY        =   7805
         Placement       =   1
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   3
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Special Effects"
               Key             =   "tsGSFX"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Timers"
               Key             =   "tsGSTimers"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Extended Features"
               Key             =   "tsGSExtFeat"
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
      Begin SmartButtonProject.SmartButton sbGIGSProp 
         Height          =   270
         Left            =   180
         TabIndex        =   110
         Top             =   5460
         Width           =   300
         _ExtentX        =   529
         _ExtentY        =   476
         Picture         =   "frmProjProp.frx":315C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OffsetBottom    =   1
      End
   End
   Begin MSComctlLib.ImageList ilIcons 
      Left            =   165
      Top             =   6315
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483633
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProjProp.frx":34F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProjProp.frx":3650
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProjProp.frx":37AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProjProp.frx":3BFC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog cDlg 
      Left            =   195
      Top             =   5745
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picAdvanced 
      Height          =   4500
      Left            =   7920
      ScaleHeight     =   4440
      ScaleWidth      =   6450
      TabIndex        =   56
      Top             =   6315
      Width           =   6510
      Begin VB.CheckBox chkGZIP 
         Caption         =   "Further compress using GZIP"
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
         Left            =   165
         TabIndex        =   119
         Top             =   795
         Width           =   4335
      End
      Begin VB.CheckBox chkRemoveIBHSCode 
         Caption         =   "Attempt to Remove Image-based HotSpot Code"
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
         Left            =   165
         TabIndex        =   74
         Top             =   1050
         Width           =   4335
      End
      Begin VB.Frame lblCodeOp 
         Caption         =   "Code Optimization"
         Height          =   1350
         Left            =   75
         TabIndex        =   69
         Top             =   45
         Width           =   4545
         Begin xFXSlider.ucSlider ucCodeOp 
            Height          =   270
            Left            =   180
            TabIndex        =   70
            Top             =   270
            Width           =   4095
            _ExtentX        =   820
            _ExtentY        =   476
            Max             =   2
            Value           =   1
            TickStyle       =   2
            TickFrequency   =   1
            SmallChange     =   1
            LargeChange     =   1
            BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomKnobImage =   "frmProjProp.frx":3F96
            CustomSelKnobImage=   "frmProjProp.frx":4270
         End
         Begin VB.Label lblCO0 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "none"
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
            Left            =   120
            TabIndex        =   73
            Top             =   525
            Width           =   360
         End
         Begin VB.Label lblCO1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "normal"
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
            Left            =   1995
            TabIndex        =   72
            Top             =   525
            Width           =   480
         End
         Begin VB.Label lblCO2 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "max"
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
            Left            =   4050
            TabIndex        =   71
            Top             =   525
            Width           =   300
         End
      End
      Begin VB.CommandButton cmdFontSubst 
         Caption         =   "Font Substitution"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1575
         TabIndex        =   67
         Top             =   4035
         Width           =   1410
      End
      Begin VB.CommandButton cmdPosOffset 
         Caption         =   "Menus Offset"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   45
         TabIndex        =   66
         Top             =   4035
         Width           =   1410
      End
      Begin VB.TextBox txtJSFileNames 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   1095
         Left            =   225
         MultiLine       =   -1  'True
         TabIndex        =   65
         TabStop         =   0   'False
         Text            =   "frmProjProp.frx":454A
         Top             =   2865
         Width           =   3615
      End
      Begin VB.TextBox txtJSFileName 
         Height          =   315
         Left            =   45
         TabIndex        =   64
         Top             =   2550
         Width           =   3150
      End
      Begin VB.ComboBox cmbAddIns 
         Height          =   330
         ItemData        =   "frmProjProp.frx":456A
         Left            =   45
         List            =   "frmProjProp.frx":4571
         Style           =   2  'Dropdown List
         TabIndex        =   58
         Top             =   1710
         Width           =   3150
      End
      Begin SmartButtonProject.SmartButton cmdInfo 
         Height          =   360
         Left            =   3255
         TabIndex        =   59
         Top             =   1695
         Width           =   420
         _ExtentX        =   741
         _ExtentY        =   635
         Picture         =   "frmProjProp.frx":457D
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
      Begin SmartButtonProject.SmartButton cmdAddInEditor 
         Height          =   360
         Left            =   3720
         TabIndex        =   60
         Top             =   1695
         Width           =   420
         _ExtentX        =   741
         _ExtentY        =   635
         Picture         =   "frmProjProp.frx":4917
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
      Begin SmartButtonProject.SmartButton cmdEditParamValues 
         Height          =   360
         Left            =   4185
         TabIndex        =   61
         Top             =   1695
         Width           =   420
         _ExtentX        =   741
         _ExtentY        =   635
         Picture         =   "frmProjProp.frx":4CB1
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
      Begin VB.Label lblAddIneErr 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "................................................."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   165
         TabIndex        =   62
         Top             =   2040
         Width           =   2940
      End
      Begin VB.Label lblJSName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Compiled JavaScript File Name"
         Height          =   210
         Left            =   45
         TabIndex        =   63
         Top             =   2310
         Width           =   2445
      End
      Begin VB.Label lblAddIn 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "AddIns"
         Height          =   210
         Left            =   45
         TabIndex        =   57
         Top             =   1485
         Width           =   555
      End
   End
   Begin VB.Timer tmrSelPathsSection 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3285
      Top             =   5145
   End
   Begin VB.PictureBox picConfigs 
      Height          =   4920
      Left            =   7815
      ScaleHeight     =   4860
      ScaleWidth      =   6705
      TabIndex        =   15
      Top             =   300
      Width           =   6765
      Begin VB.PictureBox picConfigBtns 
         BorderStyle     =   0  'None
         Height          =   2085
         Left            =   4890
         ScaleHeight     =   2085
         ScaleWidth      =   1560
         TabIndex        =   49
         Top             =   1320
         Width           =   1560
         Begin MSComctlLib.TabStrip tsConfigs 
            Height          =   1425
            Left            =   0
            TabIndex        =   50
            Top             =   150
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   2514
            TabWidthStyle   =   2
            MultiRow        =   -1  'True
            Style           =   1
            TabFixedWidth   =   2752
            TabFixedHeight  =   556
            Separators      =   -1  'True
            TabMinWidth     =   2752
            _Version        =   393216
            BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
               NumTabs         =   4
               BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "Paths"
                  Key             =   "tsPaths"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "HotSpots Editor"
                  Key             =   "tsHSE"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "Frames Support"
                  Key             =   "tsFrames"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "FTP Information"
                  Key             =   "tsFTP"
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
         Begin VB.CommandButton cmdConfigOptions 
            Caption         =   "Options"
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
            Left            =   0
            TabIndex        =   51
            Top             =   1770
            Width           =   1560
         End
      End
      Begin VB.PictureBox picConfigsOp 
         Height          =   4050
         Index           =   0
         Left            =   30
         ScaleHeight     =   3990
         ScaleWidth      =   6300
         TabIndex        =   17
         Top             =   30
         Width           =   6360
         Begin VB.PictureBox picFTP 
            Height          =   3660
            Left            =   240
            ScaleHeight     =   3600
            ScaleWidth      =   3465
            TabIndex        =   120
            Top             =   2265
            Width           =   3525
            Begin VB.Frame frameAccountInfo 
               Caption         =   "Account Information"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1005
               Left            =   45
               TabIndex        =   133
               Top             =   1680
               Width           =   3315
               Begin VB.TextBox txtFTPPassword 
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
                  IMEMode         =   3  'DISABLE
                  Left            =   1215
                  PasswordChar    =   "*"
                  TabIndex        =   137
                  Top             =   630
                  WhatsThisHelpID =   20370
                  Width           =   1440
               End
               Begin VB.TextBox txtFTPUserName 
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
                  Left            =   1215
                  TabIndex        =   135
                  Top             =   285
                  WhatsThisHelpID =   20370
                  Width           =   1440
               End
               Begin VB.Label lblFTPPassword 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Password"
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
                  Left            =   225
                  TabIndex        =   136
                  Top             =   675
                  Width           =   690
               End
               Begin VB.Label lblFTPUserName 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Username"
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
                  Left            =   225
                  TabIndex        =   134
                  Top             =   330
                  Width           =   720
               End
            End
            Begin VB.Frame framProxyInfo 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   705
               Left            =   45
               TabIndex        =   127
               Top             =   2835
               Width           =   3315
               Begin VB.TextBox txtFTPProxyPort 
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
                  IMEMode         =   3  'DISABLE
                  Left            =   2730
                  TabIndex        =   131
                  Top             =   285
                  WhatsThisHelpID =   20370
                  Width           =   465
               End
               Begin VB.TextBox txtFTPProxyServer 
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
                  Left            =   900
                  TabIndex        =   129
                  Top             =   285
                  WhatsThisHelpID =   20370
                  Width           =   1275
               End
               Begin VB.CheckBox chkFTPUseProxy 
                  Caption         =   "Use a Proxy Server"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Left            =   90
                  TabIndex        =   128
                  Top             =   -15
                  Width           =   1680
               End
               Begin VB.Label lblFTPProxyPort 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Port"
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
                  Left            =   2370
                  TabIndex        =   132
                  Top             =   330
                  Width           =   300
               End
               Begin VB.Label lblFTPProxyAddress 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Address"
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
                  Left            =   225
                  TabIndex        =   130
                  Top             =   330
                  Width           =   585
               End
            End
            Begin VB.TextBox txtFTPHostName 
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
               Left            =   45
               TabIndex        =   121
               Top             =   285
               Width           =   3315
            End
            Begin VB.TextBox txtFTPPath 
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
               Left            =   45
               TabIndex        =   122
               Top             =   1095
               WhatsThisHelpID =   20370
               Width           =   3315
            End
            Begin VB.Label lblFTPHostName 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Hostname of the FTP server"
               Height          =   210
               Left            =   45
               TabIndex        =   126
               Top             =   60
               Width           =   2325
            End
            Begin VB.Label lblFTPPath 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Absolute path to the server's root folder"
               Height          =   210
               Left            =   45
               TabIndex        =   125
               Top             =   870
               Width           =   3345
            End
            Begin VB.Label lblFTPHostNameInfo 
               AutoSize        =   -1  'True
               BackColor       =   &H00808080&
               BackStyle       =   0  'Transparent
               Caption         =   "example: ftp://myserver.com"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00808080&
               Height          =   210
               Left            =   165
               TabIndex        =   124
               Top             =   600
               Width           =   2055
            End
            Begin VB.Label lblFTPPathInfo 
               AutoSize        =   -1  'True
               BackColor       =   &H00808080&
               BackStyle       =   0  'Transparent
               Caption         =   "example: /httpdocs/"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00808080&
               Height          =   210
               Left            =   165
               TabIndex        =   123
               Top             =   1410
               Width           =   1410
            End
         End
         Begin VB.PictureBox picHotSpotEditor 
            Height          =   3855
            Left            =   1920
            ScaleHeight     =   3795
            ScaleWidth      =   3990
            TabIndex        =   20
            Top             =   1155
            Width           =   4050
            Begin VB.TextBox txtDestFile 
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
               Left            =   45
               TabIndex        =   22
               Top             =   315
               Width           =   3315
            End
            Begin SmartButtonProject.SmartButton cmdBrowseDestFile 
               Height          =   360
               Left            =   3450
               TabIndex        =   23
               Top             =   292
               Width           =   420
               _ExtentX        =   741
               _ExtentY        =   635
               Picture         =   "frmProjProp.frx":504B
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
            Begin VB.Label lblHSDocInfo 
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00808080&
               Height          =   2415
               Left            =   165
               TabIndex        =   26
               Top             =   1200
               Width           =   4155
            End
            Begin VB.Label lblHSFileErr 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "................................................."
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   195
               Left            =   165
               TabIndex        =   25
               Top             =   825
               Width           =   2940
            End
            Begin VB.Label lblDocHS 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Document Containing the HotSpots"
               Height          =   210
               Left            =   45
               TabIndex        =   21
               Top             =   90
               Width           =   2940
            End
            Begin VB.Label Label25 
               AutoSize        =   -1  'True
               BackColor       =   &H00808080&
               BackStyle       =   0  'Transparent
               Caption         =   "example: c:\inetpub\wwwroot\header.htm"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00808080&
               Height          =   210
               Left            =   165
               TabIndex        =   24
               Top             =   645
               Width           =   3030
            End
         End
         Begin VB.PictureBox picFrames 
            Height          =   3855
            Left            =   2535
            ScaleHeight     =   3795
            ScaleWidth      =   6300
            TabIndex        =   28
            Top             =   825
            Width           =   6360
            Begin SmartButtonProject.SmartButton cmdReload 
               Height          =   360
               Left            =   3855
               TabIndex        =   33
               Top             =   832
               Width           =   420
               _ExtentX        =   741
               _ExtentY        =   635
               Picture         =   "frmProjProp.frx":51A5
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
            Begin SmartButtonProject.SmartButton cmdBrowseFramesDoc 
               Height          =   360
               Left            =   3420
               TabIndex        =   32
               Top             =   832
               Width           =   420
               _ExtentX        =   741
               _ExtentY        =   635
               Picture         =   "frmProjProp.frx":553F
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
            Begin VB.TextBox txtFramesDoc 
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
               Left            =   45
               TabIndex        =   31
               Top             =   855
               WhatsThisHelpID =   20530
               Width           =   3315
            End
            Begin VB.CheckBox chkFrameSupport 
               Caption         =   "Enable Frames Support"
               Height          =   345
               Left            =   60
               TabIndex        =   29
               Top             =   90
               Width           =   4155
            End
            Begin VB.Label lblFramesDoc 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Frames Document"
               Height          =   210
               Left            =   45
               TabIndex        =   30
               Top             =   630
               Width           =   1485
            End
            Begin VB.Label lblFramesErr 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "................................................."
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   195
               Left            =   165
               TabIndex        =   35
               Top             =   1365
               Width           =   2940
            End
            Begin VB.Label Label37 
               AutoSize        =   -1  'True
               BackColor       =   &H00808080&
               BackStyle       =   0  'Transparent
               Caption         =   "example: c:\inetpub\wwwroot\index.html"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00808080&
               Height          =   210
               Left            =   165
               TabIndex        =   34
               Top             =   1185
               Width           =   2940
            End
         End
         Begin SmartButtonProject.SmartButton cmdBrowseLocalImages 
            Height          =   360
            Index           =   0
            Left            =   3405
            TabIndex        =   45
            Top             =   2647
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   635
            Picture         =   "frmProjProp.frx":5699
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
         Begin SmartButtonProject.SmartButton cmdBrowseRootWeb 
            Height          =   360
            Index           =   0
            Left            =   3405
            TabIndex        =   37
            Top             =   735
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   635
            Picture         =   "frmProjProp.frx":57F3
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
         Begin SmartButtonProject.SmartButton cmdBrowse 
            Height          =   360
            Index           =   0
            Left            =   3405
            TabIndex        =   41
            Top             =   1695
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   635
            Picture         =   "frmProjProp.frx":594D
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
         Begin VB.ComboBox cmbLocalConfigs 
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
            Index           =   0
            ItemData        =   "frmProjProp.frx":5AA7
            Left            =   0
            List            =   "frmProjProp.frx":5AAE
            Style           =   2  'Dropdown List
            TabIndex        =   48
            Top             =   3585
            Visible         =   0   'False
            Width           =   3315
         End
         Begin VB.TextBox txtImagesPath 
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
            Index           =   0
            Left            =   0
            TabIndex        =   44
            Top             =   2670
            WhatsThisHelpID =   20370
            Width           =   3315
         End
         Begin VB.TextBox txtDest 
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
            Index           =   0
            Left            =   0
            TabIndex        =   40
            Top             =   1725
            WhatsThisHelpID =   20370
            Width           =   3315
         End
         Begin VB.TextBox txtRootWeb 
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
            Index           =   0
            Left            =   0
            TabIndex        =   36
            Top             =   765
            Width           =   3315
         End
         Begin VB.Label lblLocalConfig 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Local Configuration"
            Height          =   210
            Index           =   0
            Left            =   0
            TabIndex        =   47
            Top             =   3360
            Visible         =   0   'False
            Width           =   1545
         End
         Begin VB.Label lblConfigName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Name"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Index           =   0
            Left            =   105
            TabIndex        =   18
            Top             =   15
            Width           =   495
         End
         Begin VB.Label lblConfigDesc 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Info"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   0
            Left            =   105
            TabIndex        =   19
            Top             =   210
            Width           =   300
         End
         Begin VB.Label lblImagesPath 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Folder to Store the Images"
            Height          =   210
            Index           =   0
            Left            =   0
            TabIndex        =   43
            Top             =   2445
            Width           =   2235
         End
         Begin VB.Label lblDescImagesPath 
            AutoSize        =   -1  'True
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "example: c:\inetpub\wwwroot\menus\images\"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   210
            Index           =   0
            Left            =   120
            TabIndex        =   46
            Top             =   3015
            Width           =   3300
         End
         Begin VB.Label lblDescDest 
            AutoSize        =   -1  'True
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "example: c:\inetpub\wwwroot\menus\"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   210
            Index           =   0
            Left            =   120
            TabIndex        =   42
            Top             =   2070
            Width           =   2745
         End
         Begin VB.Label lblDescRootWeb 
            AutoSize        =   -1  'True
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "example: c:\inetpub\wwwroot\"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   210
            Index           =   0
            Left            =   120
            TabIndex        =   38
            Top             =   1095
            Width           =   2220
         End
         Begin VB.Label lblDest 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Folder to Store Compiled Files"
            Height          =   210
            Index           =   0
            Left            =   0
            TabIndex        =   39
            Top             =   1500
            Width           =   2430
         End
         Begin VB.Label lblRootWeb 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Local Path to the Root Web"
            Height          =   210
            Index           =   0
            Left            =   0
            TabIndex        =   27
            Top             =   540
            Width           =   2325
         End
         Begin VB.Shape shpConfigInfoBack 
            BackColor       =   &H00808080&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            Height          =   450
            Index           =   0
            Left            =   0
            Shape           =   4  'Rounded Rectangle
            Top             =   0
            Width           =   4080
         End
      End
      Begin MSComctlLib.TabStrip tsPublishing 
         Height          =   4485
         Left            =   -15
         TabIndex        =   16
         Top             =   -15
         Width           =   6570
         _ExtentX        =   11589
         _ExtentY        =   7911
         Placement       =   1
         ImageList       =   "ilIcons"
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   1
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Local"
               Key             =   "tsLocal"
               ImageVarType    =   2
               ImageIndex      =   4
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
   End
   Begin VB.PictureBox picGeneral 
      Height          =   4500
      Left            =   150
      ScaleHeight     =   4440
      ScaleWidth      =   6450
      TabIndex        =   1
      Top             =   480
      Width           =   6510
      Begin xfxLine3D.ucLine3D uc3DLine1 
         Height          =   30
         Left            =   300
         TabIndex        =   7
         Top             =   1125
         Width           =   6045
         _ExtentX        =   10663
         _ExtentY        =   53
      End
      Begin VB.TextBox txtProjVersion 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000011&
         Height          =   285
         Left            =   1050
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   675
         WhatsThisHelpID =   20350
         Width           =   1650
      End
      Begin VB.TextBox txtUnfoldingSound 
         Height          =   315
         Left            =   45
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   2715
         Visible         =   0   'False
         WhatsThisHelpID =   20550
         Width           =   3315
      End
      Begin VB.TextBox txtName 
         Height          =   315
         Left            =   1050
         TabIndex        =   0
         Top             =   1290
         Width           =   3330
      End
      Begin VB.TextBox txtFileName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000011&
         Height          =   285
         Left            =   1050
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   360
         WhatsThisHelpID =   20350
         Width           =   5340
      End
      Begin SmartButtonProject.SmartButton cmdBrowseSound 
         Height          =   360
         Left            =   3420
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   2692
         Visible         =   0   'False
         Width           =   420
         _ExtentX        =   741
         _ExtentY        =   635
         Picture         =   "frmProjProp.frx":5ABD
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
      Begin SmartButtonProject.SmartButton cmdPlay 
         Height          =   360
         Left            =   3885
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   2692
         Visible         =   0   'False
         Width           =   420
         _ExtentX        =   741
         _ExtentY        =   635
         Picture         =   "frmProjProp.frx":5C17
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
      Begin SmartButtonProject.SmartButton cmdRemoveSound 
         Height          =   240
         Left            =   2355
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   3075
         Visible         =   0   'False
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   423
         Caption         =   "Remove"
         Picture         =   "frmProjProp.frx":5D71
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
      Begin VB.Label lblPTitle 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Title"
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
         Left            =   600
         TabIndex        =   8
         Top             =   1350
         Width           =   300
      End
      Begin VB.Label lblPVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Version"
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
         Left            =   375
         TabIndex        =   5
         Top             =   720
         Width           =   525
      End
      Begin VB.Label lblPLocation 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Location"
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
         Left            =   300
         TabIndex        =   3
         Top             =   405
         Width           =   600
      End
      Begin VB.Label lblUnfoldingSound 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unfolding Sound"
         Height          =   210
         Left            =   45
         TabIndex        =   10
         Top             =   2475
         Visible         =   0   'False
         Width           =   1350
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "example: My Project"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   210
         Left            =   1050
         TabIndex        =   9
         Top             =   1605
         Width           =   1440
      End
      Begin VB.Label lblProjectInformation 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Project Information"
         Height          =   210
         Left            =   45
         TabIndex        =   2
         Top             =   90
         Width           =   1590
      End
   End
   Begin MSComctlLib.TabStrip tsSections 
      Height          =   4965
      Left            =   75
      TabIndex        =   68
      Top             =   90
      Width           =   6675
      _ExtentX        =   11774
      _ExtentY        =   8758
      ShowTips        =   0   'False
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "General"
            Key             =   "tsGeneral"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Configurations"
            Key             =   "tsPublishing"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Global Settings"
            Key             =   "tsGlobalSettings"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Advanced"
            Key             =   "tsAdvanced"
            ImageVarType    =   2
         EndProperty
      EndProperty
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
      Left            =   5850
      TabIndex        =   53
      Top             =   5130
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
      Left            =   4830
      TabIndex        =   52
      Top             =   5130
      Width           =   900
   End
   Begin VB.Menu mnuConfig 
      Caption         =   "mnuConfig"
      Begin VB.Menu mnuConfigAdd 
         Caption         =   "Add..."
      End
      Begin VB.Menu mnuConfigEdit 
         Caption         =   "Edit..."
      End
      Begin VB.Menu mnuConfigSep01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuConfigRemove 
         Caption         =   "Remove"
      End
      Begin VB.Menu mnuConfigSep02 
         Caption         =   "-"
      End
      Begin VB.Menu mnuConfigSetAsDefault 
         Caption         =   "Set As Default"
      End
      Begin VB.Menu mnuConfigSep03 
         Caption         =   "-"
      End
      Begin VB.Menu mnuConfigOpPaths 
         Caption         =   "Optimized Paths"
         Checked         =   -1  'True
      End
   End
End
Attribute VB_Name = "frmProjProp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim IsUpdating As Boolean
Dim ProjectBack As ProjectDef

Private Sub chkFrameSupport_Click()

    txtFramesDoc.Enabled = (chkFrameSupport.Value = vbChecked)
    cmdBrowseFramesDoc.Enabled = (chkFrameSupport.Value = vbChecked)
    cmdReload.Enabled = cmdBrowseFramesDoc.Enabled
    
    Project.UserConfigs(ppSelConfig).Frames.UseFrames = IIf(chkFrameSupport.Value = vbChecked, True, False)
      
End Sub

Private Sub chkFTPUseProxy_Click()

    If IsUpdating Then Exit Sub
    UpdateFTPInfo

End Sub

Private Sub chkGZIP_Click()

    #If LITE = 0 Then

    If IsUpdating Then Exit Sub

    If chkGZIP.Value = vbChecked Then
        With TipsSys
            .CanDisable = True
            .TipTitle = "GZIP Support and Compatibility"
            .Tip = "WARNING: In order to be able to use javascript files compressed using GZIP your server must support it." + vbCrLf + vbCrLf + _
            "It is highly recommended that you first check your server's configuration or request the help of an experienced technician to make sure that your web server will support javascript files compressed using GZIP." + vbCrLf + vbCrLf + _
            "NOTE #1: When changing this option you will need to re-install the menus using the Tools->Install Menus tool." + vbCrLf + vbCrLf + _
            "NOTE #2: The GZIPed version of the menus will only work when viewing the menus from a web server. If you try to view the menus by opening an HTML document without a web server the menus will fail load and may cause javascript errors."
            .Show
        End With
    End If
    
    #End If

End Sub

Private Sub chkRemoveIBHSCode_Click()

    If IsUpdating Then Exit Sub

    If chkRemoveIBHSCode.Value = vbChecked Then
        With TipsSys
            .CanDisable = True
            .TipTitle = "Removing Image-based HotSpot Support Code"
            .Tip = "WARNING: Do not enable this option if your project uses image-based hotspots." + vbCrLf + vbCrLf + _
            "If one (or more) groups on your project are displayed using image-based hotspots (instead of a toolbar) then you should not disable this option; otherwise, the menus will stop working." + vbCrLf + _
            "If your project uses toolbars to display the menus then it should be safe to enable this optimization."
            .Show
        End With
    End If

End Sub

Private Sub cmbAddIns_Click()

    If IsUpdating Then Exit Sub
    
    lblAddIneErr.caption = ""

    cmdAddInEditor.Enabled = cmbAddIns.ListIndex > 0
    SetEditParamsButtonState

End Sub

Private Sub SetEditParamsButtonState()

    On Error Resume Next
    Dim HasParams As Boolean
    
    LoadAddInParams cmbAddIns.Text
    HasParams = (UBound(params) > 0)
    cmdEditParamValues.Enabled = HasParams

End Sub

Private Sub cmbFX_Click()

    sldAnimSpeed.Enabled = (cmbFX.ListIndex > 0)
    lblEffectSpeed.Enabled = (cmbFX.ListIndex > 0)
    lblSlow.Enabled = (cmbFX.ListIndex > 0)
    lblFast.Enabled = (cmbFX.ListIndex > 0)

End Sub

Private Sub cmbLocalConfigs_Click(Index As Integer)

    Project.UserConfigs(ppSelConfig).LocalInfo4RemoteConfig = cmbLocalConfigs(Index).Text

End Sub

Private Sub cmdAddInEditor_Click()

#If LITE = 0 Then
    Dim oAddIn As AddInDef
    Dim SelAddIn As String
    
    SelAddIn = cmbAddIns.Text
    oAddIn = Project.AddIn
    Project.AddIn.Name = cmbAddIns.Text
    Project.AddIn.Description = GetAddInDescription(cmbAddIns.Text)
    
    frmMain.ShowAIEWarning
    frmAddInEditor.Show vbModal
    
    Project.AddIn = oAddIn
    GetAddInsList SelAddIn
    
    cmbAddIns.SetFocus
#End If

End Sub

Private Sub cmdAUE_Click()

    frmAUE.Show vbModal

End Sub

Private Sub cmdBrowseFramesDoc_Click()

    On Error GoTo ExitSub
    
    With cDlg
        .DialogTitle = GetLocalizedStr(385)
        .InitDir = GetRealLocal.RootWeb
        .filter = SupportedHTMLDocs
        .CancelError = True
        .Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly + cdlOFNLongNames + cdlOFNPathMustExist
        .ShowOpen
        txtFramesDoc.Text = .FileName
        LoadFramesDoc .FileName
    End With
    
ExitSub:
    Exit Sub

End Sub

Private Sub LoadFramesDoc(FileName As String)

    Dim i As Integer
    
    If txtFramesDoc.Text = "" Then Exit Sub
    
    IsUpdating = True
    
    With FramesInfo
        .FileName = FileName
        GetFramesInfo
        If .IsValid Then
            lblFramesErr.caption = ""
            For i = 1 To UBound(.Frames)
                If IsFrameNameInvalid(.Frames(i).Name, "top") Or _
                   IsFrameNameInvalid(.Frames(i).Name, "body") Then
                    chkFrameSupport.Value = vbUnchecked
                    chkFrameSupport_Click
                    Exit Sub
                End If
            Next i
        Else
            lblFramesErr.caption = GetLocalizedStr(388)
        End If
    End With
    
    IsUpdating = False
    
End Sub

Private Function IsFrameNameInvalid(f As String, n As String) As Boolean

    If InStr(1, f, "." + n + ".", vbTextCompare) > 0 Or Right(f, Len(n) + 1) = "." + n Then
        MsgBox GetLocalizedStr(386) + " """ + n + """." + vbCrLf + GetLocalizedStr(387), vbInformation + vbOKCancel, "Invalid Frame Name"
        IsFrameNameInvalid = True
    Else
        IsFrameNameInvalid = False
    End If

End Function

Private Sub cmdBrowseLocalImages_Click(Index As Integer)

    Dim Path As String
    
    On Error Resume Next
    
    Me.Enabled = False
    
    Path = txtRootWeb(ppSelConfig).Text
    If LenB(Dir(Path)) = 0 Or Err.number Then
        Path = ""
    Else
        Path = UnqualifyPath(Path)
    End If
    Path = BrowseForFolderByPath(Path, GetLocalizedStr(389), Me)
    
    If LenB(Path) <> 0 Then txtImagesPath(ppSelConfig).Text = AddTrailingSlash(Path, IIf(Project.UserConfigs(ppSelConfig).Type = ctcRemote, "/", "\"))
    
    Me.Enabled = True
    Me.SetFocus

End Sub

Private Sub cmdBrowseRootWeb_Click(Index As Integer)

    Dim Path As String
    Dim oPath As String
    
    On Error Resume Next
    
    Me.Enabled = False
    
    Path = txtRootWeb(ppSelConfig).Text
    oPath = Path
    If LenB(Dir(Path)) = 0 Or Err.number Then
        Path = ""
    Else
        Path = UnqualifyPath(Path)
    End If
    Path = BrowseForFolderByPath(Path, GetLocalizedStr(390), Me)
    
    If LenB(Path) <> 0 Then
        txtRootWeb(ppSelConfig).Text = AddTrailingSlash(Path, IIf(Project.UserConfigs(ppSelConfig).Type = ctcRemote, "/", "\"))
        
        Path = txtRootWeb(ppSelConfig).Text
        If Left(txtDest(ppSelConfig).Text, Len(oPath)) = oPath Then
            txtDest(ppSelConfig).Text = Replace(txtDest(ppSelConfig).Text, oPath, Path)
        End If
        If Left(txtImagesPath(ppSelConfig).Text, Len(oPath)) = oPath Then
            txtImagesPath(ppSelConfig).Text = Replace(txtImagesPath(ppSelConfig).Text, oPath, Path)
        End If
        
        If Left(txtFramesDoc.Text, Len(oPath)) = oPath Then
            txtFramesDoc.Text = Replace(txtFramesDoc.Text, oPath, Path)
        End If
        
        If Left(txtDestFile.Text, Len(oPath)) = oPath Then
            txtDestFile.Text = Replace(txtDestFile.Text, oPath, Path)
        End If
    End If
    
    Me.Enabled = True
    Me.SetFocus

End Sub

Private Sub cmdBrowseSound_Click()

    On Error GoTo ExitSub
    
    With cDlg
        .DialogTitle = GetLocalizedStr(391)
        .InitDir = txtRootWeb(ppSelConfig).Text
        .filter = SupportedAudioFiles
        .CancelError = True
        .Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly + cdlOFNLongNames + cdlOFNPathMustExist
        .ShowOpen
        txtUnfoldingSound.Text = .FileName
    End With
    
ExitSub:
    Exit Sub

End Sub

Private Sub cmdConfigOptions_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    PopupMenu mnuConfig, , cmdConfigOptions.Left + (picConfigBtns.Left + picConfigs.Left), _
                           cmdConfigOptions.Top + cmdConfigOptions.Height + (picConfigBtns.Top + picConfigs.Top)

End Sub

Private Sub cmdEditParamValues_Click()

    frmAddInParamValEditor.Show vbModal
    SaveAddInParams cmbAddIns.Text
    
    cmbAddIns.SetFocus

End Sub

Private Sub cmdFontSubst_Click()

    frmFontSubst.Show vbModal

End Sub

Private Sub cmdPosOffset_Click()

    frmPosOffset.Show vbModal

End Sub

'Private Sub cmdPlay_Click()
'
'    PlaySound ByVal txtUnfoldingSound.Text, 0&, SND_FILENAME Or SND_ASYNC Or SND_NOWAIT
'
'End Sub

Private Sub cmdRemoveSound_Click()
    
    txtUnfoldingSound.Text = ""
    txtUnfoldingSound.SetFocus

End Sub

Private Sub cmdRestore_Click()

    sldAnimSpeed.Value = 35
    sldHideDelay.Value = 200
    sldRootMenusDelay.Value = 15
    sldSubMenusDelay.Value = 150
    sldSelChangeDelay.Value = 0
    sldBlinkEffect.Value = 0
    sldBlinkSpeed.Value = 50
    
    With lvExtFeatures.ListItems
        .item(1).Checked = True
        .item(2).Checked = True
        .item(3).Checked = False
        .item(4).Checked = True
        .item(5).Checked = False
        .item(6).Checked = False
        .item(7).Checked = False
        .item(8).Checked = False
        .item(9).Checked = False
        .item(10).Checked = False
        .item(11).Checked = False
        .item(12).Checked = False
    End With

End Sub

Private Sub cmdReload_Click()

    LoadFramesDoc txtFramesDoc.Text

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF1 Then
        Select Case tsSections.SelectedItem.key
            Case "tsGeneral"
                showHelp "dialogs/pp_general.htm"
            Case "tsPublishing"
                Select Case tsConfigs.SelectedItem.key
                    Case "tsPaths"
                        showHelp "dialogs/pp_conf_paths.htm"
                    Case "tsHSE"
                        showHelp "dialogs/pp_hse.htm"
                    Case "tsFrames"
                        showHelp "dialogs/pp_frames.htm"
                    Case "tsFTP"
                        showHelp "dialogs/pp_ftp.htm"
                End Select
            Case "tsGlobalSettings"
                Select Case tsGlobalOptions.SelectedItem.key
                    Case "tsGSFX"
                        showHelp "dialogs/pp_globalsettings.htm"
                    Case "tsGSTimers"
                        showHelp "dialogs/pp_globalsettings_tmr.htm"
                    Case "tsGSExtFeat"
                        showHelp "dialogs/pp_globalsettings_ext.htm"
                End Select
            Case "tsFTP"
                showHelp "dialogs/pp_ftp.htm"
            Case "tsAdvanced"
                showHelp "dialogs/pp_advanced.htm"
        End Select
    End If

End Sub

Private Sub lvExtFeatures_ItemCheck(ByVal item As MSComctlLib.ListItem)

    UpdateRootMenusDelaySliderState
    UpdateIGSPropBtn item
    
End Sub

Private Sub UpdateIGSPropBtn(item As ListItem)

    If item.Index <> 12 Then Exit Sub

    With sbGIGSProp
        .Visible = item.Checked
        .ZOrder 0
        .Move lvExtFeatures.Left + TextWidth(item.Text) + 8 * 15, _
                lvExtFeatures.Top + (TextHeight("W") + 2 * 15) * item.Index - 15
    End With

End Sub

'Private Sub ToggleDBCSSupport(State As Boolean)
'
'    Dim i As Integer
'    Dim oSetting As Boolean
'
'    oSetting = Project.DBCSSupport
'    Project.DBCSSupport = True
'
'    For i = 1 To UBound(MenuGrps)
'        With MenuGrps(i)
'            If State Then
'                .Caption = Unicode2xUNI(.Caption)
'                .WinStatus = Unicode2xUNI(.WinStatus)
'            Else
'                .Caption = xUNI2Unicode(.Caption)
'                .WinStatus = xUNI2Unicode(.WinStatus)
'            End If
'        End With
'    Next i
'    For i = 1 To UBound(MenuCmds)
'        With MenuCmds(i)
'            If State Then
'                .Caption = Unicode2xUNI(.Caption)
'                .WinStatus = Unicode2xUNI(.WinStatus)
'            Else
'                .Caption = xUNI2Unicode(.Caption)
'                .WinStatus = xUNI2Unicode(.WinStatus)
'            End If
'        End With
'    Next i
'
'    Project.DBCSSupport = oSetting
'
'End Sub

Private Sub UpdateRootMenusDelaySliderState()

    With lvExtFeatures.ListItems
        sldRootMenusDelay.Enabled = (.item(5).Checked = False) And _
                                    (.item(8).Checked = False)
    End With
    
    If Not sldRootMenusDelay.Enabled Then sldRootMenusDelay.Value = 0
    
    UpdateSecDisp
    AdjustSMDelay

End Sub

Private Sub mnuConfigAdd_Click()

    Dim NumConf As Integer

    NumConf = UBound(Project.UserConfigs)

    frmConfigAdd.Show vbModal
    
    If NumConf <> UBound(Project.UserConfigs) Then
        CreateConfigCtrls
        tsPublishing.Tabs(tsPublishing.Tabs.Count).Selected = True
        tsPublishing_Click
    End If

End Sub

Private Sub mnuConfigEdit_Click()

    With frmConfigAdd
        .txtName.Text = Project.UserConfigs(ppSelConfig).Name
        .cmbType.ListIndex = Project.UserConfigs(ppSelConfig).Type
        .txtDesc.Text = Project.UserConfigs(ppSelConfig).Description
        .cmbConfigs.Enabled = False
        .cmbType.Enabled = False
        .caption = GetLocalizedStr(394)
        .tag = CStr(ppSelConfig)
        .Show vbModal
    End With
    
    CreateConfigCtrls
    tsPublishing_Click

End Sub

Private Sub mnuConfigOpPaths_Click()

    mnuConfigOpPaths.Checked = Not mnuConfigOpPaths.Checked
    Project.UserConfigs(ppSelConfig).OptmizePaths = mnuConfigOpPaths.Checked

End Sub

Private Sub mnuConfigRemove_Click()

    Dim i As Integer
    
    If ppSelConfig = Project.DefaultConfig Then
        MsgBox GetLocalizedStr(396), vbInformation + vbOKOnly, "Unable to Delete Configuration"
        Exit Sub
    End If
    
    If Project.UserConfigs(ppSelConfig).Type <> ctcRemote Then
        For i = 1 To UBound(Project.UserConfigs)
            If Project.UserConfigs(i).Type = ctcRemote And Project.UserConfigs(i).LocalInfo4RemoteConfig = Project.UserConfigs(ppSelConfig).Name Then
                MsgBox GetLocalizedStr(397) + " " + Project.UserConfigs(i).Name + " " + GetLocalizedStr(398), vbInformation + vbOKOnly, "Error Deleting Configuration"
                Exit Sub
            End If
        Next i
    End If
    
    For i = ppSelConfig To UBound(Project.UserConfigs) - 1
        Project.UserConfigs(i) = Project.UserConfigs(i + 1)
        shpConfigInfoBack(i) = shpConfigInfoBack(i + 1)
    Next i
    ReDim Preserve Project.UserConfigs(UBound(Project.UserConfigs) - 1)
    Unload shpConfigInfoBack(shpConfigInfoBack.Count - 1)
    
    If Project.DefaultConfig > UBound(Project.UserConfigs) Then
        Project.DefaultConfig = Project.DefaultConfig - 1
    End If
    
    CreateConfigCtrls
    tsPublishing.Tabs(1).Selected = True
    tsPublishing_Click
    
End Sub

Private Sub mnuConfigSetAsDefault_Click()

    Dim i As Integer

    Project.DefaultConfig = ppSelConfig
    
    For i = 0 To shpConfigInfoBack.Count - 1
        tsPublishing.Tabs(i + 1).Image = IIf(i = Project.DefaultConfig, 4, 0)
        shpConfigInfoBack(i).BackColor = IIf(i = ppSelConfig, &H800000, &H808080)
    Next i
    
    mnuConfigSetAsDefault.Checked = True
    
    DisplayTip GetLocalizedStr(682), GetLocalizedStr(683)

End Sub

Private Sub sbGIGSProp_Click()
  
    AddMenuGroup , True
    With MenuGrps(GetID)
        .caption = "Global In-Group"
        .Name = "mnuGrpGIS_771248"
        .scrolling = Project.AutoScroll
    End With
    frmGrpScrolling.Show vbModal
    Project.AutoScroll = MenuGrps(GetID).scrolling
    frmMain.RemoveItem True

End Sub

Private Sub sldBlinkEffect_Change()

    sldBlinkSpeed.Enabled = (sldBlinkEffect.Value > 0)

End Sub

Private Sub sldHideDelay_Change()

    UpdateSecDisp

End Sub

Private Sub sldRootMenusDelay_Change()

    AdjustSMDelay
    UpdateSecDisp

End Sub

Private Sub sldSelChangeDelay_Change()

    UpdateSecDisp

End Sub

Private Sub sldSubMenusDelay_Change()

    AdjustSMDelay
    UpdateSecDisp

End Sub

Private Sub AdjustSMDelay()

    Dim min As Integer
    
    If sldRootMenusDelay.Enabled Then
        If sldRootMenusDelay.Value > sldSubMenusDelay.Value Then
            min = sldSubMenusDelay.Value
        Else
            min = sldRootMenusDelay.Value
        End If
    Else
        min = sldSubMenusDelay.Value
    End If
    
    With sldSelChangeDelay
        If .Value > min Then .Value = min
        .min = 0
        .Enabled = (min > 0)
        If .Enabled Then
            .Max = min
        Else
            .Max = 1
        End If
    End With

End Sub

Private Sub tmrSelPathsSection_Timer()

    tmrSelPathsSection.Enabled = False
    AutoSelectPaths

End Sub

Private Sub tsConfigs_Click()

    picHotSpotEditor.Visible = False
    Set picHotSpotEditor.Container = picConfigsOp(ppSelConfig)
    picFrames.Visible = False
    Set picFrames.Container = picConfigsOp(ppSelConfig)
    picFTP.Visible = False
    Set picFTP.Container = picConfigsOp(ppSelConfig)

    Select Case tsConfigs.SelectedItem.key
        Case "tsPaths"
        Case "tsHSE"
            picHotSpotEditor.Visible = True
        Case "tsFrames"
            #If LITE = 1 Then
                frmMain.ShowLITELImitationInfo 4
                tsConfigs.Tabs("tsPaths").Selected = True
            #Else
                picFrames.Visible = True
            #End If
        Case "tsFTP"
            picFTP.Visible = True
    End Select
    
    mnuConfigOpPaths.Enabled = Project.UserConfigs(ppSelConfig).Type = ctcRemote

End Sub

Private Sub AutoSelectPaths()

    tsConfigs.Tabs("tsPaths").Selected = True
    tsConfigs_Click

End Sub

Private Sub tsGlobalOptions_Click()

    On Error Resume Next
    
    sbGIGSProp.ZOrder 1

    Select Case tsGlobalOptions.SelectedItem.key
        Case "tsGSFX"
            picGSFX.ZOrder 0
            cmbFX.SetFocus
        Case "tsGSTimers"
            picGSTimers.ZOrder 0
            sldRootMenusDelay.SetFocus
        Case "tsGSExtFeat"
            picGSExtFeat.ZOrder 0
            lvExtFeatures.SetFocus
            sbGIGSProp.ZOrder 0
    End Select

End Sub

Private Sub tsPublishing_Click()
    
    On Error Resume Next
    
    Dim i As Integer
    
    ppSelConfig = tsPublishing.SelectedItem.Index - 1
    
    picConfigsOp(ppSelConfig).ZOrder 0
    txtRootWeb(ppSelConfig).SetFocus
    picConfigBtns.ZOrder 0
    tsConfigs_Click
    
    mnuConfigEdit.Enabled = ppSelConfig > 0
    mnuConfigRemove.Enabled = ppSelConfig > 0
    
    Select Case Project.UserConfigs(ppSelConfig).Type
        Case ctcRemote
            tsConfigs.Tabs.Remove tsConfigs.Tabs("tsHSE").Index
            tsConfigs.Tabs.Remove tsConfigs.Tabs("tsFrames").Index
        Case Else
            tsConfigs.Tabs.Add , "tsHSE", GetLocalizedStr(335)
            tsConfigs.Tabs.Add , "tsFrames", GetLocalizedStr(336)
    End Select
    
    With Project.UserConfigs(ppSelConfig)
        If .Type = ctcRemote Then AutoSelectPaths
        
        txtRootWeb(ppSelConfig).Text = .RootWeb
        txtDest(ppSelConfig).Text = .CompiledPath
        txtImagesPath(ppSelConfig).Text = .ImagesPath
        
        txtDestFile.Text = .HotSpotEditor.HotSpotsFile
        ChkHSFile
        
        cmbLocalConfigs(ppSelConfig).Clear
        cmbLocalConfigs(ppSelConfig).AddItem "(" + GetLocalizedStr(400) + ") Local"
        For i = 1 To UBound(Project.UserConfigs)
            If Project.UserConfigs(i).Type <> ctcRemote Then
                cmbLocalConfigs(ppSelConfig).AddItem Project.UserConfigs(i).Name
                If .LocalInfo4RemoteConfig = cmbLocalConfigs(ppSelConfig).List(cmbLocalConfigs(ppSelConfig).NewIndex) Then
                    cmbLocalConfigs(ppSelConfig).ListIndex = cmbLocalConfigs(ppSelConfig).NewIndex
                    Exit For
                End If
            End If
        Next i
        If cmbLocalConfigs(ppSelConfig).ListIndex = -1 Then
            cmbLocalConfigs(ppSelConfig).ListIndex = 0
        End If
        
        chkFrameSupport.Value = (IIf(.Frames.UseFrames, vbChecked, vbUnchecked))
        If .Frames.UseFrames Then
            txtFramesDoc.Enabled = True
            cmdBrowseFramesDoc.Enabled = True
            txtFramesDoc.Text = .Frames.FramesFile
            LoadFramesDoc .Frames.FramesFile
        Else
            txtFramesDoc.Enabled = False
            cmdBrowseFramesDoc.Enabled = False
            txtFramesDoc.Text = ""
        End If
        cmdReload.Enabled = cmdBrowseFramesDoc.Enabled
        
        If ppSelConfig = Project.DefaultConfig Then
            shpConfigInfoBack(ppSelConfig).BackColor = &H800000
            mnuConfigSetAsDefault.Checked = True
        Else
            shpConfigInfoBack(ppSelConfig).BackColor = &H808080
            mnuConfigSetAsDefault.Checked = False
        End If
        
        mnuConfigOpPaths.Checked = .OptmizePaths
                
        Dim ftpInfo() As String
        If .FTP = "" Then
            ReDim ftpInfo(6)
        Else
            ftpInfo = Split(.FTP, "*")
        End If
        If UBound(ftpInfo) >= 0 Then txtFTPHostName.Text = ftpInfo(0)
        If UBound(ftpInfo) >= 1 Then txtFTPPath.Text = ftpInfo(1)
        If UBound(ftpInfo) >= 2 Then txtFTPUserName.Text = ftpInfo(2)
        If UBound(ftpInfo) >= 3 Then txtFTPPassword.Text = ftpInfo(3)
        If UBound(ftpInfo) >= 4 Then chkFTPUseProxy.Value = Val(ftpInfo(4))
        If UBound(ftpInfo) >= 5 Then txtFTPProxyServer.Text = ftpInfo(5)
        If UBound(ftpInfo) >= 6 Then txtFTPProxyPort.Text = ftpInfo(6)
        
        UpdateFTPProxyCtrls
    End With
    
End Sub

Private Sub tsSections_Click()

    On Error Resume Next
    
    Select Case tsSections.SelectedItem.key
        Case "tsGeneral"
            picGeneral.ZOrder 0
            txtName.SetFocus
        Case "tsAdvanced"
            picAdvanced.ZOrder 0
            ucCodeOp.SetFocus
        Case "tsPublishing"
            picConfigs.ZOrder 0
            tsPublishing_Click
        'Case "tsFTP"
        '    cmbRemoteConfigs_Click
        '    picFTP.ZOrder 0
        '    txtFTPAddress.SetFocus
        Case "tsGlobalSettings"
            If Not IsDEMO Then frmMain.ChkRegInfo True
            picGlobalSettings.ZOrder 0
            tsGlobalOptions_Click
            cmbFX.SetFocus
    End Select

End Sub

Private Sub cmdBrowse_Click(Index As Integer)

    Dim Path As String
    
    On Error Resume Next
    
    Me.Enabled = False
    
    Path = txtRootWeb(ppSelConfig).Text
    If LenB(Dir(Path)) = 0 Or Err.number Then
        Path = ""
    Else
        Path = UnqualifyPath(Path)
    End If
    Path = BrowseForFolderByPath(Path, GetLocalizedStr(401), Me)
    
    If LenB(Path) <> 0 Then txtDest(ppSelConfig).Text = AddTrailingSlash(Path, IIf(Project.UserConfigs(ppSelConfig).Type = ctcRemote, "/", "\"))
    
    Me.Enabled = True
    Me.SetFocus
    
End Sub

Private Sub cmdBrowseDestFile_Click()

    On Error GoTo ExitSub
    
    With cDlg
        .DialogTitle = GetLocalizedStr(402)
        .InitDir = GetRealLocal.RootWeb
        .filter = SupportedHTMLDocs
        .CancelError = True
        .Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly + cdlOFNLongNames + cdlOFNPathMustExist
        .ShowOpen
        txtDestFile.Text = .FileName
    End With
    
ExitSub:
    Exit Sub
    
End Sub

Private Sub ChkHSFile()

    If FileExists(txtDestFile.Text) Or LenB(txtDestFile.Text) = 0 Then
        lblHSFileErr.caption = ""
    Else
        lblHSFileErr.caption = GetLocalizedStr(403)
    End If

End Sub

Private Sub cmdCancel_Click()

    Project = ProjectBack
    Unload Me

End Sub

Private Sub cmdInfo_Click()

    Dim d As String

    If cmbAddIns.ListIndex = 0 Then
        MsgBox GetLocalizedStr(405), vbInformation + vbOKOnly, "AddIn Information"
    Else
        d = GetAddInDescription(cmbAddIns.Text)
        If Trim(d) = "" Then d = "This AddIn does not contain any additional information"
        MsgBox d, vbInformation + vbOKOnly, "AddIn Information"
    End If
    
    cmbAddIns.SetFocus

End Sub

Private Sub cmdOK_Click()

    On Error Resume Next

    Dim i As Integer
    Dim g As Integer

    With Project
        .Name = txtName.Text
        .FX = cmbFX.ListIndex
        .UnfoldingSound.onmouseover = txtUnfoldingSound.Text
        .CodeOptimization = ucCodeOp.Value
        .RemoveImageAutoPosCode = (chkRemoveIBHSCode.Value = vbChecked)
        .UseGZIP = (chkGZIP.Value = vbChecked)
        
        If cmbAddIns.ListIndex = 0 Then
            .AddIn.Name = ""
            .AddIn.Description = ""
        Else
            .AddIn.Name = cmbAddIns.Text
            .AddIn.Description = GetAddInDescription(cmbAddIns.Text)
        End If
        LoadAddInParams cmbAddIns.Text
        
        .AnimSpeed = sldAnimSpeed.Value
        .HideDelay = sldHideDelay.Value
        .SubMenusDelay = sldSubMenusDelay.Value
        .RootMenusDelay = sldRootMenusDelay.Value
        .SelChangeDelay = sldSelChangeDelay.Value
        .BlinkEffect = sldBlinkEffect.Value
        .BlinkSpeed = (sldBlinkSpeed.Max - 10) - sldBlinkSpeed.Value
        
        With lvExtFeatures.ListItems
            Project.CompileIECode = .item(1).Checked
            Project.CompileNSCode = .item(2).Checked
            Project.CompilehRefFile = .item(3).Checked
            Project.DoFormsTweak = .item(4).Checked
            Project.DWSupport = .item(5).Checked
            Project.NS4ClipBug = .item(6).Checked
            'Project.OPHelperFunctions = .Item(7).Checked
            Project.ImageReadySupport = .item(7).Checked
            Project.LotusDominoSupport = .item(8).Checked
            #If LITE = 0 Then
            Project.AutoSelFunction = .item(9).Checked
            Project.KeyboardSupport = .item(10).Checked
            'Project.DBCSSupport = .Item(11).Checked
            Project.StatusTextDisplay = 2 * Abs(.item(11).Checked)
            Project.AutoScroll.maxHeight = Abs(.item(12).Checked)
            Project.SEOTweak = .item(13).Checked
            #Else
            Project.SEOTweak = .item(9).Checked
            #End If
            
        End With
        
        .JSFileName = txtJSFileName.Text: If LenB(.JSFileName) = 0 Then .JSFileName = "menu"
        
        For i = 0 To UBound(.UserConfigs)
            With .UserConfigs(i)
                .Frames.UseFrames = .Frames.UseFrames And LenB(.Frames.FramesFile) <> 0
                Select Case .Type
                    Case ctcLocal, ctcCDROM
                        .CompiledPath = AddTrailingSlash(.CompiledPath, "\")
                        .ImagesPath = AddTrailingSlash(.ImagesPath, "\")
                        .RootWeb = AddTrailingSlash(.RootWeb, "\")
                    Case ctcRemote
                        If Left(.CompiledPath, Len(.RootWeb)) = .RootWeb Then .CompiledPath = Mid(.CompiledPath, Len(.RootWeb))
                        If Left(.ImagesPath, Len(.RootWeb)) = .RootWeb Then .ImagesPath = Mid(.ImagesPath, Len(.RootWeb))
                        
                        .CompiledPath = FixPathSlashes(.CompiledPath, "/")
                        .ImagesPath = FixPathSlashes(.ImagesPath, "/")
                        .RootWeb = AddTrailingSlash(.RootWeb, "/")
                        
                        .HotSpotEditor.HotSpotsFile = Project.UserConfigs(GetConfigID(.LocalInfo4RemoteConfig)).HotSpotEditor.HotSpotsFile
                        .Frames.UseFrames = Project.UserConfigs(GetConfigID(.LocalInfo4RemoteConfig)).Frames.UseFrames
                        .Frames.FramesFile = Project.UserConfigs(GetConfigID(.LocalInfo4RemoteConfig)).Frames.FramesFile
                End Select
            End With
        Next i
        
    End With
    
    ReplaceNewPathsOnLinks
    UpdateItemsLinks
    
    frmMain.SaveState GetLocalizedStr(406)

    Unload Me

End Sub

Private Sub ReplaceNewPathsOnLinks()

    Dim i As Integer
    Dim BackConfig As ConfigDef
    Dim ThisConfig As ConfigDef
    
    BackConfig = ProjectBack.UserConfigs(ProjectBack.DefaultConfig)
    ThisConfig = Project.UserConfigs(Project.DefaultConfig)
    
    For i = 1 To UBound(MenuCmds)
        With MenuCmds(i).Actions
            .onmouseover.url = Replace(.onmouseover.url, BackConfig.RootWeb, ThisConfig.RootWeb)
            .onclick.url = Replace(.onclick.url, BackConfig.RootWeb, ThisConfig.RootWeb)
            .OnDoubleClick.url = Replace(.OnDoubleClick.url, BackConfig.RootWeb, ThisConfig.RootWeb)
        End With
    Next i
    
    For i = 1 To UBound(MenuGrps)
        With MenuGrps(i).Actions
            .onmouseover.url = Replace(.onmouseover.url, BackConfig.RootWeb, ThisConfig.RootWeb)
            .onclick.url = Replace(.onclick.url, BackConfig.RootWeb, ThisConfig.RootWeb)
            .OnDoubleClick.url = Replace(.OnDoubleClick.url, BackConfig.RootWeb, ThisConfig.RootWeb)
        End With
    Next i

End Sub

Private Function FixPathSlashes(d As String, s As String) As String

    If InStr(d, "://") = 0 Then
        If Right$(d, 1) <> s Then d = d + s
        If Left$(d, 1) <> s Then d = s + d
        If s = "/" Then
            d = SetSlashDir(d, sdFwd)
        Else
            d = SetSlashDir(d, sdBack)
        End If
    End If
    
    FixPathSlashes = d

End Function

Private Sub Form_Load()

    On Error GoTo showError

    LocalizeUI

    ProjectBack = Project

    mnuConfig.Visible = False

    Width = 6960
    Height = 5625 + GetClientTop(Me.hwnd)
    
    InitializeDialog
    
    tsPublishing.Tabs("tsLocal").Selected = True
    Select Case ProjectPropertiesPage
        Case pppcGeneral
            tsSections.Tabs("tsGeneral").Selected = True
        Case pppcConfig
            tsSections.Tabs("tsPublishing").Selected = True
        Case pppcGlobal
            tsSections.Tabs("tsGlobalSettings").Selected = True
        Case pppcAdvanced
            tsSections.Tabs("tsAdvanced").Selected = True
    End Select
    
    #If LITE = 1 Then
        lblAddIn.Visible = False
        cmbAddIns.Visible = False
        cmdInfo.Visible = False
        cmdAddInEditor.Visible = False
        cmdEditParamValues.Visible = False
        cmdPosOffset.Visible = False
        cmdFontSubst.Visible = False
        cmdAUE.Visible = False
        chkGZIP.Enabled = False
        chkGZIP.Value = vbUnchecked
    #End If
    
    Exit Sub
    
showError:
    MsgBox "Unexpected Error " & Err.number & ": " & Err.Description

End Sub

Private Sub CreateConfigCtrls()

    On Error Resume Next

    Dim i As Integer
    Dim j As Integer
    Dim IsLocal As Boolean
    
    IsUpdating = True
    
    'cmbRemoteConfigs.Clear
    'cmbRemoteConfigs.AddItem GetLocalizedStr(110)
    'cmbRemoteConfigs.ListIndex = 0
    
    For i = tsPublishing.Tabs.Count To 2 Step -1
        tsPublishing.Tabs.Remove tsPublishing.Tabs(i).Index
    Next i
    
    tsPublishing.Tabs(1).Image = IIf(Project.DefaultConfig = 0, 4, 0)
    tsPublishing.Tabs(1).caption = "Default" ' GetLocalizedStr(253)
    For i = 1 To UBound(Project.UserConfigs)
        tsPublishing.Tabs.Add , , Project.UserConfigs(i).Name, IIf(i = Project.DefaultConfig, 4, 0)
        
        Load picConfigsOp(i)
        picConfigsOp(i).Visible = True
        
        IsLocal = Project.UserConfigs(i).Type <> ctcRemote
        
        ' ----- CONFIG INFO -------
        
        Load lblConfigName(i)
        With lblConfigName(i)
            Set .Container = picConfigsOp(i)
            .Visible = True
            .Move lblConfigName(0).Left, lblConfigName(0).Top
            .caption = Project.UserConfigs(i).Name + " (" + ConfigTypeName(Project.UserConfigs(i)) + ")"
        End With
        
        Load lblConfigDesc(i)
        With lblConfigDesc(i)
            Set .Container = picConfigsOp(i)
            .Visible = True
            .Move lblConfigDesc(0).Left, lblConfigDesc(0).Top
            .caption = Project.UserConfigs(i).Description
        End With
        
        Load shpConfigInfoBack(i)
        With shpConfigInfoBack(i)
            Set .Container = picConfigsOp(i)
            .Visible = True
            .Move shpConfigInfoBack(0).Left, shpConfigInfoBack(0).Top, picConfigsOp(0).Width
        End With
        
        ' ----- ROOT WEB -------
        
        Load lblRootWeb(i)
        With lblRootWeb(i)
            Set .Container = picConfigsOp(i)
            .Visible = True
            .Move lblRootWeb(0).Left, lblRootWeb(0).Top
            .caption = IIf(IsLocal, GetLocalizedStr(326), GetLocalizedStr(330))
        End With
        
        Load txtRootWeb(i)
        With txtRootWeb(i)
            Set .Container = picConfigsOp(i)
            .Visible = True
            .Move txtRootWeb(0).Left, txtRootWeb(0).Top
        End With
        
        Load lblDescRootWeb(i)
        With lblDescRootWeb(i)
            Set .Container = picConfigsOp(i)
            .Visible = True
            .Move lblDescRootWeb(0).Left, lblDescRootWeb(0).Top
            .caption = IIf(IsLocal, GetLocalizedStr(333) + ": c:\inetpub\wwwroot\", GetLocalizedStr(333) + ": http://myweb.com/")
        End With
        
        Load cmdBrowseRootWeb(i)
        With cmdBrowseRootWeb(i)
            Set .Container = picConfigsOp(i)
            .Visible = IIf(IsLocal, True, False)
            .Move cmdBrowseRootWeb(0).Left, cmdBrowseRootWeb(0).Top
        End With
        
        ' ----- DEST FOLDER -------
        
        Load lblDest(i)
        With lblDest(i)
            Set .Container = picConfigsOp(i)
            .Visible = True
            .Move lblDest(0).Left, lblDest(0).Top
            .caption = IIf(IsLocal, GetLocalizedStr(327), GetLocalizedStr(331))
        End With
        
        Load txtDest(i)
        With txtDest(i)
            Set .Container = picConfigsOp(i)
            .Visible = True
            .Move txtDest(0).Left, txtDest(0).Top
        End With
        
        Load lblDescDest(i)
        With lblDescDest(i)
            Set .Container = picConfigsOp(i)
            .Visible = True
            .Move lblDescDest(0).Left, lblDescDest(0).Top
            .caption = IIf(IsLocal, GetLocalizedStr(333) + ": c:\inetpub\wwwroot\menus\", GetLocalizedStr(333) + ": /menus/")
        End With
        
        Load cmdBrowse(i)
        With cmdBrowse(i)
            Set .Container = picConfigsOp(i)
            .Visible = IIf(IsLocal, True, False)
            .Move cmdBrowse(0).Left, cmdBrowse(0).Top
        End With
        
        ' ----- IMGs PATH -------
        
        Load lblImagesPath(i)
        With lblImagesPath(i)
            Set .Container = picConfigsOp(i)
            .Visible = True
            .Move lblImagesPath(0).Left, lblImagesPath(0).Top
            .caption = IIf(IsLocal, GetLocalizedStr(328), GetLocalizedStr(332))
        End With
        
        Load txtImagesPath(i)
        With txtImagesPath(i)
            Set .Container = picConfigsOp(i)
            .Visible = True
            .Move txtImagesPath(0).Left, txtImagesPath(0).Top
        End With
        
        Load lblDescImagesPath(i)
        With lblDescImagesPath(i)
            Set .Container = picConfigsOp(i)
            .Visible = True
            .Move lblDescImagesPath(0).Left, lblDescImagesPath(0).Top
            .caption = IIf(IsLocal, GetLocalizedStr(333) + ": c:\inetpub\wwwroot\menus\images\", GetLocalizedStr(333) + ": /menus/images/")
        End With
        
        Load cmdBrowseLocalImages(i)
        With cmdBrowseLocalImages(i)
            Set .Container = picConfigsOp(i)
            .Visible = IIf(IsLocal, True, False)
            .Move cmdBrowseLocalImages(0).Left, cmdBrowseLocalImages(0).Top
        End With
        
        ' ----- LOCAL CONFIG 4 REMOTE CONFIGS -------
        
        Load lblLocalConfig(i)
        With lblLocalConfig(i)
            Set .Container = picConfigsOp(i)
            .Visible = Project.UserConfigs(i).Type = ctcRemote
            .Move lblLocalConfig(0).Left, lblLocalConfig(0).Top
        End With
        
        Load cmbLocalConfigs(i)
        With cmbLocalConfigs(i)
            Set .Container = picConfigsOp(i)
            .Visible = Project.UserConfigs(i).Type = ctcRemote
            .Move cmbLocalConfigs(0).Left, cmbLocalConfigs(0).Top
            .ListIndex = 0
        End With
    Next i
    
    IsUpdating = False

End Sub

Private Function GetProjectVersion() As String

    Dim v As String
    Dim Prj As ProjectDef

    Prj = GetProjectProperties(Project.FileName, False)
    Close #ff
    
    v = Prj.version
    v = Left(v, 1) + "." & Val(Mid(v, 2, 2)) & "." + Right(v, 3)
    
    GetProjectVersion = v

End Function

Private Sub InitializeDialog()

    Dim i As Integer
    Dim nItem As ListItem

    lblFramesErr.caption = ""
    lblHSFileErr.caption = ""
    lblAddIneErr.caption = ""
    
    If FileExists(Project.FileName) Then
        txtProjVersion.Text = GetProjectVersion
    Else
        txtProjVersion.Text = ""
    End If
    
    CreateConfigCtrls
    
    IsUpdating = True
    
    With picGeneral
        .BorderStyle = 0
        .ZOrder 0
        picAdvanced.Move .Left, .Top, .Width, .Height
        picAdvanced.BorderStyle = 0
        picConfigs.Move .Left, .Top, .Width, .Height
        picConfigs.BorderStyle = 0
        'picFTP.Move .Left, .Top, .Width, .Height
        'picFTP.BorderStyle = 0
        picGlobalSettings.Move .Left, .Top, .Width, .Height
        picGlobalSettings.BorderStyle = 0
    End With
    
    With picHotSpotEditor
        .Move -45, 450, 4425, 4000
        .BorderStyle = 0
        .Visible = False
        picFrames.Move .Left, .Top, .Width, .Height
        picFrames.BorderStyle = 0
        picFrames.Visible = False
        
        picFTP.Move .Left, .Top, .Width, .Height
        picFTP.BorderStyle = 0
        picFTP.Visible = False
    End With
    
    With picConfigsOp(0)
        .BorderStyle = 0
        .ZOrder 0
        For i = 1 To UBound(Project.UserConfigs)
            picConfigsOp(i).Move .Left, .Top, .Width, .Height
            picConfigsOp(i).BorderStyle = 0
        Next i
        shpConfigInfoBack(0).Width = .Width
    End With
    
    With picGSFX
        .Move 30, 30, 6360, 4050 - 240
        .BorderStyle = 0
        .ZOrder 0
        picGSTimers.Move .Left, .Top, .Width, .Height
        picGSTimers.BorderStyle = 0
        picGSExtFeat.Move .Left, .Top, .Width, .Height
        picGSExtFeat.BorderStyle = 0
    End With
    
    With Project
        txtFileName.Text = .FileName
        txtName.Text = .Name
        
        cmbFX.ListIndex = .FX
        txtUnfoldingSound.Text = .UnfoldingSound.onmouseover
        ucCodeOp.Value = .CodeOptimization
        If Not RequiresImageCode Then
            chkRemoveIBHSCode.Value = IIf(.RemoveImageAutoPosCode, vbChecked, vbUnchecked)
        Else
            chkRemoveIBHSCode.Enabled = False
        End If
        chkGZIP.Value = IIf(.UseGZIP, vbChecked, vbUnchecked)
        For i = 0 To UBound(.UserConfigs)
            ppSelConfig = i
            With .UserConfigs(i)
                txtRootWeb(i).Text = .RootWeb
                txtDest(i).Text = .CompiledPath
                txtImagesPath(i).Text = .ImagesPath
                lblConfigName(i).caption = .Name + " (" + ConfigTypeName(Project.UserConfigs(i)) + ")"
                lblConfigDesc(i).caption = .Description
                tsPublishing.Tabs(i + 1).caption = .Name
            End With
        Next i
        
        GetAddInsList .AddIn.Name
        
        txtJSFileName.Text = .JSFileName
        UpdateJSFileNames
        
        With lvExtFeatures.ListItems
            Set nItem = .Add(, , GetLocalizedStr(733))
            nItem.Checked = Project.CompileIECode
            
            Set nItem = .Add(, , GetLocalizedStr(734))
            nItem.Checked = Project.CompileNSCode
            
            Set nItem = .Add(, , GetLocalizedStr(735))
            nItem.Checked = Project.CompilehRefFile
            
            Set nItem = .Add(, , GetLocalizedStr(737))
            nItem.Checked = Project.DoFormsTweak
            
            Set nItem = .Add(, , GetLocalizedStr(738))
            nItem.Checked = Project.DWSupport
            
            Set nItem = .Add(, , GetLocalizedStr(739))
            nItem.Checked = Project.NS4ClipBug
            
            'Set nItem = .Add(, , GetLocalizedStr(740))
            'nItem.Checked = Project.OPHelperFunctions
            
            Set nItem = .Add(, , GetLocalizedStr(741))
            nItem.Checked = Project.ImageReadySupport
            
            Set nItem = .Add(, , GetLocalizedStr(961))
            nItem.Checked = Project.LotusDominoSupport
            
            #If LITE = 0 Then
            Set nItem = .Add(, , GetLocalizedStr(962))
            nItem.Checked = Project.AutoSelFunction
            
            Set nItem = .Add(, , GetLocalizedStr(963))
            nItem.Checked = Project.KeyboardSupport
            
            'Set nItem = .Add(, , "Unicode Support")
            'nItem.Checked = Project.DBCSSupport
            
            Set nItem = .Add(, , "Display Status Text as a tooltip")
            nItem.Checked = Project.StatusTextDisplay = socBoth Or Project.StatusTextDisplay = socTooltip
            
            Set nItem = .Add(, , "Enable Global In-Group Scrolling")
            nItem.Checked = (Project.AutoScroll.maxHeight <> 0)
            UpdateIGSPropBtn nItem
            #End If
            
            Set nItem = .Add(, , "Enable Search Engine Optimizations")
            nItem.Checked = Project.SEOTweak
        End With
        CoolListView lvExtFeatures
        
        sldAnimSpeed.Value = .AnimSpeed
        sldHideDelay.Value = .HideDelay
        sldRootMenusDelay.Value = .RootMenusDelay
        sldSubMenusDelay.Value = .SubMenusDelay
        UpdateRootMenusDelaySliderState
        sldSelChangeDelay.Value = .SelChangeDelay
        
        sldBlinkEffect.Value = .BlinkEffect
        sldBlinkSpeed.Value = (sldBlinkSpeed.Max - 10) - .BlinkSpeed
    End With
    
    #If DEVVER = 1 Then
        chkRemoveIBHSCode.Enabled = Not frmMain.mnuToolsDynAPI.Checked
    #End If
    
    IsUpdating = False
    
    CenterForm Me
    SetupCharset Me

End Sub

Private Sub GetAddInsList(AddInName As String)

    Dim fName As String
    Dim sName As String

    CenterForm Me
    
    cmbAddIns.Clear
    cmbAddIns.AddItem GetLocalizedStr(110)
    cmbAddIns.ListIndex = 0
    
    IsUpdating = True
    
    fName = Dir(AppPath + "AddIns\*.ext")
    Do Until LenB(fName) = 0
        sName = Left$(fName, InStrRev(fName, ".") - 1)
        cmbAddIns.AddItem sName
        If sName = AddInName Then
            cmbAddIns.ListIndex = cmbAddIns.ListCount - 1
        End If
        fName = Dir
    Loop
    
    IsUpdating = False
    
    If LenB(AddInName) <> 0 And cmbAddIns.ListIndex = 0 Then
        lblAddIneErr.caption = GetLocalizedStr(407) + " " + AddInName + " " + GetLocalizedStr(408)
    End If
    cmdAddInEditor.Enabled = cmbAddIns.ListIndex > 0
    SetEditParamsButtonState
    
    ResizeComboList cmbAddIns

End Sub

Private Sub txtDest_Change(Index As Integer)

    If IsUpdating Then Exit Sub
    Project.UserConfigs(ppSelConfig).CompiledPath = txtDest(Index).Text

End Sub

Private Sub txtDestFile_Change()

    Project.UserConfigs(ppSelConfig).HotSpotEditor.HotSpotsFile = txtDestFile.Text
    ChkHSFile

End Sub

Private Sub txtFramesDoc_Change()

    Project.UserConfigs(ppSelConfig).Frames.FramesFile = txtFramesDoc.Text

End Sub

Private Sub txtFTPHostName_Change()

    If IsUpdating Then Exit Sub
    UpdateFTPInfo

End Sub

Private Sub UpdateFTPInfo()

    Project.UserConfigs(ppSelConfig).FTP = txtFTPHostName.Text + "*" + _
                                            txtFTPPath + "*" + _
                                            txtFTPUserName.Text + "*" + _
                                            txtFTPPassword.Text + "*" + _
                                            CStr(chkFTPUseProxy.Value) + "*" + _
                                            txtFTPProxyServer.Text + "*" + _
                                            txtFTPProxyPort.Text
    UpdateFTPProxyCtrls
    
End Sub

Private Sub UpdateFTPProxyCtrls()

    txtFTPProxyServer.Enabled = (chkFTPUseProxy.Value = vbChecked)
    txtFTPProxyPort.Enabled = txtFTPProxyServer.Enabled
    lblFTPProxyAddress.Enabled = txtFTPProxyServer.Enabled
    lblFTPProxyPort.Enabled = txtFTPProxyServer.Enabled

End Sub

Private Sub txtFTPHostName_LostFocus()

    If IsUpdating Then Exit Sub
    
    With txtFTPHostName
        If .Text <> "" Then
            Dim ss As Integer
            Dim sl As Integer
            
            ss = .SelStart
            sl = .SelLength
            
            If Left(.Text, 6) <> "ftp://" Then
                .Text = "ftp://" + .Text
                ss = ss + 6
            End If
            
            Dim i As Integer
            Dim p As String
            i = InStr(8, .Text, "/")
            If i > 0 Then
                p = Mid(.Text, i)
                If Len(p) > 1 Then
                    txtFTPPath.Text = p
                    .Text = Left(.Text, i)
                End If
            End If
            
            .SelStart = ss
            .SelLength = sl
        End If
    End With

End Sub

Private Sub txtFTPPassword_Change()

    If IsUpdating Then Exit Sub
    UpdateFTPInfo

End Sub

Private Sub txtFTPPath_Change()

    If IsUpdating Then Exit Sub
    UpdateFTPInfo

End Sub

Private Sub txtFTPPath_LostFocus()

    If IsUpdating Then Exit Sub
    
    With txtFTPPath
        If .Text <> "" Then
            Dim ss As Integer
            Dim sl As Integer
            
            ss = .SelStart
            sl = .SelLength
            
            .Text = AddTrailingSlash(SetSlashDir(.Text, sdFwd), "/")
            
            .SelStart = ss
            .SelLength = sl
        End If
    End With
    
End Sub

Private Sub txtFTPProxyPort_Change()

    If IsUpdating Then Exit Sub
    UpdateFTPInfo

End Sub

Private Sub txtFTPProxyServer_Change()

    If IsUpdating Then Exit Sub
    UpdateFTPInfo

End Sub

Private Sub txtFTPUserName_Change()

    If IsUpdating Then Exit Sub
    UpdateFTPInfo

End Sub

Private Sub txtImagesPath_Change(Index As Integer)

    If IsUpdating Then Exit Sub
    Project.UserConfigs(ppSelConfig).ImagesPath = txtImagesPath(Index).Text

End Sub

Private Sub txtJSFileName_Change()

    UpdateJSFileNames

End Sub

Private Sub txtJSFileName_KeyPress(KeyAscii As Integer)

    Dim InvalidChars As String
    InvalidChars = " !@#$%^&*()|\/,.;':{}[]?<>" + Chr(34)

    If InStr(InvalidChars, Chr(KeyAscii)) <> 0 Then
        KeyAscii = 0
    End If

End Sub

Private Sub UpdateJSFileNames()

    Dim fn As String
    
    fn = txtJSFileName.Text
    If LenB(fn) = 0 Then fn = "menu"

    txtJSFileNames.Text = fn + ".js" + vbCrLf + _
                          "ie" + fn + ".js" + vbCrLf + _
                          "ns" + fn + ".js" + vbCrLf + _
                          "ie" + fn + "_frames.js" + vbCrLf + _
                          "ns" + fn + "_frames.js"
    
End Sub

Private Sub txtName_GotFocus()

    SelAll txtName

End Sub

Private Sub txtRootWeb_Change(Index As Integer)

    If IsUpdating Then Exit Sub
    Project.UserConfigs(ppSelConfig).RootWeb = txtRootWeb(Index).Text

End Sub

Private Sub LocalizeUI()

    Dim ctrl As Control

    caption = GetLocalizedStr(378)

    'General
    tsSections.Tabs(1).caption = GetLocalizedStr(321)
    lblProjectInformation.caption = GetLocalizedStr(915)
    lblPLocation.caption = GetLocalizedStr(916)
    lblPVersion.caption = GetLocalizedStr(917)
    lblPTitle.caption = GetLocalizedStr(918)
    lblUnfoldingSound.caption = GetLocalizedStr(320)
    cmdRemoveSound.caption = GetLocalizedStr(201)
    
    'Global Settings
    tsSections.Tabs(3).caption = GetLocalizedStr(727)
    tsGlobalOptions.Tabs("tsGSFX").caption = GetLocalizedStr(319)
    lblEffectType.caption = GetLocalizedStr(726)
    tsGlobalOptions.Tabs("tsGSTimers").caption = GetLocalizedStr(728)
    lblRootmDispDelay.caption = GetLocalizedStr(960)
    lblMenusHideDelay.caption = GetLocalizedStr(729)
    lblSubmDispDelay.caption = GetLocalizedStr(730)
    lblEffectSpeed.caption = GetLocalizedStr(731)
    tsGlobalOptions.Tabs("tsGSExtFeat").caption = GetLocalizedStr(732)
    cmdRestore.caption = GetLocalizedStr(742)
    lblLess1.caption = GetLocalizedStr(744)
    lblLess2.caption = GetLocalizedStr(744)
    lblLess3.caption = GetLocalizedStr(744)
    lblMore1.caption = GetLocalizedStr(745)
    lblMore2.caption = GetLocalizedStr(745)
    lblMore3.caption = GetLocalizedStr(745)
    lblSlow.caption = GetLocalizedStr(746)
    lblFast.caption = GetLocalizedStr(747)
    
    'Configurations
    tsSections.Tabs(2).caption = GetLocalizedStr(322)
    lblRootWeb(0).caption = GetLocalizedStr(326)
    lblDest(0).caption = GetLocalizedStr(327)
    lblImagesPath(0).caption = GetLocalizedStr(328)
    lblLocalConfig(0).caption = GetLocalizedStr(329)
    
    tsConfigs.Tabs(1).caption = GetLocalizedStr(334)
    tsConfigs.Tabs(2).caption = GetLocalizedStr(748)
    tsConfigs.Tabs(3).caption = GetLocalizedStr(336)
    cmdConfigOptions.caption = GetLocalizedStr(337)
    
    mnuConfigAdd.caption = GetLocalizedStr(338)
    mnuConfigEdit.caption = GetLocalizedStr(339)
    mnuConfigRemove.caption = GetLocalizedStr(201)
    mnuConfigSetAsDefault.caption = GetLocalizedStr(341)
    mnuConfigOpPaths.caption = GetLocalizedStr(743)
    
    lblDocHS.caption = GetLocalizedStr(344)
        
    chkFrameSupport.caption = GetLocalizedStr(346)
    lblFramesDoc.caption = GetLocalizedStr(347)
    
    'FTP
    tsConfigs.Tabs(4).caption = GetLocalizedStr(324)
    'lblFTPServer.Caption = GetLocalizedStr(369)
    'opLogin(0).Caption = GetLocalizedStr(370)
    'opLogin(1).Caption = GetLocalizedStr(371)
    chkFTPUseProxy.caption = GetLocalizedStr(372)
    lblFTPProxyAddress.caption = GetLocalizedStr(375)
    lblFTPProxyPort.caption = GetLocalizedStr(376)
    'lblRemoteConfig.Caption = GetLocalizedStr(377)
    frameAccountInfo.caption = GetLocalizedStr(677)
    lblFTPUserName.caption = GetLocalizedStr(373)
    lblFTPPassword.caption = GetLocalizedStr(374)
    
    'Advanced
    tsSections.Tabs(4).caption = GetLocalizedStr(325)
    lblCodeOp.caption = GetLocalizedStr(379)
    lblJSName.caption = GetLocalizedStr(380)
    cmdPosOffset.caption = GetLocalizedStr(381)
    cmdFontSubst.caption = GetLocalizedStr(713)
    
    lblCO0.caption = GetLocalizedStr(822)
    lblCO1.caption = GetLocalizedStr(823)
    lblCO2.caption = GetLocalizedStr(824)
    
    cmbFX.Clear
    cmbFX.AddItem GetLocalizedStr(455)
    cmbFX.AddItem GetLocalizedStr(451)
    cmbFX.AddItem GetLocalizedStr(452)
    cmbFX.AddItem GetLocalizedStr(453)
    cmbFX.AddItem GetLocalizedStr(454)
    
    For Each ctrl In Controls
        If TypeOf ctrl Is Label Then
            If Left(ctrl.caption, 9) = "example: " Then
                ctrl.caption = Replace(ctrl.caption, "example", GetLocalizedStr(333))
            End If
        End If
    Next ctrl
    
    cmdOK.caption = GetLocalizedStr(186)
    cmdCancel.caption = GetLocalizedStr(187)
    
    FixContolsWidth Me
    
    cmdPosOffset.Width = SetCtrlWidth(cmdPosOffset)
    cmdFontSubst.Width = SetCtrlWidth(cmdFontSubst)
    cmdFontSubst.Left = cmdPosOffset.Left + cmdPosOffset.Width + 90
    
    If Preferences.language <> "eng" Then
        cmdOK.Width = SetCtrlWidth(cmdOK)
        cmdCancel.Width = SetCtrlWidth(cmdCancel)
    End If

End Sub

Private Sub ucCodeOp_Change()

    #If LITE = 1 Then
        If ucCodeOp.Value >= 2 Then
            ucCodeOp.Value = 1
            frmMain.ShowLITELImitationInfo 3
        End If
    #Else
        chkGZIP.Enabled = (ucCodeOp.Value > 0)
    #End If

End Sub

Private Sub UpdateSecDisp()

    If sldRootMenusDelay.Enabled Then
        rmDSec.caption = NiceSec(sldRootMenusDelay.Value)
    Else
        rmDSec.caption = "n/a"
    End If
    smDSec.caption = NiceSec(sldSubMenusDelay.Value)
    scDSec.caption = NiceSec(sldSelChangeDelay.Value)
    mhDSec.caption = NiceSec(sldHideDelay.Value)

End Sub

Private Function NiceSec(v As Integer) As String

    If v < 1000 Then
        NiceSec = CStr(v) + "ms"
    Else
        NiceSec = CStr(Round(v / 1000, 2)) + "s"
    End If

End Function

