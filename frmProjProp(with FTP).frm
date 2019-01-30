VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{E1C6DB9D-BD4A-4A61-A759-0CED75D034BF}#43.0#0"; "SmartButton.ocx"
Object = "{9A2947B4-87EE-40EE-A3EF-32BDC32D5726}#1.0#0"; "xfxline3d.ocx"
Begin VB.Form frmProjProp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Project Properties"
   ClientHeight    =   11010
   ClientLeft      =   645
   ClientTop       =   2205
   ClientWidth     =   15300
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmProjProp(with FTP).frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11010
   ScaleWidth      =   15300
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picFTP 
      Height          =   3945
      Left            =   780
      ScaleHeight     =   3885
      ScaleWidth      =   6585
      TabIndex        =   79
      Top             =   6105
      Width           =   6645
      Begin VB.Frame frameAccountInfo 
         Caption         =   "Account Information"
         Height          =   1425
         Left            =   75
         TabIndex        =   118
         Top             =   1035
         Width           =   3090
         Begin VB.OptionButton opLogin 
            Caption         =   "Login Anonymously"
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
            Left            =   240
            TabIndex        =   120
            Top             =   495
            Width           =   2820
         End
         Begin VB.OptionButton opLogin 
            Caption         =   "Prompt for Account"
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
            Left            =   240
            TabIndex        =   119
            Top             =   810
            Value           =   -1  'True
            Width           =   2820
         End
      End
      Begin VB.ComboBox cmbRemoteConfigs 
         Height          =   330
         ItemData        =   "frmProjProp(with FTP).frx":058A
         Left            =   0
         List            =   "frmProjProp(with FTP).frx":0591
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   3075
         Width           =   3315
      End
      Begin VB.CheckBox chkFTPUseProxy 
         Caption         =   "Use a Proxy Server"
         Height          =   240
         Left            =   3390
         TabIndex        =   28
         Top             =   1020
         Width           =   2070
      End
      Begin VB.TextBox txtFTPAddress 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   45
         TabIndex        =   27
         Top             =   315
         Width           =   3315
      End
      Begin VB.Frame Frame5 
         Height          =   1425
         Left            =   3255
         TabIndex        =   80
         Top             =   1035
         Width           =   3165
         Begin VB.TextBox txtFTPProxyPort 
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
            Left            =   240
            TabIndex        =   30
            Top             =   990
            Width           =   510
         End
         Begin VB.TextBox txtFTPProxyAddress 
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
            Left            =   240
            TabIndex        =   29
            Top             =   465
            Width           =   2700
         End
         Begin VB.Label lblFTPProxyPort 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Port"
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
            Height          =   195
            Left            =   240
            TabIndex        =   82
            Top             =   795
            Width           =   300
         End
         Begin VB.Label lblFTPProxyAddress 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Address"
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
            Height          =   195
            Left            =   240
            TabIndex        =   81
            Top             =   255
            Width           =   585
         End
      End
      Begin VB.Label lblFTPUpImg 
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "example: ftp://ftp.xfx.net/web/"
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
         Left            =   120
         TabIndex        =   99
         Top             =   3630
         Width           =   2190
      End
      Begin VB.Label lblFTPUpJS 
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "example: ftp://ftp.xfx.net/web/"
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
         Left            =   120
         TabIndex        =   98
         Top             =   3420
         Width           =   2190
      End
      Begin VB.Label lblRemoteConfig 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remote Configuration"
         Height          =   210
         Left            =   0
         TabIndex        =   97
         Top             =   2850
         Width           =   1785
      End
      Begin VB.Label lblFTPServer 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FTP Server Address"
         Height          =   210
         Left            =   45
         TabIndex        =   84
         Top             =   90
         Width           =   1605
      End
      Begin VB.Label lblFTPServerSample 
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "example: ftp://ftp.xfx.net/web/"
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
         TabIndex        =   83
         Top             =   645
         Width           =   2190
      End
   End
   Begin VB.Timer tmrSelPathsSection 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3285
      Top             =   5145
   End
   Begin VB.PictureBox picAdvanced 
      Height          =   4500
      Left            =   165
      ScaleHeight     =   4440
      ScaleWidth      =   6450
      TabIndex        =   45
      Top             =   6990
      Width           =   6510
      Begin VB.ComboBox cmbCodeOp 
         Height          =   330
         ItemData        =   "frmProjProp(with FTP).frx":05A0
         Left            =   45
         List            =   "frmProjProp(with FTP).frx":05AA
         Style           =   2  'Dropdown List
         TabIndex        =   121
         Top             =   345
         WhatsThisHelpID =   20400
         Width           =   3150
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
         TabIndex        =   117
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
         TabIndex        =   102
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
         TabIndex        =   77
         Text            =   "frmProjProp(with FTP).frx":05E4
         Top             =   2475
         Width           =   3615
      End
      Begin VB.TextBox txtJSFileName 
         Height          =   315
         Left            =   45
         TabIndex        =   33
         Top             =   2160
         Width           =   3150
      End
      Begin VB.ComboBox cmbAddIns 
         Height          =   330
         ItemData        =   "frmProjProp(with FTP).frx":0604
         Left            =   45
         List            =   "frmProjProp(with FTP).frx":060B
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   1185
         Width           =   3150
      End
      Begin SmartButtonProject.SmartButton cmdInfo 
         Height          =   360
         Left            =   3255
         TabIndex        =   113
         Top             =   1170
         Width           =   420
         _ExtentX        =   741
         _ExtentY        =   635
         Picture         =   "frmProjProp(with FTP).frx":0617
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
         TabIndex        =   114
         Top             =   1170
         Width           =   420
         _ExtentX        =   741
         _ExtentY        =   635
         Picture         =   "frmProjProp(with FTP).frx":09B1
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
         TabIndex        =   115
         Top             =   1170
         Width           =   420
         _ExtentX        =   741
         _ExtentY        =   635
         Picture         =   "frmProjProp(with FTP).frx":0D4B
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
         TabIndex        =   95
         Top             =   1515
         Width           =   2940
      End
      Begin VB.Label lblJSName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Compiled JavaScript File Name"
         Height          =   210
         Left            =   45
         TabIndex        =   76
         Top             =   1920
         Width           =   2445
      End
      Begin VB.Label lblCodeOp 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Code Optimization"
         Height          =   210
         Left            =   45
         TabIndex        =   75
         Top             =   90
         Width           =   1485
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "AddIns"
         Height          =   210
         Left            =   45
         TabIndex        =   74
         Top             =   960
         Width           =   555
      End
   End
   Begin VB.PictureBox picConfigs 
      Height          =   6315
      Left            =   7620
      ScaleHeight     =   6255
      ScaleWidth      =   9150
      TabIndex        =   47
      Top             =   4920
      Width           =   9210
      Begin VB.PictureBox picConfigBtns 
         BorderStyle     =   0  'None
         Height          =   1710
         Left            =   4890
         ScaleHeight     =   1710
         ScaleWidth      =   1560
         TabIndex        =   85
         Top             =   1320
         Width           =   1560
         Begin MSComctlLib.TabStrip tsConfigs 
            Height          =   1050
            Left            =   0
            TabIndex        =   7
            Top             =   150
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   1852
            TabWidthStyle   =   2
            MultiRow        =   -1  'True
            Style           =   1
            TabFixedWidth   =   2752
            TabFixedHeight  =   556
            Separators      =   -1  'True
            TabMinWidth     =   2752
            _Version        =   393216
            BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
               NumTabs         =   3
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
            TabIndex        =   8
            Top             =   1395
            Width           =   1560
         End
      End
      Begin VB.PictureBox picConfigsOp 
         Height          =   4050
         Index           =   0
         Left            =   30
         ScaleHeight     =   3990
         ScaleWidth      =   6300
         TabIndex        =   48
         Top             =   30
         Width           =   6360
         Begin VB.PictureBox picHotSpotEditor 
            Height          =   3855
            Left            =   765
            ScaleHeight     =   3795
            ScaleWidth      =   3990
            TabIndex        =   86
            Top             =   45
            Width           =   4050
            Begin VB.TextBox txtDestFile 
               Height          =   315
               Left            =   45
               TabIndex        =   40
               Top             =   315
               Width           =   3315
            End
            Begin SmartButtonProject.SmartButton cmdBrowseDestFile 
               Height          =   360
               Left            =   3450
               TabIndex        =   109
               Top             =   292
               Width           =   420
               _ExtentX        =   741
               _ExtentY        =   635
               Picture         =   "frmProjProp(with FTP).frx":10E5
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
               TabIndex        =   178
               Top             =   1305
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
               TabIndex        =   89
               Top             =   825
               Width           =   2940
            End
            Begin VB.Label lblDocHS 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Document Containing the HotSpots"
               Height          =   210
               Left            =   45
               TabIndex        =   88
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
               TabIndex        =   87
               Top             =   645
               Width           =   3030
            End
         End
         Begin VB.PictureBox picFrames 
            Height          =   3855
            Left            =   420
            ScaleHeight     =   3795
            ScaleWidth      =   6300
            TabIndex        =   90
            Top             =   585
            Width           =   6360
            Begin SmartButtonProject.SmartButton cmdReload 
               Height          =   360
               Left            =   3855
               TabIndex        =   111
               Top             =   832
               Width           =   420
               _ExtentX        =   741
               _ExtentY        =   635
               Picture         =   "frmProjProp(with FTP).frx":123F
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
               TabIndex        =   110
               Top             =   832
               Width           =   420
               _ExtentX        =   741
               _ExtentY        =   635
               Picture         =   "frmProjProp(with FTP).frx":15D9
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
               Height          =   315
               Left            =   45
               TabIndex        =   6
               Top             =   855
               WhatsThisHelpID =   20530
               Width           =   3315
            End
            Begin VB.CheckBox chkFrameSupport 
               Caption         =   "Enable Frames Support"
               Height          =   345
               Left            =   60
               TabIndex        =   5
               TabStop         =   0   'False
               Top             =   90
               Width           =   4155
            End
            Begin VB.Label lblFramesDoc 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Frames Document"
               Height          =   210
               Left            =   45
               TabIndex        =   93
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
               TabIndex        =   92
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
               TabIndex        =   91
               Top             =   1185
               Width           =   2940
            End
         End
         Begin SmartButtonProject.SmartButton cmdBrowseLocalImages 
            Height          =   360
            Index           =   0
            Left            =   3405
            TabIndex        =   108
            Top             =   2647
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   635
            Picture         =   "frmProjProp(with FTP).frx":1733
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
            TabIndex        =   107
            Top             =   735
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   635
            Picture         =   "frmProjProp(with FTP).frx":188D
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
            TabIndex        =   105
            Top             =   1695
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   635
            Picture         =   "frmProjProp(with FTP).frx":19E7
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
            Height          =   330
            Index           =   0
            ItemData        =   "frmProjProp(with FTP).frx":1B41
            Left            =   0
            List            =   "frmProjProp(with FTP).frx":1B48
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   3585
            Visible         =   0   'False
            Width           =   3315
         End
         Begin VB.TextBox txtImagesPath 
            Height          =   315
            Index           =   0
            Left            =   0
            TabIndex        =   3
            Top             =   2670
            WhatsThisHelpID =   20370
            Width           =   3315
         End
         Begin VB.TextBox txtDest 
            Height          =   315
            Index           =   0
            Left            =   0
            TabIndex        =   2
            Top             =   1725
            WhatsThisHelpID =   20370
            Width           =   3315
         End
         Begin VB.TextBox txtRootWeb 
            Height          =   315
            Index           =   0
            Left            =   0
            TabIndex        =   1
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
            TabIndex        =   96
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
            ForeColor       =   &H8000000E&
            Height          =   210
            Index           =   0
            Left            =   105
            TabIndex        =   94
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
            ForeColor       =   &H8000000E&
            Height          =   195
            Index           =   0
            Left            =   105
            TabIndex        =   78
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
            TabIndex        =   55
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
            TabIndex        =   54
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
            TabIndex        =   52
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
            TabIndex        =   51
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
            TabIndex        =   50
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
            TabIndex        =   49
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
         TabIndex        =   36
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
            Picture         =   "frmProjProp(with FTP).frx":1B57
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProjProp(with FTP).frx":1CB1
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProjProp(with FTP).frx":1E0B
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProjProp(with FTP).frx":225D
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdImport 
      Caption         =   "Import..."
      Height          =   375
      Left            =   75
      TabIndex        =   39
      Top             =   5130
      Width           =   1050
   End
   Begin VB.PictureBox picGeneral 
      Height          =   4500
      Left            =   150
      ScaleHeight     =   4440
      ScaleWidth      =   6450
      TabIndex        =   41
      Top             =   480
      Width           =   6510
      Begin VB.TextBox txtUnfoldingSound 
         Height          =   315
         Left            =   45
         TabIndex        =   101
         Top             =   2715
         Visible         =   0   'False
         WhatsThisHelpID =   20550
         Width           =   3315
      End
      Begin VB.TextBox txtName 
         Height          =   315
         Left            =   45
         TabIndex        =   0
         Top             =   975
         Width           =   2205
      End
      Begin VB.TextBox txtFileName 
         BackColor       =   &H80000000&
         Height          =   315
         Left            =   45
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   315
         WhatsThisHelpID =   20350
         Width           =   6375
      End
      Begin SmartButtonProject.SmartButton cmdBrowseSound 
         Height          =   360
         Left            =   3420
         TabIndex        =   106
         Top             =   2692
         Visible         =   0   'False
         Width           =   420
         _ExtentX        =   741
         _ExtentY        =   635
         Picture         =   "frmProjProp(with FTP).frx":25F7
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
         TabIndex        =   112
         Top             =   2692
         Visible         =   0   'False
         Width           =   420
         _ExtentX        =   741
         _ExtentY        =   635
         Picture         =   "frmProjProp(with FTP).frx":2751
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
         TabIndex        =   116
         Top             =   3075
         Visible         =   0   'False
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   423
         Caption         =   "Remove"
         Picture         =   "frmProjProp(with FTP).frx":28AB
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
      Begin VB.Label lblUnfoldingSound 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unfolding Sound"
         Height          =   210
         Left            =   45
         TabIndex        =   100
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
         Left            =   165
         TabIndex        =   53
         Top             =   1290
         Width           =   1440
      End
      Begin VB.Label lblProjectName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Project Name"
         Height          =   210
         Left            =   45
         TabIndex        =   44
         Top             =   750
         Width           =   1110
      End
      Begin VB.Label lblProjectLocation 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Project Location"
         Height          =   210
         Left            =   45
         TabIndex        =   43
         Top             =   90
         Width           =   1335
      End
   End
   Begin VB.PictureBox picToolbar 
      Height          =   4485
      Left            =   7860
      ScaleHeight     =   4425
      ScaleWidth      =   7095
      TabIndex        =   46
      Top             =   225
      Width           =   7155
      Begin VB.PictureBox picTBPositioning 
         Height          =   3975
         Left            =   390
         ScaleHeight     =   3915
         ScaleWidth      =   5985
         TabIndex        =   64
         Top             =   315
         Width           =   6045
         Begin xfxLine3D.ucLine3D uc3DLine4 
            Height          =   30
            Left            =   30
            TabIndex        =   177
            Top             =   1695
            Width           =   5970
            _ExtentX        =   10530
            _ExtentY        =   53
         End
         Begin xfxLine3D.ucLine3D uc3DLine3 
            Height          =   30
            Left            =   15
            TabIndex        =   176
            Top             =   975
            Width           =   5940
            _ExtentX        =   10478
            _ExtentY        =   53
         End
         Begin VB.Frame Frame1 
            BorderStyle     =   0  'None
            Height          =   675
            Left            =   1575
            TabIndex        =   163
            Top             =   105
            Width           =   2790
            Begin VB.TextBox txtACY 
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
               Left            =   1425
               TabIndex        =   175
               Text            =   "000"
               Top             =   240
               Width           =   360
            End
            Begin VB.TextBox txtACX 
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
               Left            =   1005
               TabIndex        =   174
               Text            =   "000"
               Top             =   240
               Width           =   360
            End
            Begin VB.OptionButton opAlignment 
               Caption         =   "Custom"
               Enabled         =   0   'False
               Height          =   225
               Index           =   9
               Left            =   795
               TabIndex        =   173
               Top             =   0
               Width           =   1965
            End
            Begin VB.OptionButton opAlignment 
               Enabled         =   0   'False
               Height          =   225
               Index           =   8
               Left            =   525
               TabIndex        =   172
               Top             =   480
               Width           =   225
            End
            Begin VB.OptionButton opAlignment 
               Enabled         =   0   'False
               Height          =   225
               Index           =   7
               Left            =   255
               TabIndex        =   171
               Top             =   480
               Width           =   225
            End
            Begin VB.OptionButton opAlignment 
               Enabled         =   0   'False
               Height          =   225
               Index           =   6
               Left            =   0
               TabIndex        =   170
               Top             =   480
               Width           =   225
            End
            Begin VB.OptionButton opAlignment 
               Enabled         =   0   'False
               Height          =   225
               Index           =   5
               Left            =   525
               TabIndex        =   169
               Top             =   240
               Width           =   225
            End
            Begin VB.OptionButton opAlignment 
               Enabled         =   0   'False
               Height          =   225
               Index           =   4
               Left            =   255
               TabIndex        =   168
               Top             =   240
               Width           =   225
            End
            Begin VB.OptionButton opAlignment 
               Enabled         =   0   'False
               Height          =   225
               Index           =   3
               Left            =   0
               TabIndex        =   167
               Top             =   240
               Width           =   225
            End
            Begin VB.OptionButton opAlignment 
               Enabled         =   0   'False
               Height          =   225
               Index           =   2
               Left            =   525
               TabIndex        =   166
               Top             =   0
               Width           =   225
            End
            Begin VB.OptionButton opAlignment 
               Enabled         =   0   'False
               Height          =   225
               Index           =   1
               Left            =   255
               TabIndex        =   165
               Top             =   0
               Width           =   225
            End
            Begin VB.OptionButton opAlignment 
               Enabled         =   0   'False
               Height          =   225
               Index           =   0
               Left            =   0
               TabIndex        =   164
               Top             =   0
               Value           =   -1  'True
               Width           =   225
            End
         End
         Begin VB.TextBox txtMarginH 
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
            Left            =   1575
            TabIndex        =   22
            Text            =   "888"
            Top             =   1905
            Width           =   480
         End
         Begin VB.TextBox txtMarginV 
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
            Left            =   1575
            TabIndex        =   23
            Text            =   "888"
            Top             =   2250
            Width           =   480
         End
         Begin VB.ComboBox cmbSpanning 
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
            Height          =   315
            ItemData        =   "frmProjProp(with FTP).frx":2C45
            Left            =   1575
            List            =   "frmProjProp(with FTP).frx":2C4F
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   1200
            Width           =   1620
         End
         Begin VB.Label lblTBOffsetH 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pixels Horizontally"
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
            Left            =   2115
            TabIndex        =   69
            Top             =   1950
            Width           =   1290
         End
         Begin VB.Label lblTBOffsetV 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pixels Vertically"
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
            Left            =   2115
            TabIndex        =   68
            Top             =   2295
            Width           =   1095
         End
         Begin VB.Label lblTBOffset 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Offset"
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
            Left            =   990
            TabIndex        =   67
            Top             =   2100
            Width           =   465
         End
         Begin VB.Label lblTBSpanning 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Spanning"
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
            Left            =   795
            TabIndex        =   66
            Top             =   1260
            Width           =   660
         End
         Begin VB.Label lblTBAlignment 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Alignment"
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
            Left            =   750
            TabIndex        =   65
            Top             =   330
            Width           =   705
         End
      End
      Begin VB.PictureBox picTBAppearance 
         AutoRedraw      =   -1  'True
         Height          =   3870
         Left            =   735
         ScaleHeight     =   3810
         ScaleWidth      =   5940
         TabIndex        =   58
         Top             =   60
         Width           =   6000
         Begin VB.TextBox txtTBCMarginH 
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
            Left            =   1935
            TabIndex        =   14
            Text            =   "123"
            Top             =   1230
            Width           =   420
         End
         Begin VB.TextBox txtTBCMarginV 
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
            Left            =   3225
            TabIndex        =   15
            Text            =   "123"
            Top             =   1215
            Width           =   420
         End
         Begin MSComCtl2.UpDown udSep 
            Height          =   285
            Left            =   2010
            TabIndex        =   104
            Top             =   1575
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   503
            _Version        =   393216
            AutoBuddy       =   -1  'True
            BuddyControl    =   "txtTBSeparation"
            BuddyDispid     =   196708
            OrigLeft        =   1785
            OrigTop         =   1125
            OrigRight       =   1980
            OrigBottom      =   1440
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin VB.TextBox txtTBSeparation 
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
            Left            =   1635
            TabIndex        =   16
            Text            =   "888"
            Top             =   1575
            Width           =   390
         End
         Begin VB.PictureBox picTBImage 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   480
            Left            =   2745
            ScaleHeight     =   450
            ScaleWidth      =   450
            TabIndex        =   59
            Top             =   2760
            Width           =   480
         End
         Begin VB.ComboBox cmbBorder 
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
            Height          =   315
            ItemData        =   "frmProjProp(with FTP).frx":2C66
            Left            =   1635
            List            =   "frmProjProp(with FTP).frx":2C8B
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   480
            Width           =   945
         End
         Begin VB.CheckBox chkJustify 
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
            Left            =   1635
            TabIndex        =   17
            Top             =   1935
            Width           =   195
         End
         Begin VB.ComboBox cmbStyle 
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
            Height          =   315
            ItemData        =   "frmProjProp(with FTP).frx":2CB6
            Left            =   1635
            List            =   "frmProjProp(with FTP).frx":2CC0
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   120
            Width           =   1665
         End
         Begin SmartButtonProject.SmartButton cmdTBBackColor 
            Height          =   240
            Left            =   1635
            TabIndex        =   18
            Top             =   2430
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
            Left            =   1635
            TabIndex        =   19
            Top             =   2760
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   423
            Caption         =   "Change"
            Picture         =   "frmProjProp(with FTP).frx":2CDA
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
            Left            =   1635
            TabIndex        =   20
            Top             =   3000
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   423
            Caption         =   "Remove"
            Picture         =   "frmProjProp(with FTP).frx":3074
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
            Left            =   2655
            TabIndex        =   13
            Top             =   510
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
            Left            =   3645
            TabIndex        =   135
            Top             =   1230
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   503
            _Version        =   393216
            BuddyControl    =   "txtTBCMarginV"
            BuddyDispid     =   196707
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
            Left            =   2355
            TabIndex        =   136
            Top             =   1230
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   503
            _Version        =   393216
            BuddyControl    =   "txtTBCMarginH"
            BuddyDispid     =   196706
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
            TabIndex        =   139
            Top             =   900
            Width           =   5745
            _ExtentX        =   10134
            _ExtentY        =   53
         End
         Begin xfxLine3D.ucLine3D uc3DLine2 
            Height          =   30
            Left            =   0
            TabIndex        =   140
            Top             =   2250
            Width           =   5745
            _ExtentX        =   10134
            _ExtentY        =   53
         End
         Begin VB.Label lblJustify 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Justify HotSpots"
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
            TabIndex        =   142
            Top             =   1950
            Width           =   1185
         End
         Begin VB.Label lblTBMargins 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Margins"
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
            Left            =   1005
            TabIndex        =   141
            Top             =   1275
            Width           =   555
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
            Left            =   1635
            TabIndex        =   138
            Top             =   1005
            Width           =   720
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
            Left            =   2925
            TabIndex        =   137
            Top             =   1005
            Width           =   525
         End
         Begin VB.Image Image1 
            Height          =   240
            Left            =   1635
            Picture         =   "frmProjProp(with FTP).frx":340E
            Top             =   1245
            Width           =   240
         End
         Begin VB.Image Image2 
            Height          =   240
            Left            =   2925
            Picture         =   "frmProjProp(with FTP).frx":3798
            Top             =   1245
            Width           =   240
         End
         Begin VB.Label lblTBSeparation 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Separation"
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
            Left            =   780
            TabIndex        =   103
            Top             =   1605
            Width           =   780
         End
         Begin VB.Label lblTBBorder 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Border"
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
            Left            =   1080
            TabIndex        =   63
            Top             =   540
            Width           =   480
         End
         Begin VB.Label lblTBBackImage 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Back Image"
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
            Left            =   735
            TabIndex        =   62
            Top             =   2880
            Width           =   825
         End
         Begin VB.Label lblTBStyle 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Toolbar Style"
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
            Left            =   615
            TabIndex        =   61
            Top             =   180
            Width           =   945
         End
         Begin VB.Label lblTBBackColor 
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
            Left            =   810
            TabIndex        =   60
            Top             =   2445
            Width           =   750
         End
      End
      Begin VB.PictureBox picTBAdvanced 
         Height          =   3240
         Left            =   270
         ScaleHeight     =   3180
         ScaleWidth      =   6525
         TabIndex        =   70
         Top             =   825
         Width           =   6585
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'None
            Height          =   795
            Left            =   2670
            ScaleHeight     =   795
            ScaleWidth      =   2520
            TabIndex        =   129
            Top             =   1350
            Width           =   2520
            Begin VB.TextBox txtTBHeight 
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
               Left            =   900
               TabIndex        =   132
               Text            =   "0"
               Top             =   495
               Width           =   405
            End
            Begin VB.OptionButton opTBHeight 
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
               Index           =   1
               Left            =   615
               TabIndex        =   131
               Top             =   225
               Width           =   1620
            End
            Begin VB.OptionButton opTBHeight 
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
               Left            =   615
               TabIndex        =   130
               Top             =   0
               Width           =   1620
            End
            Begin SmartButtonProject.SmartButton cmdTBAutoHeight 
               Height          =   300
               Left            =   1350
               TabIndex        =   133
               Top             =   495
               Width           =   1170
               _ExtentX        =   2064
               _ExtentY        =   529
               Caption         =   "Calculate"
               Picture         =   "frmProjProp(with FTP).frx":3B22
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
               Left            =   75
               TabIndex        =   134
               Top             =   360
               Width           =   465
            End
            Begin VB.Image imgGSHeight 
               Appearance      =   0  'Flat
               Height          =   240
               Left            =   180
               Picture         =   "frmProjProp(with FTP).frx":3EBC
               Top             =   135
               Width           =   240
            End
         End
         Begin VB.TextBox txtTBWidth 
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
            Left            =   915
            TabIndex        =   126
            Text            =   "0"
            Top             =   2070
            Width           =   405
         End
         Begin VB.OptionButton opTBWidth 
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
            Left            =   630
            TabIndex        =   125
            Top             =   1815
            Width           =   1935
         End
         Begin VB.OptionButton opTBWidth 
            Caption         =   "Match Group's Width"
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
            Index           =   1
            Left            =   630
            TabIndex        =   124
            Top             =   1590
            Width           =   1860
         End
         Begin VB.OptionButton opTBWidth 
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
            Left            =   630
            TabIndex        =   123
            Top             =   1350
            Width           =   720
         End
         Begin VB.CheckBox chkFollowScrolling 
            Caption         =   "Follow Scrolling"
            Enabled         =   0   'False
            Height          =   345
            Left            =   0
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   150
            WhatsThisHelpID =   20410
            Width           =   4560
         End
         Begin VB.CheckBox chkHorizontal 
            Enabled         =   0   'False
            Height          =   210
            Left            =   285
            TabIndex        =   25
            Top             =   570
            Width           =   195
         End
         Begin VB.CheckBox chkVertical 
            Enabled         =   0   'False
            Height          =   210
            Left            =   1155
            TabIndex        =   26
            Top             =   570
            Width           =   195
         End
         Begin MSComCtl2.UpDown udV 
            Height          =   360
            Left            =   1395
            TabIndex        =   71
            TabStop         =   0   'False
            Top             =   495
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   635
            _Version        =   393216
            Enabled         =   0   'False
         End
         Begin MSComCtl2.UpDown udH 
            Height          =   240
            Left            =   525
            TabIndex        =   72
            TabStop         =   0   'False
            Top             =   555
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
            TabIndex        =   127
            Top             =   2070
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   529
            Caption         =   "Calculate"
            Picture         =   "frmProjProp(with FTP).frx":4246
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
         Begin VB.Image imgGSWidth 
            Appearance      =   0  'Flat
            Height          =   240
            Left            =   165
            Picture         =   "frmProjProp(with FTP).frx":45E0
            Top             =   1485
            Width           =   240
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
            Left            =   75
            TabIndex        =   128
            Top             =   1710
            Width           =   420
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Toolbar Size"
            Height          =   210
            Left            =   0
            TabIndex        =   122
            Top             =   1080
            Width           =   990
         End
      End
      Begin VB.PictureBox picTBGeneral 
         Height          =   3540
         Left            =   45
         ScaleHeight     =   3480
         ScaleWidth      =   6375
         TabIndex        =   56
         Top             =   555
         Width           =   6435
         Begin MSComctlLib.Toolbar tbGroupsOptions 
            Height          =   990
            Left            =   5700
            TabIndex        =   57
            Top             =   390
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   1746
            ButtonWidth     =   609
            ButtonHeight    =   582
            AllowCustomize  =   0   'False
            Style           =   1
            ImageList       =   "ilIcons"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   3
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "tbUp"
                  ImageIndex      =   1
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "tbDown"
                  ImageIndex      =   2
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "tbCheck"
                  ImageIndex      =   3
                  Style           =   1
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.ListView lvGroups 
            Height          =   2625
            Left            =   -15
            TabIndex        =   10
            Top             =   390
            Width           =   5700
            _ExtentX        =   10054
            _ExtentY        =   4630
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   0   'False
            HideSelection   =   0   'False
            Checkboxes      =   -1  'True
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
            Caption         =   "Select the groups to include in the Toolbar"
            Height          =   210
            Left            =   0
            TabIndex        =   73
            Top             =   150
            Width           =   3555
         End
      End
      Begin VB.CheckBox chkCreateToolbar 
         Caption         =   "Create Toolbar"
         Height          =   345
         Left            =   0
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   90
         WhatsThisHelpID =   20410
         Width           =   3840
      End
      Begin MSComctlLib.TabStrip tsToolbar 
         Height          =   3975
         Left            =   -15
         TabIndex        =   37
         Top             =   495
         Width           =   6570
         _ExtentX        =   11589
         _ExtentY        =   7011
         TabWidthStyle   =   1
         Placement       =   1
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   4
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
   End
   Begin MSComctlLib.TabStrip tsSections 
      Height          =   4965
      Left            =   75
      TabIndex        =   35
      Top             =   90
      Width           =   6675
      _ExtentX        =   11774
      _ExtentY        =   8758
      ShowTips        =   0   'False
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   6
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
            Caption         =   "Toolbar"
            Key             =   "tsToolbar"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "FTP"
            Key             =   "tsFTP"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Global Settings"
            Key             =   "tsGlobalSettings"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Advanced"
            Key             =   "tsAdvanced"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5850
      TabIndex        =   38
      Top             =   5130
      Width           =   900
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   4740
      TabIndex        =   34
      Top             =   5130
      Width           =   900
   End
   Begin MSComDlg.CommonDialog cDlg 
      Left            =   195
      Top             =   5745
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picGlobalSettings 
      Height          =   4500
      Left            =   1965
      ScaleHeight     =   4440
      ScaleWidth      =   6450
      TabIndex        =   143
      Top             =   5760
      Width           =   6510
      Begin VB.Frame frameTimers 
         Caption         =   "Timers"
         Height          =   1935
         Left            =   3330
         TabIndex        =   151
         Top             =   60
         Width           =   3090
         Begin MSComctlLib.Slider sldSubmenuDelay 
            Height          =   270
            Left            =   165
            TabIndex        =   152
            Top             =   1350
            Width           =   2730
            _ExtentX        =   4815
            _ExtentY        =   476
            _Version        =   393216
            LargeChange     =   1
            Min             =   10
            Max             =   2000
            SelStart        =   10
            TickFrequency   =   100
            Value           =   10
         End
         Begin MSComctlLib.Slider sldHideDelay 
            Height          =   270
            Left            =   165
            TabIndex        =   153
            Top             =   525
            Width           =   2730
            _ExtentX        =   4815
            _ExtentY        =   476
            _Version        =   393216
            LargeChange     =   1
            Min             =   250
            Max             =   5000
            SelStart        =   250
            TickFrequency   =   250
            Value           =   250
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
            Left            =   120
            TabIndex        =   159
            Top             =   315
            Width           =   1275
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
            Left            =   165
            TabIndex        =   158
            Top             =   810
            Width           =   300
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
            Left            =   2580
            TabIndex        =   157
            Top             =   810
            Width           =   345
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
            Left            =   120
            TabIndex        =   156
            Top             =   1125
            Width           =   1740
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
            Left            =   165
            TabIndex        =   155
            Top             =   1635
            Width           =   300
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
            Left            =   2580
            TabIndex        =   154
            Top             =   1635
            Width           =   345
         End
      End
      Begin VB.Frame frameUnfoldingEffect 
         Caption         =   "Unfolding Effect"
         Height          =   1935
         Left            =   75
         TabIndex        =   145
         Top             =   60
         Width           =   3090
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
            ItemData        =   "frmProjProp(with FTP).frx":496A
            Left            =   165
            List            =   "frmProjProp(with FTP).frx":497D
            Style           =   2  'Dropdown List
            TabIndex        =   147
            Top             =   525
            Width           =   2730
         End
         Begin MSComctlLib.Slider sldAnimSpeed 
            Height          =   270
            Left            =   165
            TabIndex        =   146
            Top             =   1350
            Width           =   2730
            _ExtentX        =   4815
            _ExtentY        =   476
            _Version        =   393216
            LargeChange     =   1
            Min             =   5
            Max             =   50
            SelStart        =   5
            TickFrequency   =   5
            Value           =   5
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
            Left            =   120
            TabIndex        =   162
            Top             =   315
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
            Left            =   165
            TabIndex        =   150
            Top             =   1635
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
            Left            =   2595
            TabIndex        =   149
            Top             =   1635
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
            Left            =   120
            TabIndex        =   148
            Top             =   1125
            Width           =   930
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
         Left            =   4845
         TabIndex        =   144
         Top             =   4050
         Width           =   1575
      End
      Begin MSComctlLib.ListView lvExtFeatures 
         Height          =   2070
         Left            =   75
         TabIndex        =   160
         Top             =   2325
         Width           =   4665
         _ExtentX        =   8229
         _ExtentY        =   3651
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
      Begin VB.Label lblExtFeatures 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Extended Features"
         Height          =   210
         Left            =   75
         TabIndex        =   161
         Top             =   2115
         Width           =   1560
      End
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

Private WithEvents msgSubClass As SmartSubClass
Attribute msgSubClass.VB_VarHelpID = -1

Private Sub chkCreateToolbar_Click()

    UpdateToolbarControls

End Sub

Private Sub UpdateToolbarControls()

    Dim i As Integer
    Dim IsOn As Boolean

    IsOn = (chkCreateToolbar.Value = vbChecked)

    For i = 0 To opAlignment.count - 1
        opAlignment(i).Enabled = IsOn
    Next i
    chkFollowScrolling.Enabled = IsOn
    chkHorizontal.Enabled = IsOn And (chkFollowScrolling.Value = vbChecked)
    chkVertical.Enabled = IsOn And (chkFollowScrolling.Value = vbChecked)
    txtMarginH.Enabled = IsOn
    txtMarginV.Enabled = IsOn
    udH.Enabled = IsOn And (chkFollowScrolling.Value = vbChecked) And (chkHorizontal.Value = vbChecked)
    udV.Enabled = IsOn And (chkFollowScrolling.Value = vbChecked) And (chkVertical.Value = vbChecked)
    txtMarginH.Enabled = IsOn
    txtMarginV.Enabled = IsOn
    cmbStyle.Enabled = IsOn
    cmbSpanning.Enabled = IsOn
    cmbBorder.Enabled = IsOn
    chkJustify.Enabled = IsOn
    cmdTBBackColor.Enabled = IsOn
    cmdTBBorderColor.Enabled = IsOn
    txtACX.Enabled = (opAlignment(9).Value = True) And IsOn
    txtACY.Enabled = (opAlignment(9).Value = True) And IsOn
    lvGroups.Enabled = IsOn And (lvGroups.ListItems.count > 0)
    lvGroups.Checkboxes = lvGroups.Enabled
    tbGroupsOptions.Enabled = lvGroups.Enabled
    cmdChange.Enabled = IsOn
    cmdRemove.Enabled = IsOn
    txtTBSeparation.Enabled = IsOn
    udSep.Enabled = IsOn
    txtTBCMarginH.Enabled = IsOn
    txtTBCMarginV.Enabled = IsOn
    udTBCMH.Enabled = IsOn
    udTBCMV.Enabled = IsOn
    
    opTBWidth(0).Enabled = IsOn
    opTBWidth(1).Enabled = IsOn
    opTBWidth(2).Enabled = IsOn
    txtTBWidth.Enabled = IsOn And opTBWidth(2).Value
    cmdTBAutoWidth.Enabled = txtTBWidth.Enabled
    
    opTBHeight(0).Enabled = IsOn
    opTBHeight(1).Enabled = IsOn
    txtTBHeight.Enabled = IsOn And opTBHeight(1).Value
    cmdTBAutoHeight.Enabled = txtTBHeight.Enabled
    
End Sub

Private Sub chkFollowScrolling_Click()

    UpdateToolbarControls

End Sub

Private Sub chkFrameSupport_Click()

    txtFramesDoc.Enabled = (chkFrameSupport.Value = vbChecked)
    cmdBrowseFramesDoc.Enabled = (chkFrameSupport.Value = vbChecked)
    cmdReload.Enabled = cmdBrowseFramesDoc.Enabled
    
    Project.UserConfigs(ppSelConfig).Frames.UseFrames = IIf(chkFrameSupport.Value = vbChecked, True, False)
      
End Sub

Private Sub chkFTPUseProxy_Click()

    txtFTPProxyAddress.Enabled = (chkFTPUseProxy.Value = vbChecked)
    lblFTPProxyAddress.Enabled = (chkFTPUseProxy.Value = vbChecked)
    txtFTPProxyPort.Enabled = (chkFTPUseProxy.Value = vbChecked)
    lblFTPProxyPort.Enabled = (chkFTPUseProxy.Value = vbChecked)

End Sub

Private Sub chkHorizontal_Click()

    UpdateToolbarControls

End Sub

Private Sub chkJustify_Click()

    UpdateToolbarControls

End Sub

Private Sub chkVertical_Click()

    UpdateToolbarControls

End Sub

Private Sub cmbAddIns_Click()

    If IsUpdating Then Exit Sub
    
    lblAddIneErr.Caption = ""

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


Private Sub cmbLocalConfigs_Click(Index As Integer)

    Project.UserConfigs(ppSelConfig).LocalInfo4RemoteConfig = cmbLocalConfigs(Index).Text

End Sub

Private Sub cmbRemoteConfigs_Click()

    Dim sName As String
    
    lblFTPUpJS.ForeColor = &H808080
    
    sName = txtFTPAddress.Text
    If InStr(sName, "ftp://") = 0 Or sName = "ftp://" Then
        If sName = "" Then
            lblFTPUpJS.Caption = ""
        Else
            lblFTPUpJS.Caption = GetLocalizedStr(384)
            lblFTPUpJS.ForeColor = vbRed
        End If
        lblFTPUpImg.Caption = ""
    Else
        sName = AddTrailingSlash(Mid(sName, 7), "/")
        lblFTPUpJS.Caption = "ftp://" + Replace(sName + IIf(cmbRemoteConfigs.ListIndex > 0, AddTrailingSlash(Project.UserConfigs(GetConfigID(cmbRemoteConfigs.Text)).CompiledPath, "/"), ""), "//", "/")
        lblFTPUpImg.Caption = "ftp://" + Replace(sName + IIf(cmbRemoteConfigs.ListIndex > 0, AddTrailingSlash(Project.UserConfigs(GetConfigID(cmbRemoteConfigs.Text)).ImagesPath, "/"), ""), "//", "/")
    End If
    If Not IsUpdating Then Project.FTP.RemoteInfo4FTP = Project.UserConfigs(cmbRemoteConfigs.ListIndex).Name

End Sub

Private Sub cmdAddInEditor_Click()

    Dim oAddIn As AddInDef
    
    oAddIn = Project.AddIn
    Project.AddIn.Name = cmbAddIns.Text
    Project.AddIn.Description = GetAddInDescription(cmbAddIns.Text)
    
    If Preferences.ShowWarningAddInEditor Then
        DlgAns = -1
        frmAIEConfirm.Show vbModal
    Else
        DlgAns = vbYes
    End If
    
    If DlgAns = vbYes Then frmAddInEditor.Show vbModal
    
    Project.AddIn = oAddIn
    
    cmbAddIns.SetFocus

End Sub

Private Sub cmdBrowseFramesDoc_Click()

    On Error GoTo ExitSub
    
    With cDlg
        .DialogTitle = GetLocalizedStr(385)
        .InitDir = GetRealLocal.RootWeb
        .filter = SupportedHTMLDocs
        .CancelError = True
        .flags = cdlOFNFileMustExist + cdlOFNHideReadOnly + cdlOFNLongNames + cdlOFNPathMustExist
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
            lblFramesErr.Caption = ""
            For i = 1 To UBound(.Frames)
                If IsFrameNameInvalid(.Frames(i), "top") Or _
                   IsFrameNameInvalid(.Frames(i), "body") Then
                    chkFrameSupport.Value = vbUnchecked
                    chkFrameSupport_Click
                    Exit Sub
                End If
            Next i
        Else
            lblFramesErr.Caption = GetLocalizedStr(388)
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
    If Dir(Path) = "" Or Err.Number Then
        Path = ""
    Else
        Path = UnqualifyPath(Path)
    End If
    Path = BrowseForFolderByPath(Path, GetLocalizedStr(389))
    
    If Path <> "" Then txtImagesPath(ppSelConfig).Text = AddTrailingSlash(Path, IIf(Project.UserConfigs(ppSelConfig).Type = ctcRemote, "/", "\"))
    
    Me.Enabled = True
    Me.SetFocus

End Sub

Private Sub cmdBrowseRootWeb_Click(Index As Integer)

    Dim Path As String
    
    On Error Resume Next
    
    Me.Enabled = False
    
    Path = txtRootWeb(ppSelConfig).Text
    If Dir(Path) = "" Or Err.Number Then
        Path = ""
    Else
        Path = UnqualifyPath(Path)
    End If
    Path = BrowseForFolderByPath(Path, GetLocalizedStr(390))
    
    If Path <> "" Then txtRootWeb(ppSelConfig).Text = AddTrailingSlash(Path, IIf(Project.UserConfigs(ppSelConfig).Type = ctcRemote, "/", "\"))
    
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
        .flags = cdlOFNFileMustExist + cdlOFNHideReadOnly + cdlOFNLongNames + cdlOFNPathMustExist
        .ShowOpen
        txtUnfoldingSound.Text = .FileName
    End With
    
ExitSub:
    Exit Sub

End Sub

Private Sub cmdChange_Click()
    
    SelImage.FileName = picTBImage.Tag
    frmRscImages.Show vbModal
    With SelImage
        If .IsValid Then
            SetTBPicture .FileName
        End If
    End With

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

Private Sub cmdImport_Click()

    On Error GoTo ExitSub
    
    If Project.FileName = "" Then
        MsgBox GetLocalizedStr(392), vbOKOnly + vbInformation, "Unable to Import"
        Exit Sub
    End If
    
    With cDlg
        .DialogTitle = GetLocalizedStr(393)
        .flags = cdlOFNFileMustExist + cdlOFNHideReadOnly + cdlOFNLongNames + cdlOFNPathMustExist
        .filter = GetLocalizedStr(256) + "|*.dmb"
        .CancelError = True
        .ShowOpen
        ImportProjectFileName = .FileName
        frmPPImportOptions.Show vbModal
        InitializeDialog
    End With
    
ExitSub:
    Exit Sub

End Sub

Private Sub cmdPlay_Click()

    PlaySound ByVal txtUnfoldingSound.Text, 0&, SND_FILENAME Or SND_ASYNC Or SND_NOWAIT

End Sub

Private Sub cmdRemove_Click()

    SetTBPicture ""

End Sub

Private Sub cmdRemoveSound_Click()
    
    txtUnfoldingSound.Text = ""
    txtUnfoldingSound.SetFocus

End Sub

Private Sub cmdRestore_Click()

    sldAnimSpeed.Value = 35
    sldHideDelay.Value = 2000
    sldSubmenuDelay.Value = 200
    
    With lvExtFeatures.ListItems
        .item(1).Checked = True
        .item(2).Checked = True
        .item(3).Checked = True
        .item(4).Checked = False
        .item(5).Checked = True
        .item(6).Checked = False
        .item(7).Checked = False
        .item(8).Checked = False
        .item(9).Checked = False
    End With

End Sub

Private Sub cmdTBAutoHeight_Click()

    Dim r() As Integer

    r = GetTBHeight(True)
    txtTBHeight.Text = r(1)

End Sub

Private Sub cmdTBAutoWidth_Click()

    Dim r() As Integer

    r = GetTBWidth(True)
    txtTBWidth.Text = r(1)

End Sub

Private Sub cmdTBBackColor_Click()
    
    BuildUsedColorsArray
    
    With cmdTBBackColor
        SelColor = .Tag
        SelColor_CanBeTransparent = True
        frmColorPicker.Show vbModal
        SetColor SelColor, cmdTBBackColor
    End With
    
End Sub

Private Sub cmdReload_Click()

    LoadFramesDoc txtFramesDoc.Text

End Sub

Private Sub cmdTBBorderColor_Click()

    BuildUsedColorsArray
    
    With cmdTBBorderColor
        SelColor = .Tag
        SelColor_CanBeTransparent = True
        frmColorPicker.Show vbModal
        SetColor SelColor, cmdTBBorderColor
    End With

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 112 Then
        Select Case tsSections.SelectedItem.key
            Case "tsGeneral"
                ShowHelp "dialogs/pp_general.htm"
            Case "tsPublishing"
                Select Case tsConfigs.SelectedItem.key
                    Case "tsPaths"
                        ShowHelp "dialogs/pp_conf_paths.htm"
                    Case "tsHSE"
                        ShowHelp "dialogs/pp_hse.htm"
                    Case "tsFrames"
                        ShowHelp "dialogs/pp_frames.htm"
                End Select
            Case "tsToolbar"
                Select Case tsToolbar.SelectedItem.key
                    Case "tsGeneral"
                        ShowHelp "dialogs/pp_tb_general.htm"
                    Case "tsAppearance"
                        ShowHelp "dialogs/pp_tb_appearance.htm"
                    Case "tsPositioning"
                        ShowHelp "dialogs/pp_tb_positioning.htm"
                    Case "tsAdvanced"
                        ShowHelp "dialogs/pp_tb_advanced.htm"
                End Select
            Case "tsGlobalSettings"
                ShowHelp "dialogs/pp_globalsettings.htm"
            Case "tsFTP"
                ShowHelp "dialogs/pp_ftp.htm"
            Case "tsAdvanced"
                ShowHelp "dialogs/pp_advanced.htm"
        End Select
    End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    msgSubClass.SubClassHwnd picTBAppearance.hwnd, False

End Sub

Private Sub lvGroups_ItemCheck(ByVal item As MSComctlLib.ListItem)

    tbGroupsOptions.Buttons("tbCheck").Value = Abs(item.Checked)
    MenuGrps(GetIDByName(item.Text)).IncludeInToolbar = item.Checked

End Sub

Private Sub lvGroups_ItemClick(ByVal item As MSComctlLib.ListItem)

    tbGroupsOptions.Buttons("tbCheck").Value = Abs(item.Checked)
    MenuGrps(GetIDByName(item.Text)).IncludeInToolbar = item.Checked

End Sub

Private Sub mnuConfigAdd_Click()

    Dim NumConf As Integer

    NumConf = UBound(Project.UserConfigs)

    frmConfigAdd.Show vbModal
    
    If NumConf <> UBound(Project.UserConfigs) Then
        CreateConfigCtrls
        tsPublishing.Tabs(tsPublishing.Tabs.count).Selected = True
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
        .Caption = GetLocalizedStr(394)
        .Tag = CStr(ppSelConfig)
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

    If Project.UserConfigs(ppSelConfig).Name = Project.FTP.RemoteInfo4FTP Then
        MsgBox GetLocalizedStr(395), vbInformation + vbOKOnly, "Unable to Delete Configuration"
        Exit Sub
    End If
    
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
    Unload shpConfigInfoBack(shpConfigInfoBack.count - 1)
    
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
    
    For i = 0 To shpConfigInfoBack.count - 1
        tsPublishing.Tabs(i + 1).Image = IIf(i = Project.DefaultConfig, 4, 0)
        shpConfigInfoBack(i).BackColor = IIf(i = ppSelConfig, &H800000, &H808080)
    Next i
    
    mnuConfigSetAsDefault.Checked = True
    
    DisplayTip GetLocalizedStr(682), GetLocalizedStr(683)

End Sub

Private Sub opAlignment_Click(Index As Integer)

    UpdateToolbarControls

End Sub

Private Sub opTBHeight_Click(Index As Integer)

    txtTBHeight.Enabled = (Index = 1)
    cmdTBAutoHeight.Enabled = (Index = 1)

End Sub

Private Sub opTBWidth_Click(Index As Integer)

    txtTBWidth.Enabled = (Index = 2)
    cmdTBAutoWidth.Enabled = (Index = 2)

End Sub

Private Sub tbGroupsOptions_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.key
         Case "tbUp"
            MoveGrpUp
         Case "tbDown"
            MoveGrpDown
         Case "tbCheck"
            If lvGroups.SelectedItem.Text <> GetLocalizedStr(399) Then
                lvGroups.SelectedItem.Checked = Button.Value
            End If
    End Select

End Sub

Private Sub MoveGrpDown()

    Dim SelItem As String

    With lvGroups
        If .SelectedItem.Index = .ListItems.count Then Exit Sub
        SelItem = .SelectedItem.Text
        MenuGrps(GetIDByName(.SelectedItem.Text)).ToolbarIndex = .SelectedItem.Index + 1
        MenuGrps(GetIDByName(.ListItems(.SelectedItem.Index + 1).Text)).ToolbarIndex = .SelectedItem.Index - 1
    End With
    CreateGroupsList
    
    With lvGroups.FindItem(SelItem, lvwText, , lvwWhole)
        .Selected = True
        .EnsureVisible
    End With
    
    lvGroups.SetFocus

End Sub

Private Sub MoveGrpUp()

    Dim SelItem As String

    With lvGroups
        If .SelectedItem.Index = 1 Then Exit Sub
        SelItem = .SelectedItem.Text
        MenuGrps(GetIDByName(.SelectedItem.Text)).ToolbarIndex = .SelectedItem.Index - 1
        MenuGrps(GetIDByName(.ListItems(.SelectedItem.Index - 1).Text)).ToolbarIndex = .SelectedItem.Index + 1
    End With
    CreateGroupsList
    
    With lvGroups.FindItem(SelItem, lvwText, , lvwWhole)
        .Selected = True
        .EnsureVisible
    End With
    
    lvGroups.SetFocus

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

    Select Case tsConfigs.SelectedItem.key
        Case "tsPaths"
        Case "tsHSE"
            lblHSDocInfo.Caption = GetLocalizedStr(749) + vbCrLf + IIf(chkCreateToolbar.Value = vbChecked, GetLocalizedStr(750), GetLocalizedStr(751))
            picHotSpotEditor.Visible = True
        Case "tsFrames"
            picFrames.Visible = True
    End Select
    
    mnuConfigOpPaths.Enabled = Project.UserConfigs(ppSelConfig).Type = ctcRemote

End Sub

Private Sub AutoSelectPaths()

    tsConfigs.Tabs("tsPaths").Selected = True
    tsConfigs_Click

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
            cmbCodeOp.SetFocus
        Case "tsToolbar"
            picToolbar.ZOrder 0
            chkCreateToolbar.SetFocus
        Case "tsPublishing"
            picConfigs.ZOrder 0
            tsPublishing_Click
        Case "tsFTP"
            cmbRemoteConfigs_Click
            picFTP.ZOrder 0
            txtFTPAddress.SetFocus
        Case "tsGlobalSettings"
            picGlobalSettings.ZOrder 0
            cmbFX.SetFocus
    End Select

End Sub

Private Sub cmdBrowse_Click(Index As Integer)

    Dim Path As String
    
    On Error Resume Next
    
    Me.Enabled = False
    
    Path = txtRootWeb(ppSelConfig).Text
    If Dir(Path) = "" Or Err.Number Then
        Path = ""
    Else
        Path = UnqualifyPath(Path)
    End If
    Path = BrowseForFolderByPath(Path, GetLocalizedStr(401))
    
    If Path <> "" Then txtDest(ppSelConfig).Text = AddTrailingSlash(Path, IIf(Project.UserConfigs(ppSelConfig).Type = ctcRemote, "/", "\"))
    
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
        .flags = cdlOFNFileMustExist + cdlOFNHideReadOnly + cdlOFNLongNames + cdlOFNPathMustExist
        .ShowOpen
        txtDestFile.Text = .FileName
    End With
    
ExitSub:
    Exit Sub
    
End Sub

Private Sub ChkHSFile()

    If FileExists(txtDestFile.Text) Or txtDestFile.Text = "" Then
        lblHSFileErr.Caption = ""
    Else
        lblHSFileErr.Caption = GetLocalizedStr(403)
    End If

End Sub

Private Sub cmdCancel_Click()

    Project = ProjectBack
    Unload Me

End Sub

Private Sub cmdInfo_Click()

    If cmbAddIns.ListIndex = 0 Then
        MsgBox GetLocalizedStr(405), vbInformation + vbOKOnly, "AddIn Information"
    Else
        MsgBox GetAddInDescription(cmbAddIns.Text), vbInformation + vbOKOnly, "AddIn Information"
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
        '.AbsPath = SetSlashDir(Mid$(.Local.CompiledPath, Len(.Local.RootWeb)), sdFwd)
        .CodeOptimization = cmbCodeOp.ListIndex
        
        .FTP.FTPAddress = AddTrailingSlash(txtFTPAddress.Text, "/")
        If chkFTPUseProxy.Value = vbChecked Then
            .FTP.ProxyAddress = txtFTPProxyAddress.Text
            .FTP.ProxyPort = Val(txtFTPProxyPort.Text)
        Else
            .FTP.ProxyAddress = ""
            .FTP.ProxyPort = 0
        End If
        .FTP.UserName = IIf(opLogin(0).Value = True, "anonymous", "nonanon")
        .FTP.Password = ""
        If cmbRemoteConfigs.ListIndex > 0 Then
            .FTP.RemoteInfo4FTP = cmbRemoteConfigs.Text
        Else
            .FTP.RemoteInfo4FTP = ""
        End If
        
        If cmbAddIns.ListIndex = 0 Then
            .AddIn.Name = ""
            .AddIn.Description = ""
        Else
            .AddIn.Name = cmbAddIns.Text
            .AddIn.Description = GetAddInDescription(cmbAddIns.Text)
        End If
        LoadAddInParams cmbAddIns.Text
        .ToolBar.CreateToolBar = (chkCreateToolbar.Value = vbChecked)
        .ToolBar.FollowHScroll = (chkHorizontal.Value = vbChecked) And (chkFollowScrolling.Value = vbChecked)
        .ToolBar.FollowVScroll = (chkVertical.Value = vbChecked) And (chkFollowScrolling.Value = vbChecked)
        .ToolBar.MarginH = txtMarginH.Text
        .ToolBar.MarginV = txtMarginV.Text
        .ToolBar.ContentsMarginH = Val(txtTBCMarginH.Text)
        .ToolBar.ContentsMarginV = Val(txtTBCMarginV.Text)
        
        If opTBWidth(0).Value Then .ToolBar.Width = 0
        If opTBWidth(1).Value Then .ToolBar.Width = 1
        If opTBWidth(2).Value Then .ToolBar.Width = -Val(txtTBWidth.Text)
        
        If opTBHeight(0).Value Then .ToolBar.Height = 0
        If opTBHeight(1).Value Then .ToolBar.Height = -Val(txtTBHeight.Text)
        
        For i = 0 To opAlignment.count - 1
            If opAlignment(i).Value = True Then
                .ToolBar.Alignment = i
                Exit For
            End If
        Next i
        
        For i = 1 To lvGroups.ListItems.count
            With MenuGrps(GetIDByName(lvGroups.ListItems(i).Text))
                .ToolbarIndex = i
                .IncludeInToolbar = lvGroups.ListItems(i).Checked
            End With
        Next i
        
        .ToolBar.Style = cmbStyle.ListIndex
        .ToolBar.Spanning = cmbSpanning.ListIndex
        .ToolBar.border = cmbBorder.ListIndex
        .ToolBar.JustifyHotSpots = (chkJustify.Value = vbChecked)
        .ToolBar.BackColor = cmdTBBackColor.Tag
        .ToolBar.BorderColor = cmdTBBorderColor.Tag
        .ToolBar.Separation = Val(txtTBSeparation.Text)
        .ToolBar.Image = picTBImage.Tag
        .ToolBar.CustX = txtACX.Text
        .ToolBar.CustY = txtACY.Text
        
        .AnimSpeed = sldAnimSpeed.Value
        .HideDelay = sldHideDelay.Value
        .SubmenuDelay = sldSubmenuDelay.Value
        
        With lvExtFeatures.ListItems
            Project.CompileIECode = .item(1).Checked
            Project.CompileNSCode = .item(2).Checked
            Project.CompilehRefFile = .item(3).Checked
            Project.SharedProject = .item(4).Checked
            Project.DoFormsTweak = .item(5).Checked
            Project.DWSupport = .item(6).Checked
            Project.NS4ClipBug = .item(7).Checked
            Project.OPHelperFunctions = .item(8).Checked
            Project.ImageReadySupport = .item(9).Checked
        End With
        
        .JSFileName = txtJSFileName.Text: If .JSFileName = "" Then .JSFileName = "menu"
        
        For i = 0 To UBound(.UserConfigs)
            With .UserConfigs(i)
                .Frames.UseFrames = .Frames.UseFrames And .Frames.FramesFile <> ""
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
            .onmouseover.URL = Replace(.onmouseover.URL, BackConfig.RootWeb, ThisConfig.RootWeb)
            .onclick.URL = Replace(.onclick.URL, BackConfig.RootWeb, ThisConfig.RootWeb)
            .OnDoubleClick.URL = Replace(.OnDoubleClick.URL, BackConfig.RootWeb, ThisConfig.RootWeb)
        End With
    Next i
    
    For i = 1 To UBound(MenuGrps)
        With MenuGrps(i).Actions
            .onmouseover.URL = Replace(.onmouseover.URL, BackConfig.RootWeb, ThisConfig.RootWeb)
            .onclick.URL = Replace(.onclick.URL, BackConfig.RootWeb, ThisConfig.RootWeb)
            .OnDoubleClick.URL = Replace(.OnDoubleClick.URL, BackConfig.RootWeb, ThisConfig.RootWeb)
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

    LocalizeUI
    
    Set msgSubClass = New SmartSubClass
    msgSubClass.SubClassHwnd picTBAppearance.hwnd, True

    ProjectBack = Project

    mnuConfig.Visible = False

    Width = 6960
    Height = 5625 + GetClientTop(Me.hwnd)
    
    InitializeDialog
    
    If Not ToolbarConfigOnly Then
        tsSections.Tabs("tsGeneral").Selected = True
        tsPublishing.Tabs("tsLocal").Selected = True
    End If

End Sub

Private Sub CreateConfigCtrls()

    On Error Resume Next

    Dim i As Integer
    Dim j As Integer
    Dim IsLocal As Boolean
    
    IsUpdating = True
    
    cmbRemoteConfigs.Clear
    cmbRemoteConfigs.AddItem GetLocalizedStr(110)
    cmbRemoteConfigs.ListIndex = 0
    
    For i = tsPublishing.Tabs.count To 2 Step -1
        tsPublishing.Tabs.Remove tsPublishing.Tabs(i).Index
    Next i
    
    tsPublishing.Tabs(1).Image = IIf(Project.DefaultConfig = 0, 4, 0)
    tsPublishing.Tabs(1).Caption = "Default" ' GetLocalizedStr(253)
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
            .Caption = Project.UserConfigs(i).Name + " (" + ConfigTypeName(Project.UserConfigs(i)) + ")"
        End With
        
        Load lblConfigDesc(i)
        With lblConfigDesc(i)
            Set .Container = picConfigsOp(i)
            .Visible = True
            .Move lblConfigDesc(0).Left, lblConfigDesc(0).Top
            .Caption = Project.UserConfigs(i).Description
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
            .Caption = IIf(IsLocal, GetLocalizedStr(326), GetLocalizedStr(330))
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
            .Caption = IIf(IsLocal, GetLocalizedStr(333) + ": c:\inetpub\wwwroot\", GetLocalizedStr(333) + ": http://myweb.com/")
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
            .Caption = IIf(IsLocal, GetLocalizedStr(327), GetLocalizedStr(331))
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
            .Caption = IIf(IsLocal, GetLocalizedStr(333) + ": c:\inetpub\wwwroot\menus\", GetLocalizedStr(333) + ": /menus/")
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
            .Caption = IIf(IsLocal, GetLocalizedStr(328), GetLocalizedStr(332))
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
            .Caption = IIf(IsLocal, GetLocalizedStr(333) + ": c:\inetpub\wwwroot\menus\images\", GetLocalizedStr(333) + ": /menus/images/")
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
        
        ' ----- FTP SUPPORT -------
        If Project.UserConfigs(i).Type = ctcRemote Then
            cmbRemoteConfigs.AddItem Project.UserConfigs(i).Name
            If Project.UserConfigs(i).Name = Project.FTP.RemoteInfo4FTP Then
                cmbRemoteConfigs.ListIndex = cmbRemoteConfigs.NewIndex
            End If
        End If
    Next i
    
    IsUpdating = False

End Sub

Private Sub InitializeDialog()

    Dim i As Integer
    Dim nItem As ListItem

    lblFramesErr.Caption = ""
    lblHSFileErr.Caption = ""
    lblAddIneErr.Caption = ""
    
    CreateConfigCtrls
    
    IsUpdating = True

    With picGeneral
        .BorderStyle = 0
        .ZOrder 0
        picAdvanced.Move .Left, .Top, .Width, .Height
        picAdvanced.BorderStyle = 0
        picToolbar.Move .Left, .Top, .Width, .Height
        picToolbar.BorderStyle = 0
        picConfigs.Move .Left, .Top, .Width, .Height
        picConfigs.BorderStyle = 0
        picFTP.Move .Left, .Top, .Width, .Height
        picFTP.BorderStyle = 0
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
    
    With picTBGeneral
        .BorderStyle = 0
        .ZOrder 0
        picTBAdvanced.Move .Left, .Top, .Width, .Height
        picTBAdvanced.BorderStyle = 0
        picTBAppearance.Move .Left, .Top, .Width, .Height
        picTBAppearance.BorderStyle = 0
        picTBPositioning.Move .Left, .Top, .Width, .Height
        picTBPositioning.BorderStyle = 0
    End With
    
    With Project
        txtFileName.Text = .FileName
        txtName.Text = .Name
        
        cmbFX.ListIndex = .FX
        txtUnfoldingSound.Text = .UnfoldingSound.onmouseover
        cmbCodeOp.ListIndex = .CodeOptimization
        For i = 0 To UBound(.UserConfigs)
            ppSelConfig = i
            With .UserConfigs(i)
                txtRootWeb(i).Text = .RootWeb
                txtDest(i).Text = .CompiledPath
                txtImagesPath(i).Text = .ImagesPath
                lblConfigName(i).Caption = .Name + " (" + ConfigTypeName(Project.UserConfigs(i)) + ")"
                lblConfigDesc(i).Caption = .Description
                tsPublishing.Tabs(i + 1).Caption = .Name
            End With
        Next i
        
        txtFTPAddress.Text = .FTP.FTPAddress
        txtFTPProxyAddress.Text = .FTP.ProxyAddress
        txtFTPProxyPort.Text = .FTP.ProxyPort
        opLogin(0).Value = (.FTP.UserName = "" Or .FTP.UserName = "anonymous")
        chkFTPUseProxy.Value = IIf(.FTP.ProxyAddress <> "", vbChecked, vbUnchecked)
        
        GetAddInsList .AddIn.Name
        
        chkCreateToolbar.Value = Abs(.ToolBar.CreateToolBar)
        chkFollowScrolling.Value = Abs(.ToolBar.FollowHScroll Or .ToolBar.FollowVScroll)
        chkHorizontal.Value = Abs(.ToolBar.FollowHScroll)
        chkVertical.Value = Abs(.ToolBar.FollowVScroll)
        Select Case .ToolBar.Width
            Case 0
                opTBWidth(0).Value = True
            Case 1
                opTBWidth(1).Value = True
            Case Else
                opTBWidth(2).Value = True
                txtTBWidth.Text = Abs(.ToolBar.Width)
        End Select
        Select Case .ToolBar.Height
            Case 0
                opTBHeight(0).Value = True
            Case Else
                opTBHeight(1).Value = True
                txtTBHeight.Text = Abs(.ToolBar.Height)
        End Select
        txtMarginH.Text = .ToolBar.MarginH
        txtMarginV.Text = .ToolBar.MarginV
        txtTBCMarginH.Text = .ToolBar.ContentsMarginH
        txtTBCMarginV.Text = .ToolBar.ContentsMarginV
        opAlignment(.ToolBar.Alignment).Value = True
        cmbStyle.ListIndex = .ToolBar.Style
        cmbSpanning.ListIndex = .ToolBar.Spanning
        cmbBorder.ListIndex = .ToolBar.border
        chkJustify.Value = Abs(.ToolBar.JustifyHotSpots)
        txtTBSeparation.Text = .ToolBar.Separation
        SetColor .ToolBar.BackColor, cmdTBBackColor
        SetColor .ToolBar.BorderColor, cmdTBBorderColor
        SetTBPicture .ToolBar.Image
        txtACX.Text = .ToolBar.CustX
        txtACY.Text = .ToolBar.CustY
        
        txtJSFileName.Text = .JSFileName
        UpdateJSFileNames
        
        sldAnimSpeed.Value = .AnimSpeed
        sldHideDelay.Value = .HideDelay
        sldSubmenuDelay.Value = .SubmenuDelay
        
        With lvExtFeatures.ListItems
            Set nItem = .Add(, , GetLocalizedStr(733))
            nItem.Checked = Project.CompileIECode
            Set nItem = .Add(, , GetLocalizedStr(734))
            nItem.Checked = Project.CompileNSCode
            Set nItem = .Add(, , GetLocalizedStr(735))
            nItem.Checked = Project.CompilehRefFile
            Set nItem = .Add(, , GetLocalizedStr(736))
            nItem.Checked = Project.SharedProject
            Set nItem = .Add(, , GetLocalizedStr(737))
            nItem.Checked = Project.DoFormsTweak
            Set nItem = .Add(, , GetLocalizedStr(738))
            nItem.Checked = Project.DWSupport
            Set nItem = .Add(, , GetLocalizedStr(739))
            nItem.Checked = Project.NS4ClipBug
            Set nItem = .Add(, , GetLocalizedStr(740))
            nItem.Checked = Project.OPHelperFunctions
            Set nItem = .Add(, , GetLocalizedStr(741))
            nItem.Checked = Project.ImageReadySupport
        End With
        CoolListView lvExtFeatures
    End With
    
    CreateGroupsList
    
    UpdateToolbarControls
    
    If ToolbarConfigOnly Then
        tsSections.Tabs.Remove "tsGeneral"
        tsSections.Tabs.Remove "tsPublishing"
        tsSections.Tabs.Remove "tsFTP"
        tsSections.Tabs.Remove "tsAdvanced"
        tsSections.Tabs.Remove "tsGlobalSettings"
        picToolbar.ZOrder 0
        cmdImport.Visible = False
    End If
    
    IsUpdating = False
    
    CenterForm Me
    SetupCharset Me

End Sub

Private Sub SetTBPicture(FileName As String)

    On Error Resume Next

    TileImage FileName, picTBImage
    picTBImage.Tag = FileName

End Sub

Private Sub CreateGroupsList()

    Dim i As Integer
    Dim g As Integer
    Dim nItem As ListItem
    
    lvGroups.ListItems.Clear

    Do Until lvGroups.ListItems.count = UBound(MenuGrps)
        For g = 1 To UBound(MenuGrps)
            If MenuGrps(g).ToolbarIndex = i Then
                Set nItem = lvGroups.ListItems.Add(, , MenuGrps(g).Name)
                nItem.SubItems(1) = MenuGrps(g).Caption
                nItem.Checked = MenuGrps(g).IncludeInToolbar
            End If
        Next g
        i = i + 1
    Loop
    If lvGroups.ListItems.count = 0 Then
        lvGroups.Checkboxes = False
        lvGroups.ListItems.Add , , GetLocalizedStr(399)
        lvGroups.Enabled = False
    Else
        lvGroups.ListItems(1).Selected = True
        tbGroupsOptions.Buttons("tbCheck").Value = Abs(lvGroups.SelectedItem.Checked)
    End If
    
    'lvGroups.ColumnHeaders(1).Width = lvGroups.Width / 2
    'lvGroups.ColumnHeaders(2).Width = lvGroups.Width / 2 - 60
    CoolListView lvGroups

End Sub

Private Sub GetAddInsList(AddInName As String)

    Dim fName As String
    Dim sName As String

    CenterForm Me
    
    cmbAddIns.ListIndex = 0
    
    IsUpdating = True
    
    fName = Dir(AppPath + "AddIns\*.ext")
    Do Until fName = ""
        sName = Left$(fName, InStrRev(fName, ".") - 1)
        cmbAddIns.AddItem sName
        If sName = AddInName Then
            cmbAddIns.ListIndex = cmbAddIns.ListCount - 1
        End If
        fName = Dir
    Loop
    
    IsUpdating = False
    
    If AddInName <> "" And cmbAddIns.ListIndex = 0 Then
        lblAddIneErr.Caption = GetLocalizedStr(407) + " " + AddInName + " " + GetLocalizedStr(408)
    End If
    cmdAddInEditor.Enabled = cmbAddIns.ListIndex > 0
    SetEditParamsButtonState

End Sub

Private Sub tsToolbar_Click()

    On Error Resume Next

    Select Case tsToolbar.SelectedItem.key
        Case "tsGeneral"
            picTBGeneral.ZOrder 0
        Case "tsPositioning"
            picTBPositioning.ZOrder 0
        Case "tsAppearance"
            picTBAppearance.ZOrder 0
        Case "tsAdvanced"
            picTBAdvanced.ZOrder 0
    End Select

End Sub

Private Sub txtACX_GotFocus()

    SelAll txtACX

End Sub

Private Sub txtACX_KeyPress(KeyAscii As Integer)

    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If

End Sub

Private Sub txtACY_GotFocus()

    SelAll txtACY

End Sub

Private Sub txtACY_KeyPress(KeyAscii As Integer)

    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If

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

Private Sub txtFTPAddress_Change()

    cmbRemoteConfigs_Click

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
    InvalidChars = " !@#$%^&*()|\/,.;':{}[]" + Chr(34)

    If InStr(InvalidChars, Chr(KeyAscii)) <> 0 Then
        KeyAscii = 0
    End If

End Sub

Private Sub UpdateJSFileNames()

    Dim fn As String
    
    fn = txtJSFileName.Text
    If fn = "" Then fn = "menu"

    txtJSFileNames.Text = fn + ".js" + vbCrLf + _
                          "ie" + fn + ".js" + vbCrLf + _
                          "ns" + fn + ".js" + vbCrLf + _
                          "ie" + fn + "_frames.js" + vbCrLf + _
                          "ns" + fn + "_frames.js"
    
End Sub

Private Sub txtMarginH_GotFocus()

    SelAll txtMarginH

End Sub

Private Sub txtMarginH_KeyPress(KeyAscii As Integer)

    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 45 Then
        KeyAscii = 0
    End If

End Sub

Private Sub txtMarginV_GotFocus()

    SelAll txtMarginV

End Sub

Private Sub txtMarginV_KeyPress(KeyAscii As Integer)

    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 45 Then
        KeyAscii = 0
    End If

End Sub

Private Sub txtName_GotFocus()

    SelAll txtName

End Sub

Private Sub txtRootWeb_Change(Index As Integer)

    If IsUpdating Then Exit Sub
    Project.UserConfigs(ppSelConfig).RootWeb = txtRootWeb(Index).Text

End Sub

Private Sub txtTBCMarginH_GotFocus()

    SelAll txtTBCMarginH

End Sub

Private Sub txtTBCMarginH_KeyPress(KeyAscii As Integer)

    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If

End Sub

Private Sub txtTBCMarginV_GotFocus()

    SelAll txtTBCMarginV

End Sub

Private Sub txtTBCMarginV_KeyPress(KeyAscii As Integer)

    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If

End Sub

Private Sub txtTBSeparation_GotFocus()

    SelAll txtTBSeparation

End Sub

Private Sub txtTBSeparation_KeyPress(KeyAscii As Integer)

    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If

End Sub

Private Sub LocalizeUI()

    Dim ctrl As Control

    Caption = GetLocalizedStr(378)

    'General
    tsSections.Tabs(1).Caption = GetLocalizedStr(321)
    lblProjectLocation.Caption = GetLocalizedStr(317)
    lblProjectName.Caption = GetLocalizedStr(318)
    lblUnfoldingSound.Caption = GetLocalizedStr(320)
    cmdRemoveSound.Caption = GetLocalizedStr(201)
    
    'Global Settings
    tsSections.Tabs(5).Caption = GetLocalizedStr(727)
    frameUnfoldingEffect.Caption = GetLocalizedStr(319)
    lblEffectType.Caption = GetLocalizedStr(726)
    frameTimers.Caption = GetLocalizedStr(728)
    lblMenusHideDelay.Caption = GetLocalizedStr(729)
    lblSubmDispDelay.Caption = GetLocalizedStr(730)
    lblEffectSpeed.Caption = GetLocalizedStr(731)
    lblExtFeatures.Caption = GetLocalizedStr(732)
    cmdRestore.Caption = GetLocalizedStr(742)
    lblLess1.Caption = GetLocalizedStr(744)
    lblLess2.Caption = GetLocalizedStr(744)
    lblMore1.Caption = GetLocalizedStr(745)
    lblMore2.Caption = GetLocalizedStr(745)
    lblSlow.Caption = GetLocalizedStr(746)
    lblFast.Caption = GetLocalizedStr(747)
    
    'Configurations
    tsSections.Tabs(2).Caption = GetLocalizedStr(322)
    lblRootWeb(0).Caption = GetLocalizedStr(326)
    lblDest(0).Caption = GetLocalizedStr(327)
    lblImagesPath(0).Caption = GetLocalizedStr(328)
    lblLocalConfig(0).Caption = GetLocalizedStr(329)
    
    tsConfigs.Tabs(1).Caption = GetLocalizedStr(334)
    tsConfigs.Tabs(2).Caption = GetLocalizedStr(748)
    tsConfigs.Tabs(3).Caption = GetLocalizedStr(336)
    cmdConfigOptions.Caption = GetLocalizedStr(337)
    
    mnuConfigAdd.Caption = GetLocalizedStr(338)
    mnuConfigEdit.Caption = GetLocalizedStr(339)
    mnuConfigRemove.Caption = GetLocalizedStr(201)
    mnuConfigSetAsDefault.Caption = GetLocalizedStr(341)
    mnuConfigOpPaths.Caption = GetLocalizedStr(743)
    
    lblDocHS.Caption = GetLocalizedStr(344)
        
    chkFrameSupport.Caption = GetLocalizedStr(346)
    lblFramesDoc.Caption = GetLocalizedStr(347)
    
    'Toolbar
    tsSections.Tabs(3).Caption = GetLocalizedStr(323)
    chkCreateToolbar.Caption = GetLocalizedStr(350)
    lblTBSelGrps.Caption = GetLocalizedStr(351)
    
    tsToolbar.Tabs(1).Caption = GetLocalizedStr(321)
    tsToolbar.Tabs(2).Caption = GetLocalizedStr(352)
    tsToolbar.Tabs(3).Caption = GetLocalizedStr(353)
    tsToolbar.Tabs(4).Caption = GetLocalizedStr(325)
    
    lblTBStyle.Caption = GetLocalizedStr(354)
    lblTBBorder.Caption = GetLocalizedStr(355)
    lblTBSeparation.Caption = GetLocalizedStr(356)
    lblJustify.Caption = GetLocalizedStr(357)
    lblTBBackColor.Caption = GetLocalizedStr(358)
    lblTBBackImage.Caption = GetLocalizedStr(359)
    lblTBMargins.Caption = GetLocalizedStr(216)
    lblH.Caption = GetLocalizedStr(211)
    lblV.Caption = GetLocalizedStr(210)
    
    cmdChange.Caption = GetLocalizedStr(189)
    cmdRemove.Caption = GetLocalizedStr(201)
    
    lblTBAlignment.Caption = GetLocalizedStr(115)
    lblTBSpanning.Caption = GetLocalizedStr(361)
    lblTBOffset.Caption = GetLocalizedStr(362)
    lblTBOffsetH.Caption = GetLocalizedStr(363)
    lblTBOffsetV.Caption = GetLocalizedStr(364)
    
    cmbStyle.Clear
    cmbStyle.AddItem GetLocalizedStr(211)
    cmbStyle.AddItem GetLocalizedStr(210)
    
    cmbSpanning.Clear
    cmbSpanning.AddItem GetLocalizedStr(365)
    cmbSpanning.AddItem GetLocalizedStr(366)
    
    opAlignment(9).Caption = GetLocalizedStr(367)
    chkFollowScrolling.Caption = GetLocalizedStr(368)
    
    'FTP
    tsSections.Tabs(4).Caption = GetLocalizedStr(324)
    lblFTPServer.Caption = GetLocalizedStr(369)
    opLogin(0).Caption = GetLocalizedStr(370)
    opLogin(1).Caption = GetLocalizedStr(371)
    chkFTPUseProxy.Caption = GetLocalizedStr(372)
    lblFTPProxyAddress.Caption = GetLocalizedStr(375)
    lblFTPProxyPort.Caption = GetLocalizedStr(376)
    lblRemoteConfig.Caption = GetLocalizedStr(377)
    frameAccountInfo.Caption = GetLocalizedStr(677)
    
    'Advanced
    tsSections.Tabs(6).Caption = GetLocalizedStr(325)
    lblCodeOp.Caption = GetLocalizedStr(379)
    lblJSName.Caption = GetLocalizedStr(380)
    cmdPosOffset.Caption = GetLocalizedStr(381)
    cmdFontSubst.Caption = GetLocalizedStr(713)
    
    cmbCodeOp.Clear
    cmbCodeOp.AddItem GetLocalizedStr(382)
    cmbCodeOp.AddItem GetLocalizedStr(383)
    
    cmbFX.Clear
    cmbFX.AddItem GetLocalizedStr(455)
    cmbFX.AddItem GetLocalizedStr(451)
    cmbFX.AddItem GetLocalizedStr(452)
    cmbFX.AddItem GetLocalizedStr(453)
    cmbFX.AddItem GetLocalizedStr(454)
    
    cmdImport.Caption = GetLocalizedStr(343)
    
    For Each ctrl In Controls
        If TypeOf ctrl Is Label Then
            If Left(ctrl.Caption, 9) = "example: " Then
                ctrl.Caption = Replace(ctrl.Caption, "example", GetLocalizedStr(333))
            End If
        End If
    Next ctrl
    
    cmdOK.Caption = GetLocalizedStr(186)
    cmdCancel.Caption = GetLocalizedStr(187)
    
    FixContolsWidth Me
    
    cmdPosOffset.Width = SetCtrlWidth(cmdPosOffset)
    cmdFontSubst.Width = SetCtrlWidth(cmdFontSubst)
    cmdFontSubst.Left = cmdPosOffset.Left + cmdPosOffset.Width + 90
    cmdImport.Width = SetCtrlWidth(cmdImport)
    
    If Preferences.language <> "eng" Then
        cmdOK.Width = SetCtrlWidth(cmdOK)
        cmdCancel.Width = SetCtrlWidth(cmdCancel)
    End If

End Sub

Private Sub msgSubClass_NewMessage(ByVal hwnd As Long, uMsg As Long, wParam As Long, lParam As Long, Cancel As Boolean)

    Select Case hwnd
        Case picTBAppearance.hwnd
            Select Case uMsg
                Case WM_CTLCOLORSTATIC, WM_CTLCOLOREDIT, WM_PAINT
                    DrawColorBoxes Me
            End Select
    End Select
    
End Sub
